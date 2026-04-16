#!/usr/bin/env python3
from __future__ import annotations

import argparse
import datetime as dt
import html
import json
import os
import re
import smtplib
import sys
import time
import traceback
import urllib.parse
import urllib.request
from dataclasses import dataclass, field
from email.message import EmailMessage
from pathlib import Path
from zoneinfo import ZoneInfo

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

SCOPES = ["https://www.googleapis.com/auth/calendar"]
DEFAULT_URL = "https://share.myfreecams.com/RinCity/calendar"

DEFAULT_CALENDAR_ID = (
    "c_6c31b0f68c7d1a73f3443e2a89d8408b940b5953f731293e37f9a1450c13bad6"
    "@group.calendar.google.com"
)
TZ = "America/Los_Angeles"
LOCAL_TZ = ZoneInfo(TZ)
SCRIPT_DIR = Path(__file__).resolve().parent
TOKEN_PATH = SCRIPT_DIR / "token.json"
CREDENTIALS_PATH = SCRIPT_DIR / "credentials.json"

# Google Calendar event color IDs. These are the closest matches to the
# MFC Share colors currently used on RinCity's calendar.
COLOR_GREEN = "2"   # sage / pale green, closest to MFC lime stream blocks
COLOR_PURPLE = "3"  # purple, used for non-stream all-day events
COLOR_YELLOW = "5"  # yellow, used for possible/potential timed events
COLOR_RED = "11"    # red, used for "No Stream" all-day events

EMAIL_COLOR_STYLES = {
    COLOR_GREEN: ("#e3fee0", "#41fb30", "#1f1f1f"),
    COLOR_PURPLE: ("#fee5ff", "#fa54fc", "#1f1f1f"),
    COLOR_YELLOW: ("#fff0df", "#fc9827", "#1f1f1f"),
    COLOR_RED: ("#fadedf", "#dc2127", "#1f1f1f"),
}
DEFAULT_EMAIL_STYLE = ("#f3f3f3", "#cccccc", "#1f1f1f")
RINCITY_AVATAR_URL = (
    "https://img.mfcimg.com/photos2/164/16418930/avatar.150x150.jpg?nc=1765486890"
)

SYNC_MARKER_KEY = "rincity_sync"
SYNC_MARKER_VALUE = "mfcshare"
SYNC_KEY_FIELD = "mfc_key"
SYNC_KIND_FIELD = "mfc_kind"
SYNC_DATE_FIELD = "mfc_date"

NO_STREAM_RE = re.compile(r"\bno\s+stream\b", re.IGNORECASE)
STREAM_RE = re.compile(r"\bstream\b", re.IGNORECASE)
POSSIBLE_RE = re.compile(
    r"\b(possible|potential|maybe|tentative|probably|probable|might|tbd)\b",
    re.IGNORECASE,
)
BIRTHDAY_RE = re.compile(r"\bbirthday\b", re.IGNORECASE)

TRANSITION_KINDS = {"no_stream", "stream", "possible_stream"}
LEGACY_MANAGED_TITLES = {"stream", "no stream"}
HTTP_USER_AGENT = "sync-rin-calendars/1.0 (+https://github.com/jasmeralia/sync-rin-calendars)"

DAY_CONTENT_RE = re.compile(
    r"<div class='day-content'[^>]*data-date='(?P<date>\d{4}-\d{2}-\d{2})'[^>]*>"
    r"(?P<body>.*?)"
    r"</div>\s*<!-- day-content -->",
    re.DOTALL,
)
EVENT_RE = re.compile(
    r"<a class=\"(?P<class>[^\"]*\bevent\b[^\"]*)\"[^>]*"
    r"data-modal-url-value=\"(?P<modal>[^\"]+)\"[^>]*>"
    r"(?P<body>.*?)"
    r"</a>",
    re.DOTALL,
)
TITLE_RE = re.compile(
    r"<span class=['\"]title(?:\s+text-truncate)?['\"]>\s*(?P<body>.*?)\s*</span>",
    re.DOTALL,
)
START_TIME_RE = re.compile(
    r"<span class=['\"]start-time['\"]>\s*(?P<body>.*?)\s*</span>",
    re.DOTALL,
)
DURATION_PART_RE = re.compile(r"(?P<value>\d+)(?P<unit>[smhd])", re.IGNORECASE)


@dataclass(frozen=True)
class MfcEvent:
    date: dt.date
    title: str
    event_type: str
    start_time: str | None
    source_url: str
    modal_url: str
    description: str = ""

    @property
    def is_all_day(self) -> bool:
        return self.event_type == "all-day"

    @property
    def normalized_title(self) -> str:
        return normalize_text(self.title)

    @property
    def contains_stream(self) -> bool:
        return bool(STREAM_RE.search(self.title))

    @property
    def kind(self) -> str:
        if NO_STREAM_RE.search(self.title):
            return "no_stream"
        if self.contains_stream and POSSIBLE_RE.search(self.title):
            return "possible_stream"
        if self.contains_stream:
            return "stream"
        if BIRTHDAY_RE.search(self.title):
            return "birthday"
        if self.is_all_day:
            return "other_all_day"
        return "other_timed"

    @property
    def color_id(self) -> str:
        if self.is_all_day:
            return COLOR_RED if self.kind == "no_stream" else COLOR_PURPLE
        return COLOR_YELLOW if self.kind == "possible_stream" else COLOR_GREEN

    @property
    def sync_key(self) -> str:
        time_part = "all-day" if self.is_all_day else normalize_text(self.start_time or "")
        return f"{self.date.isoformat()}|{self.kind}|{time_part}|{self.normalized_title}"

    def start_end(self) -> tuple[dict[str, str], dict[str, str]]:
        if self.is_all_day:
            return (
                {"date": self.date.isoformat()},
                {"date": (self.date + dt.timedelta(days=1)).isoformat()},
            )

        if not self.start_time:
            raise ValueError(f"timed event is missing start_time: {self.title}")

        start_time = parse_time_string(self.start_time)
        start_dt = dt.datetime.combine(self.date, start_time, tzinfo=LOCAL_TZ)
        end_dt = dt.datetime.combine(
            self.date, dt.time(hour=23, minute=59), tzinfo=LOCAL_TZ
        )
        return (
            {"dateTime": start_dt.isoformat(), "timeZone": TZ},
            {"dateTime": end_dt.isoformat(), "timeZone": TZ},
        )


@dataclass(frozen=True)
class NotificationConfig:
    host: str
    port: int
    username: str | None
    password: str | None
    from_addr: str | None
    to_addrs: tuple[str, ...]
    use_starttls: bool
    notify_on_success: bool
    notify_on_changes: bool
    notify_on_error: bool
    subject_prefix: str

    @property
    def enabled(self) -> bool:
        return bool(self.from_addr and self.to_addrs)


@dataclass
class SyncSummary:
    lines: list[str] = field(default_factory=list)
    events: list[dict[str, str]] = field(default_factory=list)
    created: int = 0
    updated: int = 0
    metadata_only: int = 0
    unchanged: int = 0
    skipped: int = 0
    mfc_events_found: int = 0
    range_start: dt.date | None = None
    range_end: dt.date | None = None
    calendar_url: str = DEFAULT_URL
    calendar_id: str = DEFAULT_CALENDAR_ID
    dry_run: bool = False
    preview_mode: bool = False

    @property
    def has_changes(self) -> bool:
        return (self.created + self.updated) > 0

    def render_text(self) -> str:
        return "\n".join(self.lines).strip()

    def render_html(self) -> str:
        range_start = self.range_start.isoformat() if self.range_start else "unknown"
        range_end = self.range_end.isoformat() if self.range_end else "unknown"
        mode = "DRY RUN" if self.dry_run else "LIVE"
        new_events = [event for event in self.events if event["action"] == "CREATE"]
        updated_events = [event for event in self.events if event["action"] == "UPDATE"]
        no_changes_html = ""
        if not new_events and not updated_events:
            no_changes_html = (
                "<p style='margin:20px 0 0 0;color:#666;line-height:1.5;'>"
                "No new or updated events were included in this notification."
                "</p>"
            )
        return "\n".join(
            [
                "<!DOCTYPE html>",
                "<html><head><meta charset='utf-8' />",
                "<meta name='viewport' content='width=device-width, initial-scale=1.0' />",
                "<title>RinCity Calendar Sync</title></head>",
                (
                    "<body style='width:100%;margin:0;padding:0;background:#f1f1f1;"
                    "font-size:16px;line-height:1.4;'>"
                ),
                "<div style='max-width:680px;margin:0 auto;background:#f1f1f1;'>",
                "<div style='padding:20px;background:#ffffff;'>",
                (
                    "<div style='font-size:28px;font-weight:700;color:#222;"
                    "margin-bottom:10px;'>RinCity Calendar Sync</div>"
                ),
                "<p style='margin:0 0 16px 0;color:#555;'>"
                f"{html.escape(mode)} notification for RinCity calendar sync."
                "</p>",
                "<p style='margin:0 0 20px 0;color:#555;'>"
                f"<strong>Created:</strong> {self.created} &nbsp; "
                f"<strong>Updated:</strong> {self.updated} &nbsp; "
                f"<strong>Metadata-only:</strong> {self.metadata_only} &nbsp; "
                f"<strong>MFC events found:</strong> {self.mfc_events_found}"
                "</p>",
                self._render_event_section("New Events", new_events),
                self._render_event_section("Updated Events", updated_events),
                no_changes_html,
                "<div style='margin-top:22px;font-size:12px;color:#777;line-height:1.6;'>"
                f"MFC calendar: {html.escape(self.calendar_url)}<br>"
                f"Google calendar: {html.escape(self.calendar_id)}<br>"
                f"Range: {range_start} to {range_end} (exclusive)"
                "</div>",
                "</div></div></body></html>",
            ]
        )

    def add_event(
        self, action: str, event: MfcEvent, details: str, *, timing: str | None = None
    ) -> None:
        self.events.append(
            {
                "action": action,
                "date": event.date.isoformat(),
                "timing": timing or ("All day" if event.is_all_day else (event.start_time or "")),
                "title": event.title,
                "details": details,
                "color_id": event.color_id,
                "link": event.source_url,
            }
        )

    def _render_event_row(self, event: dict[str, str]) -> str:
        background, border, foreground = EMAIL_COLOR_STYLES.get(
            event["color_id"], DEFAULT_EMAIL_STYLE
        )
        details_html = ""
        if event["details"]:
            details_html = (
                f"<div style='margin-top:6px;font-size:13px;color:#666;'>"
                f"{html.escape(event['details'])}</div>"
            )
        return (
            "<table style='width:100%;border-collapse:collapse;margin:0 0 12px 0;'>"
            "<tr>"
            "<td style='width:50px;padding-right:8px;vertical-align:middle;'>"
            f"<a href=\"{html.escape(event['link'])}\" "
            "style='display:block;width:50px;height:50px;text-decoration:none;'>"
            f"<img src=\"{html.escape(RINCITY_AVATAR_URL)}\" width='50' height='50' "
            "style='display:block;border-radius:50px;border:0;' alt='RinCity' />"
            "</a></td>"
            f"<td style='vertical-align:middle;border-left:5px solid {border};'>"
            f"<a href=\"{html.escape(event['link'])}\" "
            f"style='display:block;padding:10px 12px;background:{background};"
            f"color:{foreground};text-decoration:none;'>"
            f"<span style='font-size:18px;font-weight:700;color:{foreground};'>"
            f"{html.escape(event['title'])}</span><br />"
            "<span style='font-size:14px;color:#555;'>RinCity &bull; "
            f"{html.escape(format_notification_datetime(event['date'], event['timing']))}</span>"
            f"{details_html}"
            "</a></td>"
            "</tr></table>"
        )

    def _render_event_section(self, heading: str, events: list[dict[str, str]]) -> str:
        if not events:
            return ""
        rows = "".join(self._render_event_row(event) for event in events)
        return (
            f"<h3 style='margin:24px 0 12px 0;color:#222;'>{html.escape(heading)}</h3>"
            f"{rows}"
        )


def normalize_text(value: str) -> str:
    return re.sub(r"[^a-z0-9]+", " ", value.lower()).strip()


def getenv_str(name: str, default: str | None = None) -> str | None:
    value = os.getenv(name)
    if value is None or value == "":
        return default
    return value


def getenv_int(name: str, default: int) -> int:
    value = getenv_str(name)
    return int(value) if value is not None else default


def getenv_bool(name: str, default: bool) -> bool:
    value = getenv_str(name)
    if value is None:
        return default
    normalized = value.strip().lower()
    if normalized in {"1", "true", "yes", "on"}:
        return True
    if normalized in {"0", "false", "no", "off"}:
        return False
    raise ValueError(f"invalid boolean value for {name}: {value}")


def parse_sleep_interval(value: str) -> int:
    normalized = value.strip().lower()
    if not normalized:
        raise ValueError("SLEEP_INTERVAL must not be empty")

    total_seconds = 0
    position = 0
    for match in DURATION_PART_RE.finditer(normalized):
        if match.start() != position:
            raise ValueError(
                "invalid SLEEP_INTERVAL format; use values like 30m, 12h, or 1d"
            )
        magnitude = int(match.group("value"))
        unit = match.group("unit")
        multiplier = {"s": 1, "m": 60, "h": 3600, "d": 86400}[unit]
        total_seconds += magnitude * multiplier
        position = match.end()

    if position != len(normalized) or total_seconds <= 0:
        raise ValueError("SLEEP_INTERVAL must be a positive duration like 30m or 1d")
    return total_seconds


def load_sleep_interval_seconds() -> int | None:
    value = getenv_str("SLEEP_INTERVAL")
    if value is None:
        return None
    return parse_sleep_interval(value)


def split_recipients(value: str | None) -> tuple[str, ...]:
    if not value:
        return ()
    return tuple(part.strip() for part in value.split(",") if part.strip())


def parse_time_string(value: str) -> dt.time:
    for fmt in ("%I:%M %p", "%I %p"):
        try:
            return dt.datetime.strptime(value, fmt).time()
        except ValueError:
            continue
    raise ValueError(f"unsupported time format: {value}")


def format_notification_datetime(date_str: str, timing: str) -> str:
    event_date = dt.date.fromisoformat(date_str)
    if timing.lower() == "all day":
        return f"{event_date.strftime('%a, %b')} {event_date.day} • All Day"
    return f"{event_date.strftime('%a, %b')} {event_date.day}, {timing} PDT"


def add_months(value: dt.date, months: int) -> dt.date:
    year = value.year + (value.month - 1 + months) // 12
    month = (value.month - 1 + months) % 12 + 1
    return dt.date(year, month, 1)


def month_floor(value: dt.date) -> dt.date:
    return value.replace(day=1)


def build_fetch_url(base_url: str, list_view: bool = False) -> str:
    parsed = urllib.parse.urlparse(base_url)
    query = urllib.parse.parse_qs(parsed.query, keep_blank_values=True)
    if list_view:
        query["list_view"] = ["true"]
    new_query = urllib.parse.urlencode(query, doseq=True)
    return urllib.parse.urlunparse(parsed._replace(query=new_query))


def build_month_url(base_url: str, month_start: dt.date) -> str:
    list_url = build_fetch_url(base_url, list_view=True)
    parsed = urllib.parse.urlparse(list_url)
    query = urllib.parse.parse_qs(parsed.query, keep_blank_values=True)
    query["start_date"] = [month_start.isoformat()]
    new_query = urllib.parse.urlencode(query, doseq=True)
    return urllib.parse.urlunparse(parsed._replace(query=new_query))


def absolute_mfc_url(url: str) -> str:
    return urllib.parse.urljoin("https://share.myfreecams.com", url)


def build_event_share_url(modal_url: str, event_date: dt.date) -> str:
    parsed = urllib.parse.urlparse(modal_url)
    event_id = Path(parsed.path).name
    query = urllib.parse.urlencode({"date": event_date.isoformat(), "event": event_id})
    return urllib.parse.urlunparse(
        (
            "https",
            "share.myfreecams.com",
            "/RinCity/calendar",
            "",
            query,
            "",
        )
    )


def fetch_html(url: str) -> str:
    request = urllib.request.Request(
        url,
        headers={
            "User-Agent": HTTP_USER_AGENT,
            "Accept": "text/html,application/xhtml+xml",
        },
    )
    with urllib.request.urlopen(request, timeout=30) as response:
        return response.read().decode("utf-8")


def strip_html_fragment(fragment: str) -> str:
    normalized = re.sub(r"(?i)<br\s*/?>", "\n", fragment)
    normalized = re.sub(r"(?is)<(script|style)[^>]*>.*?</\\1>", "", normalized)
    normalized = re.sub(r"(?s)<[^>]+>", "", normalized)
    normalized = html.unescape(normalized)
    normalized = normalized.replace("\r", "")
    lines = [line.strip() for line in normalized.split("\n")]
    return "\n".join(line for line in lines if line)


def extract_event_description(modal_html: str) -> str:
    patterns = [
        re.compile(
            r"<div[^>]*class=\"[^\"]*event-description[^\"]*\"[^>]*>(?P<body>.*?)</div>",
            re.DOTALL | re.IGNORECASE,
        ),
        re.compile(
            r"<div[^>]*class='[^']*event-description[^']*'[^>]*>(?P<body>.*?)</div>",
            re.DOTALL | re.IGNORECASE,
        ),
        re.compile(
            r"<div[^>]*class=\"[^\"]*description[^\"]*\"[^>]*>(?P<body>.*?)</div>",
            re.DOTALL | re.IGNORECASE,
        ),
        re.compile(
            r"<div[^>]*class='[^']*description[^']*'[^>]*>(?P<body>.*?)</div>",
            re.DOTALL | re.IGNORECASE,
        ),
        re.compile(
            r"<p[^>]*class=\"[^\"]*\bmy-2\b[^\"]*\"[^>]*>(?P<body>.*?)</p>",
            re.DOTALL | re.IGNORECASE,
        ),
        re.compile(
            r"<p[^>]*class='[^']*\bmy-2\b[^']*'[^>]*>(?P<body>.*?)</p>",
            re.DOTALL | re.IGNORECASE,
        ),
    ]
    for pattern in patterns:
        match = pattern.search(modal_html)
        if not match:
            continue
        description = strip_html_fragment(match.group("body"))
        if description:
            return description
    return ""


def extract_events(page_html: str) -> list[dict[str, str | None]]:
    events: list[dict[str, str | None]] = []

    for day_match in DAY_CONTENT_RE.finditer(page_html):
        event_date = html.unescape(day_match.group("date"))
        day_body = day_match.group("body")

        for event_match in EVENT_RE.finditer(day_body):
            event_body = event_match.group("body")
            title_match = TITLE_RE.search(event_body)
            start_time_match = START_TIME_RE.search(event_body)
            if not title_match or not start_time_match:
                continue

            title = strip_html_fragment(title_match.group("body"))
            start_time = strip_html_fragment(start_time_match.group("body"))
            css_classes = event_match.group("class")
            event_type = "all-day" if " all-day " in f" {css_classes} " else "timed"
            modal_url = html.unescape(event_match.group("modal"))

            events.append(
                {
                    "date": event_date,
                    "title": title,
                    "event_type": event_type,
                    "start_time": None if event_type == "all-day" else start_time,
                    "modal_url": modal_url,
                }
            )

    return events


def hydrate_mfc_event(raw_event: dict[str, str | None], event_date: dt.date) -> MfcEvent:
    title = str(raw_event["title"]).strip()
    event_type = str(raw_event["event_type"])
    start_time = raw_event["start_time"]
    modal_url = absolute_mfc_url(str(raw_event["modal_url"]))
    modal_html = fetch_html(modal_url)
    description = extract_event_description(modal_html)
    return MfcEvent(
        date=event_date,
        title=title,
        event_type=event_type,
        start_time=None if start_time is None else str(start_time),
        source_url=build_event_share_url(modal_url, event_date),
        modal_url=modal_url,
        description=description,
    )


def fetch_mfc_events(
    calendar_url: str, range_start: dt.date, range_end: dt.date
) -> list[MfcEvent]:
    seen: set[tuple[str, str, str, str | None]] = set()
    events: list[MfcEvent] = []
    month_start = month_floor(range_start)

    while month_start < range_end:
        fetch_url = build_month_url(calendar_url, month_start)
        page_html = fetch_html(fetch_url)
        for raw_event in extract_events(page_html):
            event_date = dt.date.fromisoformat(str(raw_event["date"]))
            if event_date < range_start or event_date >= range_end:
                continue

            title = str(raw_event["title"]).strip()
            event_type = str(raw_event["event_type"])
            start_time = raw_event["start_time"]
            key = (event_date.isoformat(), title, event_type, start_time)
            if key in seen:
                continue
            seen.add(key)

            events.append(hydrate_mfc_event(raw_event, event_date))

        month_start = add_months(month_start, 1)

    return sorted(
        events,
        key=lambda event: (
            event.date,
            0 if event.is_all_day else 1,
            event.start_time or "",
            event.title.lower(),
        ),
    )


def auth_calendar():
    creds = None
    token_json = getenv_str("GOOGLE_TOKEN_JSON")
    token_path = Path(getenv_str("GOOGLE_TOKEN_PATH", str(TOKEN_PATH)) or TOKEN_PATH)
    credentials_json = getenv_str("GOOGLE_CREDENTIALS_JSON")
    credentials_path = Path(
        getenv_str("GOOGLE_CREDENTIALS_PATH", str(CREDENTIALS_PATH)) or CREDENTIALS_PATH
    )

    try:
        if token_json:
            creds = Credentials.from_authorized_user_info(
                json.loads(token_json), SCOPES
            )
        elif token_path.exists():
            creds = Credentials.from_authorized_user_file(str(token_path), SCOPES)
    except Exception:
        creds = None

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if credentials_json:
                flow = InstalledAppFlow.from_client_config(
                    json.loads(credentials_json), SCOPES
                )
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    str(credentials_path), SCOPES
                )
            creds = flow.run_local_server(port=0)

        if not token_json:
            token_path.parent.mkdir(parents=True, exist_ok=True)
            with open(token_path, "w", encoding="utf-8") as token:
                token.write(creds.to_json())

    return build("calendar", "v3", credentials=creds)


def build_private_props(event: MfcEvent) -> dict[str, str]:
    return {
        SYNC_MARKER_KEY: SYNC_MARKER_VALUE,
        SYNC_KEY_FIELD: event.sync_key,
        SYNC_KIND_FIELD: event.kind,
        SYNC_DATE_FIELD: event.date.isoformat(),
    }


def build_event_body(event: MfcEvent) -> dict:
    start, end = event.start_end()
    return {
        "summary": event.title,
        "description": event.description,
        "start": start,
        "end": end,
        "colorId": event.color_id,
        "extendedProperties": {"private": build_private_props(event)},
    }


def list_google_events(service, calendar_id: str, range_start: dt.date, range_end: dt.date):
    time_min = dt.datetime.combine(range_start, dt.time.min, tzinfo=LOCAL_TZ).isoformat()
    time_max = dt.datetime.combine(range_end, dt.time.min, tzinfo=LOCAL_TZ).isoformat()

    events = []
    page_token = None
    while True:
        response = (
            service.events()
            .list(
                calendarId=calendar_id,
                timeMin=time_min,
                timeMax=time_max,
                singleEvents=True,
                showDeleted=False,
                orderBy="startTime",
                pageToken=page_token,
                fields=(
                    "nextPageToken,"
                    "items(id,status,summary,description,colorId,start,end,"
                    "extendedProperties,recurringEventId)"
                ),
            )
            .execute()
        )
        events.extend(response.get("items", []))
        page_token = response.get("nextPageToken")
        if not page_token:
            return events


def is_all_day_google_event(event: dict) -> bool:
    return "date" in event.get("start", {})


def parse_google_datetime(value: str) -> dt.datetime:
    return dt.datetime.fromisoformat(value.replace("Z", "+00:00"))


def google_event_date(event: dict) -> dt.date:
    start = event.get("start", {})
    if "date" in start:
        return dt.date.fromisoformat(start["date"])
    return parse_google_datetime(start["dateTime"]).astimezone(LOCAL_TZ).date()


def google_summary(event: dict) -> str:
    return str(event.get("summary", "")).strip()


def google_normalized_title(event: dict) -> str:
    return normalize_text(google_summary(event))


def google_kind(event: dict) -> str:
    summary = google_summary(event)
    if NO_STREAM_RE.search(summary):
        return "no_stream"
    if STREAM_RE.search(summary) and POSSIBLE_RE.search(summary):
        return "possible_stream"
    if STREAM_RE.search(summary):
        return "stream"
    if BIRTHDAY_RE.search(summary):
        return "birthday"
    if is_all_day_google_event(event):
        return "other_all_day"
    return "other_timed"


def is_managed_google_event(event: dict) -> bool:
    private = event.get("extendedProperties", {}).get("private", {})
    return private.get(SYNC_MARKER_KEY) == SYNC_MARKER_VALUE


def google_sync_key(event: dict) -> str | None:
    private = event.get("extendedProperties", {}).get("private", {})
    return private.get(SYNC_KEY_FIELD)


def is_recurring_instance(event: dict) -> bool:
    return bool(event.get("recurringEventId"))


def is_protected_recurring_birthday(event: dict) -> bool:
    return is_recurring_instance(event) and google_kind(event) == "birthday"


def candidate_is_updatable(event: dict, mfc_event: MfcEvent) -> bool:
    if not is_recurring_instance(event):
        return True
    return mfc_event.contains_stream


def choose_best_exact_title_match(candidates: list[dict]) -> dict | None:
    if not candidates:
        return None
    managed = [candidate for candidate in candidates if is_managed_google_event(candidate)]
    if len(managed) == 1:
        return managed[0]
    non_recurring = [
        candidate for candidate in candidates if not is_recurring_instance(candidate)
    ]
    if len(non_recurring) == 1:
        return non_recurring[0]
    if len(candidates) == 1:
        return candidates[0]
    return None


def find_match_for_mfc_event(
    mfc_event: MfcEvent, candidates: list[dict]
) -> tuple[dict | None, str | None]:
    if mfc_event.kind == "birthday":
        protected_birthdays = [
            candidate for candidate in candidates if is_protected_recurring_birthday(candidate)
        ]
        if protected_birthdays:
            return None, "protected recurring birthday already exists"

    updatable = [
        candidate for candidate in candidates if candidate_is_updatable(candidate, mfc_event)
    ]

    exact_managed = [
        candidate
        for candidate in updatable
        if is_managed_google_event(candidate) and google_sync_key(candidate) == mfc_event.sync_key
    ]
    if len(exact_managed) == 1:
        return exact_managed[0], None

    exact_title = [
        candidate
        for candidate in updatable
        if google_normalized_title(candidate) == mfc_event.normalized_title
        and is_all_day_google_event(candidate) == mfc_event.is_all_day
    ]
    chosen_exact = choose_best_exact_title_match(exact_title)
    if chosen_exact:
        return chosen_exact, None

    same_day_managed = [
        candidate
        for candidate in updatable
        if is_managed_google_event(candidate) and not is_recurring_instance(candidate)
    ]
    if len(same_day_managed) == 1:
        return same_day_managed[0], None

    if mfc_event.kind in TRANSITION_KINDS:
        transition_candidates = [
            candidate for candidate in updatable if google_kind(candidate) in TRANSITION_KINDS
        ]
        if len(transition_candidates) == 1:
            return transition_candidates[0], None

        legacy_candidates = [
            candidate
            for candidate in transition_candidates
            if google_normalized_title(candidate) in LEGACY_MANAGED_TITLES
        ]
        if len(legacy_candidates) == 1:
            return legacy_candidates[0], None

    return None, None


def extract_existing_private_props(event: dict) -> dict[str, str]:
    return event.get("extendedProperties", {}).get("private", {})


def body_matches_existing(existing: dict, desired: dict) -> bool:
    diffs = classify_event_diffs(existing, desired)
    return not diffs["visible"] and not diffs["metadata_only"]


def classify_event_diffs(existing: dict, desired: dict) -> dict[str, dict]:
    visible: dict[str, dict] = {}
    metadata_only: dict[str, dict] = {}

    if google_summary(existing) != desired["summary"]:
        visible["summary"] = {
            "existing": google_summary(existing),
            "desired": desired["summary"],
        }
    if (existing.get("description") or "") != desired["description"]:
        visible["description"] = {
            "existing": existing.get("description") or "",
            "desired": desired["description"],
        }
    if (existing.get("colorId") or "") != desired["colorId"]:
        visible["colorId"] = {
            "existing": existing.get("colorId") or "",
            "desired": desired["colorId"],
        }

    existing_start = existing.get("start", {})
    existing_end = existing.get("end", {})
    desired_start = desired["start"]
    desired_end = desired["end"]
    if existing_start != desired_start or existing_end != desired_end:
        if existing_start != desired_start:
            visible["start"] = {"existing": existing_start, "desired": desired_start}
        if existing_end != desired_end:
            visible["end"] = {"existing": existing_end, "desired": desired_end}

    private_existing = extract_existing_private_props(existing)
    private_desired = desired["extendedProperties"]["private"]
    for key, value in private_desired.items():
        if private_existing.get(key) != value:
            metadata_only[key] = {
                "existing": private_existing.get(key),
                "desired": value,
            }

    return {"visible": visible, "metadata_only": metadata_only}


def create_google_event(service, calendar_id: str, body: dict):
    return (
        service.events()
        .insert(calendarId=calendar_id, body=body, sendUpdates="none")
        .execute()
    )


def update_google_event(service, calendar_id: str, event_id: str, body: dict):
    return (
        service.events()
        .patch(calendarId=calendar_id, eventId=event_id, body=body, sendUpdates="none")
        .execute()
    )


def load_notification_config() -> NotificationConfig:
    username = getenv_str("SMTP_USERNAME")
    password = getenv_str("SMTP_PASSWORD")
    if bool(username) != bool(password):
        raise ValueError(
            "SMTP_USERNAME and SMTP_PASSWORD must both be set when SMTP auth is configured"
        )

    return NotificationConfig(
        host=getenv_str("SMTP_HOST", "smtp.gmail.com") or "smtp.gmail.com",
        port=getenv_int("SMTP_PORT", 587),
        username=username,
        password=password,
        from_addr=getenv_str("SMTP_FROM"),
        to_addrs=split_recipients(getenv_str("SMTP_TO")),
        use_starttls=getenv_bool("SMTP_USE_STARTTLS", True),
        notify_on_success=getenv_bool("NOTIFY_ON_SUCCESS", False),
        notify_on_changes=getenv_bool("NOTIFY_ON_CHANGES", True),
        notify_on_error=getenv_bool("NOTIFY_ON_ERROR", True),
        subject_prefix=getenv_str("SMTP_SUBJECT_PREFIX", "RinCity Calendar Sync")
        or "RinCity Calendar Sync",
    )


def should_send_summary_email(config: NotificationConfig, summary: SyncSummary) -> bool:
    if not config.enabled:
        return False
    if not config.to_addrs or not config.from_addr:
        return False
    return summary.has_changes and config.notify_on_changes


def send_email(config: NotificationConfig, subject: str, html_body: str) -> None:
    if not config.enabled:
        return
    if not config.from_addr or not config.to_addrs:
        raise ValueError("SMTP_FROM and SMTP_TO are required to send notifications")

    message = EmailMessage()
    message["Subject"] = subject
    message["From"] = config.from_addr
    message["To"] = ", ".join(config.to_addrs)
    message.set_content(html_body, subtype="html")

    with smtplib.SMTP(config.host, config.port, timeout=30) as server:
        if config.use_starttls:
            server.starttls()
        if config.username and config.password:
            server.login(config.username, config.password)
        server.send_message(message)


def send_summary_email(config: NotificationConfig, summary: SyncSummary) -> None:
    if not should_send_summary_email(config, summary):
        return
    mode = "DRY RUN" if summary.dry_run else "LIVE"
    subject = (
        f"{config.subject_prefix}: {summary.created} created, "
        f"{summary.updated} updated, {summary.skipped} skipped ({mode})"
    )
    send_email(config, subject, summary.render_html())


def build_preview_summary(
    source_summary: SyncSummary, mfc_events: list[MfcEvent]
) -> SyncSummary:
    preview_summary = SyncSummary(
        lines=list(source_summary.lines),
        created=0,
        updated=0,
        unchanged=0,
        skipped=0,
        mfc_events_found=source_summary.mfc_events_found,
        range_start=source_summary.range_start,
        range_end=source_summary.range_end,
        calendar_url=source_summary.calendar_url,
        calendar_id=source_summary.calendar_id,
        dry_run=True,
        preview_mode=True,
    )
    for event in mfc_events:
        preview_summary.created += 1
        preview_summary.add_event("CREATE", event, "Preview mode forced event creation")
    preview_summary.lines.extend(
        [
            "",
            "=== Preview Summary ===",
            (
                "No creates or updates were pending, so preview mode rendered "
                "all MFC events as CREATE."
            ),
            f"Preview Created: {preview_summary.created}",
            "",
        ]
    )
    return preview_summary


def write_preview_html(path: Path, summary: SyncSummary, mfc_events: list[MfcEvent]) -> Path:
    preview_summary = summary
    if not summary.has_changes:
        preview_summary = build_preview_summary(summary, mfc_events)

    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(preview_summary.render_html(), encoding="utf-8")
    return path


def send_error_email(
    config: NotificationConfig, args: argparse.Namespace, error: Exception
) -> None:
    if not (config.enabled and config.notify_on_error):
        return
    html_body = "\n".join(
        [
            "<html><body style='font-family:Arial,sans-serif;color:#222;'>",
            "<h2>RinCity calendar sync failed</h2>",
            "<p>"
            f"<strong>MFC calendar:</strong> {html.escape(args.calendar_url)}<br>"
            f"<strong>Google calendar:</strong> {html.escape(args.calendar_id)}<br>"
            f"<strong>Dry run:</strong> {args.dry_run}<br>"
            f"<strong>Error:</strong> {html.escape(str(error))}"
            "</p>",
            (
                "<pre style='white-space:pre-wrap;background:#f6f6f6;"
                "padding:12px;border:1px solid #ddd;'>"
            ),
            f"{html.escape(traceback.format_exc().strip())}"
            "</pre>",
            "</body></html>",
        ]
    )
    subject = f"{config.subject_prefix}: FAILED"
    send_email(config, subject, html_body)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Sync RinCity's MFC Share calendar into Google Calendar."
    )
    parser.add_argument(
        "--calendar-url",
        default=getenv_str("MFC_CALENDAR_URL", DEFAULT_URL),
        help=f"MFC Share calendar URL. Default: {DEFAULT_URL}",
    )
    parser.add_argument(
        "--calendar-id",
        default=getenv_str("GOOGLE_CALENDAR_ID", DEFAULT_CALENDAR_ID),
        help="Google Calendar ID to update.",
    )
    parser.add_argument(
        "--start-date",
        default=getenv_str("SYNC_START_DATE"),
        help=(
            "Start syncing from the month containing this YYYY-MM-DD date. "
            "Default: today in Pacific time."
        ),
    )
    parser.add_argument(
        "--months",
        type=int,
        default=getenv_int("SYNC_MONTHS", 2),
        help="Number of calendar months to sync, starting with the start month. Default: 2.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        default=getenv_bool("DRY_RUN", False),
        help="Print planned changes without writing to Google Calendar.",
    )
    parser.add_argument(
        "--preview-email-html",
        default=getenv_str("PREVIEW_EMAIL_HTML_PATH"),
        help=(
            "Write the HTML notification body to this file without sending email. "
            "This implies --dry-run. If there are no pending creates or updates, "
            "the preview renders all MFC events as CREATE."
        ),
    )
    return parser.parse_args()


def parse_start_date(value: str | None) -> dt.date:
    if value:
        return dt.date.fromisoformat(value)
    return dt.datetime.now(LOCAL_TZ).date()


def run_sync(args: argparse.Namespace) -> tuple[SyncSummary, list[MfcEvent]]:
    if args.months < 1:
        raise ValueError("--months must be at least 1")

    start_date = parse_start_date(args.start_date)
    range_start = month_floor(start_date)
    range_end = add_months(range_start, args.months)

    mfc_events = fetch_mfc_events(args.calendar_url, range_start, range_end)
    service = auth_calendar()
    google_events = list_google_events(service, args.calendar_id, range_start, range_end)

    summary = SyncSummary(
        mfc_events_found=len(mfc_events),
        range_start=range_start,
        range_end=range_end,
        calendar_url=args.calendar_url,
        calendar_id=args.calendar_id,
        dry_run=args.dry_run,
    )

    def emit(line: str = "") -> None:
        print(line)
        summary.lines.append(line)

    emit()
    emit("=== RinCity Calendar Sync ===")
    emit(f"MFC calendar: {args.calendar_url}")
    emit(f"Google calendar: {args.calendar_id}")
    emit(f"Range: {range_start.isoformat()} to {range_end.isoformat()} (exclusive)")
    emit(f"MFC events found: {len(mfc_events)}")
    emit(f"Mode: {'DRY RUN' if args.dry_run else 'LIVE'}")
    emit()

    google_by_date: dict[dt.date, list[dict]] = {}
    for google_event in google_events:
        if google_event.get("status") == "cancelled":
            continue
        google_by_date.setdefault(google_event_date(google_event), []).append(google_event)

    for mfc_event in mfc_events:
        desired_body = build_event_body(mfc_event)
        same_day_candidates = google_by_date.get(mfc_event.date, [])
        matched_event, skip_reason = find_match_for_mfc_event(
            mfc_event, same_day_candidates
        )

        if skip_reason:
            summary.skipped += 1
            emit(
                f"SKIP   {mfc_event.date.isoformat()}  {mfc_event.title}"
                f"  ({skip_reason})"
            )
            continue

        if matched_event is None:
            summary.created += 1
            action = "CREATE"
            summary.add_event(action, mfc_event, "Google event created or would be created")
            emit(f"{action:<6} {mfc_event.date.isoformat()}  {mfc_event.title}")
            if not args.dry_run:
                created_event = create_google_event(service, args.calendar_id, desired_body)
                google_by_date.setdefault(mfc_event.date, []).append(created_event)
            continue

        diffs = classify_event_diffs(matched_event, desired_body)
        if not diffs["visible"] and not diffs["metadata_only"]:
            summary.unchanged += 1
            emit(f"OK     {mfc_event.date.isoformat()}  {mfc_event.title}")
            continue

        if not diffs["visible"] and diffs["metadata_only"]:
            summary.metadata_only += 1
            emit(f"META   {mfc_event.date.isoformat()}  {mfc_event.title}")
            if not args.dry_run:
                updated_event = update_google_event(
                    service, args.calendar_id, matched_event["id"], desired_body
                )
                day_events = google_by_date.get(mfc_event.date, [])
                for index, current_event in enumerate(day_events):
                    if current_event.get("id") == matched_event.get("id"):
                        day_events[index] = updated_event
                        break
            continue

        summary.updated += 1
        summary.add_event(
            "UPDATE",
            mfc_event,
            f"{google_summary(matched_event)} -> {mfc_event.title}",
        )
        emit(
            f"UPDATE {mfc_event.date.isoformat()}  {google_summary(matched_event)}"
            f" -> {mfc_event.title}"
        )
        if not args.dry_run:
            updated_event = update_google_event(
                service, args.calendar_id, matched_event["id"], desired_body
            )
            day_events = google_by_date.get(mfc_event.date, [])
            for index, current_event in enumerate(day_events):
                if current_event.get("id") == matched_event.get("id"):
                    day_events[index] = updated_event
                    break

    emit()
    emit("=== Summary ===")
    emit(f"Created: {summary.created}")
    emit(f"Updated: {summary.updated}")
    emit(f"Metadata-only: {summary.metadata_only}")
    emit(f"Unchanged: {summary.unchanged}")
    emit(f"Skipped: {summary.skipped}")
    emit()
    return summary, mfc_events


def main() -> int:
    sleep_interval_seconds = load_sleep_interval_seconds()
    args = parse_args()
    notification_config = load_notification_config()
    if args.preview_email_html:
        args.dry_run = True
        sleep_interval_seconds = None

    while True:
        try:
            summary, mfc_events = run_sync(args)
        except Exception as exc:
            print(f"error: {exc}", file=sys.stderr)
            try:
                send_error_email(notification_config, args, exc)
            except Exception as email_exc:
                print(f"warning: failed to send error email: {email_exc}", file=sys.stderr)
            if sleep_interval_seconds is None:
                return 1
            print(
                f"Sleeping for {sleep_interval_seconds} seconds before retry.",
                file=sys.stderr,
            )
            time.sleep(sleep_interval_seconds)
            continue

        try:
            if args.preview_email_html:
                preview_path = Path(args.preview_email_html).expanduser().resolve()
                written_path = write_preview_html(preview_path, summary, mfc_events)
                print(f"Preview HTML written to {written_path}")
                return 0

            send_summary_email(notification_config, summary)
        except Exception as post_run_exc:
            print(f"warning: post-run action failed: {post_run_exc}", file=sys.stderr)
            if sleep_interval_seconds is None:
                return 1

        if sleep_interval_seconds is None:
            return 0

        print(f"Sleeping for {sleep_interval_seconds} seconds before next run.")
        time.sleep(sleep_interval_seconds)


if __name__ == "__main__":
    raise SystemExit(main())
