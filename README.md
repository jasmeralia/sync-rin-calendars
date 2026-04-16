# RinCity Calendar Sync

Copies the public MFC Share calendar for `RinCity` into Google Calendar and keeps it in sync.

The sync is conservative:

- MFC Share is treated as the source of truth for matching events in the selected sync window.
- Existing unrelated Google events are left alone.
- Recurring birthday events are not modified unless the MFC title includes `stream`.
- Missing Google events are created.
- Matching Google events are updated when the MFC title, timing, description, or color changes.
- Google events that no longer appear on MFC are not deleted.

## Color Rules

- All-day `No Stream` events: red
- Other all-day events: purple
- Timed events with `possible`, `potential`, `maybe`, `tentative`, `probably`, `might`, or `tbd`: yellow
- Other timed events: green

## Timed Events

- Timed MFC events start at the MFC start time, usually `8:00 PM` Pacific.
- Timed Google events end at `11:59 PM` Pacific.

## Local Setup

1. Create the virtual environment and install dependencies:

```bash
make install
```

`make install` uses `requirements-dev.txt`, which includes the runtime dependencies from
`requirements.txt` plus `ruff`, `pylint`, and `mypy` for local development.

2. Put OAuth files in the repo root if you are not using environment variables:

- `credentials.json`
- `token.json`

3. Run a dry-run:

```bash
make dry-run
```

4. Run a live sync:

```bash
make run
```

### Preview Notification HTML

To render the HTML notification body to a file without sending email:

```bash
.venv/bin/python sync_rin_calendars.py --preview-email-html preview_email.html
```

This implies `--dry-run`. If the sync currently has no pending creates or updates, the
preview mode renders all MFC events as `CREATE` so the email styling can still be reviewed.

## Environment Variables

The script supports file-based credentials or raw JSON values.

### Calendar Sync

- `MFC_CALENDAR_URL`
- `GOOGLE_CALENDAR_ID`
- `SYNC_START_DATE`
- `SYNC_MONTHS`
- `DRY_RUN`
- `PREVIEW_EMAIL_HTML_PATH`

### Google OAuth

Use either file paths:

- `GOOGLE_CREDENTIALS_PATH`
- `GOOGLE_TOKEN_PATH`

Or raw JSON values:

- `GOOGLE_CREDENTIALS_JSON`
- `GOOGLE_TOKEN_JSON`

If `GOOGLE_TOKEN_JSON` is used, refreshed credentials stay in memory for that run. If `GOOGLE_TOKEN_PATH` is used, refreshed credentials are written back to disk.

## Email Notifications

Notifications are optional and use Gmail SMTP by default.

- `EMAIL_NOTIFICATIONS_ENABLED`
- `SMTP_HOST`
- `SMTP_PORT`
- `SMTP_USE_STARTTLS`
- `SMTP_USERNAME`
- `SMTP_PASSWORD`
- `SMTP_FROM`
- `SMTP_TO`
- `SMTP_SUBJECT_PREFIX`
- `NOTIFY_ON_SUCCESS`
- `NOTIFY_ON_CHANGES`
- `NOTIFY_ON_ERROR`

`SMTP_TO` accepts a comma-separated list.

Recommended for Gmail:

- `SMTP_HOST=smtp.gmail.com`
- `SMTP_PORT=587`
- `SMTP_USE_STARTTLS=true`
- Use a Gmail app password, not your normal login password.

## Cron

Use the virtualenv interpreter explicitly:

```cron
*/30 * * * * /home/morgan/gcal/.venv/bin/python /home/morgan/gcal/sync_rin_calendars.py >> /home/morgan/gcal/sync.log 2>&1
```

## Docker

Build:

```bash
docker build -t rincity-calendar-sync .
```

GitHub Actions publishes the image to `ghcr.io/<owner>/<repo>` on every push to `master`
and on every pushed tag. A git tag such as `v1.2.3` becomes the image tag `1.2.3`.

Run:

```bash
docker run --rm \
  -e GOOGLE_CALENDAR_ID='your-calendar-id' \
  -e GOOGLE_CREDENTIALS_JSON='{"installed":{...}}' \
  -e GOOGLE_TOKEN_JSON='{"token":"...","refresh_token":"...","token_uri":"https://oauth2.googleapis.com/token","client_id":"...","client_secret":"...","scopes":["https://www.googleapis.com/auth/calendar"]}' \
  rincity-calendar-sync --dry-run
```

See [docker-compose.example.yml](./docker-compose.example.yml) for a direct-environment-variable example suitable for TrueNAS SCALE app configuration.

## Checks

Run linting and typing:

```bash
make lint
```

Individual targets:

- `make style`
- `make typecheck`

## Files

- `sync_rin_calendars.py`: main sync script
- `extract_rin_calendar.py`: MFC Share event extractor
- `Dockerfile`: container image
- `requirements.txt`: runtime dependencies
- `requirements-dev.txt`: runtime + developer tooling dependencies
- `docker-compose.example.yml`: environment-variable deployment example
- `Makefile`: local install and quality targets
