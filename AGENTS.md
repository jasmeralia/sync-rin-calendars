# AGENTS

## Repo Purpose

This repo syncs RinCity's public MFC Share calendar into Google Calendar.

## Important Behavior

- MFC Share is the source of truth inside the selected sync window.
- The sync is idempotent for managed events.
- Do not add deletion of orphaned Google events unless explicitly requested.
- Do not update recurring birthday events unless the MFC title includes `stream`.
- Timed events end at `11:59 PM` Pacific.
- Color mapping is intentional:
  - all-day `No Stream` -> red
  - other all-day -> purple
  - possible/potential timed stream -> yellow
  - other timed stream -> green

## Working Conventions

- Use `.venv` for local execution.
- Install dependencies with `make install`.
- `requirements.txt` is runtime-only; `requirements-dev.txt` includes lint/typecheck tools.
- Run checks with `make lint`.
- Prefer environment-variable support over hard-coded secrets.
- Do not commit real OAuth secrets or live tokens.

## Main Entry Points

- `sync_rin_calendars.py`
- `extract_rin_calendar.py`

## Deployment Notes

- The script is designed for cron and container execution.
- Docker deployments can use `GOOGLE_CREDENTIALS_JSON` and `GOOGLE_TOKEN_JSON`.
- Gmail SMTP notifications are optional and env-driven.
- `--preview-email-html` / `PREVIEW_EMAIL_HTML_PATH` writes the HTML email body to disk.
