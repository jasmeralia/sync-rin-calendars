FROM python:3.12-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1

WORKDIR /app

COPY requirements.txt ./
RUN pip install --no-cache-dir --upgrade pip && \
    pip install --no-cache-dir -r requirements.txt

COPY sync_rin_calendars.py extract_rin_calendar.py pyproject.toml ./

ENTRYPOINT ["python", "sync_rin_calendars.py"]
