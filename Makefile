PYTHON ?= .venv/bin/python
PIP ?= .venv/bin/pip
RUFF ?= .venv/bin/ruff
PYLINT ?= .venv/bin/pylint
MYPY ?= .venv/bin/mypy

SOURCES = sync_rin_calendars.py extract_rin_calendar.py

.PHONY: venv install run dry-run lint style typecheck

venv:
	python3 -m venv .venv

install: venv
	$(PIP) install --upgrade pip
	$(PIP) install -r requirements-dev.txt

run:
	$(PYTHON) sync_rin_calendars.py

dry-run:
	$(PYTHON) sync_rin_calendars.py --dry-run

style:
	$(RUFF) check $(SOURCES)
	$(PYLINT) $(SOURCES)

typecheck:
	$(MYPY) $(SOURCES)

lint: style typecheck
