# UNIScheduler parser & schedule importer

A first-sheet-only .xlsx parser, date resolver, console/JSON preview and authenticated,
transactional import client for the **UNIScheduler backend**.
No Excel runtime, HTTP, or other third-party packages are needed: Python 3.12+
standard library is sufficient. Optional `pytest` runs the regression tests.

## Quick start (Windows / PowerShell)

```powershell
py -3.12 -m pip install -e .
# Preview every Excel row, dates and diagnostics (NO HTTP)
unischeduler parse tests/fixtures/4.xlsx
# Preview all calendar instances as well
unischeduler parse tests/fixtures/4.xlsx --full
# Save complete in-memory data as UTF-8 JSON
unischeduler parse tests/fixtures/4.xlsx --format json --output parsed.json
# Preview a backend import payload; the department must be provided explicitly
unischeduler import tests/fixtures/4.xlsx --department "<real department>" --dry-run --output payload.json
```

If running the app in an older PowerShell that displays mojibake, use:

```powershell
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
```

## Calendar configuration

The supplied calendar is also available in `calendars/autumn_2026.json`.
`--calendar path/to/calendar.json` works with both `parse` and `import`.
The JSON object uses `academic_year`, `week_starts` (ISO `YYYY-MM-DD` per
academic week), and `allow_saturday`. The built-in default matches 
2026-2027 autumn semester dates. For later semesters, supply a new calendar.

## Apply a schedule import

First deploy the backend (see `README.md`, `sql/SCHEDULE_IMPORT.md` there), which contains
`POST /api/admin/schedule-imports/replace` with transactional replacement.
Configure its environment:

```text
UNISCHEDULER_IMPORT_ENABLED=true
UNISCHEDULER_IMPORT_TOKEN=<import token>
```

Then, in the client PowerShell session:

```powershell
unischeduler import tests/fixtures/4.xlsx --department "<real department>" --dry-run
unischeduler import tests/fixtures/4.xlsx --department "<real department>" --token "<import token>" --apply
# If warnings were reported, examine them first, then explicitly acknowledge:
unischeduler import tests/fixtures/1.xlsx --department "<real department>" --token "<import token>" --apply --allow-warnings
```

The `--dry-run` option is **local-only**: it shows parsed/importable counts and
optionally emits the exact request JSON, but does not contact the backend.
`--apply` sends the full payload. The backend performs the actual diff and
transactional authoritative replacement, returning created/updated/archived/
deleted counts.

## Library API

```python
from unischeduler_importer import parse_schedule, SemesterCalendar
from unischeduler_importer.bridge import build_payload, apply

schedule = parse_schedule('4.xlsx', calendar=SemesterCalendar())
for item in schedule.entries:
    print(item.row, item.course_name, item.teacher, item.group, item.weeks, item.raw_cells)
    for occurrence in item.occurrences:
        print(occurrence.starts_at, occurrence.ends_at)

payload = build_payload(schedule, department='Real university department')
# apply('http://localhost:8080/api', token, payload) # only after explicit approval
```

## Data handling

- Reads the **first workbook worksheet only**, selected via .xlsx workbook
  relationships; other sheets are ignored entirely.
- Finds the six header columns by Ukrainian header names, not fixed row number.
- Keeps source strings, original Excel coordinates, full source description,
  day/time/room/group/weeks, and unknown columns. Numbers are represented as
  strings, not coerced into the meaning of a student group or week.
- Carries forward day and class time (for merged/sparse rows) but not subject,
  group, weeks or location. Multiple entries at the same time remain separate.
- Separates course and teacher only for recognizable academic titles. Otherwise
  preserves full text as course name and teacher=None. Missing teacher/group/
  location remain null in the import request.
- Resolves week ranges (`2-14`), comma lists (`6,8,10,12`), single weeks (`3`),
  explicit dates (`03.09.26`, `7.09`). Unknown expressions become errors, not
  invented dates. A week's Monday-Friday dates come from the supplied official
  calendar. Week 1 has no Monday; Saturday means the day following the listed
  week's Friday, an explicit documented assumption.
- **2025 year typos:** `01.09.25` on Tuesday rows 36/39 in `1.xlsx` are
  interpreted as **2026-09-01** because workbook academic year, weekday and
  calendar uniquely corroborate it. Both yield warnings, while original text
  remains available. `--apply` refuses warnings unless `--allow-warnings` is
  specified. Inconsistent dates without strong corroboration are errors.
- The `CourseClass` backend has 100-character course/teacher/location limits,
  50-character group limit. Parser marks violations as errors. The specialities,
  faculty and department have 100-character limits.
- Adds SHA-256 stable import keys, independent of Excel source row position.
  Keys preserve an existing `CourseClass` across teacher/location/source-row
  changes. Course name, group, time and date are part of the identity.

## Regression tests

```powershell
py -m pip install -e ".[test]"
py -m pytest -q
```

## Boundaries / production checks

1. `department` must be supplied as an actual value; it is not extractable from
   the sample Excel files. **This is the one required import input not provided.**
2. Use a backed-up PostgreSQL instance to test the patched backend first.
3. The backend does not currently filter the calendar by speciality/study year/
   group; add filters before exposing all imported classes to users.
