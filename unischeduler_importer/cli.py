"""Command-line entry point: preview/JSON/export and explicit authenticated apply."""
from __future__ import annotations
import argparse
from contextlib import redirect_stdout
from io import StringIO
from collections import Counter
from datetime import datetime
import json
import os
from pathlib import Path
import sys

from .parser import parse_schedule
from .calendar import SemesterCalendar
from .bridge import build_payload, apply, ImportErrorResponse


def _json(obj: dict) -> str:
    return json.dumps(obj,ensure_ascii=False,indent=2,default=lambda x:x.isoformat() if hasattr(x,'isoformat') else str(x))


def _print_summary(doc):
    m=doc.metadata
    print(f'Source: {doc.source} / sheet: {m.sheet_name}')
    print(f'Scope: {m.faculty} | {m.speciality} | year {m.study_year} | {m.semester} {m.academic_year}')
    print(f'Schedule rows: {len(doc.entries)} | calendar instances: {len(doc.occurrences)} | errors: {len(doc.errors)}')
    if doc.issues:
        print(f'Diagnostics: {dict(Counter(i.code for i in doc.issues))}')


def _print_entries(doc,full: bool):
    for entry in doc.entries:
        print(f'row={entry.row:3}  {entry.day:10} {entry.time:11} '
              f'{entry.course_name} | teacher={entry.teacher or "∅"} | '
              f'group={entry.group or "∅"} | weeks={entry.weeks} | '
              f'room={entry.location or "∅"} | instances={len(entry.occurrences)} '
              f'| extras={entry.extra}')
        if full:
            for occurrence in entry.occurrences:
                print(f'    {occurrence.starts_at.isoformat(timespec="minutes")} → '
                      f'{occurrence.ends_at.isoformat(timespec="minutes")} (week {occurrence.week})')
    for issue in doc.issues:
        if issue.severity != 'info':
            print(f'[{issue.severity}] row={issue.row or "-"} {issue.code}: {issue.message}')


def main(argv: list[str]|None=None) -> int:
    p=argparse.ArgumentParser(prog='unischeduler',description='Parse the first XLSX worksheet and import university timetables')
    sub=p.add_subparsers(dest='command',required=True)
    parse=sub.add_parser('parse',help='Parse and preview without touching the backend')
    parse.add_argument('xlsx',type=Path)
    parse.add_argument('--calendar',type=Path,help='JSON calendar file (defaults to autumn 2026)')
    parse.add_argument('--format',choices=['text','json'],default='text')
    parse.add_argument('--output',type=Path)
    parse.add_argument('--full',action='store_true',help='Print every date occurrence (text mode)')
    imp=sub.add_parser('import',help='Prepare/import occurrences into patched backend')
    imp.add_argument('xlsx',type=Path)
    imp.add_argument('--calendar',type=Path,help='JSON calendar file (defaults to autumn 2026)')
    imp.add_argument('--department',required=True,help='Required by existing CourseCreateDto; not in XLSX')
    mode=imp.add_mutually_exclusive_group(required=True)
    mode.add_argument('--dry-run',action='store_true')
    mode.add_argument('--apply',action='store_true')
    imp.add_argument('--backend',default='http://localhost:8080/api')
    imp.add_argument('--token',default=None,help='Defaults to UNISCHEDULER_IMPORT_TOKEN env variable')
    imp.add_argument('--output',type=Path,help='Save proposed import JSON payload')
    imp.add_argument('--allow-warnings',action='store_true',help='Explicitly accept non-info diagnostics on --apply')
    args=p.parse_args(argv)
    try:
        doc=parse_schedule(args.xlsx, calendar=SemesterCalendar.from_json(args.calendar) if args.calendar else SemesterCalendar())
        if args.command=='parse':
            if args.format=='json':
                rendered=_json(doc.as_dict())
                if args.output:
                    args.output.write_text(rendered,encoding='utf-8')
                    print(f'Wrote {args.output}',file=sys.stderr)
                else:
                    print(rendered)
            else:
                if args.output:
                    capture=StringIO()
                    with redirect_stdout(capture):
                        _print_summary(doc)
                        _print_entries(doc,args.full)
                    args.output.write_text(capture.getvalue(),encoding='utf-8')
                    print(f'Wrote {args.output}')
                else:
                    _print_summary(doc)
                    _print_entries(doc,args.full)
            return 2 if doc.errors else 0
        _print_summary(doc)
        payload=build_payload(doc,args.department)
        if args.output:
            args.output.write_text(_json(payload),encoding='utf-8')
            print(f'Wrote import payload to {args.output}')
        print(f'Courses (distinct by normalized name): {len({_normal(e.course_name) for e in doc.entries})}')
        print(f'Proposed course-class instances: {len(payload["occurrences"])}')
        if args.dry_run:
            print('DRY RUN: no HTTP requests made; no database changes. This is a local preflight, not a server-side diff.')
            return 0
        warnings=[i for i in doc.issues if i.severity=='warning']
        if warnings and not args.allow_warnings:
            print('Import blocked by warnings. Inspect diagnostics and pass --allow-warnings to acknowledge.',file=sys.stderr)
            return 2
        token=args.token or os.getenv('UNISCHEDULER_IMPORT_TOKEN','')
        result=apply(args.backend,token,payload)
        print('Import completed:',_json(result))
        return 0
    except (OSError,ValueError,ImportErrorResponse) as ex:
        print(f'ERROR: {ex}',file=sys.stderr)
        return 2


def _normal(s: str) -> str:
    import re
    return re.sub(r'\s+',' ',s).strip().casefold()


if __name__=='__main__':
    raise SystemExit(main())
