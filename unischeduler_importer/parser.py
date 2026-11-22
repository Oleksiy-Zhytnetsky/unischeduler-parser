"""Conservative Excel schedule interpretation. No HTTP calls or database writes."""
from __future__ import annotations
from datetime import date, datetime, time
from pathlib import Path
import re

from .xlsx import read_first_sheet, Worksheet
from .calendar import SemesterCalendar, normalize_day
from .models import ScheduleDocument, ScheduleEntry, ScheduleMetadata, Occurrence, Issue

TIME_RE = re.compile(r"^\s*(\d{1,2}):(\d{2})\s*[-–—−]\s*(\d{1,2}):(\d{2})\s*$")
WEEKS_RE = re.compile(r"^\s*\d{1,2}(?:\s*[-–—−]\s*\d{1,2}|(?:\s*[,;]\s*\d{1,2})*)\s*$")
DATE_RE = re.compile(r"^\s*(\d{1,2})\.(\d{1,2})(?:\.(\d{2,4}))?\.?\s*$")

ACADEMIC_TITLES = (
    r"(?:"
    r"проф\.?|"
    r"доц\.?|"
    r"ст\.?\s*викл\.?|"
    r"[сc][тn][сn]\.?\s*викл\.?|"
    r"ас\.?"
    r")"
)
TEACHER_RE = re.compile(
    rf"(?:\s*,\s*|\s+)(?={ACADEMIC_TITLES}(?:\s|\.|[А-ЯІЇЄҐ]|$))",
    re.IGNORECASE,
)
TITLE_RE = re.compile(
    rf"^{ACADEMIC_TITLES}",
    re.IGNORECASE,
)


def _canonical(s: str) -> str:
    return re.sub(r"\s+", " ", s).strip()


def _col_letter(col: int) -> str:
    text = ""
    while col:
        col, rem = divmod(col-1, 26)
        text = chr(65+rem)+text
    return text


def _title_teacher(s: str) -> tuple[str,str|None]:
    """Separate only with evidence of an academic title; preserve the original description."""
    s = _canonical(s)
    for m in TEACHER_RE.finditer(s):
        suffix = s[m.end():].lstrip()
        if TITLE_RE.match(suffix):
            name = s[:m.start()].strip(' ,')
            if name:
                return name, suffix
    return s, None


def _metadata(sheet: Worksheet, hdr: int) -> ScheduleMetadata:
    head = "\n".join(v for r,cols in sheet.rows.items() if r<hdr for v in cols.values())
    faculty_match = re.search(r"факультет\s+([^\n]+)", head, re.I)
    specialty_match = re.search(r"спеціальн(?:ість|iсть)\s*[\"«]?(.+?)[\"»]\s*,?\s*(\d+)\s*р\.?\s*н", head, re.I)
    if not specialty_match:
        specialty_match = re.search(r"спеціальн(?:ість|iсть)\s+([^\n,]+),\s*(\d+)\s*р\.?\s*н", head, re.I)
    acad_match = re.search(r"(20\d{2})\s*[-–—]\s*(20\d{2})", head)
    semester = "AUTUMN" if re.search(r"осін",head,re.I) else "SPRING" if re.search(r"веснян",head,re.I) else None
    return ScheduleMetadata(
        original_metadata={str(r): _canonical(" | ".join(str(v) for v in cols.values()))
                           for r,cols in sorted(sheet.rows.items()) if r < hdr and any(cols.values())},
        faculty=_canonical(faculty_match.group(1)) if faculty_match else None,
        speciality=_canonical(specialty_match.group(1)).strip('"«» ') if specialty_match else None,
        study_year=int(specialty_match.group(2)) if specialty_match else None,
        academic_year=f"{acad_match.group(1)}-{acad_match.group(2)}" if acad_match else None,
        semester=semester,
        sheet_name=sheet.name,
    )


def _header(sheet: Worksheet) -> tuple[int, dict[str,int]]:
    for r, cols in sorted(sheet.rows.items()):
        norm = {c:_canonical(v).casefold() for c,v in cols.items()}
        def find(prefix):
            return next((c for c,v in norm.items() if v.startswith(prefix)),None)
        m = {
            'day': find('день'), 'time': find('час'),
            'description': find('дисципліна'), 'group': find('група'),
            'weeks': find('тижні'), 'location': find('аудиторія'),
        }
        if all(v is not None for v in m.values()):
            return r,m
    raise ValueError("Could not locate the six expected schedule columns")


def _parse_time(text: str) -> tuple[time,time] | None:
    m=TIME_RE.fullmatch(text)
    if not m:
        return None
    try:
        a,b,c,d = map(int,m.groups())
        start,end=time(a,b),time(c,d)
        return (start,end) if start<end else None
    except ValueError:
        return None


def _parse_date(expression: str, expected_day: int, calendar: SemesterCalendar, row: int, issues: list[Issue]) -> date | None:
    m = DATE_RE.fullmatch(expression)
    if not m:
        return None
    dd,mm=int(m.group(1)),int(m.group(2))
    year=int(m.group(3)) if m.group(3) else int(calendar.academic_year[:4])
    if year < 100:
        year += 2000
    try:
        given = date(year,mm,dd)
    except ValueError:
        issues.append(Issue('error','bad_date',f'Invalid calendar date: {expression!r}',row,'E'))
        return None
    if given.year != int(calendar.academic_year[:4]):
        # The 2025-09-01 entries in 1.xlsx are under Tuesday, while 2026-09-01
        # is Tuesday. Correct only when the workbook year, day and term agree.
        candidate=date(int(calendar.academic_year[:4]),mm,dd)
        if candidate.weekday()==expected_day and calendar.in_term(candidate):
            issues.append(Issue('warning','corroborated_year_correction',
                f'{expression} interpreted as {candidate.isoformat()} (year disagrees with workbook); source preserved',row,'E'))
            given=candidate
        else:
            issues.append(Issue('error','year_mismatch',
                f'{expression} is outside {calendar.academic_year}; manual correction required',row,'E'))
            return None
    if not calendar.in_term(given):
        issues.append(Issue('error','date_outside_calendar',f'{given.isoformat()} outside semester calendar',row,'E'))
        return None
    if given.weekday()!=expected_day:
        issues.append(Issue('error','date_weekday_mismatch',
                            f'{expression} is {given.strftime("%A")}, not {expected_day}',row,'E'))
        return None
    return given


def _resolve_dates(expression: str, day: int, calendar: SemesterCalendar, row: int, issues: list[Issue]) -> list[tuple[date,int|None,str]]:
    expr=expression.strip()
    if not expr:
        issues.append(Issue('error','missing_weeks','Missing weeks/date expression',row,'E'))
        return []
    if DATE_RE.fullmatch(expr):
        day_date=_parse_date(expr,day,calendar,row,issues)
        return [(day_date,calendar.week_for(day_date),'explicit_date')] if day_date else []
    if not WEEKS_RE.fullmatch(expr):
        issues.append(Issue('error','unresolved_weeks',f'Unknown weeks expression {expr!r}',row,'E'))
        return []
    norm=re.sub(r'\s+','',expr).translate(str.maketrans('–—−','---'))
    if '-' in norm:
        left,right=map(int,norm.split('-',1))
        if left>right:
            issues.append(Issue('error','reversed_weeks',f'Bad week range: {expr}',row,'E'))
            return []
        week_numbers=list(range(left,right+1))
    else:
        week_numbers=[int(x) for x in re.split('[,;]',norm)]
    seen=set()
    result=[]
    for week in week_numbers:
        if week in seen:
            issues.append(Issue('warning','duplicate_week',f'Duplicate week {week}',row,'E'))
            continue
        seen.add(week)
        d=calendar.date_for(week,day)
        if d is None:
            if week==1 and day==0:
                issues.append(Issue('warning','week_one_monday','Week 1 starts Tuesday; Monday has no class',row,'E'))
            else:
                issues.append(Issue('error','week_not_in_calendar',f'Week {week}, weekday {day} not in official calendar',row,'E'))
            continue
        result.append((d,week,'academic_week'))
    return result


def parse_schedule(path: str|Path, calendar: SemesterCalendar|None=None) -> ScheduleDocument:
    """Read only the FIRST worksheet and expand each timetable row to date instances."""
    calendar=calendar or SemesterCalendar()
    sheet=read_first_sheet(path)
    header,mapping=_header(sheet)
    meta=_metadata(sheet,header)
    meta.original_headers={field:sheet.get(header,col) for field,col in mapping.items()}
    issues: list[Issue]=[]
    for field in ('faculty','speciality','study_year','academic_year','semester'):
        if not getattr(meta,field):
            issues.append(Issue('error','missing_metadata',f'Missing metadata: {field}'))
    if meta.academic_year and meta.academic_year!=calendar.academic_year:
        issues.append(Issue('error','calendar_mismatch',f'Workbook {meta.academic_year}; calendar {calendar.academic_year}'))
    current_day: str|None=None
    current_day_num: int|None=None
    current_time: str|None=None
    entries=[]
    max_row=max(sheet.rows,default=header)
    for r in range(header+1,max_row+1):
        cols=sheet.rows.get(r,{})
        day=cols.get(mapping['day'],'').strip()
        slot=cols.get(mapping['time'],'').strip()
        if day:
            current_day=day
            current_day_num=normalize_day(day)
            current_time=None  # no slot inherited across days
            if current_day_num is None:
                issues.append(Issue('error','unknown_day',f'Unrecognized weekday: {day}',r,_col_letter(mapping['day'])))
        if slot:
            current_time=slot
        desc=cols.get(mapping['description'],'').strip()
        if not desc:
            if any(cols.get(mapping[k],'').strip() for k in ('group','weeks','location')):
                issues.append(Issue('warning','orphan_row','Details without discipline/description',r))
            continue
        group=cols.get(mapping['group'],'').strip()
        weeks=cols.get(mapping['weeks'],'').strip()
        location=cols.get(mapping['location'],'').strip()
        lastcol=max(cols, default=6)
        extras={_col_letter(c):cols[c] for c in range(max(mapping.values())+1,lastcol+1) if c in cols}
        title,teacher=_title_teacher(desc)
        entry=ScheduleEntry(r,current_day or '',current_time or '',desc,group,weeks,location,extras,
                            title,teacher,group.casefold()=='лекція',
                            raw_cells={_col_letter(c):v for c,v in cols.items()})
        entries.append(entry)
        if not teacher:
            issues.append(Issue('info','teacher_unspecified','Teacher not identified; may be an event or unspecified',r))
        if not group:
            issues.append(Issue('info','group_unspecified','No group specified',r))
        if not location:
            issues.append(Issue('warning','location_unspecified','No location provided',r))
        if current_day_num is None:
            issues.append(Issue('error','missing_day','No weekday context',r))
        parsed_times=_parse_time(current_time or '')
        if parsed_times is None:
            issues.append(Issue('error','bad_time',f'Invalid/missing class time: {current_time!r}',r,_col_letter(mapping['time'])))
        if parsed_times and current_day_num is not None:
            start,end=parsed_times
            for d,week,kind in _resolve_dates(weeks,current_day_num,calendar,r,issues):
                entry.occurrences.append(Occurrence(datetime.combine(d,start), datetime.combine(d,end),week,kind,r))
        entry.unresolved=not bool(entry.occurrences)
        if len(title)>100:
            issues.append(Issue('error','course_name_length',f'Course name >100 characters ({len(title)})',r))
        if teacher and len(teacher)>100:
            issues.append(Issue('error','teacher_length','Teacher exceeds backend 100-character limit',r))
        if len(group)>50:
            issues.append(Issue('error','group_length','Group exceeds backend 50-character limit',r))
        if len(location)>100:
            issues.append(Issue('error','location_length','Location exceeds backend 100-character limit',r))
    if not entries:
        issues.append(Issue('error','no_entries','No schedule entries identified'))
    return ScheduleDocument(str(path),meta,entries,issues)
