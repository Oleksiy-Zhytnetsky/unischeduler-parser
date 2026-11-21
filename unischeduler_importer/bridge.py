"""Build an exact backend import payload. No network calls until apply()."""
from __future__ import annotations
from collections import Counter, defaultdict
from dataclasses import asdict
from hashlib import sha256
from urllib.request import Request, urlopen
from urllib.error import HTTPError, URLError
import json
import re

from .models import ScheduleDocument


def _normal(s: str) -> str:
    return re.sub(r'\s+', ' ', s).strip().casefold()


def build_payload(doc: ScheduleDocument, department: str) -> dict:
    if doc.errors:
        raise ValueError(f"Cannot import: {len(doc.errors)} parse errors (inspect parse --format json)")
    m=doc.metadata
    if not all([m.faculty,m.speciality,m.study_year,m.academic_year,m.semester]):
        raise ValueError("Incomplete schedule metadata")
    if not department.strip():
        raise ValueError("Backend requires nonblank department; provide --department")
    if len(department.strip())>100:
        raise ValueError("Department exceeds backend limit of 100 characters")
    result=[]
    expanded=[]
    for entry in doc.entries:
        for o in entry.occurrences:
            # Course + group + start time is the stable identity, allowing
            # teacher/location/end-time edits without destroying class IDs.
            base='\x1f'.join([_normal(entry.course_name), _normal(entry.group), o.starts_at.isoformat()])
            expanded.append((entry,o,base))
    base_counts=Counter(base for _,_,base in expanded)
    # If two classes have the same course/group/start time, disambiguate them
    # using the location and source weeks. Never collapse real overlapping rows.
    keys=[]
    for entry,o,base in expanded:
        if base_counts[base]>1:
            base+='\x1f'+_normal(entry.location)+'\x1f'+_normal(entry.weeks)
        keys.append(base)
    seen=defaultdict(int)
    for (entry,o,_),seed in zip(expanded,keys):
            seen[seed]+=1
            key=sha256(f'{seed}\x1f{seen[seed]}'.encode('utf-8')).hexdigest()
            context={
                'sheet':m.sheet_name,'row':entry.row,'day':entry.day,'time':entry.time,
                'description':entry.description,'group':entry.group,'weeks':entry.weeks,
                'location':entry.location,'extra':entry.extra,'rawCells':entry.raw_cells,
                'week':o.week,'origin':o.origin,
            }
            extras='; '.join(f'{column}: {value}' for column,value in entry.extra.items())
            result.append({
                'sourceKey':key,
                'courseName':entry.course_name,
                'teacher':entry.teacher or None,
                'group':entry.group or None,
                'location':entry.location or None,
                'description':extras or None,
                'startsAt':o.starts_at.isoformat(timespec='seconds'),
                'endsAt':o.ends_at.isoformat(timespec='seconds'),
                'sourceRow':entry.row,
                'sourceMetadata':json.dumps(context,ensure_ascii=False,sort_keys=True),
            })
    if not result:
        raise ValueError("Refusing an empty import, which would delete previous timetable records")
    return {
        'specialityName':m.speciality,
        'faculty':m.faculty,
        'department':department.strip(),
        'academicYear':m.academic_year,
        'studyYear':m.study_year,
        'semester':m.semester,
        'occurrences':result,
    }


class ImportErrorResponse(RuntimeError):
    pass


def apply(base_url: str, token: str, payload: dict, timeout: int=90) -> dict:
    if not token.strip():
        raise ValueError("An import token is required")
    url=base_url.rstrip('/')+'/admin/schedule-imports/replace'
    if not url.startswith(('http://','https://')):
        raise ValueError('Backend URL must be HTTP(S)')
    # Avoid printing auth header or token in exceptions, previews or logs.
    data=json.dumps(payload,ensure_ascii=False,separators=(',',':')).encode('utf-8')
    req=Request(url,data=data,headers={
        'Content-Type':'application/json; charset=utf-8',
        'Accept':'application/json',
        'X-Import-Token':token,
    },method='POST')
    try:
        with urlopen(req,timeout=timeout) as response:
            return json.loads(response.read().decode('utf-8'))
    except HTTPError as ex:
        body=ex.read(4000).decode('utf-8',errors='replace')
        raise ImportErrorResponse(f'Backend returned HTTP {ex.code}: {body}') from ex
    except URLError as ex:
        raise ImportErrorResponse(f'Cannot contact backend: {ex.reason}') from ex
