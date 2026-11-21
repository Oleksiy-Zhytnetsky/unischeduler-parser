from pathlib import Path
from datetime import datetime
from collections import Counter
import json
import pytest

from unischeduler_importer import parse_schedule, SemesterCalendar
from unischeduler_importer.bridge import build_payload, apply
from unischeduler_importer.parser import _resolve_dates
from unischeduler_importer.xlsx import read_first_sheet

BASE=Path(__file__).parent/'fixtures'

@pytest.mark.parametrize('sample,rows,occurrences',[('1',81,619),('3',58,573),('4',30,301)])
def test_sample_workbooks(sample,rows,occurrences):
    result=parse_schedule(BASE/f'{sample}.xlsx')
    assert result.metadata.sheet_name=='Аркуш1'  # first of three sheets
    assert result.metadata.academic_year=='2026-2027'
    assert result.metadata.semester=='AUTUMN'
    assert len(result.entries)==rows
    assert len(result.occurrences)==occurrences
    assert not result.errors
    assert all(o.starts_at<o.ends_at for o in result.occurrences)
    assert all(o.starts_at.year==2026 for o in result.occurrences)


def test_first_year_named_groups_one_off_events_and_saturdays():
    result=parse_schedule(BASE/'1.xlsx')
    row13=next(e for e in result.entries if e.row==13)
    assert row13.group=='А51'
    assert row13.course_name=='Англійська мова'
    assert row13.teacher=='ст. викл. С.О. Гісем'
    assert row13.occurrences[0].starts_at.isoformat()=='2026-09-07T10:00:00'
    one_off=next(e for e in result.entries if e.row==33)
    assert len(one_off.occurrences)==1
    assert one_off.occurrences[0].starts_at.date().isoformat()=='2026-09-07'
    typo=next(e for e in result.entries if e.row==36)
    assert typo.weeks=='01.09.25'   # original preserved
    assert typo.occurrences[0].starts_at.isoformat()=='2026-09-01T11:40:00'
    assert sum(i.code=='corroborated_year_correction' for i in result.issues)==2
    sat=next(e for e in result.entries if e.row==83)
    assert sat.occurrences[0].starts_at.date().isoformat()=='2026-09-19'  # week 3 Saturday
    assert any(e.teacher is None for e in result.entries)
    assert any(e.group=='' for e in result.entries)


def test_third_year_preserves_nonstandard_class_time():
    result=parse_schedule(BASE/'3.xlsx')
    c=next(e for e in result.entries if e.row==22)
    assert c.time=='11:40-13:01'
    assert all(o.ends_at.minute==1 for o in c.occurrences)
    assert len(c.occurrences)==7


def test_fourth_year_keeps_extra_column_and_detects_unseparated_teacher():
    result=parse_schedule(BASE/'4.xlsx')
    c=next(e for e in result.entries if e.row==32)
    assert c.extra['G']=='10к'
    swift=next(e for e in result.entries if e.row==18)
    assert swift.course_name=='Вступ до мікросервісної архітектури з використанням Spring Boot'
    assert 'Андрощук' in (swift.teacher or '')


def test_calendar_boundaries():
    cal=SemesterCalendar()
    assert cal.date_for(1,0) is None
    assert cal.date_for(1,1).isoformat()=='2026-09-01'
    assert cal.date_for(1,5).isoformat()=='2026-09-05'
    assert cal.date_for(14,5).isoformat()=='2026-12-05'
    assert cal.date_for(15,2) is None
    assert cal.week_for(cal.date_for(14,5))==14


def test_unknown_weeks_reports_error_without_guessing():
    issues=[]
    result=_resolve_dates('1,foo,3',2,SemesterCalendar(),42,issues)
    assert result==[]
    assert any(i.code=='unresolved_weeks' and i.row==42 for i in issues)


def test_payload_is_stable_and_nullable():
    doc=parse_schedule(BASE/'1.xlsx')
    first=build_payload(doc,'математики')
    second=build_payload(parse_schedule(BASE/'1.xlsx'),'математики')
    assert first==second
    assert len(first['occurrences'])==619
    assert len({r['sourceKey'] for r in first['occurrences']})==619
    assert any(o['teacher'] is None for o in first['occurrences'])
    assert any(o['location'] is None for o in first['occurrences'])
    assert all(o['courseName'] for o in first['occurrences'])
    metadata=json.loads(first['occurrences'][0]['sourceMetadata'])
    assert metadata['sheet']=='Аркуш1'
    assert metadata['row']==13
    assert first['studyYear']==1


def test_backend_request_shape_and_auth(monkeypatch):
    payload=build_payload(parse_schedule(BASE/'4.xlsx'),'кафедра ІПЗ')
    captured={}
    class Response:
        def __enter__(self):return self
        def __exit__(self,*args):return False
        def read(self):return b'{"scopeId":8,"classesCreated":301}'
    def mock_urlopen(req, timeout):
        captured['url']=req.full_url
        captured['method']=req.get_method()
        captured['token']=req.get_header('X-import-token')
        captured['data']=json.loads(req.data)
        return Response()
    monkeypatch.setattr('unischeduler_importer.bridge.urlopen',mock_urlopen)
    result=apply('http://localhost:8080/api','example-token',payload)
    assert captured['url']=='http://localhost:8080/api/admin/schedule-imports/replace'
    assert captured['method']=='POST'
    assert captured['token']=='example-token'
    assert captured['data']==payload
    assert result['classesCreated']==301


def test_defensive_empty_department():
    with pytest.raises(ValueError,match='department'):
        build_payload(parse_schedule(BASE/'3.xlsx'),'')


def test_cli_outputs_complete_unicode_text_and_json(tmp_path):
    from unischeduler_importer.cli import main
    output=tmp_path/'preview.txt'
    assert main(['parse',str(BASE/'4.xlsx'),'--output',str(output)])==0
    assert 'Проектування програмних систем' in output.read_text(encoding='utf-8')
    datafile=tmp_path/'parsed.json'
    assert main(['parse',str(BASE/'4.xlsx'),'--format','json','--output',str(datafile)])==0
    doc=json.loads(datafile.read_text(encoding='utf-8'))
    assert len(doc['entries'])==30
    assert doc['entries'][0]['raw_cells']['C'].startswith('Проектування')


def test_custom_calendar_matches_default():
    from unischeduler_importer import SemesterCalendar
    f=BASE.parent.parent/'calendars'/'autumn_2026.json'
    assert SemesterCalendar.from_json(f)==SemesterCalendar()
