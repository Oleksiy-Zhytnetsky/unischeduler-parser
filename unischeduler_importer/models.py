"""Backend-independent, serializable schedule domain objects.

All spreadsheet fields are stored as strings. Interpretations are ADDITIONAL,
not replacements for source data. Source Excel coordinates are one-based.
"""
from __future__ import annotations

from dataclasses import dataclass, field, asdict
from datetime import datetime
from typing import Literal


@dataclass(slots=True)
class Issue:
    severity: Literal["info", "warning", "error"]
    code: str
    message: str
    row: int | None = None
    column: str | None = None


@dataclass(slots=True)
class ScheduleMetadata:
    faculty: str | None
    speciality: str | None
    study_year: int | None
    academic_year: str | None
    semester: str | None
    sheet_name: str
    original_headers: dict[str, str] = field(default_factory=dict)
    original_metadata: dict[str, str] = field(default_factory=dict)


@dataclass(slots=True)
class Occurrence:
    starts_at: datetime
    ends_at: datetime
    week: int | None
    origin: Literal["academic_week", "explicit_date"]
    source_row: int


@dataclass(slots=True)
class ScheduleEntry:
    row: int
    day: str
    time: str
    description: str
    group: str
    weeks: str
    location: str
    extra: dict[str, str]
    course_name: str
    teacher: str | None
    lecture: bool
    occurrences: list[Occurrence] = field(default_factory=list)
    unresolved: bool = False
    raw_cells: dict[str, str] = field(default_factory=dict)


@dataclass(slots=True)
class ScheduleDocument:
    source: str
    metadata: ScheduleMetadata
    entries: list[ScheduleEntry]
    issues: list[Issue]

    @property
    def errors(self) -> list[Issue]:
        return [i for i in self.issues if i.severity == "error"]

    @property
    def occurrences(self) -> list[Occurrence]:
        return [o for entry in self.entries for o in entry.occurrences]

    def as_dict(self) -> dict:
        return asdict(self)
