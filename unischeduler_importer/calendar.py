"""Official autumn 2026-2027 study week mapping (Mon-Fri supplied; Sat inferred)."""
from __future__ import annotations
from dataclasses import dataclass
from datetime import date, timedelta
from pathlib import Path
import json
import re

WEEK_STARTS_2026 = (
    "2026-09-01", "2026-09-07", "2026-09-14", "2026-09-21",
    "2026-09-28", "2026-10-05", "2026-10-12", "2026-10-19",
    "2026-10-26", "2026-11-02", "2026-11-09", "2026-11-16",
    "2026-11-23", "2026-11-30",
)

DAYS = {
    "понеділок": 0, "вівторок": 1, "середа": 2,
    "четвер": 3, "п'ятниця": 4, "п`ятниця": 4,
    "субота": 5, "неділя": 6,
}


def normalize_day(s: str) -> int | None:
    key = re.sub(r"\s+", "", s.strip().casefold()).replace("’", "'").replace("ʼ", "'")
    return DAYS.get(key)


@dataclass(frozen=True)
class SemesterCalendar:
    academic_year: str = "2026-2027"
    weeks: tuple[date, ...] = tuple(date.fromisoformat(x) for x in WEEK_STARTS_2026)
    allow_saturday: bool = True

    def __post_init__(self):
        if not 1 <= len(self.weeks) <= 24:
            raise ValueError("Calendar must contain 1-24 academic weeks")
        for i, d in enumerate(self.weeks):
            if i and (d.weekday() != 0 or d <= self.weeks[i-1]):
                raise ValueError(f"Invalid Monday start at academic week {i+1}: {d}")
            if i == 0 and d.weekday() > 5:
                raise ValueError("The first academic week cannot start on Sunday")

    @classmethod
    def from_json(cls, path: str | Path) -> "SemesterCalendar":
        data = json.loads(Path(path).read_text(encoding='utf-8'))
        return cls(
            academic_year=data['academic_year'],
            weeks=tuple(date.fromisoformat(x) for x in data['week_starts']),
            allow_saturday=bool(data.get('allow_saturday',True)),
        )

    def date_for(self, week: int, weekday: int) -> date | None:
        if not 1 <= week <= len(self.weeks) or not 0 <= weekday <= 6:
            return None
        if weekday == 6 or (weekday == 5 and not self.allow_saturday):
            return None
        start = self.weeks[week-1]
        if week == 1 and weekday < start.weekday():
            return None  # The first academic week may start after Monday
        return start + timedelta(days=weekday - start.weekday())

    def week_for(self, day: date) -> int | None:
        for i, start in enumerate(self.weeks, 1):
            end = start + timedelta(days=(5 if self.allow_saturday else 4) - start.weekday())
            if start <= day <= end:
                return i
        return None

    def in_term(self, d: date) -> bool:
        return self.week_for(d) is not None
