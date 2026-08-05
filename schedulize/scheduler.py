"""Fair multi-duty schedule generation for one month."""

import calendar
import re
from dataclasses import dataclass, field
from datetime import date, timedelta

import holidays

from schedulize.duty_config import DutyType
from schedulize.models import Soldier


class SchedulingError(Exception):
    pass


# duty name -> shift label -> assigned soldiers
DayAssignments = dict[str, dict[str, list[Soldier]]]


@dataclass
class DaySchedule:
    date: date
    status: str  # "duty" | "weekend" | "holiday"
    holiday_name: str | None = None
    assignments: DayAssignments | None = None


@dataclass
class MonthSchedule:
    year: int
    month: int
    duties: tuple[DutyType, ...]
    days: list[DaySchedule]
    total_counts: dict[str, int]  # soldier id -> shifts across all duties
    per_duty_counts: dict[str, dict[str, int]]  # duty name -> soldier id -> shifts
    soldiers: list[Soldier]
    warnings: list[str] = field(default_factory=list)


def parse_shift_hours(shift: str) -> tuple[int, int] | None:
    """'21-00' -> (21, 24); '23-01' -> (23, 25); non-time labels -> None."""
    match = re.fullmatch(r"(\d{1,2})-(\d{1,2})", shift.strip())
    if not match:
        return None
    start, end = int(match.group(1)), int(match.group(2))
    if end <= start:
        end += 24
    return start, end


def _overlaps(a: tuple[int, int], b: tuple[int, int]) -> bool:
    return max(a[0], b[0]) < min(a[1], b[1])


def build_schedule(
    soldiers: list[Soldier],
    year: int,
    month: int,
    duties: tuple[DutyType, ...] | list[DutyType],
) -> MonthSchedule:
    duties = tuple(duties)
    kr_holidays = holidays.KR(years=year, language="en_US")
    _, last_day = calendar.monthrange(year, month)

    total_counts = {s.id: 0 for s in soldiers}
    per_duty_counts = {d.name: {s.id: 0 for s in soldiers} for d in duties}
    last_duty_day: dict[str, date] = {}
    days: list[DaySchedule] = []
    warnings: list[str] = []

    for day_number in range(1, last_day + 1):
        current = date(year, month, day_number)
        holiday_name = kr_holidays.get(current)

        if current.weekday() >= 5:
            days.append(DaySchedule(current, "weekend", holiday_name))
            continue
        if holiday_name:
            days.append(DaySchedule(current, "holiday", holiday_name))
            continue

        pool = [s for s in soldiers if s.is_eligible_on(current)]
        for duty in duties:
            if len(pool) < duty.slots_per_day:
                raise SchedulingError(
                    f"only {len(pool)} eligible soldiers on {current.isoformat()} "
                    f"but '{duty.name}' needs {duty.slots_per_day} "
                    f"({len(duty.shifts)} shifts x {duty.soldiers_per_shift})"
                )

        assignments = _assign_day(
            current, pool, duties, total_counts, per_duty_counts,
            last_duty_day, warnings,
        )
        days.append(DaySchedule(current, "duty", holiday_name, assignments))

    return MonthSchedule(
        year, month, duties, days, total_counts, per_duty_counts,
        list(soldiers), warnings,
    )


def _assign_day(
    current: date,
    pool: list[Soldier],
    duties: tuple[DutyType, ...],
    total_counts: dict[str, int],
    per_duty_counts: dict[str, dict[str, int]],
    last_duty_day: dict[str, date],
    warnings: list[str],
) -> DayAssignments:
    """Fill every duty's shifts for one day, updating fairness tracking."""
    roster_index = {s.id: i for i, s in enumerate(pool)}
    previous_day = current - timedelta(days=1)
    busy_hours: dict[str, list[tuple[int, int]]] = {}
    assignments: DayAssignments = {}

    for duty in duties:
        duty_counts = per_duty_counts[duty.name]
        assigned_this_duty: set[str] = set()
        assignments[duty.name] = {}

        def priority(soldier: Soldier):
            return (
                total_counts[soldier.id],
                duty_counts[soldier.id],
                last_duty_day.get(soldier.id, date.min),
                roster_index[soldier.id],
            )

        for shift in duty.shifts:
            hours = parse_shift_hours(shift)
            picks: list[Soldier] = []
            for _ in range(duty.soldiers_per_shift):
                candidates = [
                    s
                    for s in pool
                    if s.id not in assigned_this_duty
                    and (
                        hours is None
                        or not any(
                            _overlaps(hours, held)
                            for held in busy_hours.get(s.id, [])
                        )
                    )
                ]
                if not candidates:
                    raise SchedulingError(
                        f"no eligible soldier free for {duty.name} {shift} on "
                        f"{current.isoformat()} (all are already assigned or "
                        "have overlapping shifts)"
                    )
                rested = [
                    s for s in candidates
                    if last_duty_day.get(s.id) != previous_day
                ]
                if not rested:
                    message = (
                        f"{current.isoformat()}: pool too small to avoid "
                        "consecutive duty days"
                    )
                    if message not in warnings:
                        warnings.append(message)
                chosen = min(rested or candidates, key=priority)
                picks.append(chosen)
                assigned_this_duty.add(chosen.id)
                if hours is not None:
                    busy_hours.setdefault(chosen.id, []).append(hours)
                total_counts[chosen.id] += 1
                duty_counts[chosen.id] += 1
            assignments[duty.name][shift] = picks

    for day_assignments in assignments.values():
        for soldiers_on_shift in day_assignments.values():
            for soldier in soldiers_on_shift:
                last_duty_day[soldier.id] = current
    return assignments
