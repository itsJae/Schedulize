from datetime import date, timedelta

import pytest

from schedulize.duty_config import DEFAULT_DUTIES, DutyType
from schedulize.models import Rank, Soldier
from schedulize.scheduler import (
    SchedulingError,
    build_schedule,
    parse_shift_hours,
)

NIGHT_WATCH = DutyType("Night Watch", ("22-24", "00-02", "02-04", "04-06"), 2)


def make_roster(count, rank=Rank.CORPORAL, retirement=date(2030, 1, 1)):
    return [
        Soldier(
            id=f"24-{i:04d}",
            name=f"Soldier {i}",
            rank=rank,
            retirement_date=retirement,
            medical_condition="",
        )
        for i in range(count)
    ]


def all_assigned(day):
    """(duty_name, shift, soldier) triples for a duty day."""
    return [
        (duty_name, shift, soldier)
        for duty_name, shifts in day.assignments.items()
        for shift, soldiers in shifts.items()
        for soldier in soldiers
    ]


class TestShiftHourParsing:
    def test_plain_range(self):
        assert parse_shift_hours("06-09") == (6, 9)

    def test_range_ending_at_midnight(self):
        assert parse_shift_hours("21-00") == (21, 24)
        assert parse_shift_hours("22-24") == (22, 24)

    def test_range_crossing_midnight(self):
        assert parse_shift_hours("23-01") == (23, 25)

    def test_non_time_shift_is_none(self):
        assert parse_shift_hours("After meals") is None


class TestCalendar:
    def test_covers_every_day_of_leap_february(self):
        schedule = build_schedule(make_roster(20), 2024, 2, [NIGHT_WATCH])
        assert len(schedule.days) == 29
        assert schedule.days[-1].date == date(2024, 2, 29)

    def test_covers_every_day_of_non_leap_february(self):
        schedule = build_schedule(make_roster(20), 2026, 2, [NIGHT_WATCH])
        assert len(schedule.days) == 28

    def test_weekends_have_no_assignments(self):
        schedule = build_schedule(make_roster(20), 2026, 9, [NIGHT_WATCH])
        for day in schedule.days:
            if day.date.weekday() >= 5:
                assert day.status == "weekend"
                assert day.assignments is None

    def test_korean_holidays_have_no_assignments_and_are_named(self):
        schedule = build_schedule(make_roster(20), 2026, 9, [NIGHT_WATCH])
        chuseok = next(d for d in schedule.days if d.date == date(2026, 9, 25))
        assert chuseok.status == "holiday"
        assert chuseok.assignments is None
        assert chuseok.holiday_name  # e.g. "Chuseok"


class TestMultiDutyStructure:
    def test_every_duty_fully_staffed_each_duty_day(self):
        schedule = build_schedule(make_roster(30), 2026, 9, DEFAULT_DUTIES)
        duty_days = [d for d in schedule.days if d.status == "duty"]
        assert duty_days
        for day in duty_days:
            assert set(day.assignments) == {d.name for d in DEFAULT_DUTIES}
            for duty in DEFAULT_DUTIES:
                shifts = day.assignments[duty.name]
                assert set(shifts) == set(duty.shifts)
                for soldiers in shifts.values():
                    assert len(soldiers) == duty.soldiers_per_shift

    def test_at_most_one_shift_per_duty_per_day(self):
        schedule = build_schedule(make_roster(30), 2026, 9, DEFAULT_DUTIES)
        for day in schedule.days:
            if day.status != "duty":
                continue
            for duty in DEFAULT_DUTIES:
                ids = [
                    s.id
                    for soldiers in day.assignments[duty.name].values()
                    for s in soldiers
                ]
                assert len(ids) == len(set(ids))

    def test_no_time_overlapping_assignments_same_day(self):
        schedule = build_schedule(make_roster(12), 2026, 9, DEFAULT_DUTIES)
        for day in schedule.days:
            if day.status != "duty":
                continue
            intervals = {}
            for _, shift, soldier in all_assigned(day):
                hours = parse_shift_hours(shift)
                if hours is None:
                    continue
                for start, end in intervals.setdefault(soldier.id, []):
                    assert max(start, hours[0]) >= min(end, hours[1]), (
                        f"{soldier.name} double-booked {day.date}"
                    )
                intervals[soldier.id].append(hours)

    def test_tight_pool_succeeds_via_same_day_double_duty(self):
        # 23 slots/day but only 12 eligible soldiers: only possible because a
        # soldier may hold two non-overlapping duties the same day.
        schedule = build_schedule(make_roster(12), 2026, 9, DEFAULT_DUTIES)
        doubled = False
        for day in schedule.days:
            if day.status != "duty":
                continue
            ids = [s.id for _, _, s in all_assigned(day)]
            if len(ids) > len(set(ids)):
                doubled = True
        assert doubled


class TestFairnessAndConstraints:
    def test_total_assignments_spread_evenly(self):
        schedule = build_schedule(make_roster(20), 2026, 9, [NIGHT_WATCH])
        counts = schedule.total_counts.values()
        assert max(counts) - min(counts) <= 1

    def test_per_duty_counts_reported(self):
        schedule = build_schedule(make_roster(30), 2026, 9, DEFAULT_DUTIES)
        assert set(schedule.per_duty_counts) == {d.name for d in DEFAULT_DUTIES}
        total_from_duties = {
            s.id: sum(
                schedule.per_duty_counts[d.name][s.id] for d in DEFAULT_DUTIES
            )
            for s in schedule.soldiers
        }
        assert total_from_duties == schedule.total_counts

    def test_no_consecutive_duty_days_with_ample_pool(self):
        schedule = build_schedule(make_roster(30), 2026, 9, [NIGHT_WATCH])
        on_duty = {}
        for day in schedule.days:
            if day.status != "duty":
                continue
            for _, _, soldier in all_assigned(day):
                on_duty.setdefault(soldier.id, []).append(day.date)
        for dates in on_duty.values():
            for a, b in zip(dates, dates[1:]):
                assert b - a > timedelta(days=1)

    def test_ineligible_soldiers_never_assigned(self):
        roster = make_roster(16)
        roster.append(
            Soldier("24-9998", "Pvt Kim", Rank.PRIVATE, date(2030, 1, 1), "")
        )
        roster.append(
            Soldier("24-9999", "Sick Lee", Rank.SERGEANT, date(2030, 1, 1), "flu")
        )
        schedule = build_schedule(roster, 2026, 9, DEFAULT_DUTIES)
        assigned_ids = {
            s.id
            for day in schedule.days
            if day.status == "duty"
            for _, _, s in all_assigned(day)
        }
        assert "24-9998" not in assigned_ids
        assert "24-9999" not in assigned_ids

    def test_mid_month_retiree_not_assigned_on_or_after_retirement(self):
        roster = make_roster(20)
        roster[0] = Soldier(
            "24-0000", "Soldier 0", Rank.CORPORAL, date(2026, 9, 15), ""
        )
        schedule = build_schedule(roster, 2026, 9, DEFAULT_DUTIES)
        for day in schedule.days:
            if day.status != "duty" or day.date < date(2026, 9, 15):
                continue
            assert all(s.id != "24-0000" for _, _, s in all_assigned(day))

    def test_deterministic_output(self):
        roster = make_roster(20)
        first = build_schedule(roster, 2026, 9, DEFAULT_DUTIES)
        second = build_schedule(roster, 2026, 9, DEFAULT_DUTIES)
        for day_a, day_b in zip(first.days, second.days):
            if day_a.status == "duty":
                assert [
                    (duty, shift, s.id) for duty, shift, s in all_assigned(day_a)
                ] == [
                    (duty, shift, s.id) for duty, shift, s in all_assigned(day_b)
                ]


class TestInfeasible:
    def test_too_few_eligible_soldiers_raises(self):
        with pytest.raises(SchedulingError, match="eligible"):
            build_schedule(make_roster(5), 2026, 9, [NIGHT_WATCH])

    def test_only_privates_raises(self):
        with pytest.raises(SchedulingError, match="eligible"):
            build_schedule(make_roster(20, rank=Rank.PRIVATE), 2026, 9, [NIGHT_WATCH])

    def test_error_names_the_understaffed_duty(self):
        with pytest.raises(SchedulingError, match="Guard Duty"):
            build_schedule(make_roster(9), 2026, 9, DEFAULT_DUTIES)

    def test_relaxation_warns_at_most_once_per_day(self):
        schedule = build_schedule(make_roster(8), 2026, 9, [NIGHT_WATCH])
        assert schedule.warnings, "expected relaxation warnings"
        assert len(schedule.warnings) == len(set(schedule.warnings))

    def test_small_pool_relaxes_consecutive_rule_instead_of_failing(self):
        schedule = build_schedule(make_roster(8), 2026, 9, [NIGHT_WATCH])
        for day in schedule.days:
            if day.status == "duty":
                for soldiers in day.assignments["Night Watch"].values():
                    assert len(soldiers) == 2
