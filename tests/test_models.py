from datetime import date

import pytest

from schedulize.models import Rank, Soldier


def make_soldier(**overrides):
    defaults = dict(
        id="24-70001234",
        name="Kim Minjun",
        rank=Rank.SERGEANT,
        retirement_date=date(2027, 3, 15),
        medical_condition="",
    )
    defaults.update(overrides)
    return Soldier(**defaults)


class TestRankParsing:
    def test_parses_abbreviations_case_insensitively(self):
        assert Rank.parse("pfc") is Rank.PRIVATE_FIRST_CLASS
        assert Rank.parse("SGT") is Rank.SERGEANT
        assert Rank.parse("Cpl") is Rank.CORPORAL
        assert Rank.parse("PVT") is Rank.PRIVATE

    def test_parses_full_names(self):
        assert Rank.parse("Private First Class") is Rank.PRIVATE_FIRST_CLASS
        assert Rank.parse("private") is Rank.PRIVATE
        assert Rank.parse("Corporal") is Rank.CORPORAL
        assert Rank.parse("sergeant") is Rank.SERGEANT

    def test_rejects_unknown_rank(self):
        with pytest.raises(ValueError, match="staff sergeant"):
            Rank.parse("staff sergeant")


class TestEligibility:
    def test_private_is_never_eligible(self):
        soldier = make_soldier(rank=Rank.PRIVATE)
        assert not soldier.is_eligible_on(date(2026, 9, 1))

    def test_pfc_cpl_sgt_are_eligible(self):
        for rank in (Rank.PRIVATE_FIRST_CLASS, Rank.CORPORAL, Rank.SERGEANT):
            assert make_soldier(rank=rank).is_eligible_on(date(2026, 9, 1))

    def test_medical_condition_excludes(self):
        soldier = make_soldier(medical_condition="knee injury")
        assert not soldier.is_eligible_on(date(2026, 9, 1))

    def test_none_and_dash_medical_values_do_not_exclude(self):
        for value in ("", "none", "None", "-", "  "):
            assert make_soldier(medical_condition=value).is_eligible_on(
                date(2026, 9, 1)
            )

    def test_excluded_on_and_after_retirement_date(self):
        soldier = make_soldier(retirement_date=date(2026, 9, 15))
        assert soldier.is_eligible_on(date(2026, 9, 14))
        assert not soldier.is_eligible_on(date(2026, 9, 15))
        assert not soldier.is_eligible_on(date(2026, 9, 16))

    def test_exclusion_reason_reported(self):
        assert make_soldier(rank=Rank.PRIVATE).exclusion_reason(date(2026, 9, 1)) == (
            "rank not eligible (PVT)"
        )
        assert make_soldier(medical_condition="flu").exclusion_reason(
            date(2026, 9, 1)
        ) == "medical: flu"
        assert make_soldier(retirement_date=date(2026, 1, 1)).exclusion_reason(
            date(2026, 9, 1)
        ) == "retired 2026-01-01"
        assert make_soldier().exclusion_reason(date(2026, 9, 1)) is None
