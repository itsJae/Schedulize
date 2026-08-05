"""Soldier data model and duty eligibility rules."""

from dataclasses import dataclass
from datetime import date
from enum import Enum

FIT_MEDICAL_VALUES = {"", "none", "-"}


class Rank(Enum):
    PRIVATE = "PVT"
    PRIVATE_FIRST_CLASS = "PFC"
    CORPORAL = "CPL"
    SERGEANT = "SGT"

    @classmethod
    def parse(cls, text: str) -> "Rank":
        normalized = text.strip().lower()
        aliases = {
            "pvt": cls.PRIVATE,
            "private": cls.PRIVATE,
            "pfc": cls.PRIVATE_FIRST_CLASS,
            "private first class": cls.PRIVATE_FIRST_CLASS,
            "cpl": cls.CORPORAL,
            "corporal": cls.CORPORAL,
            "sgt": cls.SERGEANT,
            "sergeant": cls.SERGEANT,
        }
        if normalized not in aliases:
            raise ValueError(
                f"unknown rank '{text.strip()}' (expected PVT, PFC, CPL, or SGT)"
            )
        return aliases[normalized]


ELIGIBLE_RANKS = {Rank.PRIVATE_FIRST_CLASS, Rank.CORPORAL, Rank.SERGEANT}


@dataclass
class Soldier:
    id: str
    name: str
    rank: Rank
    retirement_date: date
    medical_condition: str

    @property
    def is_fit(self) -> bool:
        return self.medical_condition.strip().lower() in FIT_MEDICAL_VALUES

    def exclusion_reason(self, day: date) -> str | None:
        """Why this soldier cannot stand duty on `day`, or None if eligible."""
        if self.rank not in ELIGIBLE_RANKS:
            return f"rank not eligible ({self.rank.value})"
        if not self.is_fit:
            return f"medical: {self.medical_condition.strip()}"
        if day >= self.retirement_date:
            return f"retired {self.retirement_date.isoformat()}"
        return None

    def is_eligible_on(self, day: date) -> bool:
        return self.exclusion_reason(day) is None
