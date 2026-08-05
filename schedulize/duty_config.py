"""Duty type definitions, either built-in defaults or from a JSON file."""

import json
from dataclasses import dataclass
from pathlib import Path


class DutyConfigError(Exception):
    pass


@dataclass(frozen=True)
class DutyType:
    name: str
    shifts: tuple[str, ...]
    soldiers_per_shift: int

    @property
    def slots_per_day(self) -> int:
        return len(self.shifts) * self.soldiers_per_shift


DEFAULT_DUTIES = (
    DutyType("Night Watch", ("22-24", "00-02", "02-04", "04-06"), 2),
    DutyType("Guard Duty", ("06-09", "09-12", "12-15", "15-18", "18-21"), 2),
    DutyType("CCTV", ("21-00", "00-03", "03-06"), 1),
    DutyType("Dishwashing", ("After meals",), 2),
)


def load_duties(path: str | Path) -> tuple[DutyType, ...]:
    path = Path(path)
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError as e:
        raise DutyConfigError(f"{path}: invalid JSON: {e}") from e

    if not isinstance(data, list) or not data:
        raise DutyConfigError(f"{path}: must be a non-empty list of duty objects")

    duties: list[DutyType] = []
    seen_names: set[str] = set()
    for index, entry in enumerate(data, start=1):
        duties.append(_parse_duty(entry, index, seen_names))
    return tuple(duties)


def _parse_duty(entry, index: int, seen_names: set[str]) -> DutyType:
    if not isinstance(entry, dict):
        raise DutyConfigError(f"duty #{index}: must be an object")

    missing = [
        key
        for key in ("name", "shifts", "soldiers_per_shift")
        if key not in entry
    ]
    if missing:
        raise DutyConfigError(f"duty #{index}: missing {', '.join(missing)}")

    name = str(entry["name"]).strip()
    if not name:
        raise DutyConfigError(f"duty #{index}: name is empty")
    if name in seen_names:
        raise DutyConfigError(f"duty #{index}: duplicate duty name '{name}'")
    seen_names.add(name)

    shifts = entry["shifts"]
    if not isinstance(shifts, list) or not shifts:
        raise DutyConfigError(f"duty '{name}': shifts must be a non-empty list")

    staffing = entry["soldiers_per_shift"]
    if not isinstance(staffing, int) or staffing < 1:
        raise DutyConfigError(
            f"duty '{name}': soldiers_per_shift must be a positive integer"
        )

    return DutyType(name, tuple(str(s) for s in shifts), staffing)
