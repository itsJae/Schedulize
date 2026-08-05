"""Load and validate a soldier roster from a CSV or JSON file."""

import csv
import json
from datetime import date, datetime
from pathlib import Path

from schedulize.models import Rank, Soldier

REQUIRED_FIELDS = ("id", "name", "rank", "retirement_date", "medical_condition")


class LoaderError(Exception):
    pass


def load_roster(path: str | Path) -> list[Soldier]:
    path = Path(path)
    suffix = path.suffix.lower()
    if suffix == ".csv":
        records = _read_csv(path)
    elif suffix == ".json":
        records = _read_json(path)
    else:
        raise LoaderError(
            f"unsupported file type '{suffix}': expected .csv or .json"
        )
    return _build_soldiers(records)


def _read_csv(path: Path) -> list[tuple[int, dict]]:
    with path.open(newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        header = reader.fieldnames or []
        missing = [field for field in REQUIRED_FIELDS if field not in header]
        if missing:
            raise LoaderError(f"missing column(s): {', '.join(missing)}")
        # Row numbers are 1-based file lines; header is line 1.
        return [(line, row) for line, row in enumerate(reader, start=2)]


def _read_json(path: Path) -> list[tuple[int, dict]]:
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError as e:
        raise LoaderError(f"invalid JSON: {e}") from e
    if not isinstance(data, list):
        raise LoaderError("JSON roster must be a list of objects")
    return [(index, record) for index, record in enumerate(data, start=1)]


def _build_soldiers(records: list[tuple[int, dict]]) -> list[Soldier]:
    soldiers: list[Soldier] = []
    seen_ids: set[str] = set()
    errors: list[str] = []

    for row_number, record in records:
        try:
            soldiers.append(_parse_record(record, seen_ids))
        except ValueError as e:
            errors.append(f"row {row_number}: {e}")

    if errors:
        raise LoaderError("roster has errors:\n  " + "\n  ".join(errors))
    if not soldiers:
        raise LoaderError("roster contains no soldiers")
    return soldiers


def _parse_record(record: dict, seen_ids: set[str]) -> Soldier:
    missing = [
        field
        for field in REQUIRED_FIELDS
        if field != "medical_condition" and not str(record.get(field) or "").strip()
    ]
    if "medical_condition" not in record:
        missing.append("medical_condition")
    if missing:
        raise ValueError(f"missing field(s): {', '.join(missing)}")

    soldier_id = str(record["id"]).strip()
    if soldier_id in seen_ids:
        raise ValueError(f"duplicate id '{soldier_id}'")
    seen_ids.add(soldier_id)

    raw_date = str(record["retirement_date"]).strip()
    try:
        retirement = datetime.strptime(raw_date, "%Y-%m-%d").date()
    except ValueError:
        raise ValueError(
            f"invalid retirement_date '{raw_date}' (expected YYYY-MM-DD)"
        ) from None

    return Soldier(
        id=soldier_id,
        name=str(record["name"]).strip(),
        rank=Rank.parse(str(record["rank"])),
        retirement_date=retirement,
        medical_condition=str(record.get("medical_condition") or "").strip(),
    )
