import json

import pytest

from schedulize.duty_config import (
    DEFAULT_DUTIES,
    DutyConfigError,
    DutyType,
    load_duties,
)


def test_default_duties_match_agreed_config():
    assert [d.name for d in DEFAULT_DUTIES] == [
        "Night Watch",
        "Guard Duty",
        "CCTV",
        "Dishwashing",
    ]
    night_watch = DEFAULT_DUTIES[0]
    assert night_watch.shifts == ("22-24", "00-02", "02-04", "04-06")
    assert night_watch.soldiers_per_shift == 2
    assert DEFAULT_DUTIES[2].soldiers_per_shift == 1


def test_slots_per_day():
    assert DutyType("X", ("a", "b", "c"), 2).slots_per_day == 6


def test_loads_valid_json(tmp_path):
    path = tmp_path / "duties.json"
    path.write_text(
        json.dumps(
            [
                {"name": "Night Watch", "shifts": ["22-24"], "soldiers_per_shift": 2},
                {"name": "CCTV", "shifts": ["00-06"], "soldiers_per_shift": 1},
            ]
        ),
        encoding="utf-8",
    )
    duties = load_duties(path)
    assert [d.name for d in duties] == ["Night Watch", "CCTV"]
    assert duties[0].shifts == ("22-24",)


def test_missing_field_is_error(tmp_path):
    path = tmp_path / "duties.json"
    path.write_text(json.dumps([{"name": "CCTV"}]), encoding="utf-8")
    with pytest.raises(DutyConfigError, match="shifts"):
        load_duties(path)


def test_duplicate_name_is_error(tmp_path):
    path = tmp_path / "duties.json"
    duty = {"name": "CCTV", "shifts": ["00-06"], "soldiers_per_shift": 1}
    path.write_text(json.dumps([duty, duty]), encoding="utf-8")
    with pytest.raises(DutyConfigError, match="duplicate"):
        load_duties(path)


def test_non_positive_staffing_is_error(tmp_path):
    path = tmp_path / "duties.json"
    path.write_text(
        json.dumps([{"name": "CCTV", "shifts": ["00-06"], "soldiers_per_shift": 0}]),
        encoding="utf-8",
    )
    with pytest.raises(DutyConfigError, match="soldiers_per_shift"):
        load_duties(path)


def test_non_list_json_is_error(tmp_path):
    path = tmp_path / "duties.json"
    path.write_text(json.dumps({"name": "CCTV"}), encoding="utf-8")
    with pytest.raises(DutyConfigError, match="list"):
        load_duties(path)
