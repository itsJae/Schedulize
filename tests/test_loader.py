import json
from datetime import date

import pytest

from schedulize.loader import LoaderError, load_roster
from schedulize.models import Rank

CSV_HEADER = "id,name,rank,retirement_date,medical_condition\n"

VALID_ROWS = (
    "24-70001234,Kim Minjun,SGT,2027-03-15,\n"
    "24-70005678,Lee Junho,CPL,2027-06-20,none\n"
    "25-70009012,Park Jihoon,PFC,2027-11-01,knee injury\n"
    "25-70003456,Choi Hyunwoo,PVT,2028-02-10,-\n"
)


def write_csv(tmp_path, body=VALID_ROWS, header=CSV_HEADER):
    path = tmp_path / "roster.csv"
    path.write_text(header + body, encoding="utf-8")
    return path


class TestCsvLoading:
    def test_loads_valid_csv(self, tmp_path):
        soldiers = load_roster(write_csv(tmp_path))
        assert len(soldiers) == 4
        first = soldiers[0]
        assert first.id == "24-70001234"
        assert first.name == "Kim Minjun"
        assert first.rank is Rank.SERGEANT
        assert first.retirement_date == date(2027, 3, 15)
        assert first.medical_condition == ""

    def test_missing_column_is_error(self, tmp_path):
        path = tmp_path / "roster.csv"
        path.write_text(
            "id,name,rank,retirement_date\n24-1,Kim,SGT,2027-01-01\n",
            encoding="utf-8",
        )
        with pytest.raises(LoaderError, match="medical_condition"):
            load_roster(path)

    def test_duplicate_id_is_error(self, tmp_path):
        body = (
            "24-1,Kim Minjun,SGT,2027-03-15,\n"
            "24-1,Lee Junho,CPL,2027-06-20,\n"
        )
        with pytest.raises(LoaderError, match="duplicate id '24-1'"):
            load_roster(write_csv(tmp_path, body))

    def test_bad_date_reports_row(self, tmp_path):
        body = "24-1,Kim Minjun,SGT,15-03-2027,\n"
        with pytest.raises(LoaderError, match="row 2"):
            load_roster(write_csv(tmp_path, body))

    def test_unknown_rank_reports_row(self, tmp_path):
        body = "24-1,Kim Minjun,GENERAL,2027-03-15,\n"
        with pytest.raises(LoaderError, match="GENERAL"):
            load_roster(write_csv(tmp_path, body))

    def test_empty_roster_is_error(self, tmp_path):
        with pytest.raises(LoaderError, match="no soldiers"):
            load_roster(write_csv(tmp_path, body=""))

    def test_csv_with_utf8_bom_loads(self, tmp_path):
        path = tmp_path / "roster.csv"
        path.write_text(
            CSV_HEADER + "24-1,Kim Minjun,SGT,2027-03-15,\n",
            encoding="utf-8-sig",
        )
        soldiers = load_roster(path)
        assert soldiers[0].id == "24-1"

    def test_multiple_errors_reported_together(self, tmp_path):
        body = (
            "24-1,Kim Minjun,GENERAL,2027-03-15,\n"
            "24-2,Lee Junho,CPL,not-a-date,\n"
        )
        with pytest.raises(LoaderError) as excinfo:
            load_roster(write_csv(tmp_path, body))
        message = str(excinfo.value)
        assert "row 2" in message and "row 3" in message


class TestJsonLoading:
    def test_loads_valid_json(self, tmp_path):
        path = tmp_path / "roster.json"
        path.write_text(
            json.dumps(
                [
                    {
                        "id": "24-1",
                        "name": "Kim Minjun",
                        "rank": "Sergeant",
                        "retirement_date": "2027-03-15",
                        "medical_condition": "none",
                    }
                ]
            ),
            encoding="utf-8",
        )
        soldiers = load_roster(path)
        assert len(soldiers) == 1
        assert soldiers[0].rank is Rank.SERGEANT

    def test_missing_key_is_error(self, tmp_path):
        path = tmp_path / "roster.json"
        path.write_text(
            json.dumps([{"id": "24-1", "name": "Kim"}]), encoding="utf-8"
        )
        with pytest.raises(LoaderError, match="rank"):
            load_roster(path)


def test_unsupported_extension_is_error(tmp_path):
    path = tmp_path / "roster.txt"
    path.write_text("hello", encoding="utf-8")
    with pytest.raises(LoaderError, match=".txt"):
        load_roster(path)
