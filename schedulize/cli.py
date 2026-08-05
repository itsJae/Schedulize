"""Command-line entry point: roster file in, Excel schedule out."""

import argparse
import re
import sys
from pathlib import Path

from schedulize.duty_config import DEFAULT_DUTIES, DutyConfigError, load_duties
from schedulize.excel_writer import write_workbook
from schedulize.loader import LoaderError, load_roster
from schedulize.scheduler import SchedulingError, build_schedule


def _parse_month(text: str) -> tuple[int, int]:
    match = re.fullmatch(r"(\d{4})-(0[1-9]|1[0-2])", text)
    if not match:
        raise ValueError(f"invalid month '{text}' (expected YYYY-MM)")
    return int(match.group(1)), int(match.group(2))


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        prog="schedulize",
        description="Generate a monthly night-watch duty schedule as Excel.",
    )
    parser.add_argument("roster", help="roster file (.csv or .json)")
    parser.add_argument(
        "--month", required=True, help="schedule month, e.g. 2026-09"
    )
    parser.add_argument(
        "-o", "--output", help="output .xlsx path (default: duty_schedule_<month>.xlsx)"
    )
    parser.add_argument(
        "--duties",
        help="duty types JSON file (default: ./duties.json if present, "
        "otherwise built-in defaults)",
    )
    args = parser.parse_args(argv)

    try:
        year, month = _parse_month(args.month)
    except ValueError as e:
        print(f"error: {e}", file=sys.stderr)
        return 1

    roster_path = Path(args.roster)
    if not roster_path.exists():
        print(f"error: roster file not found: {roster_path}", file=sys.stderr)
        return 1

    output = Path(args.output or f"duty_schedule_{year}-{month:02d}.xlsx")

    try:
        duties, duties_source = _resolve_duties(args.duties)
        soldiers = load_roster(roster_path)
        schedule = build_schedule(soldiers, year, month, duties)
    except (LoaderError, SchedulingError, DutyConfigError) as e:
        print(f"error: {e}", file=sys.stderr)
        return 1

    for warning in schedule.warnings:
        print(f"warning: {warning}", file=sys.stderr)

    write_workbook(schedule, output)
    duty_days = sum(1 for d in schedule.days if d.status == "duty")
    duty_names = ", ".join(d.name for d in schedule.duties)
    print(
        f"Wrote {output} ({duty_days} duty days, "
        f"{len(soldiers)} soldiers, duties from {duties_source}: {duty_names})"
    )
    return 0


def _resolve_duties(duties_arg: str | None):
    if duties_arg:
        return load_duties(duties_arg), duties_arg
    local_config = Path("duties.json")
    if local_config.exists():
        return load_duties(local_config), str(local_config)
    return DEFAULT_DUTIES, "built-in defaults"


if __name__ == "__main__":
    sys.exit(main())
