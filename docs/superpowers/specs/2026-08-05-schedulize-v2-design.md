# Schedulize v2 — Automated Night-Watch Duty Scheduler

**Date:** 2026-08-05
**Status:** Approved by user

## Purpose

Replace the legacy interactive menu program with a single-command tool that reads a
soldier roster from CSV or JSON, generates a fair night-watch duty schedule for a
given month, and writes it to a formatted Excel workbook.

## Requirements

### Input

One roster file, `.csv` or `.json` (list of objects). Required fields per soldier:

| Field | Format | Notes |
|---|---|---|
| `id` | string | Service number; must be unique |
| `name` | string | |
| `rank` | string | `PVT`, `PFC`, `CPL`, `SGT` (case-insensitive); full names also accepted (`Private`, `Private First Class`, `Corporal`, `Sergeant`) |
| `retirement_date` | `YYYY-MM-DD` | Expected discharge date |
| `medical_condition` | string | Empty, `none`, or `-` = fit for duty; any other value = medically excluded for the scheduled month |

Loader validation: missing columns/keys, duplicate ids, unparseable dates, and
unknown ranks are reported as clear errors naming the row and field. A file with
no valid soldiers is an error.

### Scheduling rules

- Schedule covers every calendar day of the requested month (`--month YYYY-MM`).
  Month length (incl. February 28/29) comes from `calendar.monthrange`.
- **Duty days:** weekdays only. Saturdays, Sundays, and Korean public holidays
  (via the `holidays` package, `holidays.KR`, which includes lunar-calendar and
  substitute holidays) have **no duties**. Holiday rows still appear in the
  output, labeled with the holiday name.
- **Shifts:** 4 per duty night — `22-24`, `00-02`, `02-04`, `04-06` — each
  staffed by exactly 2 soldiers (8 slots per night).
- **Eligibility:**
  - Rank must be PFC, CPL, or SGT. PVT is never eligible.
  - Any non-empty medical condition excludes the soldier for the whole month.
  - A soldier is excluded from days **on or after** their retirement date
    (mid-month retirees serve until they leave). Soldiers retired before the
    month starts never appear in assignments.
- **Fairness algorithm:** for each slot, pick the eligible soldier with the
  fewest total assignments this month; tie-break by longest time since last
  duty, then roster order (deterministic output).
  - Hard constraint: a soldier stands at most one shift per night.
  - Soft constraint: avoid consecutive duty nights; relax with a printed
    warning only if the eligible pool is too small for the night.
  - If fewer than 2 eligible soldiers exist for a slot even after relaxing,
    fail with a clear error stating how many eligible soldiers are needed.

### Output (Excel, via openpyxl, English)

Workbook with three sheets:

1. **Schedule** — `Date | Day | 22-24 | 00-02 | 02-04 | 04-06`. Each shift cell
   holds the two assignees as `RANK Name / RANK Name`. Weekend rows shaded gray
   with "Weekend" across the shift columns; holiday rows shaded and labeled
   with the holiday name. Styled header, frozen panes, sensible column widths.
2. **Duty Count** — per-soldier totals (`id, rank, name, shifts assigned`) so
   fairness is verifiable at a glance.
3. **Roster** — every soldier from the input with eligibility status and, if
   excluded, the reason (rank / medical / retired).

### CLI

```
python -m schedulize <roster.csv|roster.json> --month YYYY-MM [-o output.xlsx]
```

Default output name: `duty_schedule_<YYYY-MM>.xlsx`. Missing file, bad month
format, and validation errors exit non-zero with a readable message.

## Architecture

```
schedulize/
  __init__.py
  models.py        # Rank enum, Soldier dataclass, per-day eligibility
  loader.py        # CSV/JSON parsing + validation (LoaderError with row context)
  scheduler.py     # DayPlan/MonthSchedule types + fair assignment algorithm
  excel_writer.py  # openpyxl workbook rendering
  cli.py           # argparse entry point (__main__)
sample_data/roster.csv
tests/             # pytest
```

Each module is independently testable: loader returns `list[Soldier]`,
scheduler consumes soldiers + a month and returns a `MonthSchedule` (pure logic,
no I/O), excel_writer consumes a `MonthSchedule` and writes a file.

`pandas`/`numpy` are not needed — stdlib `csv`/`json` suffice for flat rosters
(YAGNI; can be added later if inputs get messy).

## Error handling

- Loader: per-row errors collected and reported together; abort before scheduling.
- Scheduler: infeasible months (too few eligible soldiers) raise `SchedulingError`.
- CLI: catches domain errors, prints message, exits 1; unexpected exceptions propagate.

## Testing

pytest, TDD:
- models: rank parsing, eligibility (rank / medical / retirement boundary — on
  retirement day = excluded).
- loader: happy path CSV + JSON, missing column, duplicate id, bad date, unknown rank.
- scheduler: February leap/non-leap lengths, weekend/holiday skipping (fixed
  known month), one-shift-per-night, fairness spread (max−min ≤ 1 when pool is
  ample), consecutive-night avoidance, mid-month retirement, infeasible pool error.
- excel_writer: workbook opens, sheet names, header row, holiday row content.

## Out of scope (per user: iterate later)

Multiple duty types, weekend duties, Korean-language output, GUI/menu,
persistence of past months for cross-month fairness.

## Legacy code

The old `Schedulize/myenv` program (menu, sign-up/sign-in, Person/Officer
classes) is superseded and left untouched; new code lives at the repo root.
