# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

Schedulize generates a monthly duty schedule for a Korean army unit from a soldier
roster (CSV/JSON) and writes it to a styled Excel workbook — one sheet per duty
type plus "Duty Count" and "Roster" summary sheets.

## Commands

Dependencies live in a local venv at `.venv` (Python 3.14, see `requirements.txt`).

```bash
.venv/bin/python -m pytest tests/ -q                 # full test suite
.venv/bin/python -m pytest tests/test_scheduler.py -q                  # one module
.venv/bin/python -m pytest tests/test_scheduler.py::TestInfeasible -q  # one class/test

# Run the tool (writes duty_schedule_<month>.xlsx by default)
.venv/bin/python -m schedulize sample_data/roster.csv --month 2026-09
```

The root `conftest.py` exists solely to put the project root on `sys.path` for
pytest; there is no packaging/install step.

## Architecture

A pipeline of four independent modules, orchestrated by `cli.py`:

```
loader.py ──► list[Soldier] ──► scheduler.py ──► MonthSchedule ──► excel_writer.py ──► .xlsx
                 duty_config.py ──► tuple[DutyType] ─┘
```

- **`models.py`** — `Soldier`, `Rank`. All eligibility policy lives here:
  PVT is never eligible; any non-empty `medical_condition` (other than
  `none`/`-`) excludes for the whole month; soldiers are excluded on/after
  their `retirement_date` (checked per-day, so mid-month retirees serve until
  they leave).
- **`duty_config.py`** — `DutyType` (name, shift labels, soldiers_per_shift)
  and `DEFAULT_DUTIES`. The CLI resolves duties in this order: `--duties path`
  → `./duties.json` if present → `DEFAULT_DUTIES`. The root `duties.json` is
  the user-editable config and matches the defaults.
- **`scheduler.py`** — pure logic, no I/O. Walks every calendar day of the
  month; Saturdays/Sundays and Korean public holidays (via the `holidays`
  package, `holidays.KR`, English names) get status `weekend`/`holiday` and no
  assignments (weekend takes precedence when both). For duty days it fills
  every duty/shift greedily by priority tuple: fewest total assignments, then
  fewest in that duty, then longest since last duty, then roster order — which
  makes output deterministic. Hard constraints: max one shift per duty per
  day, and no time-overlapping assignments across duties the same day
  (`parse_shift_hours` parses labels like `"21-00"`; non-time labels like
  `"After meals"` never conflict). Soft constraint: avoid consecutive duty
  days, relaxed with a warning (collected on `MonthSchedule.warnings`) when
  the pool is too small. Same-day multi-duty assignment is intentionally
  allowed.
- **`excel_writer.py`** — renders `MonthSchedule` only; one sheet per duty
  with that duty's shift columns, weekend/holiday rows merged+shaded.

Design decisions (duty shapes, one-vs-many duties per day, etc.) were made with
the user and recorded in `docs/superpowers/specs/2026-08-05-schedulize-v2-design.md`.

## Conventions and gotchas

- TDD is the established workflow: every behavior change starts with a failing
  test in `tests/` (one test file per module).
- `legacy/` is the superseded 2-year-old interactive program — never edit or
  import it. Note the filesystem is case-insensitive: `legacy/` was renamed
  from `Schedulize/`, which used to collide with the `schedulize/` package.
- Domain errors are typed per module (`LoaderError`, `DutyConfigError`,
  `SchedulingError`); `cli.py` catches exactly these, prints to stderr, and
  returns exit code 1. Loader/config errors report row numbers and collect
  multiple errors before aborting.
- Feasibility rule: every duty needs `slots_per_day` eligible soldiers on every
  duty day (plus non-overlap headroom); `sample_data/roster.csv` has 32
  soldiers because the default four duties consume 23 slots/day.
- `pandas`/`numpy` were deliberately not used — stdlib `csv`/`json` handle the
  flat roster.
