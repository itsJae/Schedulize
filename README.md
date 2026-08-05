# Schedulize

Automated duty scheduler for a Korean army unit. Reads a soldier roster from a
CSV or JSON file, builds a fair monthly schedule for each duty type, and writes
a styled Excel workbook — one sheet per duty, plus fairness and roster summary
sheets.

## Quick start

```bash
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt

.venv/bin/python -m schedulize sample_data/roster.csv --month 2026-09
# → duty_schedule_2026-09.xlsx
```

Options:

| Option | Meaning |
|---|---|
| `--month YYYY-MM` | (required) month to schedule |
| `-o, --output PATH` | output file (default `duty_schedule_<month>.xlsx`) |
| `--duties PATH` | duty types config (default: `./duties.json` if present, else built-in defaults) |

## Input roster

CSV columns or JSON object keys (see `sample_data/roster.csv`):

| Field | Format |
|---|---|
| `id` | unique service number |
| `name` | soldier name |
| `rank` | `PVT`, `PFC`, `CPL`, `SGT` (case-insensitive; full names also accepted) |
| `retirement_date` | `YYYY-MM-DD` |
| `medical_condition` | empty, `none`, or `-` = fit; anything else = excluded that month |

## Duty types (`duties.json`)

Each duty becomes its own Excel sheet with its own shift columns:

```json
[
  { "name": "Night Watch", "shifts": ["22-24", "00-02", "02-04", "04-06"], "soldiers_per_shift": 2 },
  { "name": "Guard Duty",  "shifts": ["06-09", "09-12", "12-15", "15-18", "18-21"], "soldiers_per_shift": 2 },
  { "name": "CCTV",        "shifts": ["21-00", "00-03", "03-06"], "soldiers_per_shift": 1 },
  { "name": "Dishwashing", "shifts": ["After meals"], "soldiers_per_shift": 2 }
]
```

Shift labels shaped like `HH-HH` are treated as time ranges (used for conflict
detection); any other label (e.g. `"After meals"`) is just a slot name.

## Output workbook

- **One sheet per duty** — `Date | Day | <shift columns>`, two assignees shown
  as `SGT Kim Minjun / CPL Han Jiwon`. Weekend rows are shaded gray, holiday
  rows orange with the holiday's name.
- **Duty Count** — per-soldier assignment totals, broken down by duty, sorted
  most-to-least, so fairness is verifiable at a glance.
- **Roster** — every soldier from the input with `Eligible`/`Excluded` status
  and the exclusion reason.

## Scheduling rules

- Eligible ranks: PFC, CPL, SGT. Privates are never assigned.
- No duties on Saturdays, Sundays, or Korean public holidays.
- Fairness: each slot goes to the eligible soldier with the fewest total
  assignments (ties: fewest in that duty, then longest-rested, then roster
  order), so output is deterministic — same input, same schedule.
- A soldier may hold more than one duty on the same day, but never two
  time-overlapping shifts.

## Edge cases covered

**Calendar**
- February length is computed per year — 28 days (e.g. 2026) vs 29 in leap
  years (e.g. 2024); both are under test.
- Korean holidays come from the [`holidays`](https://pypi.org/project/holidays/)
  library, including lunar-calendar holidays (Seollal, Chuseok blocks) and
  government substitute holidays — nothing is hard-coded.
- A holiday that lands on a weekend is shown as a weekend with the holiday's
  name appended (e.g. `Weekend — The second day of Chuseok`).
- Holiday/weekend rows still appear on every sheet (shaded and labeled), so a
  month is always a complete calendar.

**Eligibility**
- Privates are excluded by rank; any non-empty medical condition (other than
  `none`/`-`) excludes a soldier for the entire month — both appear on the
  Roster sheet with the reason, and show zero counts on Duty Count.
- Retirement is checked **per day**: a soldier retiring mid-month serves until
  the day before their retirement date and is never assigned on or after it.
- Soldiers already retired before the month simply never receive assignments.

**Assignment constraints**
- At most one shift per duty per soldier per day.
- No time-overlapping assignments across duties (e.g. Night Watch `22-24` and
  CCTV `21-00` can never go to the same person the same night). Ranges that
  cross midnight (`21-00`, `23-01`) are handled.
- Consecutive duty days are avoided when the pool allows; when it doesn't, the
  rule is relaxed and a warning is printed (once per affected day) instead of
  failing.
- If a duty needs more eligible soldiers than exist on some day, scheduling
  fails with an error naming the day, the duty, and the required headcount.
- A small pool can still fill a large slate of duties through same-day
  double duty — verified by test with 12 soldiers covering 23 daily slots.

**Input validation**
- Missing columns/keys, unknown ranks, malformed dates, duplicate ids, and
  empty rosters are rejected with the offending row number; multiple row
  errors are collected and reported together, not one at a time.
- CSV files with a UTF-8 BOM (as Excel exports them) load correctly.
- Unsupported file extensions and invalid JSON are reported clearly.

**Excel**
- Duty names are sanitized for sheet-title rules (illegal characters removed,
  31-char limit).
- Excluded soldiers still appear on Duty Count and Roster sheets, so the
  workbook accounts for the whole unit, not just those assigned.

## Development

```bash
.venv/bin/python -m pytest tests/ -q        # 65 tests
```

Code layout: `schedulize/` (package: `loader` → `scheduler` → `excel_writer`,
`duty_config`, `models`, `cli`), `tests/` (one file per module),
`docs/superpowers/specs/` (design spec), `legacy/` (superseded original
program, kept for reference). See `CLAUDE.md` for architecture details.
