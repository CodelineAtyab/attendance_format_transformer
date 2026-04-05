# Attendance Format Transformer

Transforms monthly attendance data from `source_format_with_data.xlsm` into
the client-facing report format saved as
`opal_batch1_agile_oracles_attendance_<month>_<year>.xlsx`.

---

## Source File — `source_format_with_data.xlsm`

### Sheet structure
| Sheet | Contents |
|-------|----------|
| `Helper` | Holiday calendar for the year |
| `April`, `March`, `February`, … | One sheet per month of attendance data |
| `SickLeaves` | Sick-leave tracking |

### Month sheet layout (e.g. `March`)
| Row | Contents |
|-----|----------|
| 1 | Title: "Attendance for the month of … `<MonthName>` … `<Year>`" |
| 2 | Start Date / End Date formulas |
| 3 | Blank |
| 4 | Holiday-check formulas (hidden logic) |
| 5 | Column headers: `Can. ID`, `Candidates Name`, then day-name formulas |
| 6 | Date formulas (start date auto-fills across columns) |
| 7 … N | Employee attendance rows (one row per employee) |

### Day columns
- Column **C** = Day 1, **D** = Day 2, … **AG** = Day 31
- `None` in a cell on a **weekend day** means that day is not a working day
- `None` on a **working day** means missing/unrecorded data
- `'-'` means **public holiday** (consistent across ALL employees for that day)
- Weekend and holiday days are identified empirically:
  days where **all employees** have `None` → weekend;
  days where **all employees** have `'-'` → public holiday

### Source attendance codes
| Code | Meaning |
|------|---------|
| `P` | Present |
| `A` | Absent (no reason) |
| `AP` | Absent with Permission |
| `SL` | Sick Leave |
| `L` | Leave (full day) |
| `L--Xh [Ym]` | Late by X hours Y minutes (still counted as present) |
| `'-'` | Public Holiday |
| `None` | Weekend (or unrecorded) |

### Summary columns (cols AH–AM per employee row)
- `AH`–`AK`: ArrayFormula-calculated counts
- `AL`: Presence percentage formula
- `AM`: Absence percentage formula

---

## Destination / Output Format

### File naming
```
opal_batch1_agile_oracles_attendance_<month>_<year>.xlsx
```

### Sheet: "Tamayuz Program Trainees Attendance Report"

#### Column layout (for a 31-day month such as March)
| Col(s) | Letter(s) | Contents |
|--------|-----------|----------|
| 1 | A | Sequential number (N) |
| 2 | B | Trainee Name |
| 3 | C | ID *(fill manually — not in source)* |
| 4 | D | GSM *(fill manually — not in source)* |
| 5 | E | Company *(fill manually — not in source)* |
| 6–36 | F–AK | Daily attendance codes (one column per day, Day 1 … Day 31) |
| 37 | AL | Attendance Days (count of `P` days) |
| 38 | AM | Actual working days in the month |
| 39 | AN | Stipend Amount *(fill manually — not in source)* |
| 40 | AO | Total Amount (formula: `=AN/AM*AL`) |
| 41 | AP | Notes |
| 42 | AQ | Key Code (legend) |
| 43 | AR | Key Description (legend) |

> For a **30-day month** (e.g. April) the date columns end one column earlier
> and all summary columns shift left by one.

#### Row layout
| Row | Contents |
|-----|----------|
| 1 | Title (merged across data columns + key-code panel) |
| 2 | Column headers (dates formatted `D\nMMM`, summary labels) |
| 3 … N | One data row per employee |
| N+1 … | Important Note (merged in key-code panel area) |

#### Source → Destination code mapping
| Source code | Destination code | Notes |
|-------------|-----------------|-------|
| `P` | `P` | Present |
| `A` | `A` | Absent with no reason |
| `AP` | `OA` | Official Absence (absent with permission) |
| `SL` | `SL` | Sick Leave |
| `L` | `AL` | Annual Leave (full-day leave) |
| `L--Xh …` | `P` | Late arrival — still counted as present |
| `'-'` | `PH` | Public Holiday |
| `None` on weekend | `W` | Weekend |
| `None` on working day | *(blank)* | Missing / unrecorded |

#### Destination attendance key legend
| Code | Description |
|------|-------------|
| `P` | Present |
| `W` | Weekend |
| `AL` | Annual Leave |
| `R` | Rotation / days off |
| `PH` | Public Holiday |
| `A` | Absent with no reason |
| `SL` | Sick Leave |
| `OA` | Official Absence |

#### Styling
| Element | Fill | Font |
|---------|------|------|
| Title row | `#A7D9E7` (light blue) | Bold 22pt `#003366` dark navy |
| Header row | `#0066CC` (blue) | Bold 12pt white |
| Weekend date header | `#D9D9D9` (grey) | Bold 10pt white |
| Holiday date header | `#FFE699` (gold) | Bold 10pt white |
| `P` cells | `#C6EFCE` (green) | — |
| `A` cells | `#FFC7CE` (red) | — |
| `OA` cells | `#FFEB9C` (yellow) | — |
| `SL` cells | `#BDD7EE` (blue) | — |
| `AL` cells | `#E2EFDA` (light green) | — |
| `PH` cells | `#FFE699` (gold) | — |
| `W` cells | `#D9D9D9` (grey) | — |

---

## Oman Calendar Notes

- **Work week**: Sunday–Thursday (Friday + Saturday are the normal weekend)
- **Weekend identification**: Always derive from source data — days where
  *all employees* have `None`. This catches exceptions such as a Friday being
  worked (e.g. March 27 2026 was a working day for this cohort).
- **Public holidays**: Days where *all employees* have `'-'`.
  Example — March 2026: Days 19, 22, 23 = Eid Al-Fitr compensatory holidays.

---

## Helper Sheet — Holiday Calendar
Columns: `Months | Official Holidays | Date-from | Date-to | Total Days`

Known 2026 holidays recorded in the file:
| Month | Holiday | Dates |
|-------|---------|-------|
| January | Accession Day | 15 Jan |
| January | Al-Isra and Al-Miraj | 18 Jan |
| June | Hijri New Year | 18 Jun |
| August | Prophet Muhammad's Birthday | 27 Aug |
| November | National Day | 25–26 Nov |

> Note: Eid Al-Fitr (March 2026) was **not** recorded in the Helper sheet but
> is visible as `'-'` markers in the March attendance data.

---

## Python Scripts

### `main.py`
Placeholder entry point created by `uv init`. Currently only prints a hello
message. Intended as a future unified CLI entry point for the transformer.

```
uv run python main.py
```

### `generate_march_attendance.py`
Generates the March 2026 attendance report Excel file.

**Input:** `source_format_with_data.xlsm` → `March` sheet  
**Output:** `opal_batch1_agile_oracles_attendance_march_2026.xlsx`

```
uv run python generate_march_attendance.py
```

**What the script does:**
1. Reads all employee rows (rows 7 onwards) from the `March` sheet
2. Derives weekend days (all-`None` columns) and public holidays (all-`'-'` columns) empirically from the source data — no hardcoded calendar assumption
3. Maps each source attendance code to the destination code (see mapping table above)
4. Writes a fully formatted `.xlsx` with:
   - Colour-coded attendance cells per code
   - Date header row with weekend/holiday highlighting
   - Attendance Days count, Actual Working Days, and a Total Amount formula (`=Stipend/ActualDays*AttDays`)
   - Key-code legend panel on the right
   - Freeze panes at F3, landscape page setup

**Key constants to update when reusing for a different month:**

| Constant | Description |
|----------|-------------|
| `MONTH_NAME` | Month name string (e.g. `"April"`) |
| `YEAR` | Year (e.g. `2026`) |
| `MONTH_NUM` | Month number (e.g. `4`) |
| `TOTAL_DAYS` | Number of days in the month |
| `WEEKEND_DAYS` | Set of day-numbers (1-based) that are weekends — derive from source data |
| `PUBLIC_HOLIDAYS` | Set of day-numbers (1-based) that are public holidays — derive from source data |

**Fields left blank for manual entry after generation:**
- Column C (ID), D (GSM), E (Company) — not present in the source file
- Column AN (Stipend Amount) — the Total Amount formula calculates automatically once filled

---

## Project Setup

Dependencies are managed with [uv](https://github.com/astral-sh/uv) and Python 3.11.

```
uv add openpyxl xlrd   # install dependencies
uv run python <script> # run any script
```

Dependencies declared in `pyproject.toml`:
- `openpyxl >= 3.1.5` — read/write `.xlsx` files
- `xlrd >= 2.0.2` — legacy `.xls` support (available for future use)
