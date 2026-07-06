"""Timesheet fill logic.

Runs identically in plain CPython (for tests) and in the browser via Pyodide.
Only cell *values* are written, so openpyxl preserves every other feature of the
workbook untouched -- styling, the K/M formulas, and critically the O2/O1
integrity token the institution stamps on each sheet.
"""

import io
import math
import openpyxl

# Layout constants for the NTUA / EUREKA3D-XR monthly timesheet.
FIRST_ROW = 13          # first day row
LAST_ROW = 43           # last possible day row (31-day month)
COL_DATE = 1            # A
COL_DAY = 2            # B  (Greek day name)
COL_RESEARCH = 3        # C  (hours land here)
COL_NOTE = 10           # J  (WP label / pre-marked holiday)
COL_MONTH_CELL = "B10"
COL_YEAR_CELL = "B9"
DAILY_CAP = 8

WEEKEND = {"Σαββάτο", "Κυριακή"}


def _workday_rows(ws):
    """Rows that are real working days: a weekday with an empty note column.

    A non-empty note in column J means the sheet already flags that day as a
    public holiday (e.g. 'Πρωτομαγιά', 'Αγίου Πνεύματος'), so it is skipped.
    """
    rows = []
    for r in range(FIRST_ROW, LAST_ROW + 1):
        if ws.cell(r, COL_DATE).value is None:
            continue
        if ws.cell(r, COL_DAY).value in WEEKEND:
            continue
        if ws.cell(r, COL_NOTE).value not in (None, ""):
            continue
        rows.append(r)
    return rows


def _even_spread(hours, ndays):
    """Split `hours` across `ndays` as evenly as possible, each day <= DAILY_CAP."""
    base = hours // ndays
    rem = hours - base * ndays
    return [base + (1 if i < rem else 0) for i in range(ndays)]


def analyze(data: bytes) -> dict:
    """Read a workbook and report month/year and how many working days it has."""
    wb = openpyxl.load_workbook(io.BytesIO(data))
    ws = wb.active
    rows = _workday_rows(ws)
    holidays = []
    for r in range(FIRST_ROW, LAST_ROW + 1):
        note = ws.cell(r, COL_NOTE).value
        if note not in (None, "") and ws.cell(r, COL_DAY).value not in WEEKEND:
            holidays.append({"day": ws.cell(r, COL_DATE).value, "note": note})
    return {
        "month": ws[COL_MONTH_CELL].value,
        "year": ws[COL_YEAR_CELL].value,
        "workdays": len(rows),
        "capacity": len(rows) * DAILY_CAP,
        "holidays": holidays,
    }


def fill(data: bytes, allocations: list) -> bytes:
    """Fill a workbook.

    `allocations` is a list of {"wp": str, "hours": number}. Each WP is placed
    on its own consecutive block of working days (label in column J, hours in
    column C), even-spread and capped at DAILY_CAP per day.
    Returns the new .xlsx as bytes. Raises ValueError if it cannot fit.
    """
    allocations = [a for a in allocations if a.get("hours")]
    wb = openpyxl.load_workbook(io.BytesIO(data))
    ws = wb.active
    rows = _workday_rows(ws)

    plan = []
    for a in allocations:
        hours = int(a["hours"])
        if hours <= 0:
            continue
        ndays = math.ceil(hours / DAILY_CAP)
        plan.append((str(a["wp"]).strip(), hours, ndays))

    needed = sum(d for _, _, d in plan)
    if needed > len(rows):
        total = sum(h for _, h, _ in plan)
        raise ValueError(
            f"Needs {needed} working days ({total}h) but the month only has "
            f"{len(rows)} available ({len(rows) * DAILY_CAP}h capacity)."
        )

    idx = 0
    for wp, hours, ndays in plan:
        block = rows[idx:idx + ndays]
        idx += ndays
        for row, hrs in zip(block, _even_spread(hours, ndays)):
            ws.cell(row, COL_RESEARCH).value = hrs
            ws.cell(row, COL_NOTE).value = wp

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()
