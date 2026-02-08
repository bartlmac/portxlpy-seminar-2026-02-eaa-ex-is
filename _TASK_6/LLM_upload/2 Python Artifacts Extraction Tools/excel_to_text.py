#!/usr/bin/env python3
"""
EXCEL_TO_TEXT: Export Excel workbook content to CSV.

Input:
  input/Tariff_Calculator.xlsm

Outputs (written to project root / current working directory):
  - excelcell.csv  columns: Sheet, Address, Formula, Value
  - excelrange.csv columns: Sheet, Name, Address

Notes:
  - Uses xlwings (requires local Excel installation).
  - Ignores empty cells (both formula and value empty/None).
"""

from __future__ import annotations

import csv
import json
import re
from pathlib import Path
from typing import Any, Iterable, Optional, Tuple

import xlwings as xw


INPUT_PATH = Path("input") / "Tariff_Calculator.xlsm"
OUT_CELLS = Path("excelcell.csv")
OUT_RANGES = Path("excelrange.csv")


def col_to_letters(col_num: int) -> str:
    """1 -> A, 26 -> Z, 27 -> AA"""
    if col_num < 1:
        raise ValueError(f"Invalid column number: {col_num}")
    letters = []
    while col_num:
        col_num, rem = divmod(col_num - 1, 26)
        letters.append(chr(65 + rem))
    return "".join(reversed(letters))


def a1_address(row: int, col: int) -> str:
    return f"{col_to_letters(col)}{row}"


def stringify_value(v: Any) -> str:
    if v is None:
        return ""
    # Excel errors sometimes come through as strings like '#N/A', keep as-is
    if isinstance(v, (str, int, float, bool)):
        return str(v)
    # Datetimes, arrays, etc.
    try:
        return json.dumps(v, ensure_ascii=False, default=str)
    except TypeError:
        return str(v)


def normalize_formula(f: Any) -> str:
    if f is None:
        return ""
    if isinstance(f, str):
        return f
    # xlwings can return nested lists for multi-area; serialize to JSON for safety
    try:
        return json.dumps(f, ensure_ascii=False, default=str)
    except TypeError:
        return str(f)


_REFERS_TO_SHEET_RE = re.compile(
    r"""^=?'?(?P<sheet>[^']+?)'?!""", re.IGNORECASE
)


def parse_sheet_from_refers_to(refers_to: str) -> str:
    """
    Try to extract the sheet name from a Name.RefersTo string like:
      ="'Calculation'!$A$1:$B$2"
      ="=Calculation!$A$1"
    """
    if not refers_to:
        return ""
    s = refers_to.strip()
    if s.startswith("="):
        s = s[1:].lstrip()
    m = _REFERS_TO_SHEET_RE.match(s)
    return m.group("sheet") if m else ""


def is_empty_cell(formula: str, value: Any) -> bool:
    # Treat as empty if no formula and value is None/"".
    if formula and formula != "":
        return False
    if value is None:
        return True
    if isinstance(value, str) and value.strip() == "":
        return True
    return False


def iter_used_range_cells(
    sheet: xw.Sheet,
) -> Iterable[Tuple[str, str, str, Any]]:
    """
    Yield (sheet_name, address, formula, value) for non-empty cells in sheet.used_range.
    Uses array reads for performance and to preserve array formulas via Excel.
    """
    used = sheet.used_range
    # If the sheet is truly empty, used_range may still return A1; we will filter empties.
    top_row = used.row
    left_col = used.column
    nrows = used.rows.count
    ncols = used.columns.count

    # Read formulas & values as 2D arrays in a single call each
    formulas = used.formula  # can be scalar if single cell
    values = used.value

    # Normalize to 2D lists
    if nrows == 1 and ncols == 1:
        formulas_2d = [[formulas]]
        values_2d = [[values]]
    else:
        # xlwings returns list-of-lists for ranges; ensure shape
        formulas_2d = formulas
        values_2d = values

    for r in range(nrows):
        row_idx = top_row + r
        for c in range(ncols):
            col_idx = left_col + c
            f = normalize_formula(formulas_2d[r][c])
            v = values_2d[r][c]
            if is_empty_cell(f, v):
                continue
            addr = a1_address(row_idx, col_idx)
            yield (sheet.name, addr, f, v)


def export_cells(book: xw.Book) -> int:
    count = 0
    with OUT_CELLS.open("w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["Sheet", "Address", "Formula", "Value"])
        for sht in book.sheets:
            for sheet_name, addr, formula, value in iter_used_range_cells(sht):
                w.writerow([sheet_name, addr, formula, stringify_value(value)])
                count += 1
    return count


def export_named_ranges(book: xw.Book) -> int:
    count = 0
    with OUT_RANGES.open("w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["Sheet", "Name", "Address"])
        for nm in book.names:
            try:
                name = nm.name
                refers_to = nm.refers_to  # e.g., ="'Sheet'!$A$1:$B$2"
            except Exception:
                # Skip names we cannot read (rare)
                continue
            sheet = parse_sheet_from_refers_to(refers_to)
            if not name or not refers_to:
                continue
            w.writerow([sheet, name, refers_to])
            count += 1
    return count


def main() -> None:
    if not INPUT_PATH.exists():
        raise FileNotFoundError(f"Excel input not found: {INPUT_PATH.resolve()}")

    app: Optional[xw.App] = None
    book: Optional[xw.Book] = None
    try:
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        app.screen_updating = False

        book = app.books.open(str(INPUT_PATH), update_links=False, read_only=True)

        cell_rows = export_cells(book)
        range_rows = export_named_ranges(book)

        # Basic sanity output for humans
        print(f"Wrote {OUT_CELLS} with {cell_rows} data rows.")
        print(f"Wrote {OUT_RANGES} with {range_rows} data rows.")
        print(f"Workbook: {INPUT_PATH}")

    finally:
        try:
            if book is not None:
                book.close()
        finally:
            if app is not None:
                app.quit()


if __name__ == "__main__":
    main()
