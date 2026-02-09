#!/usr/bin/env python3
"""
DATA_EXTRACT: Create normalized input datasets from excel exports.

Inputs (expected in project root):
  - excelcell.csv   columns: Sheet, Address, Formula, Value
  - excelrange.csv  columns: Sheet, Name, Address   (not required for the specific outputs, but validated)

If not found in CWD, the script also checks:
  - /mnt/data/excelcell.csv
  - /mnt/data/excelrange.csv

Outputs (written to project root / current working directory):
  - var.csv     (Name, Value) from Calculation!A4:B9
  - tariff.csv  (Name, Value) from Calculation!D4:E11
  - limits.csv  (Name, Value) from Calculation!G4:H5
  - tables.csv  (Name, Value) from MortalityTables!A:E (headers row 3, data from row 4)
  - tariff.py   implements ModalSurcharge(PayFreq) exactly per Calculation!E12

Success intent:
  - All files exist & have >= 1 data row (tables >= 100 expected for this workbook)
  - import tariff; tariff.ModalSurcharge(12) matches the Excel E12 logic.
"""

from __future__ import annotations

import csv
import json
import re
from pathlib import Path
from typing import Dict, List, Tuple

# ----------------------------
# Paths
# ----------------------------
CANDIDATE_EXCELCELL = [Path("excelcell.csv"), Path("/mnt/data/excelcell.csv")]
CANDIDATE_EXCELRANGE = [Path("excelrange.csv"), Path("/mnt/data/excelrange.csv")]

OUT_VAR = Path("var.csv")
OUT_TARIFF = Path("tariff.csv")
OUT_LIMITS = Path("limits.csv")
OUT_TABLES = Path("tables.csv")
OUT_TARIFF_PY = Path("tariff.py")

A1_RE = re.compile(r"^\$?([A-Z]+)\$?(\d+)$", re.IGNORECASE)


def pick_existing(candidates: List[Path]) -> Path:
    for p in candidates:
        if p.exists():
            return p
    raise FileNotFoundError(f"None of these input files exist: {', '.join(map(str, candidates))}")


# ----------------------------
# A1 helpers
# ----------------------------
def col_letters_to_num(letters: str) -> int:
    letters = letters.strip().upper()
    n = 0
    for ch in letters:
        if not ("A" <= ch <= "Z"):
            raise ValueError(f"Invalid column letters: {letters}")
        n = n * 26 + (ord(ch) - ord("A") + 1)
    return n


def col_num_to_letters(n: int) -> str:
    if n < 1:
        raise ValueError(f"Invalid column number: {n}")
    out = []
    while n:
        n, rem = divmod(n - 1, 26)
        out.append(chr(ord("A") + rem))
    return "".join(reversed(out))


def a1(row: int, col: int) -> str:
    return f"{col_num_to_letters(col)}{row}"


# ----------------------------
# Load excelcell.csv into mapping
# ----------------------------
def load_excelcell(path: Path) -> Dict[Tuple[str, str], Dict[str, str]]:
    """
    Mapping:
      (sheet, address_upper) -> {"value": str, "formula": str}
    """
    out: Dict[Tuple[str, str], Dict[str, str]] = {}
    with path.open("r", encoding="utf-8-sig", newline="") as f:
        r = csv.DictReader(f)
        required = {"Sheet", "Address", "Formula", "Value"}
        if not r.fieldnames or not required.issubset(set(r.fieldnames)):
            raise ValueError(f"{path} must have columns {sorted(required)}; got {r.fieldnames}")

        for row in r:
            sheet = (row.get("Sheet") or "").strip()
            addr = (row.get("Address") or "").strip().upper()
            if not sheet or not addr:
                continue
            out[(sheet, addr)] = {
                "value": row.get("Value") or "",
                "formula": row.get("Formula") or "",
            }
    return out


def validate_excelrange(path: Path) -> None:
    with path.open("r", encoding="utf-8-sig", newline="") as f:
        r = csv.DictReader(f)
        required = {"Sheet", "Name", "Address"}
        if not r.fieldnames or not required.issubset(set(r.fieldnames)):
            raise ValueError(f"{path} must have columns {sorted(required)}; got {r.fieldnames}")


def get_value(cells: Dict[Tuple[str, str], Dict[str, str]], sheet: str, addr: str) -> str:
    rec = cells.get((sheet, addr.upper()))
    return "" if rec is None else (rec.get("value") or "")


def get_formula(cells: Dict[Tuple[str, str], Dict[str, str]], sheet: str, addr: str) -> str:
    rec = cells.get((sheet, addr.upper()))
    return "" if rec is None else (rec.get("formula") or "")


# ----------------------------
# Block extraction helpers
# ----------------------------
def write_name_value(path: Path, rows: List[Tuple[str, str]]) -> None:
    with path.open("w", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        w.writerow(["Name", "Value"])
        for name, value in rows:
            name_s = (name or "").strip()
            value_s = (value or "").strip()
            if name_s == "" and value_s == "":
                continue
            w.writerow([name_s, value_s])


def extract_two_col_rows(
    cells: Dict[Tuple[str, str], Dict[str, str]],
    sheet: str,
    name_col: str,
    value_col: str,
    row_start: int,
    row_end: int,
) -> List[Tuple[str, str]]:
    nc = col_letters_to_num(name_col)
    vc = col_letters_to_num(value_col)
    out: List[Tuple[str, str]] = []
    for r in range(row_start, row_end + 1):
        name = get_value(cells, sheet, a1(r, nc))
        value = get_value(cells, sheet, a1(r, vc))
        if (name or "").strip() == "" and (value or "").strip() == "":
            continue
        out.append((name, value))
    return out


# ----------------------------
# MortalityTables extraction
# ----------------------------
def parse_row_col(addr: str) -> Tuple[int, int]:
    m = A1_RE.match(addr.strip().upper())
    if not m:
        raise ValueError(f"Bad A1: {addr}")
    col = col_letters_to_num(m.group(1))
    row = int(m.group(2))
    return row, col


def extract_mortality_tables(
    cells: Dict[Tuple[str, str], Dict[str, str]],
    sheet: str = "MortalityTables",
    header_row: int = 3,
    data_row_start: int = 4,
    col_start: str = "A",
    col_end: str = "E",
) -> List[Tuple[str, str]]:
    c1 = col_letters_to_num(col_start)
    c2 = col_letters_to_num(col_end)

    # headers from row 3
    headers: List[str] = []
    for c in range(c1, c2 + 1):
        h = (get_value(cells, sheet, a1(header_row, c)) or "").strip()
        headers.append(h if h else f"Col{col_num_to_letters(c)}")

    # find max data row that has any non-empty value within A:E
    max_row = 0
    for (sh, addr), rec in cells.items():
        if sh != sheet:
            continue
        try:
            r, c = parse_row_col(addr)
        except ValueError:
            continue
        if r < data_row_start or not (c1 <= c <= c2):
            continue
        v = (rec.get("value") or "").strip()
        f = (rec.get("formula") or "").strip()
        if v != "" or f != "":
            max_row = max(max_row, r)

    if max_row < data_row_start:
        return []

    out: List[Tuple[str, str]] = []
    for r in range(data_row_start, max_row + 1):
        row_dict: Dict[str, str] = {}
        any_nonempty = False
        first_val = ""
        for i, c in enumerate(range(c1, c2 + 1)):
            v = (get_value(cells, sheet, a1(r, c)) or "").strip()
            row_dict[headers[i]] = v
            if i == 0:
                first_val = v
            if v != "":
                any_nonempty = True

        if not any_nonempty:
            continue

        # Stable row key: include first column value if present, else row number
        key = first_val if first_val else f"ROW{r}"
        name = f"{key}|{r}"
        value = json.dumps(row_dict, ensure_ascii=False)
        out.append((name, value))

    return out


# ----------------------------
# ModalSurcharge extraction & codegen
# ----------------------------
E12_PATTERN = re.compile(
    r"""^=IF\s*\(\s*PayFreq\s*=\s*(\d+)\s*,\s*([\d.]+)%\s*,\s*IF\s*\(\s*PayFreq\s*=\s*(\d+)\s*,\s*([\d.]+)%\s*,\s*IF\s*\(\s*PayFreq\s*=\s*(\d+)\s*,\s*([\d.]+)%\s*,\s*0\s*\)\s*\)\s*\)\s*$""",
    re.IGNORECASE,
)


def render_tariff_py_from_e12(e12_formula: str) -> str:
    """
    Exact implementation for this workbook's E12 formula:
      =IF(PayFreq=2,2%,IF(PayFreq=4,3%,IF(PayFreq=12,5%,0)))
    We generate code directly from the formula, and validate it matches this pattern.
    """
    f = (e12_formula or "").strip()
    if not f:
        raise ValueError("Missing E12 formula text.")

    m = E12_PATTERN.match(f.replace(" ", ""))
    if not m:
        # Fallback: keep formula as comment and implement the known behavior only if it matches expected set.
        # But we fail hard to avoid silent mismatches.
        raise ValueError(f"Unexpected E12 formula format: {f!r}")

    pf1, p1, pf2, p2, pf3, p3 = m.groups()
    mapping = {
        int(pf1): float(p1) / 100.0,
        int(pf2): float(p2) / 100.0,
        int(pf3): float(p3) / 100.0,
    }

    return f'''"""
Auto-generated by data_extract.py from excelcell.csv.

Implements ModalSurcharge(PayFreq) exactly as the Excel formula in Calculation!E12:
  {f}
"""

from __future__ import annotations


EXCEL_E12_FORMULA = {f!r}
_MODAL_MAP = {mapping!r}


def ModalSurcharge(PayFreq: int) -> float:
    """
    Modal surcharge for payment frequency.

    Mirrors the nested IF in Excel cell Calculation!E12.
    """
    try:
        pf = int(PayFreq)
    except Exception as e:
        raise TypeError("PayFreq must be convertible to int") from e
    return float(_MODAL_MAP.get(pf, 0.0))
'''


# ----------------------------
# Main
# ----------------------------
def main() -> None:
    excelcell = pick_existing(CANDIDATE_EXCELCELL)
    excelrange = pick_existing(CANDIDATE_EXCELRANGE)

    validate_excelrange(excelrange)
    cells = load_excelcell(excelcell)

    # var.csv: Calculation A4:B9
    var_rows = extract_two_col_rows(cells, "Calculation", "A", "B", 4, 9)
    write_name_value(OUT_VAR, var_rows)

    # tariff.csv: Calculation D4:E11
    tariff_rows = extract_two_col_rows(cells, "Calculation", "D", "E", 4, 11)
    write_name_value(OUT_TARIFF, tariff_rows)

    # limits.csv: Calculation G4:H5
    limits_rows = extract_two_col_rows(cells, "Calculation", "G", "H", 4, 5)
    write_name_value(OUT_LIMITS, limits_rows)

    # tables.csv: MortalityTables A:E, headers row 3, data from row 4
    tables_rows = extract_mortality_tables(cells)
    write_name_value(OUT_TABLES, tables_rows)

    # tariff.py: ModalSurcharge(PayFreq) from Calculation!E12 formula
    e12_formula = get_formula(cells, "Calculation", "E12")
    if not e12_formula.strip():
        raise ValueError("Could not find Calculation!E12 formula in excelcell.csv (missing or empty Formula column).")
    OUT_TARIFF_PY.write_text(render_tariff_py_from_e12(e12_formula), encoding="utf-8")

    # Minimal runtime checks (non-fatal prints)
    def data_rows(p: Path) -> int:
        with p.open("r", encoding="utf-8-sig", newline="") as f:
            return max(0, sum(1 for _ in f) - 1)

    print(f"Inputs: {excelcell} , {excelrange}")
    print(f"Wrote {OUT_VAR} rows={data_rows(OUT_VAR)}")
    print(f"Wrote {OUT_TARIFF} rows={data_rows(OUT_TARIFF)}")
    print(f"Wrote {OUT_LIMITS} rows={data_rows(OUT_LIMITS)}")
    print(f"Wrote {OUT_TABLES} rows={data_rows(OUT_TABLES)}")
    print(f"Wrote {OUT_TARIFF_PY} (ModalSurcharge from Calculation!E12)")


if __name__ == "__main__":
    main()
