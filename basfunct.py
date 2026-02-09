# basfunct.py
"""
BASISFUNCT – 1-to-1 port of VBA base functions to Python.

Source VBA modules:
- Mod_mPresentValues.txt
- Mod_mCommValues.txt
- Mod_mConstants.txt

Data sources (CSV):
- tables.csv (MortalityTables) is used for Act_qx / commutation functions.

Notes:
- Uses pandas for CSV access.
- Caching mirrors VBA's Scripting.Dictionary usage.
- Excel/VBA rounding differs from Python's bankers rounding; we implement Excel-like ROUND.
- Fix: avoids deprecated pd.to_numeric(errors="ignore") to silence FutureWarning.
"""

from __future__ import annotations

import json
import math
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict, Optional

import pandas as pd

# ----------------------------
# Constants (from mConstants)
# ----------------------------
round_lx: int = 16
round_tx: int = 16
round_Dx: int = 16
round_Cx: int = 16
round_Nx: int = 16
round_Mx: int = 16
round_Rx: int = 16
max_Age: int = 123


# ----------------------------
# Data access
# ----------------------------
@dataclass
class DataRepo:
    root: Path = Path.cwd()
    _cache: Dict[str, pd.DataFrame] = None  # type: ignore[assignment]

    def __post_init__(self) -> None:
        if self._cache is None:
            self._cache = {}

    def _path(self, name: str) -> Path:
        p = self.root / name
        if p.exists():
            return p
        p2 = Path("/mnt/data") / name
        if p2.exists():
            return p2
        raise FileNotFoundError(f"Missing data source: {name}")

    def read_csv(self, name: str) -> pd.DataFrame:
        if name not in self._cache:
            self._cache[name] = pd.read_csv(self._path(name), encoding="utf-8")
        return self._cache[name]

    @staticmethod
    def _safe_numeric_coerce(series: pd.Series) -> pd.Series:
        """
        Convert string-like series to numeric where possible.
        - Empty strings / 'nan' / 'None' become NaN.
        - Values that cannot be converted remain as their original (string) value.
        This avoids deprecated `errors="ignore"` behavior while preserving non-numeric text.
        """
        if series.dtype != object:
            return series

        s_str = series.astype(str)
        empties = s_str.str.strip().isin(["", "nan", "None"])
        s_clean = s_str.where(~empties, None)

        numeric = pd.to_numeric(s_clean, errors="coerce")
        # Keep original where conversion failed but original wasn't empty
        failed = numeric.isna() & ~empties

        # If nothing failed, return numeric
        if not failed.any():
            return numeric

        # Mixed: preserve originals for failed values, numeric for successful values
        out = series.copy()
        out.loc[~failed] = numeric.loc[~failed]
        return out

    def mortality_tables_df(self) -> pd.DataFrame:
        """
        Build a DataFrame from tables.csv (Name, Value where Value is JSON per row).
        Expected JSON keys correspond to MortalityTables headers.
        """
        df = self.read_csv("tables.csv")
        if "Value" not in df.columns:
            raise ValueError("tables.csv must have column 'Value'")

        rows = []
        for v in df["Value"].astype(str):
            try:
                rows.append(json.loads(v))
            except json.JSONDecodeError:
                # allow already-plain rows (shouldn't happen)
                rows.append({"_raw": v})

        tdf = pd.DataFrame(rows)

        # Coerce numeric columns safely (no FutureWarning)
        for c in tdf.columns:
            tdf[c] = self._safe_numeric_coerce(tdf[c])

        return tdf


_DATA = DataRepo()


# ----------------------------
# Excel-like rounding helpers
# ----------------------------
def _excel_round(x: float, digits: int = 0) -> float:
    """
    Excel/VBA WorksheetFunction.Round: halves away from zero.
    Python round() is bankers rounding; do NOT use it here.
    """
    if digits >= 0:
        factor = 10.0**digits
        return math.copysign(math.floor(abs(x) * factor + 0.5) / factor, x)
    factor = 10.0 ** (-digits)
    return math.copysign(math.floor(abs(x) / factor + 0.5) * factor, x)


# ----------------------------
# Cache (from mCommValues)
# ----------------------------
cache: Optional[Dict[str, float]] = None


def InitializeCache() -> None:
    """Create a new Dictionary object (Python dict)."""
    global cache
    cache = {}


# ----------------------------
# Mortality / commutation functions (from mCommValues)
# ----------------------------
def Act_qx(
    Age: int,
    Sex: str,
    TableId: str,
    BirthYear: int = 0,
    RetirementAge: int = 0,
    Layer: int = 1,
) -> float:
    """
    Equivalent to VBA:
      Select Case TableId: "DAV1994_T", "DAV2008_T"
      tableVector = UCase(TableId) & "_" & Sex
      Index(m_Tables, Age+1, Match(tableVector, v_Tables, 0))
    """
    sex = (Sex or "").upper()
    if sex != "M":
        sex = "F"

    table_id = (TableId or "").upper()
    if table_id not in ("DAV1994_T", "DAV2008_T"):
        raise ValueError(f"Unsupported TableId: {TableId}")

    table_vector = f"{table_id}_{sex}"

    tdf = _DATA.mortality_tables_df()

    # Find an "Age" column (common), otherwise assume the first column is age-like.
    age_col = None
    for cand in ("Age", "AGE", "alter", "ALTER"):
        if cand in tdf.columns:
            age_col = cand
            break

    if table_vector not in tdf.columns:
        raise KeyError(f"Column '{table_vector}' not found in tables.csv-derived DataFrame")

    if age_col is not None:
        match = tdf.loc[tdf[age_col] == Age, table_vector]
        if match.empty:
            # fallback: treat row index as age (0-based)
            if 0 <= Age < len(tdf):
                return float(tdf.iloc[Age][table_vector])
            raise IndexError(f"Age {Age} not found in Age column and out of bounds.")
        return float(match.iloc[0])

    # No age column: assume 0-based row corresponds to Age=0
    if 0 <= Age < len(tdf):
        return float(tdf.iloc[Age][table_vector])
    raise IndexError(f"Age {Age} out of bounds for tables data (len={len(tdf)}).")


def Vec_lx(
    EndAge: int,
    Sex: str,
    TableId: str,
    BirthYear: int = 0,
    RetirementAge: int = 0,
    Layer: int = 1,
) -> list[float]:
    """Creates vector of lx; if EndAge = -1 then it is created up to max_Age."""
    limit = max_Age if EndAge == -1 else EndAge
    vec: list[float] = [0.0] * (limit + 1)
    vec[0] = 1_000_000.0
    for i in range(1, limit + 1):
        vec[i] = vec[i - 1] * (1.0 - Act_qx(i - 1, Sex, TableId, BirthYear, RetirementAge, Layer))
        vec[i] = float(_excel_round(vec[i], round_lx))
    return vec


def Act_lx(
    Age: int,
    Sex: str,
    TableId: str,
    BirthYear: int = 0,
    RetirementAge: int = 0,
    Layer: int = 1,
) -> float:
    vec = Vec_lx(Age, Sex, TableId, BirthYear, RetirementAge, Layer)
    return float(vec[Age])


def Vec_tx(
    EndAge: int,
    Sex: str,
    TableId: str,
    BirthYear: int = 0,
    RetirementAge: int = 0,
    Layer: int = 1,
) -> list[float]:
    """Creates vector of tx (# deaths)."""
    limit = max_Age if EndAge == -1 else EndAge
    vec: list[float] = [0.0] * (limit + 1)
    temp_lx = Vec_lx(limit, Sex, TableId, BirthYear, RetirementAge, Layer)
    for i in range(0, limit):
        vec[i] = temp_lx[i] - temp_lx[i + 1]
        vec[i] = float(_excel_round(vec[i], round_tx))
    return vec


def Act_tx(
    Age: int,
    Sex: str,
    TableId: str,
    BirthYear: int = 0,
    RetirementAge: int = 0,
    Layer: int = 1,
) -> float:
    vec = Vec_tx(Age, Sex, TableId, BirthYear, RetirementAge, Layer)
    return float(vec[Age])


def Vec_Dx(
    EndAge: int,
    Sex: str,
    TableId: str,
    InterestRate: float,
    BirthYear: int = 0,
    RetirementAge: int = 0,
    Layer: int = 1,
) -> list[float]:
    """Creates vector of Dx."""
    limit = max_Age if EndAge == -1 else EndAge
    vec: list[float] = [0.0] * (limit + 1)
    v = 1.0 / (1.0 + float(InterestRate))
    temp_lx = Vec_lx(limit, Sex, TableId, BirthYear, RetirementAge, Layer)
    for i in range(0, limit + 1):
        vec[i] = temp_lx[i] * (v**i)
        vec[i] = float(_excel_round(vec[i], round_Dx))
    return vec


def BuildCacheKey(
    Kind: str,
    Age: int,
    Sex: str,
    TableId: str,
    InterestRate: float,
    BirthYear: int,
    RetirementAge: int,
    Layer: int,
) -> str:
    return f"{Kind}_{Age}_{Sex}_{TableId}_{InterestRate}_{BirthYear}_{RetirementAge}_{Layer}"


def Act_Dx(
    Age: int,
    Sex: str,
    TableId: str,
    InterestRate: float,
    BirthYear: int = 0,
    RetirementAge: int = 0,
    Layer: int = 1,
) -> float:
    global cache
    if cache is None:
        InitializeCache()
    assert cache is not None

    key = BuildCacheKey("Dx", Age, Sex, TableId, float(InterestRate), int(BirthYear), int(RetirementAge), int(Layer))
    if key in cache:
        return float(cache[key])

    vec = Vec_Dx(Age, Sex, TableId, float(InterestRate), BirthYear, RetirementAge, Layer)
    val = float(vec[Age])
    cache[key] = val
    return val


def Vec_Cx(
    EndAge: int,
    Sex: str,
    TableId: str,
    InterestRate: float,
    BirthYear: int = 0,
    RetirementAge: int = 0,
    Layer: int = 1,
) -> list[float]:
    """Creates vector of Cx."""
    limit = max_Age if EndAge == -1 else EndAge
    vec: list[float] = [0.0] * (limit + 1)
    v = 1.0 / (1.0 + float(InterestRate))
    temp_tx = Vec_tx(limit, Sex, TableId, BirthYear, RetirementAge, Layer)
    for i in range(0, limit):
        vec[i] = temp_tx[i] * (v ** (i + 1))
        vec[i] = float(_excel_round(vec[i], round_Cx))
    return vec


def Act_Cx(
    Age: int,
    Sex: str,
    TableId: str,
    InterestRate: float,
    BirthYear: int = 0,
    RetirementAge: int = 0,
    Layer: int = 1,
) -> float:
    global cache
    if cache is None:
        InitializeCache()
    assert cache is not None

    key = BuildCacheKey("Cx", Age, Sex, TableId, float(InterestRate), int(BirthYear), int(RetirementAge), int(Layer))
    if key in cache:
        return float(cache[key])

    vec = Vec_Cx(Age, Sex, TableId, float(InterestRate), BirthYear, RetirementAge, Layer)
    val = float(vec[Age])
    cache[key] = val
    return val


def Vec_Nx(
    Sex: str,
    TableId: str,
    InterestRate: float,
    BirthYear: int = 0,
    RetirementAge: int = 0,
    Layer: int = 1,
) -> list[float]:
    """Creates vector of Nx."""
    vec: list[float] = [0.0] * (max_Age + 1)
    temp_dx = Vec_Dx(-1, Sex, TableId, float(InterestRate), BirthYear, RetirementAge, Layer)
    vec[max_Age] = temp_dx[max_Age]
    for i in range(max_Age - 1, -1, -1):
        vec[i] = vec[i + 1] + temp_dx[i]
        vec[i] = float(_excel_round(vec[i], round_Dx))  # kept as in original
    return vec


def Act_Nx(
    Age: int,
    Sex: str,
    TableId: str,
    InterestRate: float,
    BirthYear: int = 0,
    RetirementAge: int = 0,
    Layer: int = 1,
) -> float:
    global cache
    if cache is None:
        InitializeCache()
    assert cache is not None

    key = BuildCacheKey("Nx", Age, Sex, TableId, float(InterestRate), int(BirthYear), int(RetirementAge), int(Layer))
    if key in cache:
        return float(cache[key])

    vec = Vec_Nx(Sex, TableId, float(InterestRate), BirthYear, RetirementAge, Layer)
    val = float(vec[Age])
    cache[key] = val
    return val


def Vec_Mx(
    Sex: str,
    TableId: str,
    InterestRate: float,
    BirthYear: int = 0,
    RetirementAge: int = 0,
    Layer: int = 1,
) -> list[float]:
    """Creates vector of Mx."""
    vec: list[float] = [0.0] * (max_Age + 1)
    temp_cx = Vec_Cx(-1, Sex, TableId, float(InterestRate), BirthYear, RetirementAge, Layer)
    vec[max_Age] = temp_cx[max_Age]
    for i in range(max_Age - 1, -1, -1):
        vec[i] = vec[i + 1] + temp_cx[i]
        vec[i] = float(_excel_round(vec[i], round_Mx))
    return vec


def Act_Mx(
    Age: int,
    Sex: str,
    TableId: str,
    InterestRate: float,
    BirthYear: int = 0,
    RetirementAge: int = 0,
    Layer: int = 1,
) -> float:
    global cache
    if cache is None:
        InitializeCache()
    assert cache is not None

    key = BuildCacheKey("Mx", Age, Sex, TableId, float(InterestRate), int(BirthYear), int(RetirementAge), int(Layer))
    if key in cache:
        return float(cache[key])

    vec = Vec_Mx(Sex, TableId, float(InterestRate), BirthYear, RetirementAge, Layer)
    val = float(vec[Age])
    cache[key] = val
    return val


def Vec_Rx(
    Sex: str,
    TableId: str,
    InterestRate: float,
    BirthYear: int = 0,
    RetirementAge: int = 0,
    Layer: int = 1,
) -> list[float]:
    """Creates vector of Rx."""
    vec: list[float] = [0.0] * (max_Age + 1)
    temp_mx = Vec_Mx(Sex, TableId, float(InterestRate), BirthYear, RetirementAge, Layer)
    vec[max_Age] = temp_mx[max_Age]
    for i in range(max_Age - 1, -1, -1):
        vec[i] = vec[i + 1] + temp_mx[i]
        vec[i] = float(_excel_round(vec[i], round_Rx))
    return vec


def Act_Rx(
    Age: int,
    Sex: str,
    TableId: str,
    InterestRate: float,
    BirthYear: int = 0,
    RetirementAge: int = 0,
    Layer: int = 1,
) -> float:
    global cache
    if cache is None:
        InitializeCache()
    assert cache is not None

    key = BuildCacheKey("Rx", Age, Sex, TableId, float(InterestRate), int(BirthYear), int(RetirementAge), int(Layer))
    if key in cache:
        return float(cache[key])

    vec = Vec_Rx(Sex, TableId, float(InterestRate), BirthYear, RetirementAge, Layer)
    val = float(vec[Age])
    cache[key] = val
    return val


def Act_AgeCalculation(BirthDate: date, ValuationDate: date, Method: str) -> int:
    """Age calculation based on calendar-year method (K) or half-year method (H)."""
    method = Method if Method == "K" else "H"

    # Accept datetime as well
    if isinstance(BirthDate, datetime):
        BirthDate = BirthDate.date()
    if isinstance(ValuationDate, datetime):
        ValuationDate = ValuationDate.date()

    y_birth = BirthDate.year
    y_val = ValuationDate.year
    m_birth = BirthDate.month
    m_val = ValuationDate.month

    if method == "K":
        return int(y_val - y_birth)
    # "H"
    return int(math.floor(y_val - y_birth + (1.0 / 12.0) * (m_val - m_birth + 5)))


# ----------------------------
# Present value functions (from mPresentValues)
# ----------------------------
def Act_DeductionTerm(k: int, InterestRate: float) -> float:
    """Deduction term."""
    acc = 0.0
    if k > 0:
        for l in range(0, k):
            acc += (l / k) / (1.0 + (l / k) * float(InterestRate))
        acc = acc * (1.0 + float(InterestRate)) / k
    return float(acc)


def Act_ax_k(
    Age: int,
    Sex: str,
    TableId: str,
    InterestRate: float,
    k: int,
    BirthYear: int = 0,
    RetirementAge: int = 0,
    Layer: int = 1,
) -> float:
    if k > 0:
        return float(
            Act_Nx(Age, Sex, TableId, float(InterestRate), BirthYear, RetirementAge, Layer)
            / Act_Dx(Age, Sex, TableId, float(InterestRate), BirthYear, RetirementAge, Layer)
            - Act_DeductionTerm(k, float(InterestRate))
        )
    return 0.0


def Act_axn_k(
    Age: int,
    n: int,
    Sex: str,
    TableId: str,
    InterestRate: float,
    k: int,
    BirthYear: int = 0,
    RetirementAge: int = 0,
    Layer: int = 1,
) -> float:
    if k <= 0:
        return 0.0

    nx_age = Act_Nx(Age, Sex, TableId, float(InterestRate), BirthYear, RetirementAge, Layer)
    nx_agen = Act_Nx(Age + n, Sex, TableId, float(InterestRate), BirthYear, RetirementAge, Layer)
    dx_age = Act_Dx(Age, Sex, TableId, float(InterestRate), BirthYear, RetirementAge, Layer)
    dx_agen = Act_Dx(Age + n, Sex, TableId, float(InterestRate), BirthYear, RetirementAge, Layer)

    return float(
        (nx_age - nx_agen) / dx_age
        - Act_DeductionTerm(k, float(InterestRate)) * (1.0 - dx_agen / dx_age)
    )


def Act_nax_k(
    Age: int,
    n: int,
    Sex: str,
    TableId: str,
    InterestRate: float,
    k: int,
    BirthYear: int = 0,
    RetirementAge: int = 0,
    Layer: int = 1,
) -> float:
    if k > 0:
        return float(
            Act_Dx(Age + n, Sex, TableId, float(InterestRate), BirthYear, RetirementAge, Layer)
            / Act_Dx(Age, Sex, TableId, float(InterestRate), BirthYear, RetirementAge, Layer)
            * Act_ax_k(Age + n, Sex, TableId, float(InterestRate), k, BirthYear, RetirementAge, Layer)
        )
    return 0.0


def act_nGrAx(
    Age: int,
    n: int,
    Sex: str,
    TableId: str,
    InterestRate: float,
    BirthYear: int = 0,
    RetirementAge: int = 0,
    Layer: int = 1,
) -> float:
    return float(
        (
            Act_Mx(Age, Sex, TableId, float(InterestRate), BirthYear, RetirementAge, Layer)
            - Act_Mx(Age + n, Sex, TableId, float(InterestRate), BirthYear, RetirementAge, Layer)
        )
        / Act_Dx(Age, Sex, TableId, float(InterestRate), BirthYear, RetirementAge, Layer)
    )


def act_nGrEx(
    Age: int,
    n: int,
    Sex: str,
    TableId: str,
    InterestRate: float,
    BirthYear: int = 0,
    RetirementAge: int = 0,
    Layer: int = 1,
) -> float:
    return float(
        Act_Dx(Age + n, Sex, TableId, float(InterestRate), BirthYear, RetirementAge, Layer)
        / Act_Dx(Age, Sex, TableId, float(InterestRate), BirthYear, RetirementAge, Layer)
    )


def Act_ag_k(g: int, InterestRate: float, k: int) -> float:
    v = 1.0 / (1.0 + float(InterestRate))
    if k > 0:
        if float(InterestRate) > 0:
            return float((1.0 - v**g) / (1.0 - v) - Act_DeductionTerm(k, float(InterestRate)) * (1.0 - v**g))
        return float(g)
    return 0.0
