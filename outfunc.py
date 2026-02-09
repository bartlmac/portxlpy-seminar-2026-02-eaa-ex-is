"""outfunc.py

Output functions (tariff calculator), Task 6A.

This module implements premium calculation functions derived from the Excel workbook.

TASK 6A deliverable:
- NormGrossAnnualPrem(sa, age, sex, n, t, PayFreq, tariff)

The implementation mirrors the Excel formula in Calculation!K5 exactly.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict

from basfunct import DataRepo, Act_Dx, Act_axn_k, act_nGrAx


@dataclass(frozen=True)
class TariffParams:
    """Tariff parameters needed for the premium formula."""

    InterestRate: float
    MortalityTable: str
    alpha: float
    beta1: float
    gamma1: float
    gamma2: float


def _as_float(v: Any) -> float:
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip()
    # handle values that may be quoted like """100000"""
    if len(s) >= 2 and ((s[0] == s[-1] == '"') or (s[0] == s[-1] == "'")):
        s = s[1:-1]
    return float(s)


def _read_name_value_csv(repo: DataRepo, filename: str) -> Dict[str, Any]:
    df = repo.read_csv(filename)
    if "Name" not in df.columns or "Value" not in df.columns:
        raise ValueError(f"{filename} must have columns Name and Value")
    return {str(k): v for k, v in zip(df["Name"], df["Value"], strict=False)}


def _load_tariff_params(repo: DataRepo) -> TariffParams:
    d = _read_name_value_csv(repo, "tariff.csv")
    return TariffParams(
        InterestRate=_as_float(d["InterestRate"]),
        MortalityTable=str(d["MortalityTable"]),
        alpha=_as_float(d["alpha"]),
        beta1=_as_float(d["beta1"]),
        gamma1=_as_float(d["gamma1"]),
        gamma2=_as_float(d["gamma2"]),
    )


def NormGrossAnnualPrem(
    sa: float,
    age: int,
    sex: str,
    n: int,
    t: int,
    PayFreq: int,
    tariff: str,
) -> float:
    """Normalized gross annual premium rate.

    Mirrors Excel Calculation!K5 (named range "NormGrossAnnualPrem"):

    =(act_nGrAx(x,n,Sex,MortalityTable,InterestRate)
      +Act_Dx(x+n,Sex,MortalityTable,InterestRate)/Act_Dx(x,Sex,MortalityTable,InterestRate)
      +gamma1*Act_axn_k(x,t,Sex,MortalityTable,InterestRate,1)
      +gamma2*(Act_axn_k(x,n,Sex,MortalityTable,InterestRate,1)-Act_axn_k(x,t,Sex,MortalityTable,InterestRate,1)))
     /((1-beta1)*Act_axn_k(x,t,Sex,MortalityTable,InterestRate,1)-alpha*t)

    Parameters sa, PayFreq, tariff are part of the public calculator signature but are not
    referenced by the Excel K5 formula.
    """

    repo = DataRepo(Path(__file__).resolve().parent)
    p = _load_tariff_params(repo)

    x = int(age)
    Sex = str(sex)
    nn = int(n)
    tt = int(t)

    # Common actuarial terms as in Excel
    ax_t = Act_axn_k(x, tt, Sex, p.MortalityTable, p.InterestRate, 1)
    ax_n = Act_axn_k(x, nn, Sex, p.MortalityTable, p.InterestRate, 1)

    numerator = (
        act_nGrAx(x, nn, Sex, p.MortalityTable, p.InterestRate)
        + Act_Dx(x + nn, Sex, p.MortalityTable, p.InterestRate)
        / Act_Dx(x, Sex, p.MortalityTable, p.InterestRate)
        + p.gamma1 * ax_t
        + p.gamma2 * (ax_n - ax_t)
    )

    denominator = (1.0 - p.beta1) * ax_t - p.alpha * tt

    return float(numerator / denominator)
