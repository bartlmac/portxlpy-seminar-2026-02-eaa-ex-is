from __future__ import annotations

import importlib.util
import sys
from pathlib import Path

import pytest


def _load_outfunc() -> object:
    project_root = Path(__file__).resolve().parents[1]
    outfunc_path = project_root / "outfunc.py"
    assert outfunc_path.exists(), f"Missing outfunc.py at {outfunc_path}"

    # Ensure sibling modules (e.g., basfunct.py) are importable during module execution
    sys.path.insert(0, str(project_root))

    spec = importlib.util.spec_from_file_location("outfunc_testcopy", str(outfunc_path))
    assert spec is not None and spec.loader is not None
    mod = importlib.util.module_from_spec(spec)

    # Needed so dataclasses can resolve string annotations while the module executes
    sys.modules[spec.name] = mod

    spec.loader.exec_module(mod)  # type: ignore[attr-defined]
    return mod


def test_NormGrossAnnualPrem_reference_case() -> None:
    outfunc = _load_outfunc()
    assert hasattr(outfunc, "NormGrossAnnualPrem")
    NormGrossAnnualPrem = getattr(outfunc, "NormGrossAnnualPrem")

    # Reference input from TASK 6A
    sa = 100_000
    age = 40
    sex = "M"
    n = 30
    t = 20
    PayFreq = 12
    tariff = "KLV"

    expected = 0.04226001
    tol = 1e-8

    got = NormGrossAnnualPrem(sa, age, sex, n, t, PayFreq, tariff)
    assert got == pytest.approx(expected, abs=tol)
