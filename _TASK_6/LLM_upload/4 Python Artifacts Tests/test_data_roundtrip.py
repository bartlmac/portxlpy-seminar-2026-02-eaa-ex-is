# tests/test_data_roundtrip.py
from __future__ import annotations

import csv
import importlib.util
from pathlib import Path


def _read_header(path: Path) -> list[str]:
    with path.open("r", encoding="utf-8-sig", newline="") as f:
        r = csv.reader(f)
        try:
            return next(r)
        except StopIteration:
            return []


def _count_rows(path: Path) -> int:
    with path.open("r", encoding="utf-8-sig", newline="") as f:
        r = csv.reader(f)
        try:
            next(r)  # header
        except StopIteration:
            return 0
        return sum(1 for _ in r)


def test_csv_roundtrip_smoke(data_dir: Path) -> None:
    for filename in ("var.csv", "tariff.csv", "limits.csv", "tables.csv"):
        p = data_dir / filename
        assert p.exists(), f"Missing file: {p}"

        header = _read_header(p)
        assert header, f"{filename}: missing header row"
        assert "Name" in header and "Value" in header, f"{filename}: expected Name/Value columns, got {header}"
        assert len(header) >= 2, f"{filename}: expected >=2 columns, got {len(header)}"

        assert _count_rows(p) >= 1, f"{filename}: expected at least 1 data row"

    # tables.csv should be non-trivial even in mini sample
    assert _count_rows(data_dir / "tables.csv") >= 10


def test_tariff_module_import_and_modal_surcharge(data_dir: Path) -> None:
    tariff_path = data_dir / "tariff.py"
    assert tariff_path.exists()

    spec = importlib.util.spec_from_file_location("tariff_testcopy", str(tariff_path))
    assert spec is not None and spec.loader is not None
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)  # type: ignore[attr-defined]

    assert hasattr(mod, "ModalSurcharge")
    v = mod.ModalSurcharge(12)
    assert isinstance(v, float)
