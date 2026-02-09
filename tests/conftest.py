# tests/conftest.py
from __future__ import annotations

import csv
import shutil
from pathlib import Path

import pytest

PROJECT_ROOT = Path(__file__).resolve().parents[1]
FALLBACK_DIR = Path("/mnt/data")


def _source_file(name: str) -> Path:
    """Prefer project root, fall back to /mnt/data (used by some LLM environments)."""
    p = PROJECT_ROOT / name
    if p.exists():
        return p
    p2 = FALLBACK_DIR / name
    if p2.exists():
        return p2
    raise FileNotFoundError(f"Required input file not found in project root or /mnt/data: {name}")


def _mini_csv(src: Path, dst: Path, max_rows: int) -> None:
    """Copy CSV with header + up to max_rows data rows."""
    with src.open("r", encoding="utf-8-sig", newline="") as fin:
        reader = csv.reader(fin)
        try:
            header = next(reader)
        except StopIteration as e:
            raise ValueError(f"CSV has no header: {src}") from e

        rows = [header]
        for i, row in enumerate(reader):
            if i >= max_rows:
                break
            rows.append(row)

    with dst.open("w", encoding="utf-8", newline="") as fout:
        writer = csv.writer(fout)
        writer.writerows(rows)


@pytest.fixture()
def data_dir(tmp_path: Path) -> Path:
    """
    Temp dir containing small, deterministic samples of the product data.
    """
    mini_specs = {
        "var.csv": 50,
        "tariff.csv": 50,
        "limits.csv": 50,
        "tables.csv": 200,
    }

    for filename, max_rows in mini_specs.items():
        src = _source_file(filename)
        _mini_csv(src, tmp_path / filename, max_rows=max_rows)

    shutil.copyfile(_source_file("tariff.py"), tmp_path / "tariff.py")
    return tmp_path
