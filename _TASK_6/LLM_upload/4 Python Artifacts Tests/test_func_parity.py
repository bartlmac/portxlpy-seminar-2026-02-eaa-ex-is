# tests/test_func_parity.py
from __future__ import annotations

import ast
import re
from pathlib import Path
from typing import Iterable, Set

# VBA: Public if not explicitly Private
VBA_DECL_RE = re.compile(
    r"(?im)^\s*(?:(public|private|friend)\s+)?(function|sub)\s+([A-Za-z_][A-Za-z0-9_]*)\b"
)

# Some VBA members we intentionally ignore (e.g., worksheet event stubs, if any appear)
VBA_IGNORE_NAMES: Set[str] = {
    # add known non-base / event handlers here if they appear in Mod_*.txt
}

# Python helpers inside basfunct.py we don't want to treat as "ported VBA names"
PY_IGNORE_NAMES: Set[str] = {
    "DataRepo",
    "_excel_round",
}


def _find_module_txt_files() -> list[Path]:
    # Prefer project root, but support /mnt/data (LLM sandbox)
    roots = [Path.cwd(), Path("/mnt/data")]
    files: list[Path] = []
    for root in roots:
        files.extend(sorted(root.glob("Mod_*.txt")))
    # de-dupe by resolved path string
    uniq: dict[str, Path] = {}
    for p in files:
        try:
            uniq[str(p.resolve())] = p
        except Exception:
            uniq[str(p)] = p
    return list(uniq.values())


def _vba_public_names_from_text(text: str) -> Set[str]:
    names: Set[str] = set()
    for m in VBA_DECL_RE.finditer(text):
        vis = (m.group(1) or "").strip().lower()
        kind = (m.group(2) or "").strip().lower()
        name = (m.group(3) or "").strip()

        # Only Functions/Subs
        if kind not in ("function", "sub"):
            continue

        # Exclude Private
        if vis == "private":
            continue

        if name in VBA_IGNORE_NAMES:
            continue

        names.add(name)
    return names


def _collect_vba_public_names() -> Set[str]:
    mod_files = _find_module_txt_files()
    assert mod_files, "No Mod_*.txt files found (run TASK 2 VBA export)."

    names: Set[str] = set()
    for p in mod_files:
        text = p.read_text(encoding="utf-8", errors="ignore")
        names |= _vba_public_names_from_text(text)
    assert names, "No public VBA Function/Sub names found in Mod_*.txt files."
    return names


def _collect_python_def_names(basfunct_path: Path) -> Set[str]:
    src = basfunct_path.read_text(encoding="utf-8")
    tree = ast.parse(src, filename=str(basfunct_path))
    names: Set[str] = set()
    for node in tree.body:
        if isinstance(node, ast.FunctionDef):
            if node.name in PY_IGNORE_NAMES:
                continue
            if node.name.startswith("_"):
                continue
            names.add(node.name)
    return names


def test_public_vba_names_have_python_defs() -> None:
    vba_names = _collect_vba_public_names()

    # Locate basfunct.py in project root (or /mnt/data fallback)
    basfunct_candidates = [Path.cwd() / "basfunct.py", Path("/mnt/data") / "basfunct.py"]
    basfunct_path = next((p for p in basfunct_candidates if p.exists()), None)
    assert basfunct_path is not None, "basfunct.py not found (create it in TASK 5A)."

    py_names = _collect_python_def_names(basfunct_path)

    missing = sorted(n for n in vba_names if n not in py_names)
    assert not missing, f"Missing Python defs for VBA public names: {missing}"


def test_no_duplicate_python_defs_for_vba_names() -> None:
    """
    Ensures each VBA public name maps to exactly one Python def.
    (In Python, duplicate def names in the same module would overwrite; we detect that by AST only
    yielding final name set. So we do a stricter check by scanning raw text for 'def <name>(' counts.)
    """
    vba_names = _collect_vba_public_names()

    basfunct_candidates = [Path.cwd() / "basfunct.py", Path("/mnt/data") / "basfunct.py"]
    basfunct_path = next((p for p in basfunct_candidates if p.exists()), None)
    assert basfunct_path is not None, "basfunct.py not found (create it in TASK 5A)."

    src = basfunct_path.read_text(encoding="utf-8")

    duplicates = []
    for name in sorted(vba_names):
        # exact 'def Name(' occurrences
        cnt = len(re.findall(rf"(?m)^\s*def\s+{re.escape(name)}\s*\(", src))
        if cnt != 1:
            duplicates.append((name, cnt))

    assert not duplicates, f"Expected exactly one Python def for each VBA name; mismatches: {duplicates}"
