#!/usr/bin/env python3
"""
VBA_TO_TEXT: Export all non-empty VBA modules from an .xlsm into text files.

Input:
  input/Tariff_Calculator.xlsm

Output (written to project root / current working directory):
  - Mod_<Name>.txt for each non-empty VBA code module

Rules implemented:
  - Uses oletools.olevba to extract modules.
  - Writes one file per module (deduplicates by output filename).
  - Ignores empty modules / code objects with no code text.
  - Includes modules even if they only contain constants (no Sub/Function).
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Dict

from oletools.olevba import VBA_Parser  # type: ignore


INPUT_PATH = Path("input") / "Tariff_Calculator.xlsm"
OUT_PREFIX = "Mod_"

INVALID_FILENAME_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1F]+')
SUBFUNC_RE = re.compile(r"(?im)^\s*(public|private|friend)?\s*(sub|function)\b")


def sanitize_filename(name: str) -> str:
    name = (name or "").strip()
    name = INVALID_FILENAME_CHARS.sub("_", name)
    name = re.sub(r"\s+", "_", name)
    name = name.strip("._")
    return name or "Unnamed"


def normalize_newlines(text: str) -> str:
    return text.replace("\r\n", "\n").replace("\r", "\n")


def pick_module_name(vba_filename: str, stream_path: str) -> str:
    name = (vba_filename or "").strip()
    if name:
        return name
    parts = (stream_path or "").replace("\\", "/").split("/")
    return parts[-1] if parts and parts[-1] else "Unnamed"


def unique_path(base: Path, used: set[str]) -> Path:
    """Ensure filename uniqueness within this run."""
    if base.name not in used:
        used.add(base.name)
        return base
    stem, suffix = base.stem, base.suffix
    i = 2
    while True:
        candidate = base.with_name(f"{stem}_{i}{suffix}")
        if candidate.name not in used:
            used.add(candidate.name)
            return candidate
        i += 1


def main() -> None:
    if not INPUT_PATH.exists():
        raise FileNotFoundError(f"Excel input not found: {INPUT_PATH.resolve()}")

    out_dir = Path(".")
    used_filenames: set[str] = set()
    written: Dict[str, Path] = {}
    exported_count = 0
    skipped_empty = 0
    no_subfunc_count = 0

    vba = VBA_Parser(str(INPUT_PATH))
    try:
        if not vba.detect_vba_macros():
            print("No VBA macros detected.")
            return

        for (_container, stream_path, vba_filename, code) in vba.extract_macros():
            if code is None:
                skipped_empty += 1
                continue

            code_text = normalize_newlines(str(code)).strip()
            if not code_text:
                skipped_empty += 1
                continue

            module_name = pick_module_name(str(vba_filename or ""), str(stream_path or ""))
            safe_name = sanitize_filename(module_name)

            base_path = out_dir / f"{OUT_PREFIX}{safe_name}.txt"
            out_path = unique_path(base_path, used_filenames)

            out_path.write_text(code_text + "\n", encoding="utf-8")
            written[module_name] = out_path
            exported_count += 1

            if not SUBFUNC_RE.search(code_text):
                no_subfunc_count += 1
                print(
                    f"Warning: exported module '{module_name}' but it contains no Sub/Function."
                )

    finally:
        # oletools versions differ; close() exists on newer versions
        close = getattr(vba, "close", None)
        if callable(close):
            close()

    print(f"Exported {exported_count} non-empty VBA modules to Mod_*.txt")
    print(f"Skipped {skipped_empty} empty code objects/modules.")
    if exported_count:
        print(f"Modules without Sub/Function: {no_subfunc_count}")


if __name__ == "__main__":
    main()
