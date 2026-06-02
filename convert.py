#!/usr/bin/env python3
"""
Legacy entry point — delegates to the full pipeline.

Prefer:
    python pipeline.py -i
    python pipeline.py
"""

from __future__ import annotations

import re
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

import config
from common import get_sheet_names, output_path, save_dataframe
from converters import CONVERTERS, LAYOUT_LABELS

LAYOUTS = config.LAYOUTS


def _prompt(message: str) -> str:
    try:
        return input(message).strip()
    except (EOFError, KeyboardInterrupt):
        print("\nCancelled.")
        sys.exit(0)


def _choose_layout() -> str:
    print("\nRate workbook type:")
    for i, name in enumerate(LAYOUTS, 1):
        count = sum(
            len(list(d.glob("*.xls*")))
            for d in config.input_dirs_for(name)
            if d.is_dir()
        )
        label = LAYOUT_LABELS.get(name, name)
        print(f"  {i}. {label} ({count} files)")
    print("  0. Exit")

    while True:
        choice = _prompt("\nSelect number: ")
        if choice == "0":
            sys.exit(0)
        if choice.isdigit() and 1 <= int(choice) <= len(LAYOUTS):
            return LAYOUTS[int(choice) - 1]
        print("Invalid choice. Enter a number from the list.")


def _list_files(layout: str) -> list[Path]:
    files: list[Path] = []
    seen: set[str] = set()
    for folder in config.input_dirs_for(layout):
        for p in folder.iterdir():
            if p.suffix.lower() not in (".xlsx", ".xls") or not p.is_file():
                continue
            key = str(p.resolve())
            if key in seen:
                continue
            seen.add(key)
            files.append(p)
    return sorted(files, key=lambda p: p.name.lower())


def _choose_file(files: list[Path]) -> list[Path]:
    folder_name = files[0].parent.name if files else ""
    print(f"\nFiles in input/{folder_name}:")
    for i, path in enumerate(files, 1):
        print(f"  {i}. {path.name}")
    print(f"  {len(files) + 1}. ALL files (auto price tabs only)")
    print("  0. Back / exit")

    while True:
        choice = _prompt("\nSelect file number: ")
        if choice == "0":
            return []
        if choice.isdigit():
            n = int(choice)
            if n == len(files) + 1:
                return files
            if 1 <= n <= len(files):
                return [files[n - 1]]
        print("Invalid choice.")


def _parse_tab_selection(raw: str, sheet_names: list[str]) -> list[str] | None:
    text = raw.strip().lower()
    if not text or text in ("0", "a", "auto", "all"):
        return None

    if re.fullmatch(r"[\d,\s]+", text):
        indices: list[int] = []
        for part in re.split(r"[,]+", text):
            part = part.strip()
            if not part:
                continue
            if not part.isdigit():
                raise ValueError(f"Invalid tab number: {part}")
            indices.append(int(part))
        resolved: list[str] = []
        for idx in indices:
            if idx < 1 or idx > len(sheet_names):
                raise ValueError(
                    f"Tab number {idx} out of range (1–{len(sheet_names)})"
                )
            name = sheet_names[idx - 1]
            if name not in resolved:
                resolved.append(name)
        return resolved

    parts = [p.strip() for p in raw.split(",") if p.strip()]
    by_lower = {n.strip().lower(): n for n in sheet_names}
    resolved = []
    for part in parts:
        key = part.strip().lower()
        if key in by_lower:
            resolved.append(by_lower[key])
            continue
        matches = [n for n in sheet_names if key in n.strip().lower()]
        if len(matches) == 1:
            resolved.append(matches[0])
        elif len(matches) > 1:
            raise ValueError(
                f"Ambiguous tab '{part}'. Matches: {', '.join(matches)}"
            )
        else:
            raise ValueError(f"Unknown tab: {part}")
    return resolved


def _choose_sheets(path: Path) -> list[str] | None:
    try:
        sheet_names = get_sheet_names(path)
    except Exception as exc:
        print(f"  ERROR reading workbook: {exc}")
        return None

    if not sheet_names:
        print("  No sheets found in workbook.")
        return []

    print(f"\nTabs in {path.name}:")
    for i, name in enumerate(sheet_names, 1):
        print(f"  {i}. {name}")
    print("  0 / auto — auto-detect price tabs only")
    print("\nEnter tab number(s) or name(s), comma-separated (e.g. 3,4 or FTL,LTL):")

    while True:
        raw = _prompt("Tabs to convert: ")
        try:
            return _parse_tab_selection(raw, sheet_names)
        except ValueError as exc:
            print(f"  {exc}")
            retry = _prompt("Try again? [Y/n]: ").lower()
            if retry in ("n", "no"):
                return None


def _convert_one(
    layout: str,
    source: Path,
    sheets: list[str] | None,
    sheet_configs: dict | None = None,
) -> Path | None:
    layout_key = config.normalize_layout(layout)
    converter = CONVERTERS[layout_key]
    tab_desc = "auto price tabs" if sheets is None else ", ".join(sheets)
    print(f"\nConverting: {source.name}")
    print(f"  Tabs: {tab_desc}")

    kwargs: dict = {}
    if layout_key == "usual_rate" and sheet_configs:
        kwargs["sheet_configs"] = sheet_configs

    try:
        df = converter(source, sheets=sheets, **kwargs)
    except ValueError as exc:
        print(f"  ERROR: {exc}")
        return None

    if df.empty:
        print("  WARNING: No price rows extracted.")
        return None

    dest = output_path(config.PROCESSING_DIR, layout_key, source)
    save_dataframe(df, dest)
    try:
        rel = dest.relative_to(config.PROJECT_ROOT)
    except ValueError:
        rel = dest
    print(f"  Saved {len(df):,} rows -> {rel}")
    return dest


def main() -> None:
    from pipeline import run_full_processing

    print("Note: convert.py now runs the full pipeline (rates + accessorial + matrix).")
    print("      Same as:  python pipeline.py -i\n")
    run_full_processing()


if __name__ == "__main__":
    main()
