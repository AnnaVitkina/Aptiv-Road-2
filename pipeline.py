#!/usr/bin/env python3
"""
Full rate processing pipeline: convert → accessorial (optional) → matrix export.

Entry points:
  run_full_processing()   — interactive: pick type, file, tabs, accessorial, matrix
  run_pipeline()          — batch: all files, no prompts
  python pipeline.py      — interactive (local & Colab default)
  python pipeline.py --batch   — process every file in input/ without prompts

Colab:
    import sys
    sys.path.insert(0, "/content/Aptiv_Road")
    from pipeline import setup_environment, run_full_processing
    setup_environment()
    run_full_processing()
"""

from __future__ import annotations

import argparse
import os
import re
import subprocess
import sys
from datetime import datetime
from dataclasses import dataclass, field
from pathlib import Path

import pandas as pd


def _bootstrap_project_path() -> Path | None:
    """Find project root (folder with config.py) and add it to sys.path."""
    candidates: list[Path] = []
    env_root = os.environ.get("APTIV_ROAD_ROOT")
    if env_root:
        candidates.append(Path(env_root))
    try:
        candidates.append(Path(__file__).resolve().parent)
    except NameError:
        pass
    candidates.extend(
        [
            Path("/content/Aptiv-Road-2"),
            Path("/content/drive/MyDrive/Aptiv_Road"),
            Path.cwd(),
        ]
    )
    seen: set[str] = set()
    for root in candidates:
        key = str(root.resolve()) if root.exists() else str(root)
        if key in seen:
            continue
        seen.add(key)
        if (root / "config.py").is_file():
            root_s = str(root.resolve())
            if root_s not in sys.path:
                sys.path.insert(0, root_s)
            return root.resolve()
    return None


_bootstrap_project_path()

import config

if str(config.PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(config.PROJECT_ROOT))

from common import (
    append_excel_sheet,
    copy_excel_sheet,
    get_sheet_names,
    is_blank,
    normalize_header,
    output_path,
    read_sheet_rows,
    save_dataframe,
)
from converters import CONVERTERS, LAYOUT_LABELS
from converters.accessorial import ACCESSORIAL_COLUMNS, parse_accessorial_file
from export_matrix import export_converted_file_to_matrix

ACCESSORIAL_OUTPUT_SHEET = "Accessorial Costs"


def _prompt(message: str) -> str:
    try:
        return input(message).strip()
    except (EOFError, KeyboardInterrupt):
        print("\nCancelled.")
        raise SystemExit(0) from None


@dataclass
class PipelineResult:
    converted: list[Path] = field(default_factory=list)
    matrices: list[Path] = field(default_factory=list)
    accessorial_rows: int = 0
    convert_skipped: list[str] = field(default_factory=list)
    convert_errors: list[str] = field(default_factory=list)
    export_errors: list[str] = field(default_factory=list)


@dataclass
class WorkbookResult:
    """Result of full processing for one input workbook."""

    source: Path
    converted: Path | None = None
    matrix: Path | None = None
    accessorial_rows: int = 0
    skipped: bool = False


def setup_environment(*, mount_drive: bool | None = None) -> None:
    """
    Prepare Colab or local run: mount Drive (Colab), ensure folders exist.

    Call once at the top of a Colab notebook before run_pipeline().
    """
    if mount_drive is None:
        mount_drive = config.IS_COLAB
    if mount_drive:
        from google.colab import drive

        drive.mount("/content/drive", force_remount=False)
    if not config.PROJECT_ROOT.is_dir():
        raise FileNotFoundError(
            f"Scripts directory not found: {config.PROJECT_ROOT}\n"
            "On Colab, sync this repo to COLAB_SCRIPTS_DIR in config.py."
        )
    config.ensure_dirs()
    print(f"Environment: {'Colab' if config.IS_COLAB else 'local'}")
    print(f"  Scripts:    {config.PROJECT_ROOT}")
    print(f"  Input:      {config.INPUT_DIR}")
    print(f"  Processing: {config.PROCESSING_DIR}")
    print(f"  Output:     {config.OUTPUT_DIR}")


def install_dependencies() -> None:
    """Install pinned requirements (useful in a fresh Colab runtime)."""
    req = config.PROJECT_ROOT / "requirements.txt"
    if not req.is_file():
        subprocess.check_call(
            [sys.executable, "-m", "pip", "install", "-q", "pandas>=2.0", "openpyxl>=3.1", "xlrd>=2.0"]
        )
        return
    subprocess.check_call([sys.executable, "-m", "pip", "install", "-q", "-r", str(req)])


def _list_excel_files(layout: str) -> list[Path]:
    files: list[Path] = []
    seen: set[str] = set()
    for folder in config.input_dirs_for(layout):
        for p in folder.iterdir():
            if p.suffix.lower() not in (".xlsx", ".xls") or not p.is_file():
                continue
            if p.name.startswith("~$"):
                continue
            key = str(p.resolve())
            if key in seen:
                continue
            seen.add(key)
            files.append(p)
    return sorted(files, key=lambda p: p.name.lower())


_DATE_TOKEN_RE = re.compile(
    r"\b(\d{1,2}[./-]\d{1,2}[./-]\d{2,4})\b"
)
_NEW_LANE_RE = re.compile(r"new\s+lane\s+(.+?)\s+add(?:ed)?\b", re.I)
_LANE_CODE_RE = re.compile(r"^\s*([A-Za-z]{2})\s*[-\s]*([0-9A-Za-z]+)\s*[-\s]*([A-Za-z]{2})\s*[-\s]*([0-9A-Za-z]+)\s*$")


def _normalize_lane_key(value: object) -> str:
    if value is None:
        return ""
    return re.sub(r"[^A-Za-z0-9]+", "", str(value)).upper()


def _parse_any_date(value: object) -> datetime | None:
    if value is None:
        return None
    if isinstance(value, datetime):
        return value
    text = str(value).strip()
    if not text:
        return None
    for fmt in (
        "%d.%m.%Y",
        "%d.%m.%y",
        "%m/%d/%Y",
        "%m/%d/%y",
        "%d/%m/%Y",
        "%d/%m/%y",
        "%Y-%m-%d",
        "%Y-%m-%d %H:%M:%S",
    ):
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            continue
    return None


def _format_date_ddmmyyyy(value: object) -> str | None:
    dt = _parse_any_date(value)
    if dt is None:
        return None
    return dt.strftime("%d.%m.%Y")


def _normalize_country(value: object) -> str:
    if value is None:
        return ""
    return re.sub(r"[^A-Za-z]+", "", str(value)).upper()


def _normalize_postal(value: object) -> str:
    if value is None:
        return ""
    return re.sub(r"[^0-9A-Za-z]+", "", str(value)).upper()


def _parse_lane_code(value: object) -> tuple[str, str, str, str] | None:
    if value is None:
        return None
    m = _LANE_CODE_RE.match(str(value).strip())
    if not m:
        return None
    return (
        _normalize_country(m.group(1)),
        _normalize_postal(m.group(2)),
        _normalize_country(m.group(3)),
        _normalize_postal(m.group(4)),
    )


def _extract_ratebook_validity(source: Path) -> tuple[str | None, str | None]:
    for sheet in get_sheet_names(source):
        rows = read_sheet_rows(source, sheet, as_displayed=True)
        scan_limit = min(40, len(rows))
        for i in range(scan_limit):
            row = rows[i]
            row_text = " | ".join(str(c).strip() for c in row if not is_blank(c))
            if "rate book -valid from" not in normalize_header(row_text):
                continue
            dates = _DATE_TOKEN_RE.findall(row_text)
            if len(dates) >= 2:
                start = _format_date_ddmmyyyy(dates[0])
                end = _format_date_ddmmyyyy(dates[1])
                if start or end:
                    return start, end
            for j in range(i + 1, min(i + 4, scan_limit)):
                near_text = " | ".join(
                    str(c).strip() for c in rows[j] if not is_blank(c)
                )
                dates = _DATE_TOKEN_RE.findall(near_text)
                if len(dates) >= 2:
                    start = _format_date_ddmmyyyy(dates[0])
                    end = _format_date_ddmmyyyy(dates[1])
                    if start or end:
                        return start, end
    return None, None


def _extract_latest_revision_lane_effective_date(
    source: Path,
) -> tuple[str | None, str | None]:
    revision_sheet = next(
        (s for s in get_sheet_names(source) if "revision" in normalize_header(s)),
        None,
    )
    if revision_sheet is None:
        return None, None

    rows = read_sheet_rows(source, revision_sheet, as_displayed=True)
    if not rows:
        return None, None

    header_idx: int | None = None
    effective_col: int | None = None
    for i in range(min(20, len(rows))):
        for col_i, cell in enumerate(rows[i]):
            if "effective date" in normalize_header(cell):
                header_idx = i
                effective_col = col_i
                break
        if effective_col is not None:
            break

    if effective_col is None:
        return None, None

    for r in range(len(rows) - 1, (header_idx or 0), -1):
        row = rows[r]
        if not any(not is_blank(c) for c in row):
            continue
        text = " | ".join(str(c).strip() for c in row if not is_blank(c))
        m = _NEW_LANE_RE.search(text)
        if not m:
            continue
        lane_key = _normalize_lane_key(m.group(1))
        if not lane_key:
            continue
        effective_val = row[effective_col] if effective_col < len(row) else None
        valid_from = _format_date_ddmmyyyy(effective_val)
        if valid_from:
            return lane_key, valid_from
    return None, None


def _apply_validity_columns(df: pd.DataFrame, source: Path) -> pd.DataFrame:
    out = df.copy()
    default_valid_from, default_valid_to = _extract_ratebook_validity(source)
    out["valid_from"] = default_valid_from
    out["valid_to"] = default_valid_to

    lane_key, lane_valid_from = _extract_latest_revision_lane_effective_date(source)
    if lane_key and lane_valid_from and "lane_id" in out.columns:
        lane_keys = out["lane_id"].apply(_normalize_lane_key)
        exact_mask = lane_keys == lane_key
        contains_mask = lane_keys.apply(
            lambda k: bool(k) and (lane_key in k or k in lane_key)
        )
        target_mask = exact_mask | contains_mask

        parsed_code = _parse_lane_code(lane_key)
        if parsed_code is not None:
            o_country, o_postal, d_country, d_postal = parsed_code
            by_fields = pd.Series(True, index=out.index)
            if "origin_country" in out.columns:
                by_fields &= out["origin_country"].apply(_normalize_country) == o_country
            if "origin_zip" in out.columns:
                by_fields &= out["origin_zip"].apply(_normalize_postal) == o_postal
            if "dest_country" in out.columns:
                by_fields &= out["dest_country"].apply(_normalize_country) == d_country
            if "dest_zip" in out.columns:
                by_fields &= out["dest_zip"].apply(_normalize_postal) == d_postal
            target_mask = target_mask | by_fields

        out.loc[target_mask, "valid_from"] = lane_valid_from
    return out


def convert_one(
    layout: str,
    source: Path,
    *,
    sheets: list[str] | None = None,
    sheet_configs: dict | None = None,
    new_grid_roundtrip_as_row: bool = False,
    include_validity_columns: bool = False,
) -> Path | None:
    """Convert one workbook to long-format *_converted.xlsx under processing/."""
    layout_key = config.normalize_layout(layout)
    converter = CONVERTERS[layout_key]
    kwargs: dict = {}
    layout_key = config.normalize_layout(layout)
    if layout_key == "usual_rate" and sheet_configs:
        kwargs["sheet_configs"] = sheet_configs
    if layout_key == "new_grid":
        kwargs["roundtrip_as_row"] = new_grid_roundtrip_as_row
    try:
        df = converter(source, sheets=sheets, **kwargs)
    except ValueError as exc:
        print(f"  ERROR converting {source.name}: {exc}")
        return None

    if df.empty:
        print(f"  WARNING: No price rows — {source.name}")
        return None
    if include_validity_columns:
        df = _apply_validity_columns(df, source)

    dest = output_path(config.PROCESSING_DIR, layout_key, source)
    save_dataframe(df, dest)
    print(f"  Converted {source.name} -> {dest} ({len(df):,} rows)")
    return dest


def export_one(converted: Path) -> Path | None:
    """Export one *_converted.xlsx to matrix layout under output/."""
    out = config.matrix_output_path(converted)
    try:
        export_converted_file_to_matrix(converted, output_path=out)
    except Exception as exc:
        print(f"  ERROR exporting {converted.name}: {exc}")
        return None
    if copy_excel_sheet(converted, out, ACCESSORIAL_OUTPUT_SHEET):
        from common import format_accessorial_sheet

        format_accessorial_sheet(out, ACCESSORIAL_OUTPUT_SHEET)
        print("  Accessorial sheet copied to matrix output")
    print(f"  Matrix {converted.name} -> {out}")
    return out


def _choose_accessorial_tabs(source: Path) -> list[str] | None:
    """Ask whether to add accessorial costs and which tab(s) to use."""
    answer = _prompt("\n[Step 2/3] Add accessorial costs? [y/N]: ").strip().lower()
    if answer not in ("y", "yes"):
        return None
    try:
        sheet_names = get_sheet_names(source)
    except Exception as exc:
        print(f"  ERROR reading workbook tabs: {exc}")
        return None
    if not sheet_names:
        print("  No sheets found in workbook.")
        return None
    print("\nAccessorial tab(s) in this workbook:")
    for i, name in enumerate(sheet_names, 1):
        print(f"  {i}. {name}")
    print("\nEnter tab number(s) or name(s), comma-separated:")
    while True:
        raw = _prompt("Accessorial tab(s): ")
        try:
            selected = _parse_tab_selection(raw, sheet_names)
        except ValueError as exc:
            print(f"  {exc}")
            retry = _prompt("Try again? [Y/n]: ").lower()
            if retry in ("n", "no"):
                return None
            continue
        if not selected:
            print("  Select at least one tab.")
            continue
        return selected


def add_accessorial(source: Path, converted_path: Path, sheet_names: list[str]) -> int:
    """Parse accessorial tab(s) and append sheet to the converted workbook."""
    df = parse_accessorial_file(source, sheet_names)
    if df.empty:
        print("  WARNING: No accessorial rows extracted.")
        return 0
    df = df[list(ACCESSORIAL_COLUMNS)]
    append_excel_sheet(
        converted_path,
        df,
        ACCESSORIAL_OUTPUT_SHEET,
        apply_accessorial_format=True,
    )
    print(
        f"  Accessorial: {len(df)} row(s) -> sheet '{ACCESSORIAL_OUTPUT_SHEET}'"
    )
    return len(df)


def process_one_workbook(
    layout: str,
    source: Path,
    *,
    sheets: list[str] | None = None,
    sheet_configs: dict | None = None,
    new_grid_roundtrip_as_row: bool = False,
    include_validity_columns: bool = False,
    prompt_accessorial: bool = True,
    export_matrix: bool = True,
    input_fn=None,
) -> WorkbookResult:
    """
    Full rate processing for one file:
      1. Convert main rate tabs → processing/*_converted.xlsx (sheet ``rates``)
      2. Optional accessorial tab(s) → sheet ``Accessorial Costs``
      3. Matrix export → output/*_matrix.xlsx
    """
    result = WorkbookResult(source=source)
    print(f"\n{'=' * 60}\n  Processing: {source.name}\n{'=' * 60}")

    print("\n--- Step 1/3: Convert main rates ---")
    dest = convert_one(
        layout,
        source,
        sheets=sheets,
        sheet_configs=sheet_configs,
        new_grid_roundtrip_as_row=new_grid_roundtrip_as_row,
        include_validity_columns=include_validity_columns,
    )
    if dest is None:
        result.skipped = True
        return result
    result.converted = dest

    if prompt_accessorial:
        tabs = _choose_accessorial_tabs(source)
        if tabs:
            print("\n--- Step 2/3: Accessorial costs ---")
            result.accessorial_rows = add_accessorial(source, dest, tabs)
        else:
            print("\n--- Step 2/3: Accessorial costs — skipped ---")
    elif input_fn is None:
        print("\n--- Step 2/3: Accessorial costs — skipped ---")

    if export_matrix:
        print("\n--- Step 3/3: Export matrix Excel ---")
        out = export_one(dest)
        result.matrix = out
    else:
        print("\n--- Step 3/3: Matrix export — skipped ---")

    return result


def find_converted_files(layouts: list[str] | None = None) -> list[Path]:
    layouts = layouts or list(config.LAYOUTS)
    files: list[Path] = []
    seen: set[str] = set()
    for layout in layouts:
        layout_key = config.normalize_layout(layout)
        folders = [config.PROCESSING_DIR / layout_key]
        folders.extend(
            config.PROCESSING_DIR / legacy
            for legacy in config.LEGACY_PROCESSING_FOLDERS.get(layout_key, ())
            if (config.PROCESSING_DIR / legacy).is_dir()
        )
        for folder in folders:
            if not folder.is_dir():
                continue
            for p in folder.glob("*_converted.xlsx"):
                if not p.is_file() or p.name.startswith("~$"):
                    continue
                key = str(p.resolve())
                if key in seen:
                    continue
                seen.add(key)
                files.append(p)
    return sorted(files)


def run_pipeline(
    *,
    layouts: list[str] | None = None,
    files: list[Path] | None = None,
    sheets: list[str] | None = None,
    convert: bool = True,
    export: bool = True,
) -> PipelineResult:
    """
    Run convert and/or matrix export for all matching workbooks.

    Parameters
    ----------
    layouts:
        Subset of usual_rate / new_grid. Default: all layouts in config.LAYOUTS.
    files:
        If set, only these input paths are converted (layout inferred from parent).
    sheets:
        Sheet names to convert; None = auto-detect price tabs per file.
    convert:
        When True, read input/*.xlsx and write processing/*_converted.xlsx.
    export:
        When True, read processing/*_converted.xlsx and write output/*_matrix.xlsx.
    """
    result = PipelineResult()
    layouts = layouts or list(config.LAYOUTS)

    if convert:
        print("\n--- Step 1: Convert ratebooks ---")
        for layout in layouts:
            if files is not None:
                allowed = {d.resolve() for d in config.input_dirs_for(layout)}
                sources = [p for p in files if p.parent.resolve() in allowed]
            else:
                sources = _list_excel_files(layout)
            if not sources:
                print(f"  [{layout}] no input files")
                continue
            print(f"  [{layout}] {len(sources)} file(s)")
            for source in sources:
                dest = convert_one(layout, source, sheets=sheets)
                if dest is None:
                    result.convert_skipped.append(str(source))
                else:
                    result.converted.append(dest)

    if export:
        print("\n--- Step 2: Export matrix Excel ---")
        converted_files = result.converted if convert else find_converted_files(layouts)
        if not converted_files:
            print("  No *_converted.xlsx files to export.")
        for converted in converted_files:
            out = export_one(converted)
            if out is None:
                result.export_errors.append(str(converted))
            else:
                result.matrices.append(out)

    print("\n--- Summary ---")
    print(f"  Converted: {len(result.converted)}")
    print(f"  Matrices:  {len(result.matrices)}")
    if result.convert_skipped:
        print(f"  Skipped:   {len(result.convert_skipped)}")
    if result.export_errors:
        print(f"  Export errors: {len(result.export_errors)}")
    return result


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
                raise ValueError(f"Tab number {idx} out of range (1–{len(sheet_names)})")
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
            raise ValueError(f"Ambiguous tab '{part}'. Matches: {', '.join(matches)}")
        else:
            raise ValueError(f"Unknown tab: {part}")
    return resolved


def _choose_layout() -> str | None:
    print("\nRate workbook type:")
    for i, name in enumerate(config.LAYOUTS, 1):
        count = len(_list_excel_files(name))
        label = LAYOUT_LABELS.get(name, name)
        print(f"  {i}. {label} ({count} files)")
    print("  0. Exit")
    while True:
        choice = _prompt("\nSelect layout number: ")
        if choice == "0":
            return None
        if choice.isdigit() and 1 <= int(choice) <= len(config.LAYOUTS):
            return config.LAYOUTS[int(choice) - 1]
        print("Invalid choice. Enter a number from the list.")


def _choose_files(files: list[Path]) -> list[Path]:
    print(f"\nFiles in input/{files[0].parent.name if files else ''}:")
    for i, path in enumerate(files, 1):
        print(f"  {i}. {path.name}")
    print(f"  {len(files) + 1}. ALL files in this layout (auto price tabs only)")
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


def run_full_processing() -> PipelineResult:
    """
    Interactive full rate processing: layout → file → rate tabs →
    convert → accessorial (optional) → matrix export.
    """
    result = PipelineResult()
    print("=" * 60)
    print("  Aptiv Road — full rate processing")
    print("=" * 60)
    print("  Steps per file: (1) convert rates  (2) accessorial  (3) matrix")
    print(f"  Input:      {config.INPUT_DIR}")
    print(f"  Processing: {config.PROCESSING_DIR}")
    print(f"  Output:     {config.OUTPUT_DIR}")

    while True:
        layout = _choose_layout()
        if layout is None:
            break

        files = _list_excel_files(layout)
        if not files:
            print(f"No Excel files found for {layout}")
            continue

        selected = _choose_files(files)
        if not selected:
            continue

        for path in selected:
            sheets: list[str] | None = None
            sheet_configs = None
            new_grid_roundtrip_as_row = False
            include_validity_columns = False
            if len(selected) == 1:
                print(f"\n--- Rate tabs for {path.name} ---")
                sheets = _choose_sheets(path)
                if sheets == []:
                    continue
                layout_key = config.normalize_layout(layout)
                if layout_key == "usual_rate":
                    from converters.usual_rate import gather_column_overrides

                    sheet_configs = gather_column_overrides(
                        path, sheets, input_fn=_prompt
                    )
                if layout_key == "new_grid":
                    rt_choice = _prompt(
                        "Roundtrip in matrix: cost column or shipment row? [column/row]: "
                    ).strip().lower()
                    new_grid_roundtrip_as_row = rt_choice in ("row", "r")
                validity_choice = _prompt(
                    "Add validity columns (Valid From / Valid To)? [y/N]: "
                ).strip().lower()
                include_validity_columns = validity_choice in ("y", "yes")

            wb = process_one_workbook(
                layout,
                path,
                sheets=sheets,
                sheet_configs=sheet_configs,
                new_grid_roundtrip_as_row=new_grid_roundtrip_as_row,
                include_validity_columns=include_validity_columns,
                prompt_accessorial=True,
                export_matrix=True,
            )
            if wb.skipped or wb.converted is None:
                result.convert_skipped.append(str(path))
                continue
            result.converted.append(wb.converted)
            result.accessorial_rows += wb.accessorial_rows
            if wb.matrix is None:
                result.export_errors.append(str(wb.converted))
            else:
                result.matrices.append(wb.matrix)

        print(
            f"\n--- Round summary ---\n"
            f"  Converted:   {len(result.converted)}\n"
            f"  Matrices:    {len(result.matrices)}\n"
            f"  Accessorial: {result.accessorial_rows} row(s) total"
        )
        again = _prompt("\nProcess another file? [y/N]: ").lower()
        if again not in ("y", "yes"):
            break

    print("\nGoodbye.")
    return result


def run_interactive() -> PipelineResult:
    """Alias for run_full_processing()."""
    return run_full_processing()


def _parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Aptiv Road: convert ratebooks and export matrix Excel."
    )
    parser.add_argument(
        "--layout",
        action="append",
        dest="layouts",
        choices=config.LAYOUTS,
        help="Process only this layout (repeatable). Default: all layouts.",
    )
    parser.add_argument(
        "--file",
        action="append",
        dest="files",
        type=Path,
        help="Process only this input file (repeatable; path under input/<layout>/).",
    )
    parser.add_argument(
        "--sheets",
        help="Comma-separated sheet names (default: auto-detect price tabs).",
    )
    parser.add_argument(
        "--convert-only",
        action="store_true",
        help="Only run conversion step.",
    )
    parser.add_argument(
        "--export-only",
        action="store_true",
        help="Only export matrices from existing *_converted.xlsx files.",
    )
    parser.add_argument(
        "--setup",
        action="store_true",
        help="Mount Colab Drive and create folders, then exit.",
    )
    parser.add_argument(
        "--install",
        action="store_true",
        help="pip install requirements.txt, then exit.",
    )
    parser.add_argument(
        "-i",
        "--interactive",
        action="store_true",
        help="Full interactive processing (convert + accessorial + matrix).",
    )
    parser.add_argument(
        "--batch",
        action="store_true",
        help="Process all input files without prompts (skip file/tab picker).",
    )
    if argv is None:
        # Colab/Jupyter passes `-f kernel-*.json`; ignore unknown args.
        args, _ = parser.parse_known_args()
    else:
        args = parser.parse_args(argv)
    return args


def _in_notebook() -> bool:
    try:
        get_ipython()  # type: ignore[name-defined]
        return True
    except NameError:
        return False


def _use_interactive_mode(args: argparse.Namespace, argv: list[str] | None) -> bool:
    """Default: interactive (choose file, tabs, accessorial). Use --batch to skip prompts."""
    if args.interactive:
        return True
    if args.batch:
        return False
    if args.export_only or args.convert_only:
        return False
    # CLI filters without --batch → batch run on that subset only
    if args.layouts or args.files:
        return False
    return True


def main(argv: list[str] | None = None) -> int:
    args = _parse_args(argv)

    if args.install:
        install_dependencies()
        return 0

    interactive = _use_interactive_mode(args, argv)
    if args.setup or config.IS_COLAB or interactive:
        drive_mounted = Path("/content/drive/MyDrive").exists() or Path(
            "/content/drive/Shareddrives"
        ).exists()
        setup_environment(mount_drive=config.IS_COLAB and not drive_mounted)

    if interactive:
        run_full_processing()
        return 0

    sheets: list[str] | None = None
    if args.sheets:
        sheets = [s.strip() for s in args.sheets.split(",") if s.strip()]

    convert = not args.export_only
    export = not args.convert_only
    if args.setup and not convert and not export:
        return 0

    result = run_pipeline(
        layouts=args.layouts,
        files=args.files,
        sheets=sheets,
        convert=convert,
        export=export,
    )
    failed = len(result.convert_skipped) + len(result.export_errors)
    return 1 if failed and not result.converted and not result.matrices else 0


if __name__ == "__main__":
    exit_code = main()
    if not _in_notebook():
        raise SystemExit(exit_code)
