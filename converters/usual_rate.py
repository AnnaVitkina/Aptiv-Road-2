"""
Usual rate converter — merged layout1 + layout2 + layout4.

Lane ratebooks, WMT Romania matrices, and simple FTL/LTL pricelists.
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Any

import pandas as pd

from common import (
    add_metadata,
    is_blank,
    normalize_header,
    read_sheet_rows,
    sheets_to_convert,
    should_skip_tab,
)
from converters.layout1 import (
    SheetColumnConfig,
    _detect_lane_structure,
    _is_price_tab as _is_lane_price_tab,
    _parse_lane_sheet,
    gather_column_overrides,
)
from converters.layout2 import (
    PRICE_TABS as LAYOUT2_PRICE_TABS,
    _is_fixed_rates_sheet,
    _parse_fixed_rates_sheet,
    _parse_matrix_sheet,
    _parse_tunisia_sheet,
)
from converters.layout4 import (
    SKIP_LAYOUT4,
    _parse_corridor_matrix,
)
from converters.schema import USUAL_RATE_DATA_COLUMNS

LAYOUT_KEY = "usual_rate"

_SKIP_USUAL = re.compile(
    r"tt\s+details|accessorial|accesorial|fsc|revision",
    re.I,
)


def _is_price_tab(name: str) -> bool:
    """Auto-include price tabs from former layout1, layout2, and layout4."""
    if _SKIP_USUAL.search(name):
        return False
    if SKIP_LAYOUT4.search(name):
        return False
    lower = name.strip().lower()
    if lower in LAYOUT2_PRICE_TABS:
        return True
    if _is_lane_price_tab(name):
        return True
    if should_skip_tab(name, LAYOUT_KEY):
        return False
    # layout4-style short tabs (after generic skip rules)
    if any(
        x in lower
        for x in (
            "ftl",
            "ltl",
            "pallet",
            "ldm",
            "milkrun",
            "xd_",
            "x-dock",
            "corbas",
            "groupage",
        )
    ):
        return True
    return False


def _parse_corridor_if_match(rows: list[list[Any]]) -> pd.DataFrame:
    r0 = " ".join(normalize_header(c) for c in rows[0]) if rows else ""
    if "trailer" in r0 or (
        rows
        and normalize_header(rows[0][0] if rows[0] else "") == ""
        and any("origin" in normalize_header(c) for c in rows[1] if rows[1:])
    ):
        return _parse_corridor_matrix(rows)
    return pd.DataFrame()


def _parse_usual_rate_sheet(
    rows: list[list[Any]],
    sheet_name: str,
    *,
    column_overrides: dict[str, int] | None = None,
    skip_columns: set[int] | None = None,
) -> pd.DataFrame:
    """Dispatch to layout2, layout4 corridor, or layout1 lane parsers."""
    name = sheet_name.strip().lower()

    if _is_fixed_rates_sheet(rows):
        return _parse_fixed_rates_sheet(rows)
    if "tunisia" in name:
        return _parse_tunisia_sheet(rows)

    corridor = _parse_corridor_if_match(rows)
    if not corridor.empty:
        return corridor

    if _detect_lane_structure(rows) is not None:
        return _parse_lane_sheet(
            rows,
            sheet_name,
            column_overrides=column_overrides,
            skip_columns=skip_columns,
        )

    # WMT-style country × band matrix
    return _parse_matrix_sheet(rows)


def convert_file(
    path: Path,
    sheets: list[str] | None = None,
    sheet_configs: dict[str, SheetColumnConfig] | None = None,
) -> pd.DataFrame:
    frames: list[pd.DataFrame] = []
    for sheet in sheets_to_convert(path, sheets=sheets, auto_include=_is_price_tab):
        rows = read_sheet_rows(path, sheet, as_displayed=True)
        cfg = (sheet_configs or {}).get(sheet, {})
        df = _parse_usual_rate_sheet(
            rows,
            sheet,
            column_overrides=cfg.get("overrides"),
            skip_columns=cfg.get("skip_columns"),
        )
        if df.empty:
            continue
        frames.append(
            add_metadata(
                df,
                source_file=path.name,
                layout=LAYOUT_KEY,
                sheet_name=sheet,
            )
        )
    if not frames:
        return pd.DataFrame()
    from common import reorder_converted_df

    return reorder_converted_df(
        pd.concat(frames, ignore_index=True), USUAL_RATE_DATA_COLUMNS
    )


__all__ = [
    "LAYOUT_KEY",
    "convert_file",
    "gather_column_overrides",
    "USUAL_RATE_DATA_COLUMNS",
]
