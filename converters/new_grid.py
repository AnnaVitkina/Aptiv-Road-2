"""New grid converter — Rate Grid FTL/LTL/Multi-stop tabs (former layout3)."""

from __future__ import annotations

from pathlib import Path

import pandas as pd

from common import add_metadata, read_sheet_rows, reorder_converted_df, sheets_to_convert
from converters.layout3 import _is_price_tab, _parse_grid_sheet
from converters.schema import USUAL_RATE_DATA_COLUMNS

LAYOUT_KEY = "new_grid"


def convert_file(path: Path, sheets: list[str] | None = None) -> pd.DataFrame:
    frames: list[pd.DataFrame] = []
    for sheet in sheets_to_convert(path, sheets=sheets, auto_include=_is_price_tab):
        rows = read_sheet_rows(path, sheet, as_displayed=True)
        df = _parse_grid_sheet(rows)
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
    return reorder_converted_df(
        pd.concat(frames, ignore_index=True), USUAL_RATE_DATA_COLUMNS
    )


__all__ = ["LAYOUT_KEY", "convert_file"]
