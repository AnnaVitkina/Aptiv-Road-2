"""Parse accessorial / additional-cost tabs into a normalized table."""

from __future__ import annotations

import re
from typing import Any

import pandas as pd

from common import is_blank, read_sheet_rows

ACCESSORIAL_COLUMNS = (
    "Rate Card Name",
    "Rate Agreement Name",
    "cost",
    "currency",
    "condition",
    "measurement",
)

_NUMBERED_RE = re.compile(r"^(\d+(?:\.\d+)*)\.?\s*$")
_CURRENCY_RE = re.compile(r"\b(EUR|USD|GBP|PLN|RON|CHF|CZK|HUF)\b", re.I)
_MEASUREMENT_TAIL_RE = re.compile(r"^(.+?)\s+(per\s+.+)$", re.I)
_NUMERIC_COST_RE = re.compile(
    r"^[\s]*[\d]+[.,][\d]+(?:\s*€)?\s*$|^[\s]*[\d]+(?:\s*€)?\s*$",
    re.I,
)
_PER_MEASUREMENT_RE = re.compile(r"\bper\s+[\w\s]+", re.I)


def _cell(row: list[Any], idx: int) -> Any:
    if idx >= len(row):
        return None
    return row[idx]


def _text(value: Any) -> str | None:
    if is_blank(value):
        return None
    return str(value).strip()


def _numbering_depth(value: Any) -> int | None:
    text = _text(value)
    if not text:
        return None
    m = _NUMBERED_RE.match(text)
    if not m:
        return None
    return m.group(1).count(".")


def _looks_measurement(value: Any) -> bool:
    text = _text(value)
    if not text:
        return False
    if _NUMERIC_COST_RE.match(text.replace(" ", "")):
        return False
    return bool(_PER_MEASUREMENT_RE.search(text))


def _is_numeric_cost(value: Any) -> bool:
    text = _text(value)
    if not text:
        return False
    return bool(_NUMERIC_COST_RE.match(text.replace(" ", "")))


def _normalize_cost(value: Any) -> Any:
    if is_blank(value):
        return None
    if isinstance(value, (int, float)):
        return value
    text = str(value).strip()
    if text.upper() in ("NA", "N/A", "-", ""):
        return "NA"
    m = re.search(r"([\d]+[.,][\d]+)", text)
    if m:
        return m.group(1).replace(",", ".")
    return text


def _detect_currency(rows: list[list[Any]]) -> str:
    for row in rows[:15]:
        for cell in row:
            text = _text(cell)
            if not text:
                continue
            m = _CURRENCY_RE.search(text)
            if m:
                return m.group(1).upper()
            if "€" in text or "eur" in text.lower():
                return "EUR"
    return "EUR"


def _split_label_measurement(name: str) -> tuple[str, str | None]:
    """'Docs fees per shipment' -> ('Docs fees', 'per shipment')."""
    s = name.strip()
    m = re.search(r"^(.+?)\s+\(\s*(per\s+[^)]+)\s*\)\s*$", s, re.I)
    if m:
        return m.group(1).strip(), m.group(2).strip()
    m = _MEASUREMENT_TAIL_RE.match(s)
    if m:
        return m.group(1).strip(), m.group(2).strip()
    return s, None


def _build_card_name(parent: str | None, label: str) -> str:
    if parent and label and label.lower() != parent.lower():
        return f"{parent}({label})"
    return label or parent or ""


def _append_record(
    records: list[dict[str, Any]],
    *,
    currency: str,
    rate_card_name: str | None,
    cost: Any = None,
    measurement: str | None = None,
    condition: str | None = None,
) -> None:
    if is_blank(rate_card_name):
        return
    if is_blank(cost) and is_blank(measurement) and is_blank(condition):
        return
    records.append(
        {
            "Rate Card Name": rate_card_name,
            "Rate Agreement Name": None,
            "cost": _normalize_cost(cost) if not is_blank(cost) else None,
            "currency": currency if not is_blank(cost) or measurement else None,
            "condition": condition,
            "measurement": measurement,
        }
    )


def _scan_cost_and_measurement(
    row: list[Any], start_col: int = 2
) -> tuple[Any, str | None, str | None]:
    """Return (cost, measurement, condition) from columns to the right of the name."""
    cost: Any = None
    measurement: str | None = None
    condition: str | None = None
    for i in range(start_col, min(len(row), 10)):
        cell = _cell(row, i)
        if is_blank(cell):
            continue
        text = _text(cell) or ""
        if _looks_measurement(cell):
            measurement = measurement or text
        elif _is_numeric_cost(cell):
            cost = cell
        elif re.match(r"^\d+\s*h(?:ours?)?$", text, re.I):
            condition = text
        elif cost is None and len(text) < 120:
            if any(
                k in text.lower()
                for k in ("included", "n/a", "mechanism", "calculated", "max ", "%")
            ):
                cost = cell
            elif not _looks_measurement(cell):
                cost = cell
    return cost, measurement, condition


def parse_accessorial_rows(
    rows: list[list[Any]],
    *,
    sheet_name: str = "",
    default_currency: str | None = None,
) -> pd.DataFrame:
    """
  Parse Aptiv/DSV-style accessorial grids.

  - Numbered row with name (col B) + price on the same row → one charge line.
  - Numbered subsection (e.g. 1.1) without price → group parent for following rows.
  - Unnumbered rows under a parent → ``Parent(Label)`` with cost from the right.
    """
    currency = default_currency or _detect_currency(rows)
    records: list[dict[str, Any]] = []

    main_card: str | None = None
    group_parent: str | None = None
    inherited_measurement: str | None = None

    for row in rows:
        num = _cell(row, 0)
        name = _text(_cell(row, 1)) or _text(_cell(row, 2))
        if is_blank(name) and is_blank(num):
            continue

        depth = _numbering_depth(num)
        cost, measurement, condition = _scan_cost_and_measurement(row, start_col=2)
        label, meas_in_name = _split_label_measurement(name) if name else ("", None)
        row_measurement = meas_in_name or measurement

        # --- Top-level numbered row (1., 2., 12.) ---
        if depth == 0 and name:
            main_card = label
            group_parent = None
            inherited_measurement = None
            if _looks_measurement(_cell(row, 4)) and not _is_numeric_cost(_cell(row, 4)):
                inherited_measurement = _text(_cell(row, 4))
            if cost is not None or _is_numeric_cost(cost):
                _append_record(
                    records,
                    currency=currency,
                    rate_card_name=main_card,
                    cost=cost,
                    measurement=row_measurement,
                    condition=condition,
                )
            continue

        # --- Subsection header (1.1., 2.1.) — group for following detail rows ---
        if depth == 1 and name:
            group_parent = label
            if row_measurement and not cost:
                inherited_measurement = row_measurement
            if cost is not None or _is_numeric_cost(cost):
                _append_record(
                    records,
                    currency=currency,
                    rate_card_name=main_card or label,
                    cost=cost,
                    measurement=inherited_measurement or row_measurement,
                    condition=label if main_card else None,
                )
            continue

        # --- Deep sub-number (3.1., 13.1) — condition under main card ---
        if depth is not None and depth >= 2 and name:
            _append_record(
                records,
                currency=currency,
                rate_card_name=main_card,
                cost=cost,
                measurement=inherited_measurement or row_measurement,
                condition=label,
            )
            continue

        # --- Detail row (no number): under group_parent or main_card ---
        if name:
            parent = group_parent or main_card
            card_name = _build_card_name(parent, label)
            _append_record(
                records,
                currency=currency,
                rate_card_name=card_name,
                cost=cost,
                measurement=row_measurement or inherited_measurement,
                condition=condition if not group_parent else None,
            )

    if not records:
        return pd.DataFrame(columns=list(ACCESSORIAL_COLUMNS))
    df = pd.DataFrame(records)
    for col in ACCESSORIAL_COLUMNS:
        if col not in df.columns:
            df[col] = None
    return df[list(ACCESSORIAL_COLUMNS)]


def parse_accessorial_file(path, sheet_names: list[str]) -> pd.DataFrame:
    """Parse one or more tabs from a workbook into one accessorial DataFrame."""
    frames: list[pd.DataFrame] = []
    for sheet in sheet_names:
        rows = read_sheet_rows(path, sheet, as_displayed=True)
        df = parse_accessorial_rows(rows, sheet_name=sheet)
        if df.empty:
            continue
        frames.append(df)
    if not frames:
        return pd.DataFrame(columns=list(ACCESSORIAL_COLUMNS))
    return pd.concat(frames, ignore_index=True)[list(ACCESSORIAL_COLUMNS)]
