"""Normalized column schema for usual-rate and new-grid converters."""

USUAL_RATE_DATA_COLUMNS = (
    "lane_id",
    "lane_number",
    "paid_by",
    "origin_name",
    "origin_zip",
    "origin_city",
    "origin_country",
    "dest_name",
    "dest_zip",
    "dest_city",
    "dest_country",
    "lane_description",
    "currency",
    "price_per",
    "description",
    "cost_component",
    "equipment_type",
    "roundtrip",
    "rate_column",
    "rate_group",
    "price",
    "row_number",
)

# Backward-compatible alias
LAYOUT1_DATA_COLUMNS = USUAL_RATE_DATA_COLUMNS
