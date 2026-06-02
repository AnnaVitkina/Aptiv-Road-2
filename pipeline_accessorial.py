"""Backward compatibility — use ``pipeline.run_full_processing`` instead."""

from common import copy_excel_sheet
from pipeline import (
    ACCESSORIAL_OUTPUT_SHEET,
    add_accessorial,
    process_one_workbook,
    run_full_processing,
)

add_accessorial_to_converted = add_accessorial


def copy_accessorial_to_matrix(converted_path, matrix_path) -> bool:
    return copy_excel_sheet(converted_path, matrix_path, ACCESSORIAL_OUTPUT_SHEET)
