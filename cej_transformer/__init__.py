"""Core package exposing Excel transformation utilities."""

from .config import (
    APP_NAME,
    VERSION,
    DUAL_LANG_INPUT_SHEET_NAME,
    SINGLE_LANG_INPUT_SHEET_NAME,
    OUTPUT_FILE_BASENAME,
    OUTPUT_SHEET_NAME_DUAL_LANG,
    OUTPUT_SHEET_NAME_SINGLE_LANG,
)

from .transformer import process_workbook
from .validator import validate_output

__all__ = [
    "APP_NAME",
    "VERSION",
    "DUAL_LANG_INPUT_SHEET_NAME",
    "SINGLE_LANG_INPUT_SHEET_NAME",
    "OUTPUT_FILE_BASENAME",
    "OUTPUT_SHEET_NAME_DUAL_LANG",
    "OUTPUT_SHEET_NAME_SINGLE_LANG",
    "process_workbook",
    "validate_output",
]
