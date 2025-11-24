"""Configuration values used across the CEJ transformer package."""

from dataclasses import dataclass
from typing import Dict, List


APP_NAME = "Excel Transformer"
VERSION = "2.4.0"

LOG_FILE = "transformer.log"
LOG_LEVEL = "INFO"
LOG_FORMAT = "%(asctime)s - %(name)s - %(levelname)s - %(message)s"
LOG_MAX_BYTES = 10 * 1024 * 1024  # 10MB rotating log size
LOG_BACKUP_COUNT = 3

DUAL_LANG_INPUT_SHEET_NAME = "Tracker (Dual Lang)"
SINGLE_LANG_INPUT_SHEET_NAME = "Tracker (Single Lang)"

OUTPUT_FILE_BASENAME = "transformed_CEJ_master_specsheet"
OUTPUT_SHEET_NAME_DUAL_LANG = "Transformed_Dual_Lang"
OUTPUT_SHEET_NAME_SINGLE_LANG = "Transformed_Single_Lang"

OUTPUT_COLUMNS_BASE: List[str] = [
    "Platform",
    "Funnel Stage",
    "Format",
    "Duration",
    "Aspect Ratio / Format",
]
OUTPUT_LANGUAGE_COLUMN = "Languages"

FUNNEL_STAGE_HEADER = "Funnel Stage"
FORMAT_HEADER = "Format"
DURATION_HEADER = "Duration"

MAIN_HEADER_ASPECT_RATIO_GROUP_PRIMARY = "Aspect Ratio"
MAIN_HEADER_ASPECT_RATIO_GROUP_SECONDARY = "Format"
MAIN_HEADER_LANGUAGES_GROUP = "Languages"
MAIN_HEADER_TOTAL_COL = "TOTAL"

START_ROW_SEARCH_FOR_PLATFORM = 7
SUB_HEADER_ROW_OFFSET = 3
DATA_START_ROW_OFFSET = 4

PLATFORM_NAMES: Dict[str, str] = {
    "youtube": "YouTube",
    "meta": "META",
    "tiktok": "TikTok",
    "linkedin": "LinkedIn",
    "programmatic": "Programmatic",
    "audio": "Audio",
    "gaming": "Gaming",
    "amazon": "Amazon",
}

FUNNEL_STAGES = ["Awareness", "Consideration", "Purchase"]
EXPAND_ALL_TO_ACP = True


@dataclass(frozen=True)
class SheetSpecification:
    """Defines how a worksheet should be processed."""

    sheet_name: str
    is_dual_language: bool
    output_sheet_name: str
    output_columns: List[str]


SHEET_SPECS: List[SheetSpecification] = [
    SheetSpecification(
        sheet_name=DUAL_LANG_INPUT_SHEET_NAME,
        is_dual_language=True,
        output_sheet_name=OUTPUT_SHEET_NAME_DUAL_LANG,
        output_columns=OUTPUT_COLUMNS_BASE + [OUTPUT_LANGUAGE_COLUMN],
    ),
    SheetSpecification(
        sheet_name=SINGLE_LANG_INPUT_SHEET_NAME,
        is_dual_language=False,
        output_sheet_name=OUTPUT_SHEET_NAME_SINGLE_LANG,
        output_columns=OUTPUT_COLUMNS_BASE,
    ),
]
