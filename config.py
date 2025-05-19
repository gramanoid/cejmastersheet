"""
Configuration settings for the Excel Transformer application.
"""
from typing import Dict, List, Tuple, Optional
from pathlib import Path

# Application constants
APP_NAME = "Excel Transformer"
VERSION = "2.0.0"

# File handling
DUAL_LANG_INPUT_SHEET_NAME = 'Tracker (Dual Lang)'
SINGLE_LANG_INPUT_SHEET_NAME = 'Tracker (Single Lang)'
OUTPUT_FILE_BASENAME = 'transformed_CEJ_master_specsheet'
LOG_FILE = 'transformer.log'
DEFAULT_OUTPUT_FORMAT = 'xlsx'  # Can be 'xlsx' or 'csv'

# Output sheet names for the transformed data
OUTPUT_SHEET_NAME_DUAL_LANG = 'Transformed_Dual_Lang'
OUTPUT_SHEET_NAME_SINGLE_LANG = 'Transformed_Single_Lang'

# Platform configuration
PLATFORM_NAMES: Dict[str, str] = {
    "youtube": "YouTube",
    "meta": "META",
    "tiktok": "TikTok",
    "programmatic": "Programmatic",
    "audio": "Audio",
    "gaming": "Gaming",
    "amazon": "Amazon"
}

# Excel parsing constants
MAIN_HEADER_ASPECT_RATIO_GROUP_PRIMARY = "Aspect Ratio"
MAIN_HEADER_ASPECT_RATIO_GROUP_SECONDARY = "Format"
MAIN_HEADER_LANGUAGES_GROUP = "Languages"
MAIN_HEADER_TOTAL_COL = "TOTAL"

# Column header constants used across the application
PLATFORM_COLUMN_HEADER = "PLATFORM"
FUNNEL_STAGE_HEADER = "Funnel Stage"
FORMAT_HEADER = "Format"
DURATION_HEADER = "Duration"
TOTAL_HEADER = MAIN_HEADER_TOTAL_COL
LANGUAGE_HEADER = MAIN_HEADER_LANGUAGES_GROUP

# Core headers expected in the main header row for validation
CORE_MAIN_HEADERS = [
    "Funnel Stage",
    "Format",
    "Duration",
    MAIN_HEADER_LANGUAGES_GROUP,
    MAIN_HEADER_TOTAL_COL
]

# Row offsets (0-indexed)
START_ROW_SEARCH_FOR_PLATFORM = 7  # Excel row 8 (1-based)
PLATFORM_TITLE_ROW_OFFSET = 0
MAIN_HEADER_ROW_OFFSET = 2
SUB_HEADER_ROW_OFFSET = 3
DATA_START_ROW_OFFSET = 4

# Search window size for platform detection
PLATFORM_SEARCH_WINDOW = 10

# Logging configuration
LOG_LEVEL = "INFO"  # DEBUG, INFO, WARNING, ERROR, CRITICAL
LOG_FORMAT = '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
LOG_MAX_BYTES = 10 * 1024 * 1024  # 10MB
LOG_BACKUP_COUNT = 3

# Output configuration
OUTPUT_COLUMNS_BASE = [
    'Platform',
    'Funnel Stage',
    'Format',
    'Duration',
    'Aspect Ratio / Format'
]
OUTPUT_LANGUAGE_COLUMN = 'Languages'
