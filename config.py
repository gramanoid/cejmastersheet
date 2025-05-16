"""
Configuration settings for the Excel Transformer application.
"""
from typing import Dict, List, Tuple, Optional
from pathlib import Path

# Application constants
APP_NAME = "Excel Transformer"
VERSION = "2.0.0"

# File handling
INPUT_SHEET_NAME = 'Tracker (Dual Lang)'
OUTPUT_FILE_BASENAME = 'transformed_CEJ_master_specsheet'
LOG_FILE = 'transformer.log'
DEFAULT_OUTPUT_FORMAT = 'xlsx'  # Can be 'xlsx' or 'csv'

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
MAIN_HEADER_LANGUAGES_GROUP = "Languages"  # Original name, kept for broader compatibility if needed
LANGUAGE_HEADER = "Languages"            # Corrected name to match streamlit_app.py usage
TOTAL_HEADER = "TOTAL" # Renamed from MAIN_HEADER_TOTAL_COL
PLATFORM_COLUMN_HEADER = "PLATFORM" # Header name in the first column to identify platform sections
FUNNEL_STAGE_HEADER = "Funnel Stage" # Header name for the funnel stage column
FORMAT_HEADER = "Format"             # Header name for the format column
DURATION_HEADER = "Duration"           # Header name for the duration column

# Core headers expected in the main header row for validation
CORE_MAIN_HEADERS = [
    FUNNEL_STAGE_HEADER,
    FORMAT_HEADER,
    DURATION_HEADER,
    LANGUAGE_HEADER, # Using the corrected LANGUAGE_HEADER
    TOTAL_HEADER
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
OUTPUT_COLUMNS = [
    'Platform',
    'Funnel Stage',
    'Format',
    'Duration',
    'Aspect Ratio / Format',
    'Languages'
]
