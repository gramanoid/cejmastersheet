"""Excel parsing utilities shared across transformer workflows."""

from __future__ import annotations

import logging
from dataclasses import dataclass
from typing import Iterable, List, Optional, Sequence, Tuple

import pandas as pd

from . import config


logger = logging.getLogger(__name__)


@dataclass(frozen=True)
class ColumnSpec:
    column_index: int
    display_name: str


@dataclass(frozen=True)
class PlatformSection:
    platform_name: str
    is_dual_language: bool
    data_row_start: int
    funnel_stage_col: int
    format_col: int
    duration_col: int
    total_col: int
    aspect_ratio_columns: Sequence[ColumnSpec]
    language_columns: Sequence[ColumnSpec]


def safe_to_numeric(value, row_idx: int, column_name: str) -> int:
    if pd.isna(value) or str(value).strip() == "":
        return 0
    try:
        return int(pd.to_numeric(value))
    except (ValueError, TypeError):
        logger.warning(
            "Row %s, Column '%s': Unable to coerce '%s' to numeric; treating as 0.",
            row_idx + 1,
            column_name,
            value,
        )
        return 0


def iter_platform_sections(df_full_sheet: pd.DataFrame, *, is_dual_language: bool) -> Iterable[PlatformSection]:
    current_row_idx = max(0, config.START_ROW_SEARCH_FOR_PLATFORM - 3)

    while current_row_idx < len(df_full_sheet):
        platform_name = None
        main_header_row_idx = -1

        for offset in range(30):
            check_row_idx = current_row_idx + offset
            if check_row_idx >= len(df_full_sheet):
                break

            row_values = df_full_sheet.iloc[check_row_idx]
            for cell_value in row_values:
                if pd.notna(cell_value) and str(cell_value).strip().lower() == config.FUNNEL_STAGE_HEADER.lower():
                    main_header_row_idx = check_row_idx
                    platform_title_row_idx = main_header_row_idx - 2
                    if platform_title_row_idx >= 0:
                        platform_cell = df_full_sheet.iloc[platform_title_row_idx, 1]
                        if pd.notna(platform_cell):
                            name_candidate = str(platform_cell).strip()
                            platform_name = _normalize_platform(name_candidate)
                    break
            if platform_name:
                break

        if not platform_name:
            current_row_idx += 10
            if current_row_idx >= len(df_full_sheet) - 5:
                break
            continue

        main_headers = df_full_sheet.iloc[main_header_row_idx].astype(str).str.strip().tolist()
        funnel_stage_col = _get_header_index(main_headers, config.FUNNEL_STAGE_HEADER)
        format_col_primary = _get_header_index(main_headers, config.FORMAT_HEADER)
        duration_col = _get_header_index(main_headers, config.DURATION_HEADER)
        total_col = _get_header_index(main_headers, config.MAIN_HEADER_TOTAL_COL)

        ar_group_info = _resolve_ar_group(main_headers, format_col_primary)
        if ar_group_info is None:
            logger.warning("%s: Unable to resolve Aspect Ratio group; skipping platform.", platform_name)
            current_row_idx = main_header_row_idx + 5
            continue

        ar_start_idx, ar_header = ar_group_info
        lang_group_start = (
            _get_header_index(main_headers, config.MAIN_HEADER_LANGUAGES_GROUP)
            if is_dual_language and config.MAIN_HEADER_LANGUAGES_GROUP in main_headers
            else None
        )

        sub_header_row_idx = main_header_row_idx + 1
        if sub_header_row_idx >= len(df_full_sheet):
            logger.warning("%s: Sub-header row out of bounds; skipping platform.", platform_name)
            current_row_idx = main_header_row_idx + 5
            continue

        sub_headers = df_full_sheet.iloc[sub_header_row_idx].astype(str).tolist()

        ar_columns = _collect_sub_headers(
            sub_headers,
            main_headers,
            start_idx=ar_start_idx,
            end_idx=lang_group_start or total_col,
            fallback_header=ar_header,
        )

        if not ar_columns:
            logger.warning("%s: No Aspect Ratio/Format columns found; skipping platform.", platform_name)
            current_row_idx = main_header_row_idx + 5
            continue

        language_columns: Sequence[ColumnSpec] = []
        if is_dual_language and lang_group_start is not None:
            language_columns = _collect_sub_headers(
                sub_headers,
                main_headers,
                start_idx=lang_group_start,
                end_idx=total_col,
                fallback_header=config.MAIN_HEADER_LANGUAGES_GROUP,
            )

        data_start = (main_header_row_idx - 2) + config.DATA_START_ROW_OFFSET

        yield PlatformSection(
            platform_name=platform_name,
            is_dual_language=is_dual_language,
            data_row_start=data_start,
            funnel_stage_col=funnel_stage_col,
            format_col=format_col_primary,
            duration_col=duration_col,
            total_col=total_col,
            aspect_ratio_columns=ar_columns,
            language_columns=language_columns,
        )

        current_row_idx = sub_header_row_idx + 3


def _normalize_platform(candidate: str) -> Optional[str]:
    normalized = candidate.strip().upper()
    # Also match against a base form without any trailing parenthetical (e.g., "LINKEDIN (EXPERT)")
    normalized_base = normalized.split(" (")[0].strip()
    for key, value in config.PLATFORM_NAMES.items():
        key_upper = key.upper()
        value_upper = value.upper()
        if normalized in {key_upper, value_upper} or normalized_base in {key_upper, value_upper}:
            return value
    logger.debug("Unknown platform label '%s'; skipping.", candidate)
    return None


def _get_header_index(headers: List[str], header_name: str) -> int:
    try:
        return headers.index(header_name)
    except ValueError as exc:
        raise ValueError(f"Header '{header_name}' not found in row: {headers}") from exc


def _resolve_ar_group(headers: List[str], primary_format_index: int) -> Optional[Tuple[int, str]]:
    if config.MAIN_HEADER_ASPECT_RATIO_GROUP_PRIMARY in headers:
        return headers.index(config.MAIN_HEADER_ASPECT_RATIO_GROUP_PRIMARY), config.MAIN_HEADER_ASPECT_RATIO_GROUP_PRIMARY

    indices = [idx for idx, value in enumerate(headers) if value == config.MAIN_HEADER_ASPECT_RATIO_GROUP_SECONDARY]
    for idx in indices:
        if idx != primary_format_index and idx > primary_format_index:
            return idx, config.MAIN_HEADER_ASPECT_RATIO_GROUP_SECONDARY
    return None


def _collect_sub_headers(
    sub_headers: Sequence[str],
    main_headers: Sequence[str],
    *,
    start_idx: int,
    end_idx: int,
    fallback_header: str,
) -> List[ColumnSpec]:
    collected: List[ColumnSpec] = []
    for idx in range(start_idx, end_idx):
        sub_value = str(sub_headers[idx]).strip()
        main_value = str(main_headers[idx]).strip()
        if not sub_value or sub_value.lower() == "nan":
            continue
        if main_value and main_value not in {fallback_header, "nan", ""} and idx > start_idx:
            break
        collected.append(ColumnSpec(column_index=idx, display_name=sub_value))
    return collected
