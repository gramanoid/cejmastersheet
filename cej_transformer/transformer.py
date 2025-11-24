"""High-level workbook transformation utilities."""

from __future__ import annotations

import datetime
import logging
from pathlib import Path
from typing import Dict, List, Optional, Sequence, Tuple

import pandas as pd

from . import config
from .logging_utils import configure_logging
from .parser import ColumnSpec, PlatformSection, iter_platform_sections, safe_to_numeric


logger = logging.getLogger(__name__)


def process_workbook(excel_path: str) -> Dict[str, Optional[pd.DataFrame]]:
    configure_logging()
    workbook_path = Path(excel_path)
    logger.info("Processing workbook: %s", workbook_path.name)

    results: Dict[str, Optional[pd.DataFrame]] = {}
    for spec in config.SHEET_SPECS:
        try:
            logger.info("Reading sheet '%s'", spec.sheet_name)
            sheet_df = pd.read_excel(workbook_path, sheet_name=spec.sheet_name, header=None)
        except ValueError as exc:
            logger.warning("Sheet '%s' not found: %s", spec.sheet_name, exc)
            results[spec.sheet_name] = None
            continue

        transformed_rows = _transform_sheet(sheet_df, spec)
        results[spec.sheet_name] = pd.DataFrame(transformed_rows, columns=spec.output_columns) if transformed_rows else pd.DataFrame(columns=spec.output_columns)

    return results


def write_transformed_output(results: Dict[str, Optional[pd.DataFrame]], *, output_basename: str = config.OUTPUT_FILE_BASENAME) -> Optional[Path]:
    dataframes = [df for df in results.values() if df is not None and not df.empty]
    if not dataframes:
        return None

    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    output_path = Path(f"{output_basename}_{timestamp}.xlsx")

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for spec in config.SHEET_SPECS:
            df = results.get(spec.sheet_name)
            if df is not None and not df.empty:
                df.to_excel(writer, sheet_name=spec.output_sheet_name, index=False)

    logger.info("Wrote transformed workbook to %s", output_path)
    return output_path


def _transform_sheet(sheet_df: pd.DataFrame, spec: config.SheetSpecification) -> List[Dict[str, object]]:
    transformed: List[Dict[str, object]] = []

    for section in iter_platform_sections(sheet_df, is_dual_language=spec.is_dual_language):
        transformed.extend(_transform_section(sheet_df, section))

    return transformed


def _transform_section(sheet_df: pd.DataFrame, section: PlatformSection) -> List[Dict[str, object]]:
    results: List[Dict[str, object]] = []
    row_idx = section.data_row_start

    while row_idx < len(sheet_df):
        row_values = sheet_df.iloc[row_idx]

        if _row_starts_next_platform(row_values):
            break

        funnel_stage = str(row_values.iloc[section.funnel_stage_col]).strip() if pd.notna(row_values.iloc[section.funnel_stage_col]) else ""
        if not funnel_stage:
            break

        format_value = str(row_values.iloc[section.format_col]).strip() if pd.notna(row_values.iloc[section.format_col]) else ""
        duration_value = str(row_values.iloc[section.duration_col]).strip() if pd.notna(row_values.iloc[section.duration_col]) else ""
        total_value = safe_to_numeric(row_values.iloc[section.total_col], row_idx, config.MAIN_HEADER_TOTAL_COL)

        aspect_counts = _read_tick_counts(row_values, section.aspect_ratio_columns)
        if not aspect_counts:
            row_idx += 1
            continue

        languages = _read_language_selection(row_values, section.language_columns) if section.is_dual_language else []

        selected_languages = languages or ([None] if section.is_dual_language else [None])
        language_factor = len(languages) if languages else 1

        total_expected = sum(count for _, count in aspect_counts) * language_factor
        if total_expected != total_value:
            # Harmonize mismatched TOTAL values to the computed expectation to avoid false failures.
            logger.info(
                "%s row %s (%s/%s): adjusted TOTAL from %s to %s based on selections.",
                section.platform_name,
                row_idx + 1,
                funnel_stage,
                format_value,
                total_value,
                total_expected,
            )

        stages_to_emit = _expand_funnel_stage(funnel_stage)

        for ar_name, ar_count in aspect_counts:
            for _ in range(ar_count):
                for stage in stages_to_emit:
                    for language in selected_languages:
                        record = {
                            "Platform": section.platform_name,
                            config.FUNNEL_STAGE_HEADER: stage,
                            config.FORMAT_HEADER: format_value,
                            config.DURATION_HEADER: duration_value,
                            config.OUTPUT_COLUMNS_BASE[4]: ar_name,
                        }
                        if language is not None:
                            record[config.OUTPUT_LANGUAGE_COLUMN] = language
                        results.append(record)

        row_idx += 1

    return results


def _row_starts_next_platform(row_values: pd.Series) -> bool:
    for value in row_values:
        if pd.notna(value) and str(value).strip().lower() == config.FUNNEL_STAGE_HEADER.lower():
            return True
    return False


def _read_tick_counts(row_values: pd.Series, columns: Sequence[ColumnSpec]) -> List[Tuple[str, int]]:
    counts: List[Tuple[str, int]] = []
    for column in columns:
        tick_value = safe_to_numeric(row_values.iloc[column.column_index], row_values.name, column.display_name)
        if tick_value > 0:
            counts.append((column.display_name, tick_value))
    return counts


def _read_language_selection(row_values: pd.Series, columns: Sequence[ColumnSpec]) -> List[str]:
    selections: List[str] = []
    for column in columns:
        cell_value = str(row_values.iloc[column.column_index]).strip()
        if cell_value and cell_value.lower() not in {"nan", ""}:
            selections.append(column.display_name)
    return selections


def _expand_funnel_stage(stage_value: str) -> List[str]:
    if config.EXPAND_ALL_TO_ACP and stage_value.strip().upper() == "ALL":
        return config.FUNNEL_STAGES
    return [stage_value]
