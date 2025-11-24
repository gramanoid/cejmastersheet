"""Validation helpers comparing input expectations against generated output."""

from __future__ import annotations

import logging
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Dict, Optional

import pandas as pd

from . import config
from .parser import PlatformSection, iter_platform_sections, safe_to_numeric


logger = logging.getLogger(__name__)


@dataclass
class PlatformComparison:
    expected: int
    actual: int

    @property
    def status(self) -> str:
        return "PASS" if self.expected == self.actual else "FAIL"

    @property
    def difference(self) -> int:
        return self.actual - self.expected


def validate_output(input_path: str, output_path: str) -> Dict:
    input_file = Path(input_path)
    output_file = Path(output_path)

    expected = _collect_expected_totals(input_file)
    actual = _collect_actual_counts(output_file)

    comparison: Dict[str, Dict[str, PlatformComparison]] = {}
    summary = {
        "total_platforms_checked": 0,
        "platforms_passed": 0,
        "platforms_failed": 0,
        "overall_status": "UNKNOWN",
    }

    for spec in config.SHEET_SPECS:
        spec_expected = expected.get(spec.sheet_name, {})
        spec_actual = actual.get(spec.output_sheet_name, {})
        sheet_results: Dict[str, PlatformComparison] = {}

        for platform_name, expected_count in spec_expected.items():
            actual_count = spec_actual.get(platform_name, 0)
            comparison_entry = PlatformComparison(expected=expected_count, actual=actual_count)
            sheet_results[platform_name] = comparison_entry

            summary["total_platforms_checked"] += 1
            if comparison_entry.status == "PASS":
                summary["platforms_passed"] += 1
            else:
                summary["platforms_failed"] += 1

            logger.info(
                "%s -> %s: expected %s, actual %s (%s)",
                spec.sheet_name,
                platform_name,
                expected_count,
                actual_count,
                comparison_entry.status,
            )

        comparison[spec.sheet_name] = sheet_results

    summary["overall_status"] = "PASS" if summary["platforms_failed"] == 0 else "FAIL"

    return {
        "timestamp": datetime.now().isoformat(),
        "input_file": input_file.name,
        "output_file": output_file.name,
        "validation_results": {
            sheet: {
                platform: {
                    "expected_count": result.expected,
                    "actual_count": result.actual,
                    "difference": result.difference,
                    "status": result.status,
                }
                for platform, result in sheet_results.items()
            }
            for sheet, sheet_results in comparison.items()
        },
        "summary": summary,
    }


def _collect_expected_totals(input_file: Path) -> Dict[str, Dict[str, int]]:
    workbook: Dict[str, Dict[str, int]] = {}
    for spec in config.SHEET_SPECS:
        try:
            sheet_df = pd.read_excel(input_file, sheet_name=spec.sheet_name, header=None)
        except ValueError:
            logger.warning("Sheet '%s' not found in input workbook.", spec.sheet_name)
            continue

        totals: Dict[str, int] = {}
        for section in iter_platform_sections(sheet_df, is_dual_language=spec.is_dual_language):
            totals[section.platform_name] = _sum_totals(sheet_df, section)
        workbook[spec.sheet_name] = totals

    return workbook


def _collect_actual_counts(output_file: Path) -> Dict[str, Dict[str, int]]:
    try:
        excel_file = pd.ExcelFile(output_file)
    except Exception as exc:
        logger.error("Unable to read output file '%s': %s", output_file, exc)
        return {}

    counts: Dict[str, Dict[str, int]] = {}
    for sheet_name in excel_file.sheet_names:
        df = pd.read_excel(output_file, sheet_name=sheet_name)
        if "Platform" in df.columns:
            counts[sheet_name] = df["Platform"].value_counts().to_dict()
        else:
            counts[sheet_name] = {}
    return counts


def _sum_totals(sheet_df: pd.DataFrame, section: PlatformSection) -> int:
    total = 0
    row_idx = section.data_row_start

    while row_idx < len(sheet_df):
        row_values = sheet_df.iloc[row_idx]

        if any(
            pd.notna(value) and str(value).strip().lower() == config.FUNNEL_STAGE_HEADER.lower()
            for value in row_values
        ):
            break

        funnel_value = row_values.iloc[section.funnel_stage_col]
        if pd.isna(funnel_value) or str(funnel_value).strip() == "":
            break

        total += safe_to_numeric(row_values.iloc[section.total_col], row_idx, config.MAIN_HEADER_TOTAL_COL)
        row_idx += 1

    return total
