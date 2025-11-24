"""Validation Script for CEJ Master Spec Sheet Transformer."""

from __future__ import annotations

import json
import logging
import os
import sys
from datetime import datetime

from cej_transformer.logging_utils import configure_logging
from cej_transformer.validator import validate_output


def main() -> None:
    if len(sys.argv) != 3:
        print("Usage: python validation_script.py <input_excel_file> <output_excel_file>")
        sys.exit(1)

    input_file, output_file = sys.argv[1:3]
    if not os.path.exists(input_file):
        print(f"Error: Input file '{input_file}' not found")
        sys.exit(1)
    if not os.path.exists(output_file):
        print(f"Error: Output file '{output_file}' not found")
        sys.exit(1)

    log_filename = f"validation_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
    configure_logging(log_file=log_filename)

    try:
        report = validate_output(input_file, output_file)
    except Exception as exc:  # pragma: no cover - defensive logging
        logging.exception("Validation failed: %s", exc)
        print(f"Validation failed with error: {exc}")
        sys.exit(1)

    _print_validation_report(report)

    report_filename = f"validation_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
    with open(report_filename, "w", encoding="utf-8") as json_file:
        json.dump(report, json_file, indent=2)

    print(f"\nDetailed validation report saved to: {report_filename}")
    print(f"Validation log saved to: {log_filename}")
    sys.exit(0 if report["summary"]["overall_status"] == "PASS" else 1)


def quick_validate(input_file: str, output_file: str) -> bool:
    try:
        report = validate_output(input_file, output_file)
        return report["summary"]["overall_status"] == "PASS"
    except Exception:
        return False


def _print_validation_report(report):
    print("\n" + "=" * 80)
    print("VALIDATION REPORT")
    print("=" * 80)
    print(f"Timestamp: {report['timestamp']}")
    print(f"Input File: {report['input_file']}")
    print(f"Output File: {report['output_file']}")
    print(f"Overall Status: {report['summary']['overall_status']}")
    print(f"Total Platforms Checked: {report['summary']['total_platforms_checked']}")
    print(f"Platforms Passed: {report['summary']['platforms_passed']}")
    print(f"Platforms Failed: {report['summary']['platforms_failed']}")

    print("\nDETAILED RESULTS:")
    print("-" * 80)
    for sheet_name, platforms in report["validation_results"].items():
        print(f"\n{sheet_name}:")
        for platform_name, details in platforms.items():
            status_symbol = "✅" if details["status"] == "PASS" else "❌"
            print(f"  {status_symbol} {platform_name}")
            print(f"     Expected: {details['expected_count']}")
            print(f"     Actual:   {details['actual_count']}")
            if details["difference"]:
                print(f"     Difference: {details['difference']:+d}")

    print("\n" + "=" * 80)


if __name__ == "__main__":
    main()