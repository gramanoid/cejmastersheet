from __future__ import annotations

import logging
import os

try:
    import tkinter as tk
    from tkinter import filedialog, messagebox

    _TK_AVAILABLE = True
except Exception:  # pragma: no cover - optional dependency
    tk = None
    filedialog = None
    messagebox = None
    _TK_AVAILABLE = False

from cej_transformer.logging_utils import configure_logging
from cej_transformer.transformer import process_workbook, write_transformed_output

logger = logging.getLogger(__name__)


def select_excel_file() -> str | None:
    if _TK_AVAILABLE:
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        file_path = filedialog.askopenfilename(
            parent=root,
            title="Select the Excel file to process",
            filetypes=(("Excel files", "*.xlsx *.xls"), ("All files", "*.*")),
        )
        root.destroy()
        return file_path

    print("Tkinter is unavailable. Enter the full path to the Excel file:")
    return input().strip() or None


def process_excel_file_for_streamlit(input_excel_file_path: str):
    configure_logging()
    return process_workbook(input_excel_file_path)


def main() -> None:
    configure_logging()

    input_path = select_excel_file()
    if not input_path:
        logger.info("No file selected; exiting.")
        if _TK_AVAILABLE:
            _show_message("Cancelled", "No file selected. Exiting script.")
        return

    logger.info("Starting transformation for %s", os.path.basename(input_path))
    results = process_workbook(input_path)
    output_path = write_transformed_output(results)

    if output_path is None:
        logger.info("No transformed data generated; skipping output file creation.")
        if _TK_AVAILABLE:
            _show_message("No Data", "No data transformed. Output not generated.", severity="warning")
        return

    logger.info("Transformation complete: %s", output_path)
    if _TK_AVAILABLE:
        _show_message("Success", f"Data written to:\n{output_path}")


def _show_message(title: str, message: str, *, severity: str = "info") -> None:
    if not _TK_AVAILABLE:
        return
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    if severity == "warning":
        messagebox.showwarning(title, message, parent=root)
    else:
        messagebox.showinfo(title, message, parent=root)
    root.destroy()


if __name__ == "__main__":
    main()
