# Identified Issues and Fixes

This document tracks issues encountered during development and their resolutions.

## PyInstaller Executable Creation

*   **Issue:** Initial PyInstaller commands with various options (e.g., `--upx`, `--icon=NONE`) failed with generic usage errors or specific flag errors.
    *   **Fix 1:** Removed the `--icon=NONE` flag as it was causing errors. PyInstaller uses a default icon if none is specified and valid.
    *   **Fix 2:** Ensured `upx.exe` was in the project root and used `pyinstaller --upx-dir .` to explicitly tell PyInstaller where to find UPX. This resolved issues where simply using `--upx` failed.
    *   **Fix 3:** Implemented clean build steps by deleting `build/`, `dist/`, and `*.spec` files before critical build attempts. Ensured correct commands for directory removal in PowerShell (`Remove-Item -Recurse -Force`) vs. CMD (`rd /s /q`).

*   **Issue:** Long startup time for the one-file executable due to decompression and pandas/library initialization.
    *   **Fix/Mitigation 1:** Switched to "one-folder" build mode (`pyinstaller --windowed --name ExcelTransformer --upx-dir . excel_transformer.py`). While pandas initialization still takes time (~13 seconds), this mode avoids the initial large decompression step of one-file mode.
    *   **Fix/Mitigation 2:** Implemented a splash screen (`--splash splash_screenv2.png`) to provide immediate visual feedback to the user during the startup period, improving perceived performance.

*   **Issue:** Default PyInstaller console window appearing for a GUI application.
    *   **Fix:** Consistently used the `--windowed` (or `-w`) flag in PyInstaller commands to suppress the console window.

## Issue: Incorrect Aspect Ratio/Format Group Header Identification with Multiple 'Format' Columns (Streamlit App)

*   **Symptom:** The Streamlit application was not generating the correct number of unique creative combinations (expected 240). Platforms like PROGRAMMATIC, AUDIO, GAMING, and AMAZON, which use 'Format' as a fallback for the Aspect Ratio group header and have two 'Format' columns in their headers, were not being processed correctly. The `ar_cols_start_idx` was often picking the first 'Format' column instead of the second one intended for aspect ratios/sub-formats.
*   **Root Cause:** The logic for finding `ar_cols_start_idx` in `streamlit_app.py` did not adequately handle cases where the `AR_GROUP_HEADER` was 'Format' and multiple columns with this name existed. It tended to pick the first instance, which was the primary 'Format' column, not the one representing different aspect ratios or sub-formats.
*   **Fix Applied (streamlit_app.py):** 
    *   Refactored the logic within the `process_uploaded_file` function, specifically the section determining `ar_cols_start_idx`.
    *   When `MAIN_HEADER_ASPECT_RATIO_GROUP_PRIMARY` ('Aspect Ratio') is not found and `MAIN_HEADER_ASPECT_RATIO_GROUP_SECONDARY` ('Format') is used as the `AR_GROUP_HEADER`:
        *   The code now iterates through the `platform_main_header_row` to find all occurrences of 'Format'.
        *   If multiple 'Format' columns are found, and the primary `format_col_idx` (for the main creative format) is one of them, the `ar_cols_start_idx` is set to the index of the *next* 'Format' column after the primary `format_col_idx`.
        *   This ensures that for platforms like PROGRAMMATIC, the second 'Format' column (e.g., index 4 when primary format is at index 2) is correctly chosen as the start of the aspect ratio/sub-format group.
*   **Verification:** Console logs confirmed that for platforms with dual 'Format' headers, `ar_cols_start_idx` was correctly set to the second 'Format' column's index. The application successfully generated 240 unique combinations, indicating all platforms were processed as expected.
*   **Status:** Resolved.
