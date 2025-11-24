# Excel Transformation Plan for 'Tracker (Dual Lang)' Sheet

This document outlines the plan for the Python script `excel_transformer.py` to process the 'Tracker (Dual Lang)' sheet from the `Haleon CEJ Master Spec Sheet 3.1.xlsx` file and generate a structured output Excel file.

## 1. Setup

*   **Libraries:** `pandas`, `datetime`, `os`, `logging`.
*   **Constants:**
    *   Input Excel file path: `Haleon CEJ Master Spec Sheet 3.1.xlsx`
    *   Sheet name: `Tracker (Dual Lang)`
    *   Known platform names (case-insensitive search): YouTube, META, TikTok, LinkedIn, Programmatic, Audio, Gaming, Amazon. (LinkedIn added for the 3.2+ templates where the platform may appear as “LinkedIn” or “LinkedIn (Expert)”.)
    *   Core main header names for identification: "Funnel Stage", "Format" (primary one, not aspect ratio type), "Duration", "Aspect Ratio", "Languages", "TOTAL".
*   **Logging:** Setup logging to both a file (e.g., `transformer.log`) and the console.

## 2. Load Data

*   Read the entire 'Tracker (Dual Lang)' sheet into a pandas DataFrame without assuming any specific header row initially (e.g., `pd.read_excel(..., header=None)`). This allows for flexible header detection.

## 3. Iterate and Find Platform Tables

*   Start scanning rows from a configurable row index (e.g., around row 8 or 9, 0-indexed, which is Excel row 9 or 10), allowing for some flexibility (a few rows up/down).
*   In each scanned row, look for platform names in the first few columns (e.g., first 3 columns).
*   **Table Structure Assumption (0-indexed based on `header=None` load):**
    *   If Platform Name is found on Row `P`.
    *   Main Headers (Funnel Stage, etc.) are expected on Row `P+2`.
    *   Sub-Headers (specific aspect ratios, specific languages) are expected on Row `P+3`.
    *   Data for that platform table starts on Row `P+4`.

## 4. For Each Identified Platform Table

*   **Extract Platform Name:** Store the identified platform name.
*   **Verify and Parse Headers:**
    *   Check for the presence of core main headers ("Funnel Stage", "Aspect Ratio", "Languages", "TOTAL") on the expected main header row (Row `P+2`). If not found, log an error, skip this table, and continue searching for the next platform.
    *   From the sub-header row (Row `P+3`):
        *   Dynamically identify the column indices and names for **Aspect Ratio Types**. These are the columns under the "Aspect Ratio" main header group, up to where the "Languages" main header group begins.
        *   Dynamically identify the column indices and names for **Language Types**. These are the columns under the "Languages" main header group, up to where the "TOTAL" main header begins.
*   **Process Data Rows:**
    *   Iterate through data rows for the current platform, starting from Row `P+4`.
    *   A data row is considered the end of the table if it's completely empty (all cells NaN or blank).
    *   For each data row:
        1.  Extract values for `Funnel Stage`, `Format` (primary), `Duration` from their respective columns (identified by the main headers).
        2.  Initialize `SumOfAspectRatioVariations = 0`.
        3.  Create a list of `selected_aspect_ratios_details` (e.g., `[{'name': 'IMG', 'count': 1}, {'name': 'Video', 'count': 5}]`). Iterate through the identified aspect ratio type columns:
            *   Attempt to convert cell value to a number. If non-numeric or < 0, treat as 0.
            *   If numeric cell value > 0, add this value to `SumOfAspectRatioVariations` and add its details (sub-header name and this count) to `selected_aspect_ratios_details`.
        4.  If `SumOfAspectRatioVariations == 0`, this data row generates no output; skip to the next data row.
        5.  Create a list of `selected_languages` (e.g., `['EN', 'AR']`). Iterate through the identified language type columns:
            *   Attempt to convert cell value to a number. If non-numeric or < 0, treat as 0.
            *   If numeric cell value > 0 and it strictly equals `SumOfAspectRatioVariations`, add the language name (its sub-header) to `selected_languages`.
        6.  **Validation (TOTAL column):**
            *   Calculate `ExpectedTotal = SumOfAspectRatioVariations * len(selected_languages)`.
            *   Compare `ExpectedTotal` with the numeric value in the `TOTAL` column for the current data row. Log any discrepancies.
        7.  **Generate Output Rows:**
            *   For each aspect ratio detail in `selected_aspect_ratios_details` (e.g., `{'name': 'Video', 'count': 5}`):
                *   For each language name in `selected_languages` (e.g., `'EN'`):
                    *   Create one output dictionary: `{'Platform': current_platform_name, 'Funnel Stage': funnel_stage_val, 'Format': format_val, 'Duration': duration_val, 'Aspect Ratio': aspect_ratio_detail['name'], 'Languages': language_name}`.
                    *   Append this dictionary to a global list that accumulates all transformed rows.
        8.  **Error Handling during row processing:** If non-numeric data is found where numbers are crucial (aspect ratio counts, language counts, TOTAL), log the specific cell/row issue. If a critical value for parsing the row (like aspect ratio counts) cannot be determined, skip generating output for that specific data row and log it, then proceed to the next data row.

## 5. Finalize Output

*   After processing all platform tables, convert the global list of output dictionaries into a pandas DataFrame.
*   Generate the output Excel filename with a timestamp: `transformed_CEJ_master_specsheet_YYYYMMDD_HHMMSS.xlsx` (e.g., `transformed_CEJ_master_specsheet_20250515_164000.xlsx`).
*   Save this DataFrame to the new Excel file. Ensure the DataFrame index is not written to the Excel file.

## 6. Executable Creation and Distribution

*   **Objective:** Create a user-friendly, standalone executable for non-technical users.
*   **Tool Used:** PyInstaller version 6.13.0.
*   **Final Build Mode:** "One-folder" mode was chosen. While this results in multiple files in a directory, it offers faster startup compared to "one-file" mode, which is beneficial given the pandas library's initialization time.
    *   The output is located in the `dist/ExcelTransformer` folder (relative to the project root).
*   **Key PyInstaller Options Used:**
    *   `--windowed`: Ensures no console window appears for the GUI application.
    *   `--name ExcelTransformer`: Sets the name of the output executable and folder.
    *   `--upx-dir .`: Specifies the current directory for `upx.exe`, enabling compression of binaries within the output folder.
    *   `--splash splash_screenv2.png`: Displays a splash screen image immediately on launch to improve perceived startup performance.
*   **Final Command:** `pyinstaller --windowed --name ExcelTransformer --upx-dir . --splash splash_screenv2.png excel_transformer.py`
*   **Distribution:**
    1.  The entire `dist/ExcelTransformer` folder should be zipped (e.g., `ExcelTransformer_App.zip`).
    2.  This zip file is then distributed to users.
    3.  Users must unzip the file and then run `ExcelTransformer.exe` from within the created `ExcelTransformer` folder.
*   **User Experience Notes:**
    *   A startup time of approximately 13 seconds was observed (from double-click to file dialog appearing). The splash screen helps mitigate the perception of this delay.
    *   All dialogs (file selection, messages) are configured to appear in the foreground.

## Key Assumptions for Modularity:

*   The relative positions of Platform Title, Main Headers, Sub-Headers, and Data Start are consistent.
*   The main header group names ("Aspect Ratio", "Languages", "TOTAL") are consistent and can be used to dynamically find the columns for specific aspect ratio types and languages.
*   Completely empty rows delimit tables or signify the end of data.

---

## 7. Toggle Controls and Presence Reporting (Added 2025‑08‑08)

Configuration toggles are defined in `config.py` and used by `excel_transformer.py`.

- EXPAND_ALL_TO_ACP: bool
  - When True, Funnel Stage "ALL" is expanded into Awareness, Consideration, Purchase during emission.
  - Default: True (stakeholder directive: “all platforms should cover all funnel stages”).

- FUNNEL_STAGES: list[str]
  - Ordered funnel stages used for expansion: ["Awareness", "Consideration", "Purchase"].

- PRESENCE_REPORTING_MODE: enum string
  - Values: "off" | "qa_only" | "emit_rows"
  - "off": No extra presence diagnostics beyond logs.
  - "qa_only": Generate coverage diagnostics (to be written by validation/QA module), but do not emit placeholder rows.
  - "emit_rows": Emit presence rows to a separate sheet (not mixed into transformed output). Sheet names:
    - Presence_Dual (Dual)
    - Presence_Single (Single)
  - Default: "qa_only".

Behavioral notes:
- TOTAL validation occurs before ANY expansion; only rows that pass original sheet totals will be expanded.
- Expansion multiplies stage representation only; it does not fabricate AR/Lang combinations beyond what is selected in source rows.

## 8. Validation, Coverage & QA Artifacts (Planned)

- Enhance `validation_script.py` to:
  - Robustly handle column naming/indexing (avoid KeyErrors).
  - Produce a Coverage Report per sheet:
    - Platforms present vs. emitted.
    - Stages seen vs. emitted; primary reasons for 0 emissions (e.g., “no AR ticks”).
  - Respect EXPAND_ALL_TO_ACP in comparisons (treat "ALL" as A/C/P when enabled).
  - Save consolidated `QA_Report.xlsx` with:
    - Summary sheet (pass/fail by platform/stage).
    - Per-platform details (rows_scanned, rows_no_ar_ticks, emitted_count, mismatches).

## 9. Streamlit UI (Planned)

- Show a “Platform Coverage” panel with badges per platform/stage and warnings for 0 emissions.
- Provide downloads for transformed output and QA_Report.xlsx.
- Display a banner when ALL→A/C/P expansion is enabled.

## 10. Runbook & WSL Note

- Development run (Linux/WSL): `python3 excel_transformer.py` or `streamlit run streamlit_app.py`.
- Logs: `transformer.log`.
- Typical input: `SSD_F&D_EG B2_V1_Haleon CEJ Master Spec Sheet 3.2_Master Template_25Jun.xlsx`.
- Use `python3` (not `python`) on WSL/Linux environments.
