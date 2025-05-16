# Improvement Plan

## Features & Improvements

*   **[✅] Create User-Friendly Executable:**
    *   **Status:** Completed.
    *   **Details:** Packaged `excel_transformer.py` into a standalone executable using PyInstaller. The final version is a "one-folder" build with UPX compression and a splash screen. Users can run `ExcelTransformer.exe` from the `dist/ExcelTransformer` directory.

*   **[✅] Debugging Streamlit Excel Tool**
    *   **Description:** Address issues in the Streamlit application (`streamlit_app.py`) to ensure it correctly identifies and processes headers (Funnel Stage, Format, Duration, TOTAL, Aspect Ratio/Format group) from the 'Tracker (Dual Lang)' sheet of the `Haleon CEJ Master Spec Sheet 3.1.xlsx`.
    *   **Goal:** Achieve consistent generation of 240 unique creative combinations and accurate validation against the TOTAL column.
    *   **Key Sub-Tasks:**
        *   ✅ Refine logic for `funnel_stage_col_idx`, `format_col_idx`, `duration_col_idx`, `total_col_idx`.
        *   ✅ Enhance logic for `ar_cols_start_idx` (Aspect Ratio/Format group header), especially for platforms with multiple 'Format' columns (PROGRAMMATIC, AUDIO, GAMING, AMAZON).
        *   ⏳ Ensure correct validation against the TOTAL column (implicitly tested with 240 combinations, but can be explicitly reviewed if needed).
        *   ✅ Test with various input files (primary focus on `Haleon CEJ Master Spec Sheet 3.1.xlsx`).
    *   **Status:** ✅ **Completed** (Core issue of 240 combinations and header identification resolved as of 2025-05-16)
    *   **Priority:** High
    *   **Effort:** Medium
    *   **Dependencies:** `excel_transformer.py` (core logic), `config.py` (constants)

*   **(Future Considerations - Optional)**
    *   **[ ] Advanced Logging GUI:** If users need to see logs, consider a simple scrolled text widget within the Tkinter app itself instead of separate log files for very non-technical users.
    *   **[ ] Configuration File:** For settings like known platform names, input sheet name (if it might change), allow configuration via a simple text/JSON file rather than hardcoding, if requirements evolve.
