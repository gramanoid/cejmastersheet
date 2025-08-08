import pandas as pd
import os
import datetime
import logging
import re

# Import from config
from config import (
    DUAL_LANG_INPUT_SHEET_NAME, SINGLE_LANG_INPUT_SHEET_NAME,
    OUTPUT_FILE_BASENAME, LOG_FILE, PLATFORM_NAMES,
    MAIN_HEADER_ASPECT_RATIO_GROUP_PRIMARY, MAIN_HEADER_ASPECT_RATIO_GROUP_SECONDARY,
    MAIN_HEADER_LANGUAGES_GROUP, MAIN_HEADER_TOTAL_COL,
    START_ROW_SEARCH_FOR_PLATFORM, PLATFORM_TITLE_ROW_OFFSET,
    MAIN_HEADER_ROW_OFFSET, SUB_HEADER_ROW_OFFSET, DATA_START_ROW_OFFSET,
    LOG_LEVEL, LOG_FORMAT, LOG_MAX_BYTES, LOG_BACKUP_COUNT, # Assuming these are used by setup_logging
    OUTPUT_COLUMNS_BASE, OUTPUT_LANGUAGE_COLUMN, # For output structure
    FUNNEL_STAGE_HEADER, FORMAT_HEADER, DURATION_HEADER, # Specific header names from config
    OUTPUT_SHEET_NAME_DUAL_LANG, OUTPUT_SHEET_NAME_SINGLE_LANG, # Output sheet names
    FUNNEL_STAGES, EXPAND_ALL_TO_ACP # Funnel expansion config
)

# Tkinter is optional (GUI-only). Gracefully degrade if not available (e.g., on Streamlit Cloud)
try:
    import tkinter as tk
    from tkinter import filedialog, messagebox
    _TK_AVAILABLE = True
except Exception:  # ImportError or _tkinter errors
    tk = None
    filedialog = None
    messagebox = None
    _TK_AVAILABLE = False

def setup_logging():
    # If root logger already has handlers, assume it's configured (e.g., by Streamlit)
    if logging.root.handlers:
        logging.info("Root logger already has handlers. excel_transformer's setup_logging will not reconfigure.")
        # Optionally, add a file handler if it doesn't exist, for standalone runs,
        # but be careful not to duplicate if Streamlit adds one.
        # For now, let's keep it simple: if handlers exist, do nothing more from here.
        return

    # Original setup for standalone execution:
    # Remove existing handlers (should be none if the above check passes for first-time standalone)
    for handler in logging.root.handlers[:]: # This line will now only run if no handlers were present
        logging.root.removeHandler(handler)
    logging.basicConfig(filename=LOG_FILE,
                        level=LOG_LEVEL, 
                        format=LOG_FORMAT,
                        filemode='w') # Overwrite log file each run
    console_handler = logging.StreamHandler()
    console_handler.setLevel(LOG_LEVEL)
    formatter = logging.Formatter(LOG_FORMAT)
    console_handler.setFormatter(formatter)
    logging.getLogger().addHandler(console_handler)
    logging.info("excel_transformer.setup_logging: Configured for standalone run.")

def select_excel_file():
    """Opens a dialog for the user to select an Excel file.
    Falls back to console prompt if Tkinter GUI is unavailable."""
    if _TK_AVAILABLE:
        root = tk.Tk()
        root.withdraw()  # Hide main window
        root.attributes('-topmost', True)
        file_path = filedialog.askopenfilename(
            parent=root,
            title="Select the Excel file to process",
            filetypes=(("Excel files", "*.xlsx *.xls"), ("All files", "*.*"))
        )
        root.attributes('-topmost', False)
        root.destroy()
        return file_path
    else:
        # Headless mode: ask via stdin or environment
        print("Tkinter is not available in this environment. Please enter the full path to the Excel file:")
        return input().strip() or None

def safe_to_numeric(value, context_row_idx, context_col_name):
    """Attempts to convert value to numeric, handling potential errors and non-numeric strings."""
    if pd.isna(value) or str(value).strip() == '':
        return 0
    try:
        return pd.to_numeric(value)
    except ValueError:
        logging.warning(f"Row {context_row_idx+1}, Col '{context_col_name}': Could not convert '{value}' to numeric. Treating as 0.")
        return 0

def find_platform_tables_and_transform(df_full_sheet, is_dual_lang: bool):
    transformed_data_all_platforms = []
    current_row_idx = START_ROW_SEARCH_FOR_PLATFORM - 3 # Start search a bit earlier
    if current_row_idx < 0: current_row_idx = 0
    
    logging.info(f"Starting platform search for {'Dual Lang' if is_dual_lang else 'Single Lang'} sheet from row {current_row_idx + 1}")
    logging.info(f"Looking for platforms: {list(PLATFORM_NAMES.values())}")

    # Define core main headers dynamically based on sheet type
    core_main_headers_check = [FUNNEL_STAGE_HEADER, FORMAT_HEADER, DURATION_HEADER, MAIN_HEADER_TOTAL_COL]
    if is_dual_lang:
        # Insert Languages group header before TOTAL for dual language sheets
        core_main_headers_check.insert(3, MAIN_HEADER_LANGUAGES_GROUP)

    while current_row_idx < len(df_full_sheet):
        platform_name_found = None
        platform_header_row_actual_idx = -1
        main_header_row_actual_idx = -1

        # Search for "Funnel Stage" in the next rows
        for search_offset in range(30): # Larger window to find main headers
            check_row_idx = current_row_idx + search_offset
            if check_row_idx >= len(df_full_sheet):
                break
            
            # Check all columns for "Funnel Stage"
            row_values = df_full_sheet.iloc[check_row_idx]
            for col_idx, cell_value in enumerate(row_values):
                if pd.notna(cell_value) and str(cell_value).strip().lower() == "funnel stage":
                    # Found "Funnel Stage" - this is the main header row
                    main_header_row_actual_idx = check_row_idx
                    
                    # Platform name should be 2 rows above in column B
                    platform_header_row_actual_idx = check_row_idx - 2
                    
                    if platform_header_row_actual_idx >= 0:
                        platform_cell = df_full_sheet.iloc[platform_header_row_actual_idx, 1]  # Column B (index 1)
                        if pd.notna(platform_cell):
                            platform_cell_str = str(platform_cell).strip().upper()
                            
                            # Check if it matches any known platform
                            for p_key, p_value in PLATFORM_NAMES.items():
                                if platform_cell_str == p_value.upper() or platform_cell_str == p_key.upper():
                                    platform_name_found = p_value
                                    logging.info(f"Found platform '{platform_name_found}' at row {platform_header_row_actual_idx + 1} (2 rows above 'Funnel Stage' at row {main_header_row_actual_idx + 1})")
                                    break
                            
                            if not platform_name_found:
                                # Unknown platform name, but structure is valid
                                logging.warning(f"Found unknown platform '{platform_cell_str}' at row {platform_header_row_actual_idx + 1}. Skipping.")
                    break
            
            if platform_name_found:
                break
        
        if not platform_name_found:
            logging.debug(f"No 'Funnel Stage' found in search window starting at sheet row {current_row_idx + 1}.")
            # Move forward to continue searching
            current_row_idx += 10  # Skip ahead more to avoid getting stuck
            if current_row_idx >= len(df_full_sheet) - 10:  # Near end of sheet
                logging.info(f"Reached near end of sheet. Stopping platform search.")
                break
            continue

        # We already have main_header_row_actual_idx from our search
        # No need to recalculate it

        main_header_values = df_full_sheet.iloc[main_header_row_actual_idx].str.strip().astype(str).tolist()
        logging.debug(f"Platform '{platform_name_found}': Potential main header at row {main_header_row_actual_idx+1}: {main_header_values}")

        # Validate core headers presence
        all_core_headers_present = all(header in main_header_values for header in core_main_headers_check)
        aspect_ratio_group_header_present = MAIN_HEADER_ASPECT_RATIO_GROUP_PRIMARY in main_header_values or \
                                          MAIN_HEADER_ASPECT_RATIO_GROUP_SECONDARY in main_header_values

        if not (all_core_headers_present and aspect_ratio_group_header_present):
            logging.warning(f"Platform '{platform_name_found}' at row {platform_header_row_actual_idx+1}: Missing one or more critical headers. Core headers present: {all_core_headers_present}, AR group present: {aspect_ratio_group_header_present}. Headers sought: {core_main_headers_check} & AR group. Found: {main_header_values}. Skipping platform.")
            current_row_idx = platform_header_row_actual_idx + 1 # Move to next row after platform title
            continue
            
        logging.info(f"Platform '{platform_name_found}': Successfully identified main headers at row {main_header_row_actual_idx+1}.")

        # Identify column indices for key headers
        funnel_stage_col_idx = main_header_values.index(FUNNEL_STAGE_HEADER)
        format_col_idx = main_header_values.index(FORMAT_HEADER)
        duration_col_idx = main_header_values.index(DURATION_HEADER)
        total_col_idx = main_header_values.index(MAIN_HEADER_TOTAL_COL)

        # Aspect Ratio / Secondary Format group columns
        ar_group_header_actual = ""
        ar_group_header_col_idx = -1
        
        if MAIN_HEADER_ASPECT_RATIO_GROUP_PRIMARY in main_header_values:
            ar_group_header_actual = MAIN_HEADER_ASPECT_RATIO_GROUP_PRIMARY
            ar_group_header_col_idx = main_header_values.index(ar_group_header_actual)
        elif MAIN_HEADER_ASPECT_RATIO_GROUP_SECONDARY in main_header_values: # Fallback for platforms like Programmatic
            ar_group_header_actual = MAIN_HEADER_ASPECT_RATIO_GROUP_SECONDARY
            # Find ALL occurrences of "Format" 
            format_indices = [i for i, h in enumerate(main_header_values) if h == MAIN_HEADER_ASPECT_RATIO_GROUP_SECONDARY]
            if len(format_indices) > 1:
                # Use the second "Format" (not the primary format column)
                ar_group_header_col_idx = format_indices[1] if format_indices[0] == format_col_idx else format_indices[0]
            else:
                ar_group_header_col_idx = format_indices[0]
        sub_header_row_actual_idx = platform_header_row_actual_idx + SUB_HEADER_ROW_OFFSET
        ar_sub_header_values = df_full_sheet.iloc[sub_header_row_actual_idx].astype(str).tolist()
        
        ar_cols_indices = []
        ar_col_names = []
        # Find columns under the Aspect Ratio/Format group until the 'Languages' or 'TOTAL' header is met
        # This logic needs to correctly identify the span of AR columns
        next_main_header_boundary_idx = len(main_header_values) # Default to end of row
        if MAIN_HEADER_LANGUAGES_GROUP in main_header_values:
            next_main_header_boundary_idx = main_header_values.index(MAIN_HEADER_LANGUAGES_GROUP)
        # If Languages isn't there (e.g. single lang sheet), TOTAL is the boundary
        elif MAIN_HEADER_TOTAL_COL in main_header_values:
             next_main_header_boundary_idx = main_header_values.index(MAIN_HEADER_TOTAL_COL)

        for i in range(ar_group_header_col_idx, next_main_header_boundary_idx):
            # Ensure sub-header is not 'nan', not empty, and the main header for this column is empty (merged cell)
            # For AR group, main_header_values[i] should be part of the merged AR group or empty if it's not the first AR col
            if pd.notna(ar_sub_header_values[i]) and str(ar_sub_header_values[i]).strip() not in ['', 'nan'] and \
               (main_header_values[i] == ar_group_header_actual or str(main_header_values[i]).strip() in ['', 'nan']):
                ar_cols_indices.append(i)
                ar_col_names.append(str(ar_sub_header_values[i]).strip())
            elif str(main_header_values[i]).strip() not in ['', 'nan'] and i > ar_group_header_col_idx: # Stop if we hit another main header
                break 
        logging.info(f"Platform '{platform_name_found}': Identified Aspect Ratio/Format columns: {ar_col_names} at indices {ar_cols_indices}")
        
        if not ar_col_names and ar_group_header_col_idx >= 0:
            # Debug why no columns were found
            logging.debug(f"Platform '{platform_name_found}': AR group header '{ar_group_header_actual}' at col {ar_group_header_col_idx}")
            logging.debug(f"Platform '{platform_name_found}': Checking columns {ar_group_header_col_idx} to {next_main_header_boundary_idx}")
            for i in range(ar_group_header_col_idx, min(ar_group_header_col_idx + 5, next_main_header_boundary_idx)):
                logging.debug(f"  Col {i}: main='{main_header_values[i] if i < len(main_header_values) else 'OOB'}', sub='{ar_sub_header_values[i] if i < len(ar_sub_header_values) else 'OOB'}')")

        # Languages group columns (only if dual_lang is true)
        lang_group_header_col_idx = -1
        lang_cols_indices = []
        lang_col_names = []
        if is_dual_lang:
            try:
                lang_group_header_col_idx = main_header_values.index(MAIN_HEADER_LANGUAGES_GROUP)
                # Find columns under the Languages group until 'TOTAL' is met
                total_header_boundary_idx = main_header_values.index(MAIN_HEADER_TOTAL_COL)
                for i in range(lang_group_header_col_idx, total_header_boundary_idx):
                    if pd.notna(ar_sub_header_values[i]) and str(ar_sub_header_values[i]).strip() not in ['', 'nan'] and \
                       (main_header_values[i] == MAIN_HEADER_LANGUAGES_GROUP or str(main_header_values[i]).strip() in ['', 'nan']):
                        lang_cols_indices.append(i)
                        lang_col_names.append(str(ar_sub_header_values[i]).strip())
                    elif str(main_header_values[i]).strip() not in ['', 'nan'] and i > lang_group_header_col_idx:
                        break
                logging.info(f"Platform '{platform_name_found}': Identified Language columns: {lang_col_names} at indices {lang_cols_indices}")
            except ValueError:
                logging.warning(f"Platform '{platform_name_found}' at row {main_header_row_actual_idx+1}: '{MAIN_HEADER_LANGUAGES_GROUP}' header not found, though is_dual_lang is True. Processing as if no language columns specified for this platform.")
                # lang_cols_indices remains empty, lang_col_names remains empty
        
        if not ar_cols_indices:
            logging.warning(f"Platform '{platform_name_found}': No Aspect Ratio/Format sub-columns found. Skipping platform.")
            # Continue searching from a reasonable position after this platform
            current_row_idx = main_header_row_actual_idx + 5  # Skip past this platform's header area
            continue

        # Process data rows for this platform
        data_start_row_for_platform = platform_header_row_actual_idx + DATA_START_ROW_OFFSET
        data_row_idx = data_start_row_for_platform

        while data_row_idx < len(df_full_sheet):
            # Stop if we encounter "Funnel Stage" in any column (indicates start of next platform)
            found_next_platform = False
            for col_idx in range(len(df_full_sheet.columns)):
                cell_val = df_full_sheet.iloc[data_row_idx, col_idx]
                if pd.notna(cell_val) and str(cell_val).strip().lower() == "funnel stage":
                    logging.debug(f"Platform '{platform_name_found}': Found 'Funnel Stage' at row {data_row_idx+1}. End of current platform data.")
                    found_next_platform = True
                    break
            
            if found_next_platform:
                # Move back 2 rows to position at the next platform header
                current_row_idx = data_row_idx - 2
                break 
            
            # Check if the row is empty or Funnel Stage is empty (heuristic for end of table section)
            # Consider a row empty if the 'Funnel Stage' column is empty for that row
            if pd.isna(df_full_sheet.iloc[data_row_idx, funnel_stage_col_idx]) or str(df_full_sheet.iloc[data_row_idx, funnel_stage_col_idx]).strip() == '':
                logging.debug(f"Platform '{platform_name_found}': Encountered empty 'Funnel Stage' at row {data_row_idx+1}. Assuming end of data for this platform.")
                break

            data_values = df_full_sheet.iloc[data_row_idx]
            logging.debug(f"Processing data row {data_row_idx+1}: {data_values.to_dict()}")

            funnel_stage = str(data_values.iloc[funnel_stage_col_idx]).strip()
            format_primary = str(data_values.iloc[format_col_idx]).strip()
            duration = str(data_values.iloc[duration_col_idx]).strip()
            total_val_from_sheet = safe_to_numeric(data_values.iloc[total_col_idx], data_row_idx, MAIN_HEADER_TOTAL_COL)

            # Expand 'ALL' funnel stage into A/C/P if enabled
            stages_to_emit = [funnel_stage]
            if EXPAND_ALL_TO_ACP and str(funnel_stage).strip().upper() == "ALL":
                stages_to_emit = FUNNEL_STAGES

            sum_of_ar_format_ticks = 0
            selected_ar_formats_for_row = [] # List of tuples (col_idx, ar_format_name)
            for ar_idx, ar_name in zip(ar_cols_indices, ar_col_names):
                tick_value = safe_to_numeric(data_values.iloc[ar_idx], data_row_idx, ar_name)
                if tick_value > 0:
                    sum_of_ar_format_ticks += tick_value
                    selected_ar_formats_for_row.append((ar_idx, ar_name))
            
            count_of_selected_languages = 0
            selected_language_names_for_row = []
            if is_dual_lang and lang_cols_indices: # Only consider languages if dual lang and lang columns were found
                for lang_idx, lang_name in zip(lang_cols_indices, lang_col_names):
                    # Language is selected if there's any non-empty, non-NA value (e.g., '1', 'x', 'yes')
                    lang_cell_value = str(data_values.iloc[lang_idx]).strip()
                    if lang_cell_value and lang_cell_value.lower() not in ['nan', '']:
                        count_of_selected_languages += 1
                        selected_language_names_for_row.append(lang_name)
                if not selected_language_names_for_row: # If dual_lang, but no languages ticked for this specific row
                    logging.debug(f"Row {data_row_idx+1} ({funnel_stage}, {format_primary}): Dual language sheet, but no languages selected for this row. Treating as 1 implicit language for combination count, but no language column will be added.")
                    count_of_selected_languages = 1 # Still counts as 1 for total calculation if row is valid
                    # selected_language_names_for_row remains empty, so no language-specific rows generated later if this path is taken.
            else: # Single language sheet or dual_lang where lang group wasn't found for platform
                count_of_selected_languages = 1
                # selected_language_names_for_row remains empty

            if sum_of_ar_format_ticks == 0:
                logging.debug(f"Row {data_row_idx+1} ({funnel_stage}, {format_primary}): Skipping, no AR/Format items ticked for this row.")
                data_row_idx += 1
                continue

            expected_combinations = sum_of_ar_format_ticks * count_of_selected_languages

            if expected_combinations != total_val_from_sheet:
                logging.warning(f"Row {data_row_idx+1} ({funnel_stage}, {format_primary}): Mismatch! Expected combinations '{expected_combinations}' (AR_sum:{sum_of_ar_format_ticks} * Lang_count:{count_of_selected_languages}) vs TOTAL in sheet '{total_val_from_sheet}'. Skipping row.")
                data_row_idx += 1
                continue
                
            logging.info(f"Row {data_row_idx+1} ({funnel_stage}, {format_primary}): Counts MATCH. Expected: {expected_combinations}, Sheet Total: {total_val_from_sheet}. Generating combinations.")

            # Generate creative combinations
            for ar_idx_tuple, ar_format_name_tuple in zip(enumerate(selected_ar_formats_for_row), selected_ar_formats_for_row):
                # ar_idx_tuple is (original_index_in_selected_ar_formats, (actual_col_idx, ar_name))
                # ar_format_name_tuple is (actual_col_idx, ar_name)
                actual_ar_col_idx, ar_format_name = ar_format_name_tuple
                num_ticks_for_ar = safe_to_numeric(data_values.iloc[actual_ar_col_idx], data_row_idx, ar_format_name)
                
                if num_ticks_for_ar > 0:
                    for _ in range(int(num_ticks_for_ar)):
                        for stage_name in stages_to_emit:
                            base_row_data = {
                                'Platform': platform_name_found,
                                FUNNEL_STAGE_HEADER: stage_name,
                                FORMAT_HEADER: format_primary,
                                DURATION_HEADER: duration,
                                OUTPUT_COLUMNS_BASE[4]: ar_format_name # Use config for 'Aspect Ratio / Format' key
                            }

                            if is_dual_lang and selected_language_names_for_row: # If dual lang and specific languages were ticked for the row
                                for lang_name in selected_language_names_for_row:
                                    specific_row = base_row_data.copy()
                                    specific_row[OUTPUT_LANGUAGE_COLUMN] = lang_name
                                    transformed_data_all_platforms.append(specific_row)
                            else: # Single language OR Dual Lang with no specific languages selected for this row (implicit single language for this row)
                                transformed_data_all_platforms.append(base_row_data) # No language column for single lang or if no langs selected
               
            data_row_idx += 1
        # After processing all data rows for the current platform
        logging.info(f"Completed processing platform '{platform_name_found}'. Moving to search from row {data_row_idx + 1}")
        current_row_idx = data_row_idx # Move to the start of the next potential platform or end of sheet

    # Summary of platforms processed
    platforms_found = set()
    for row in transformed_data_all_platforms:
        platforms_found.add(row['Platform'])
    
    logging.info(f"Platform search complete for {'Dual Lang' if is_dual_lang else 'Single Lang'} sheet")
    logging.info(f"Platforms successfully processed: {sorted(list(platforms_found))}")
    logging.info(f"Total rows generated: {len(transformed_data_all_platforms)}")
    
    return transformed_data_all_platforms


def process_excel_file_for_streamlit(input_excel_file_path: str) -> dict:
    """
    Processes the given Excel file for Streamlit app usage.
    Attempts to read and transform 'Tracker (Dual Lang)' and 'Tracker (Single Lang)' sheets.
    Returns a dictionary of DataFrames, keyed by sheet type name.
    """
    setup_logging() # Ensure logging is configured when called as a module
    logging.info(f"Processing Excel file for Streamlit: {os.path.basename(input_excel_file_path)}")

    results_by_sheet_type = {}

    sheets_to_process_config = [
        {
            "name": DUAL_LANG_INPUT_SHEET_NAME,
            "is_dual": True,
            "output_key": DUAL_LANG_INPUT_SHEET_NAME, # Or could use config.OUTPUT_SHEET_NAME_DUAL_LANG
            "output_cols": OUTPUT_COLUMNS_BASE + [OUTPUT_LANGUAGE_COLUMN]
        },
        {
            "name": SINGLE_LANG_INPUT_SHEET_NAME,
            "is_dual": False,
            "output_key": SINGLE_LANG_INPUT_SHEET_NAME, # Or could use config.OUTPUT_SHEET_NAME_SINGLE_LANG
            "output_cols": OUTPUT_COLUMNS_BASE
        }
    ]

    for config_item in sheets_to_process_config:
        sheet_name_to_process = config_item["name"]
        is_dual = config_item["is_dual"]
        output_key = config_item["output_key"]
        final_columns = config_item["output_cols"]
        df_transformed = None

        try:
            logging.info(f"Attempting to read sheet: '{sheet_name_to_process}'")
            df_sheet = pd.read_excel(input_excel_file_path, sheet_name=sheet_name_to_process, header=None)
            logging.info(f"Successfully read sheet: '{sheet_name_to_process}'. Starting transformation.")
            
            transformed_rows = find_platform_tables_and_transform(df_sheet, is_dual_lang=is_dual)
            
            if transformed_rows:
                df_transformed = pd.DataFrame(transformed_rows)
                # Ensure correct column order and presence
                df_transformed = df_transformed.reindex(columns=final_columns)
                logging.info(f"Sheet '{sheet_name_to_process}': Transformation successful. Generated {len(df_transformed)} rows.")
            else:
                logging.info(f"Sheet '{sheet_name_to_process}': No data transformed.")
                # Store an empty DataFrame with correct columns if no rows transformed but sheet existed
                df_transformed = pd.DataFrame(columns=final_columns)

        except ValueError as e:
            # Typically occurs if sheet_name is not found
            if "sheet_name" in str(e).lower() and f"'{sheet_name_to_process}'" in str(e):
                logging.warning(f"Sheet '{sheet_name_to_process}' not found in the Excel file. Skipping.")
            else:
                logging.error(f"Error processing sheet '{sheet_name_to_process}': {e}")
        except Exception as e:
            logging.error(f"Unexpected error processing sheet '{sheet_name_to_process}': {e}")
        
        results_by_sheet_type[output_key] = df_transformed

    return results_by_sheet_type


# Old main function is being replaced by process_excel_file_for_streamlit for module use,
# and the __main__ block will handle standalone execution.

if __name__ == "__main__":
    # This block executes when the script is run directly.
    setup_logging() # Ensure logging is setup for standalone run

    input_file = select_excel_file()
    if not input_file:
        logging.info("User cancelled file selection or no file provided. Exiting script.")
        if _TK_AVAILABLE:
            # Simplified Tkinter message for standalone cancellation
            root = tk.Tk()
            root.withdraw()
            root.attributes('-topmost', True)
            messagebox.showinfo("Cancelled", "No file selected. Exiting script.", parent=root)
            root.attributes('-topmost', False)
            root.destroy()
        exit()

    logging.info(f"Starting standalone Excel transformation for file: {os.path.basename(input_file)}")
    
    processed_data_map = process_excel_file_for_streamlit(input_file)
    
    output_dfs_to_write = []
    sheet_names_for_output = []

    # Check Dual Lang results
    df_dual = processed_data_map.get(DUAL_LANG_INPUT_SHEET_NAME)
    if df_dual is not None and not df_dual.empty:
        output_dfs_to_write.append(df_dual)
        sheet_names_for_output.append(OUTPUT_SHEET_NAME_DUAL_LANG) # Use config for output sheet name
        logging.info(f"'{DUAL_LANG_INPUT_SHEET_NAME}' processed: {len(df_dual)} rows.")
    elif df_dual is not None: # Exists but empty
        logging.info(f"'{DUAL_LANG_INPUT_SHEET_NAME}' processed, but resulted in 0 rows.")
    else: # Not processed (e.g. sheet not found)
        logging.info(f"'{DUAL_LANG_INPUT_SHEET_NAME}' was not processed or not found.")

    # Check Single Lang results
    df_single = processed_data_map.get(SINGLE_LANG_INPUT_SHEET_NAME)
    if df_single is not None and not df_single.empty:
        output_dfs_to_write.append(df_single)
        sheet_names_for_output.append(OUTPUT_SHEET_NAME_SINGLE_LANG) # Use config for output sheet name
        logging.info(f"'{SINGLE_LANG_INPUT_SHEET_NAME}' processed: {len(df_single)} rows.")
    elif df_single is not None: # Exists but empty
        logging.info(f"'{SINGLE_LANG_INPUT_SHEET_NAME}' processed, but resulted in 0 rows.")
    else: # Not processed
        logging.info(f"'{SINGLE_LANG_INPUT_SHEET_NAME}' was not processed or not found.")

    if output_dfs_to_write:
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        # Use OUTPUT_FILE_BASENAME from config
        output_filename = f"{OUTPUT_FILE_BASENAME}_{timestamp}.xlsx"
        
        try:
            with pd.ExcelWriter(output_filename, engine='openpyxl') as writer:
                for df_out, sheet_name_out in zip(output_dfs_to_write, sheet_names_for_output):
                    df_out.to_excel(writer, sheet_name=sheet_name_out, index=False)
            logging.info(f"Successfully transformed data written to {output_filename} with sheets: {sheet_names_for_output}")
            if _TK_AVAILABLE:
                root = tk.Tk()
                root.withdraw()
                root.attributes('-topmost', True)
                messagebox.showinfo("Success", f"Data written to:\n{output_filename}", parent=root)
                root.attributes('-topmost', False)
                root.destroy()
        except Exception as e:
            logging.error(f"Error writing output file '{output_filename}': {e}")
            if _TK_AVAILABLE:
                root = tk.Tk()
                root.withdraw()
                root.attributes('-topmost', True)
                messagebox.showerror("Error", f"Error writing to '{output_filename}':\n{e}", parent=root)
                root.attributes('-topmost', False)
                root.destroy()
    else:
        logging.info("No data was transformed from any sheet. Output file not created.")
        if _TK_AVAILABLE:
            root = tk.Tk()
            root.withdraw()
            root.attributes('-topmost', True)
            messagebox.showwarning("No Data", "No data transformed. Output not generated.", parent=root)
            root.attributes('-topmost', False)
            root.destroy()

    logging.info("Transformation script finished.")
    
    # Optional: Run validation if validation_script is available
    try:
        import validation_script
        if output_dfs_to_write and 'output_filename' in locals():
            logging.info("Running validation to verify output accuracy...")
            is_valid = validation_script.quick_validate(input_file, output_filename)
            if is_valid:
                logging.info("✅ VALIDATION PASSED - Output combinations match input totals!")
            else:
                logging.warning("⚠️ VALIDATION FAILED - Output combinations don't match input totals!")
    except ImportError:
        logging.debug("Validation script not available. Skipping validation.")
