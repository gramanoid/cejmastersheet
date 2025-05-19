import pandas as pd
import os
import datetime
import logging
import re
import sys # Keep for __main__ if needed for other things, or remove if only for old logging
import traceback # Keep for error logging
from typing import List, Dict, Tuple, Optional, Any

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
    FUNNEL_STAGE_HEADER, FORMAT_HEADER, DURATION_HEADER # Specific header names from config
)

# Module-specific logger. All functions in this module should use this logger instance.
logger = logging.getLogger(__name__)

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
        logger.warning(f"Row {context_row_idx+1}, Col '{context_col_name}': Could not convert '{value}' to numeric. Treating as 0.")
        return 0

def find_platform_tables_and_transform(df_full_sheet, is_dual_lang: bool):
    transformed_data_all_platforms = []
    current_row_idx = START_ROW_SEARCH_FOR_PLATFORM - 3 # Start search a bit earlier
    if current_row_idx < 0: current_row_idx = 0

    # Define core main headers dynamically based on sheet type
    core_main_headers_check = [FUNNEL_STAGE_HEADER, FORMAT_HEADER, DURATION_HEADER, MAIN_HEADER_TOTAL_COL]
    if is_dual_lang:
        # Insert Languages group header before TOTAL for dual language sheets
        core_main_headers_check.insert(3, MAIN_HEADER_LANGUAGES_GROUP)

    while current_row_idx < len(df_full_sheet):
        platform_name_found = None
        platform_header_row_actual_idx = -1

        # Search for platform name in the next few rows
        for search_offset in range(10): # Search a window of 10 rows
            check_row_idx = current_row_idx + search_offset
            if check_row_idx >= len(df_full_sheet):
                break
            
            row_values = df_full_sheet.iloc[check_row_idx].astype(str).values
            for cell_value in row_values[:3]: # Check first 3 cells for platform name
                if pd.isna(cell_value) or cell_value.strip() == '':
                    continue
                for p_name in PLATFORM_NAMES:
                    if p_name.lower() in cell_value.lower():
                        platform_name_found = p_name
                        platform_header_row_actual_idx = check_row_idx
                        logger.info(f"Found platform '{platform_name_found}' title on sheet row {platform_header_row_actual_idx + 1}.")
                        break
                if platform_name_found:
                    break
            if platform_name_found:
                break
        
        if not platform_name_found:
            logger.info(f"No more platform titles found after sheet row {current_row_idx + 1}.")
            break # End of platforms

        current_row_idx = platform_header_row_actual_idx # Move main scan index to where platform was found

        main_header_row_actual_idx = platform_header_row_actual_idx + MAIN_HEADER_ROW_OFFSET
        if main_header_row_actual_idx >= len(df_full_sheet):
            logger.warning(f"Platform '{platform_name_found}': Main header row index {main_header_row_actual_idx+1} is out of bounds. Skipping platform.")
            current_row_idx = platform_header_row_actual_idx + 1 # Move to next row after platform title
            continue

        main_header_values = df_full_sheet.iloc[main_header_row_actual_idx].str.strip().astype(str).tolist()
        logger.debug(f"Platform '{platform_name_found}': Potential main header at row {main_header_row_actual_idx+1}: {main_header_values}")

        # Validate core headers presence
        all_core_headers_present = all(header in main_header_values for header in core_main_headers_check)
        aspect_ratio_group_header_present = MAIN_HEADER_ASPECT_RATIO_GROUP_PRIMARY in main_header_values or \
                                          MAIN_HEADER_ASPECT_RATIO_GROUP_SECONDARY in main_header_values

        if not (all_core_headers_present and aspect_ratio_group_header_present):
            logger.warning(f"Platform '{platform_name_found}' at row {platform_header_row_actual_idx+1}: Missing one or more critical headers. Core headers present: {all_core_headers_present}, AR group present: {aspect_ratio_group_header_present}. Headers sought: {core_main_headers_check} & AR group. Found: {main_header_values}. Skipping platform.")
            current_row_idx = platform_header_row_actual_idx + 1 # Move to next row after platform title
            continue
            
        logger.info(f"Platform '{platform_name_found}': Successfully identified main headers at row {main_header_row_actual_idx+1}.")

        # Identify column indices for key headers
        funnel_stage_col_idx = main_header_values.index(FUNNEL_STAGE_HEADER)
        format_col_idx = main_header_values.index(FORMAT_HEADER)
        duration_col_idx = main_header_values.index(DURATION_HEADER)
        total_col_idx = main_header_values.index(MAIN_HEADER_TOTAL_COL)

        # Aspect Ratio / Secondary Format group columns
        ar_group_header_actual = ""
        if MAIN_HEADER_ASPECT_RATIO_GROUP_PRIMARY in main_header_values:
            ar_group_header_actual = MAIN_HEADER_ASPECT_RATIO_GROUP_PRIMARY
        elif MAIN_HEADER_ASPECT_RATIO_GROUP_SECONDARY in main_header_values: # Fallback for platforms like Programmatic
            ar_group_header_actual = MAIN_HEADER_ASPECT_RATIO_GROUP_SECONDARY
        
        ar_group_header_col_idx = main_header_values.index(ar_group_header_actual)
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
        logger.info(f"Platform '{platform_name_found}': Identified Aspect Ratio/Format columns: {ar_col_names} at indices {ar_cols_indices}")

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
                logger.info(f"Platform '{platform_name_found}': Identified Language columns: {lang_col_names} at indices {lang_cols_indices}")
            except ValueError:
                logger.warning(f"Platform '{platform_name_found}' at row {main_header_row_actual_idx+1}: '{MAIN_HEADER_LANGUAGES_GROUP}' header not found, though is_dual_lang is True. Processing as if no language columns specified for this platform.")
                # lang_cols_indices remains empty, lang_col_names remains empty
        
        if not ar_cols_indices:
            logger.warning(f"Platform '{platform_name_found}': No Aspect Ratio/Format sub-columns found. Skipping platform.")
            current_row_idx = platform_header_row_actual_idx + DATA_START_ROW_OFFSET # Move to next platform section or end of data
            data_row_idx = main_header_row_actual_idx + 1 
            while data_row_idx < len(df_full_sheet) and pd.isna(df_full_sheet.iloc[data_row_idx, 0]): # Iterate through data rows of current table
                data_row_idx +=1
            current_row_idx = data_row_idx 
            continue

        # Process data rows for this platform
        data_start_row_for_platform = platform_header_row_actual_idx + DATA_START_ROW_OFFSET
        data_row_idx = data_start_row_for_platform

        while data_row_idx < len(df_full_sheet):
            # Stop if we encounter another platform title or a completely empty row (signaling end of data for this platform)
            # Check for new platform title in the first column (approx)
            potential_next_platform_title = str(df_full_sheet.iloc[data_row_idx, 0]).strip().upper()
            if any(platform.upper() in potential_next_platform_title for platform in PLATFORM_NAMES.values()) and not pd.isna(df_full_sheet.iloc[data_row_idx, 0]):
                logger.debug(f"Platform '{platform_name_found}': Encountered new platform title '{df_full_sheet.iloc[data_row_idx, 0]}' at row {data_row_idx+1}. Ending processing for current platform.")
                break 
            
            # Check if the row is empty or Funnel Stage is empty (heuristic for end of table section)
            # Consider a row empty if the 'Funnel Stage' column is empty for that row
            if pd.isna(df_full_sheet.iloc[data_row_idx, funnel_stage_col_idx]) or str(df_full_sheet.iloc[data_row_idx, funnel_stage_col_idx]).strip() == '':
                logger.debug(f"Platform '{platform_name_found}': Encountered empty 'Funnel Stage' at row {data_row_idx+1}. Assuming end of data for this platform.")
                break

            data_values = df_full_sheet.iloc[data_row_idx]
            logger.debug(f"Processing data row {data_row_idx+1}: {data_values.to_dict()}")

            funnel_stage = str(data_values.iloc[funnel_stage_col_idx]).strip()
            format_primary = str(data_values.iloc[format_col_idx]).strip()
            duration = str(data_values.iloc[duration_col_idx]).strip()
            total_val_from_sheet = safe_to_numeric(data_values.iloc[total_col_idx], data_row_idx, MAIN_HEADER_TOTAL_COL)

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
                    logger.debug(f"Row {data_row_idx+1} ({funnel_stage}, {format_primary}): Dual language sheet, but no languages selected for this row. Treating as 1 implicit language for combination count, but no language column will be added.")
                    count_of_selected_languages = 1 # Still counts as 1 for total calculation if row is valid
                    # selected_language_names_for_row remains empty, so no language-specific rows generated later if this path is taken.
            else: # Single language sheet or dual_lang where lang group wasn't found for platform
                count_of_selected_languages = 1
                # selected_language_names_for_row remains empty

            if sum_of_ar_format_ticks == 0:
                logger.debug(f"Row {data_row_idx+1} ({funnel_stage}, {format_primary}): Skipping, no AR/Format items ticked for this row.")
                data_row_idx += 1
                continue

            expected_combinations = sum_of_ar_format_ticks * count_of_selected_languages

            if expected_combinations != total_val_from_sheet:
                logger.warning(f"Row {data_row_idx+1} ({funnel_stage}, {format_primary}): Mismatch! Expected combinations '{expected_combinations}' (AR_sum:{sum_of_ar_format_ticks} * Lang_count:{count_of_selected_languages}) vs TOTAL in sheet '{total_val_from_sheet}'. Skipping row.")
                data_row_idx += 1
                continue
                
            logger.info(f"Row {data_row_idx+1} ({funnel_stage}, {format_primary}): Counts MATCH. Expected: {expected_combinations}, Sheet Total: {total_val_from_sheet}. Generating combinations.")

            # Generate creative combinations
            for ar_idx_tuple, ar_format_name_tuple in zip(enumerate(selected_ar_formats_for_row), selected_ar_formats_for_row):
                # ar_idx_tuple is (original_index_in_selected_ar_formats, (actual_col_idx, ar_name))
                # ar_format_name_tuple is (actual_col_idx, ar_name)
                actual_ar_col_idx, ar_format_name = ar_format_name_tuple
                num_ticks_for_ar = safe_to_numeric(data_values.iloc[actual_ar_col_idx], data_row_idx, ar_format_name)
                
                if num_ticks_for_ar > 0:
                    for _ in range(int(num_ticks_for_ar)):
                        base_row_data = {
                            'Platform': platform_name_found,
                            FUNNEL_STAGE_HEADER: funnel_stage,
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
        current_row_idx = data_row_idx # Move to the start of the next potential platform or end of sheet
    else:
        current_row_idx += 1 # Increment to search for the next platform title

    return transformed_data_all_platforms


def process_excel_file_for_streamlit(input_excel_file_path: str) -> Dict[str, Optional[pd.DataFrame]]:
    """
    Processes the given Excel file, expecting 'Tracker (Dual Lang)' and/or 'Tracker (Single Lang)' sheets.
    Returns a dictionary with sheet names as keys and their transformed DataFrames (or None) as values.
    """
    logger.setLevel(logging.INFO) # Explicitly set level for this specific logger instance

    logger.info(f"EXCEL_TRANSFORMER_UI_LOG: Starting processing for file: {input_excel_file_path} (via module logger)")
    # logging.getLogger().info("EXCEL_TRANSFORMER_UI_LOG: Test message from root logger inside excel_transformer.") # Test root logger propagation

    results = {}
    found_any_sheet = False

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
            logger.info(f"Attempting to read sheet: '{sheet_name_to_process}'")
            df_sheet = pd.read_excel(input_excel_file_path, sheet_name=sheet_name_to_process, header=None)
            logger.info(f"Successfully read sheet: '{sheet_name_to_process}'. Starting transformation.")
            
            transformed_rows = find_platform_tables_and_transform(df_sheet, is_dual_lang=is_dual)
            
            if transformed_rows:
                df_transformed = pd.DataFrame(transformed_rows)
                # Ensure correct column order and presence
                df_transformed = df_transformed.reindex(columns=final_columns)
                logger.info(f"Sheet '{sheet_name_to_process}': Transformation successful. Generated {len(df_transformed)} rows.")
            else:
                logger.info(f"Sheet '{sheet_name_to_process}': No data transformed.")
                # Store an empty DataFrame with correct columns if no rows transformed but sheet existed
                df_transformed = pd.DataFrame(columns=final_columns)

        except ValueError as e:
            # Typically occurs if sheet_name is not found
            if "sheet_name" in str(e).lower() and f"'{sheet_name_to_process}'" in str(e):
                logger.warning(f"Sheet '{sheet_name_to_process}' not found in the Excel file. Skipping.")
            else:
                logger.error(f"Error processing sheet '{sheet_name_to_process}': {e}")
        except Exception as e:
            logger.error(f"Unexpected error processing sheet '{sheet_name_to_process}': {e}")
        
        results[output_key] = df_transformed

    return results


if __name__ == "__main__":
    # Configure basic logging for standalone execution to show messages in the console.
    logging.basicConfig(level=logging.INFO,
                        format='%(asctime)s - %(levelname)s - %(name)s - %(message)s',
                        handlers=[logging.StreamHandler(sys.stdout)])

    import config 
    logger.info("EXCEL_TRANSFORMER (standalone): Script started.") # Use module logger for consistency

    input_file = select_excel_file()
    if input_file:
        # Use config.OUTPUT_FILE_BASENAME for the base name
        output_filename = f"{config.OUTPUT_FILE_BASENAME}_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        # Ensure output_path is correctly formed, e.g., save in the same directory as the input file or current dir
        output_dir = os.path.dirname(input_file) or '.' # Use '.' for current directory if input_file has no path
        output_path = os.path.join(output_dir, output_filename)

        logger.info(f"Starting standalone Excel transformation for: {input_file}")
        processed_data_map = process_excel_file_for_streamlit(input_file)
        
        output_dfs_to_write = {}
        if processed_data_map.get(DUAL_LANG_INPUT_SHEET_NAME) is not None and not processed_data_map[DUAL_LANG_INPUT_SHEET_NAME].empty:
            output_dfs_to_write[config.OUTPUT_SHEET_NAME_DUAL_LANG] = processed_data_map[DUAL_LANG_INPUT_SHEET_NAME]
        else:
            logger.info(f"No transformed data for sheet: {DUAL_LANG_INPUT_SHEET_NAME}")

        if processed_data_map.get(SINGLE_LANG_INPUT_SHEET_NAME) is not None and not processed_data_map[SINGLE_LANG_INPUT_SHEET_NAME].empty:
            output_dfs_to_write[config.OUTPUT_SHEET_NAME_SINGLE_LANG] = processed_data_map[SINGLE_LANG_INPUT_SHEET_NAME]
        else:
            logger.info(f"No transformed data for sheet: {SINGLE_LANG_INPUT_SHEET_NAME}")

        if output_dfs_to_write:
            try:
                with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                    for sheet_name, df_to_write in output_dfs_to_write.items():
                        df_to_write.to_excel(writer, sheet_name=sheet_name, index=False)
                        logger.info(f"Data written to sheet: {sheet_name} in {output_path}")
                logger.info(f"Successfully saved transformed data to {output_path}")
            except Exception as e:
                logger.error(f"Error writing Excel file {output_path}: {e}")
                logger.error(traceback.format_exc())
        else:
            logger.info("No dataframes to write to Excel.")
    else:
        logger.warning("No file selected for transformation.")

    logger.info("Transformation script finished.")
