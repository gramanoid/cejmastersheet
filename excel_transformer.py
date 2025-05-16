import pandas as pd
import os
import datetime
import logging
import re
import tkinter as tk
from tkinter import filedialog, messagebox

# --- Configuration ---
INPUT_SHEET_NAME = 'Tracker (Dual Lang)'
OUTPUT_FILE_BASENAME = 'transformed_CEJ_master_specsheet'
LOG_FILE = 'transformer.log'

# Platform names to search for (case-insensitive)
PLATFORM_NAMES = ["YOUTUBE", "META", "TIKTOK", "PROGRAMMATIC", "AUDIO", "GAMING", "AMAZON"]

# Expected main headers (for locating and verifying header rows)
# Order matters for finding column groups
MAIN_HEADER_ASPECT_RATIO_GROUP_PRIMARY = "Aspect Ratio"
MAIN_HEADER_ASPECT_RATIO_GROUP_SECONDARY = "Format" # Used if primary isn't found, for platforms like Programmatic
MAIN_HEADER_LANGUAGES_GROUP = "Languages"
MAIN_HEADER_TOTAL_COL = "TOTAL"
# All core headers expected in the main header row for validation
CORE_MAIN_HEADERS = ["Funnel Stage", "Format", "Duration", MAIN_HEADER_LANGUAGES_GROUP, MAIN_HEADER_TOTAL_COL]
# Note: Aspect Ratio/Secondary Format group header is checked separately due to its variability

START_ROW_SEARCH_FOR_PLATFORM = 7 # 0-indexed, so Excel row 8. Search a bit around this.
PLATFORM_TITLE_ROW_OFFSET = 0 # Relative to found platform title cell
MAIN_HEADER_ROW_OFFSET = 2    # Relative to platform title row
SUB_HEADER_ROW_OFFSET = 3     # Relative to platform title row
DATA_START_ROW_OFFSET = 4     # Relative to platform title row

def setup_logging():
    # Remove existing handlers to prevent duplicate logs if re-run in same session
    for handler in logging.root.handlers[:]:
        logging.root.removeHandler(handler)
    logging.basicConfig(filename=LOG_FILE,
                        level=logging.INFO, 
                        format='%(asctime)s - %(levelname)s - %(message)s',
                        filemode='w') # Overwrite log file each run
    # Add a console handler to also see logs in the console
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO) # Console shows INFO and above
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    console_handler.setFormatter(formatter)
    logging.getLogger().addHandler(console_handler)

def select_excel_file():
    """Opens a dialog for the user to select an Excel file."""
    root = tk.Tk()
    root.withdraw()  # Hide the main tkinter window
    root.attributes('-topmost', True) # Make sure dialog comes to front
    file_path = filedialog.askopenfilename(
        parent=root, # Explicitly set parent
        title="Select the Excel file to process",
        filetypes=(("Excel files", "*.xlsx *.xls"), ("All files", "*.*"))
    )
    root.attributes('-topmost', False) # Reset topmost attribute
    root.destroy()
    return file_path

def safe_to_numeric(value, context_row_idx, context_col_name):
    """Attempts to convert value to numeric, handling potential errors and non-numeric strings."""
    if pd.isna(value) or str(value).strip() == '':
        return 0
    try:
        return pd.to_numeric(value)
    except ValueError:
        logging.warning(f"Row {context_row_idx+1}, Col '{context_col_name}': Could not convert '{value}' to numeric. Treating as 0.")
        return 0

def find_platform_tables_and_transform(df_full_sheet):
    transformed_data_all_platforms = []
    current_row_idx = START_ROW_SEARCH_FOR_PLATFORM - 3 # Start search a bit earlier
    if current_row_idx < 0: current_row_idx = 0

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
                        logging.info(f"Found platform '{platform_name_found}' title on sheet row {platform_header_row_actual_idx + 1}.")
                        break
                if platform_name_found:
                    break
            if platform_name_found:
                break
        
        if not platform_name_found:
            logging.info(f"No more platform titles found after sheet row {current_row_idx + 1}.")
            break # End of platforms

        current_row_idx = platform_header_row_actual_idx # Move main scan index to where platform was found

        main_header_sheet_idx = platform_header_row_actual_idx + MAIN_HEADER_ROW_OFFSET
        sub_header_sheet_idx = platform_header_row_actual_idx + SUB_HEADER_ROW_OFFSET
        data_start_sheet_idx = platform_header_row_actual_idx + DATA_START_ROW_OFFSET

        if main_header_sheet_idx >= len(df_full_sheet) or sub_header_sheet_idx >= len(df_full_sheet):
            logging.error(f"Platform '{platform_name_found}': Header rows extend beyond sheet limits. Skipping.")
            current_row_idx += 1
            continue

        # --- Extract and Validate Main Headers ---
        main_headers_series = df_full_sheet.iloc[main_header_sheet_idx].fillna('')
        main_headers_list = [str(h).strip() for h in main_headers_series.tolist()]
        logging.info(f"Platform '{platform_name_found}': Potential Main Headers on row {main_header_sheet_idx+1}: {main_headers_list}")

        # Validate core main headers are present (excluding the variable Aspect Ratio/Format group header for now)
        if not all(core_h in main_headers_list for core_h in CORE_MAIN_HEADERS):
            logging.error(f"Platform '{platform_name_found}': Missing one or more core main headers (excluding Aspect Ratio/Format group) like {CORE_MAIN_HEADERS} on row {main_header_sheet_idx+1}. Headers found: {main_headers_list}. Skipping table.")
            current_row_idx += 1
            continue
        
        # --- Identify Column Indices for Main Headers ---
        col_idx_aspect_ratio_group_start = -1
        actual_aspect_ratio_group_header_name = ""

        try:
            col_idx_funnel = main_headers_list.index("Funnel Stage")
            # Find the primary 'Format' column index first
            col_idx_format_primary = main_headers_list.index("Format") 
            col_idx_duration = main_headers_list.index("Duration")
            
            # Try to find 'Aspect Ratio' group header
            if MAIN_HEADER_ASPECT_RATIO_GROUP_PRIMARY in main_headers_list:
                col_idx_aspect_ratio_group_start = main_headers_list.index(MAIN_HEADER_ASPECT_RATIO_GROUP_PRIMARY)
                actual_aspect_ratio_group_header_name = MAIN_HEADER_ASPECT_RATIO_GROUP_PRIMARY
            # If 'Aspect Ratio' not found, try to find a *second* 'Format' column to act as the group header
            elif MAIN_HEADER_ASPECT_RATIO_GROUP_SECONDARY in main_headers_list:
                # Find all occurrences of 'Format'
                format_indices = [i for i, h in enumerate(main_headers_list) if h == MAIN_HEADER_ASPECT_RATIO_GROUP_SECONDARY]
                if len(format_indices) > 1: # We need at least two 'Format' occurrences
                    # Assume the one that is NOT col_idx_format_primary is the group header
                    # This assumes the group 'Format' is to the right of the primary 'Format'
                    for fi in format_indices:
                        if fi != col_idx_format_primary and fi > col_idx_format_primary:
                            col_idx_aspect_ratio_group_start = fi
                            actual_aspect_ratio_group_header_name = MAIN_HEADER_ASPECT_RATIO_GROUP_SECONDARY
                            break 
                if col_idx_aspect_ratio_group_start == -1: # Still not found a suitable secondary 'Format'
                    logging.error(f"Platform '{platform_name_found}': Found 'Format' but could not distinguish a secondary 'Format' to use as Aspect Ratio group header. Skipping table.")
                    current_row_idx +=1
                    continue
            else:
                logging.error(f"Platform '{platform_name_found}': Neither '{MAIN_HEADER_ASPECT_RATIO_GROUP_PRIMARY}' nor a secondary '{MAIN_HEADER_ASPECT_RATIO_GROUP_SECONDARY}' group header found. Skipping table.")
                current_row_idx += 1
                continue

            col_idx_languages_group_start = main_headers_list.index(MAIN_HEADER_LANGUAGES_GROUP)
            col_idx_total = main_headers_list.index(MAIN_HEADER_TOTAL_COL)

            # Post-check: Ensure Aspect Ratio/Format group is before Languages group, and Languages before TOTAL
            if not (col_idx_aspect_ratio_group_start < col_idx_languages_group_start < col_idx_total):
                logging.error(f"Platform '{platform_name_found}': Header group order incorrect ({actual_aspect_ratio_group_header_name} at {col_idx_aspect_ratio_group_start}, {MAIN_HEADER_LANGUAGES_GROUP} at {col_idx_languages_group_start}, {MAIN_HEADER_TOTAL_COL} at {col_idx_total}). Skipping table.")
                current_row_idx += 1
                continue

        except ValueError as ve:
            logging.error(f"Platform '{platform_name_found}': Could not find index for a core main header on row {main_header_sheet_idx+1}. Error: {ve}. Skipping table.")
            current_row_idx += 1
            continue

        # --- Extract Sub-Headers (Aspect Ratios and Languages) ---
        sub_headers_series = df_full_sheet.iloc[sub_header_sheet_idx].fillna('')
        
        aspect_ratio_cols = [] # List of {'name': str, 'idx': int}
        # Columns from Aspect Ratio group start up to (but not including) Languages group start
        for i in range(col_idx_aspect_ratio_group_start, col_idx_languages_group_start):
            sub_header_name = str(sub_headers_series.iloc[i]).strip()
            if sub_header_name: # Only add if there's a sub-header name
                aspect_ratio_cols.append({'name': sub_header_name, 'idx': i})
        
        language_cols = [] # List of {'name': str, 'idx': int}
        # Columns from Languages group start up to (but not including) TOTAL col
        for i in range(col_idx_languages_group_start, col_idx_total):
            sub_header_name = str(sub_headers_series.iloc[i]).strip()
            if sub_header_name:
                language_cols.append({'name': sub_header_name, 'idx': i})

        if not aspect_ratio_cols:
            logging.warning(f"Platform '{platform_name_found}': No aspect ratio sub-headers found under '{actual_aspect_ratio_group_header_name}'. Skipping table.")
            current_row_idx += 1
            continue
        if not language_cols:
            logging.warning(f"Platform '{platform_name_found}': No language sub-headers found under '{MAIN_HEADER_LANGUAGES_GROUP}'. Skipping table.")
            current_row_idx += 1
            continue
            
        logging.info(f"Platform '{platform_name_found}': Aspect Ratio columns: {aspect_ratio_cols}")
        logging.info(f"Platform '{platform_name_found}': Language columns: {language_cols}")

        platform_generated_rows_count = 0 # Counter for this specific platform

        # --- Process Data Rows for this Platform ---
        for data_row_idx in range(data_start_sheet_idx, len(df_full_sheet)):
            current_data_row = df_full_sheet.iloc[data_row_idx]
            # Check for end of table (completely empty row based on key columns)
            if current_data_row.iloc[[col_idx_funnel, col_idx_format_primary]].isnull().all():
                logging.info(f"Platform '{platform_name_found}': Detected end of table data at sheet row {data_row_idx + 1}.")
                current_row_idx = data_row_idx # Move scan to end of this table
                break
            
            funnel_stage_val = str(current_data_row.iloc[col_idx_funnel]).strip() if pd.notna(current_data_row.iloc[col_idx_funnel]) else ""
            format_primary_val = str(current_data_row.iloc[col_idx_format_primary]).strip() if pd.notna(current_data_row.iloc[col_idx_format_primary]) else ""
            duration_val = str(current_data_row.iloc[col_idx_duration]).strip() if pd.notna(current_data_row.iloc[col_idx_duration]) else ""

            ar_sum = 0
            lang_sum = 0
            ar_counts_for_row = []
            lang_counts_for_row = []

            for ar_col_info in aspect_ratio_cols:
                ar_val = safe_to_numeric(current_data_row.iloc[ar_col_info['idx']], data_row_idx, ar_col_info['name'])
                if ar_val > 0:
                    ar_sum += ar_val
                    ar_counts_for_row.append({'name': ar_col_info['name'], 'tick': ar_val})
                logging.debug(f"  AR/Format '{ar_col_info['name']}': read '{current_data_row.iloc[ar_col_info['idx']]}' as {ar_val}")

            for lang_col_info in language_cols:
                lang_val = safe_to_numeric(current_data_row.iloc[lang_col_info['idx']], data_row_idx, lang_col_info['name'])
                if lang_val > 0:
                    lang_sum += lang_val
                    lang_counts_for_row.append({'name': lang_col_info['name'], 'tick': lang_val})
                logging.debug(f"  Language '{lang_col_info['name']}': read '{current_data_row.iloc[lang_col_info['idx']]}' as {lang_val}")

            logging.debug(f"  AR/Format ticks sum for this row: {ar_sum}")
            logging.debug(f"  Number of distinct selected languages: {len(lang_counts_for_row)}")
            logging.debug(f"  Raw AR/Format details: {ar_counts_for_row}")
            logging.debug(f"  Raw Language details: {lang_counts_for_row}")

            # TOTAL validation (using original Excel row numbers for logging clarity)
            excel_sheet_row_num_for_log = current_data_row.name + 1 # .name is the original DataFrame index (0-indexed)
            total_from_sheet_raw = current_data_row.iloc[col_idx_total]
            total_from_sheet = 0
            try:
                if pd.notna(total_from_sheet_raw) and str(total_from_sheet_raw).strip() != "":
                    total_from_sheet = int(float(str(total_from_sheet_raw)))
            except ValueError:
                logging.warning(f"Platform '{platform_name_found}', Row {excel_sheet_row_num_for_log}: Non-numeric value '{total_from_sheet_raw}' in TOTAL column. Assuming 0 for validation.")
            
            logging.debug(f"  TOTAL from sheet (Excel Row {excel_sheet_row_num_for_log}): {total_from_sheet_raw} (parsed as {total_from_sheet})")

            # If no aspect ratios are ticked or no languages were selected for this row, it can't produce combinations
            if not ar_counts_for_row or not lang_counts_for_row:
                logging.debug(f"Platform '{platform_name_found}', Excel Row {excel_sheet_row_num_for_log}: Skipping. No AR/Format types ticked (count={len(ar_counts_for_row)}) OR no Languages selected (count={len(lang_counts_for_row)}).")
                if total_from_sheet > 0:
                    logging.warning(f"Platform '{platform_name_found}', Excel Row {excel_sheet_row_num_for_log}: TOTAL column is {total_from_sheet} but no AR/Format items or Languages found/selected. This row will be skipped.")
                continue
                
            # Calculate combinations based on: (Sum of AR/Format Ticks) * (Count of Selected Languages)
            # This is based on user clarification: (e.g. (Video_ticks + Banner_ticks) * num_languages)
            count_of_selected_languages = len(lang_counts_for_row)
            expected_combinations_for_this_row = ar_sum * count_of_selected_languages
            
            logging.debug(f"  Calculated combinations for this row (sum_AR_ticks * count_selected_langs): {ar_sum} * {count_of_selected_languages} = {expected_combinations_for_this_row}")

            if total_from_sheet != expected_combinations_for_this_row:
                logging.warning(f"Platform '{platform_name_found}', Excel Row {excel_sheet_row_num_for_log}: TOTAL column value '{total_from_sheet}' MISMATCHES calculated combinations '{expected_combinations_for_this_row}'. This input row will be SKIPPED.")
                continue # Skip this data row if totals don't align

            # Generate combinations if validation passed
            row_combinations_generated_count = 0
            for lang_detail in lang_counts_for_row:  # Iterate through each selected Language
                current_language_name = lang_detail['name']
                for ar_detail in ar_counts_for_row:  # Iterate through each selected AR/Format type
                    current_ar_format_name = ar_detail['name']
                    num_assets_for_this_ar_type = ar_detail['tick']
                    
                    if num_assets_for_this_ar_type > 0: # Only generate if the tick is positive
                        for _ in range(num_assets_for_this_ar_type):
                            transformed_data_all_platforms.append({
                                'Platform': platform_name_found,
                                'Funnel Stage': funnel_stage_val,
                                'Format': format_primary_val,
                                'Duration': duration_val,
                                'Aspect Ratio / Format': current_ar_format_name,
                                'Languages': current_language_name
                            })
                            row_combinations_generated_count += 1
            
            logging.debug(f"  Successfully generated {row_combinations_generated_count} combinations for this input row (Excel Row {excel_sheet_row_num_for_log}).")
            platform_generated_rows_count += row_combinations_generated_count

        logging.info(f"Platform '{platform_name_found}': Finished processing. Generated {platform_generated_rows_count} creative combinations for this platform.")
        current_row_idx = data_row_idx # Move to the start of the next potential platform

    return transformed_data_all_platforms

def main():
    setup_logging()

    input_excel_file = select_excel_file()
    if not input_excel_file:
        logging.info("User cancelled file selection. Exiting script.")
        # Show messagebox on top
        temp_root = tk.Tk()
        temp_root.withdraw()
        temp_root.attributes('-topmost', True)
        messagebox.showinfo("Cancelled", "No file selected. Exiting script.", parent=temp_root)
        temp_root.attributes('-topmost', False)
        temp_root.destroy()
        return

    logging.info(f"Starting Excel transformation for file: {os.path.basename(input_excel_file)}, sheet: {INPUT_SHEET_NAME}")

    try:
        df_full_sheet = pd.read_excel(input_excel_file, sheet_name=INPUT_SHEET_NAME, header=None)
    except FileNotFoundError:
        logging.error(f"Error: The file '{input_excel_file}' was not found.")
        temp_root = tk.Tk()
        temp_root.withdraw()
        temp_root.attributes('-topmost', True)
        messagebox.showerror("Error", f"File Not Found: '{os.path.basename(input_excel_file)}'", parent=temp_root)
        temp_root.attributes('-topmost', False)
        temp_root.destroy()
        return
    except Exception as e:
        logging.error(f"Error reading Excel file '{input_excel_file}': {e}")
        temp_root = tk.Tk()
        temp_root.withdraw()
        temp_root.attributes('-topmost', True)
        messagebox.showerror("Error", f"Error reading Excel file '{os.path.basename(input_excel_file)}':\n{e}", parent=temp_root)
        temp_root.attributes('-topmost', False)
        temp_root.destroy()
        return

    transformed_data = find_platform_tables_and_transform(df_full_sheet)

    if transformed_data:
        output_df = pd.DataFrame(transformed_data)
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        output_filename = f"{OUTPUT_FILE_BASENAME}_{timestamp}.xlsx"
        
        # Ensure correct column order for the output
        output_columns_ordered = ['Platform', 'Funnel Stage', 'Format', 'Duration', 'Aspect Ratio / Format', 'Languages']
        output_df = output_df.reindex(columns=output_columns_ordered)

        try:
            output_df.to_excel(output_filename, index=False)
            logging.info(f"Successfully transformed data written to {output_filename}")
            logging.info(f"Total unique creative combinations generated: {len(output_df)}")
            temp_root = tk.Tk()
            temp_root.withdraw()
            temp_root.attributes('-topmost', True)
            messagebox.showinfo("Success", f"Successfully transformed data written to:\n{output_filename}", parent=temp_root)
            temp_root.attributes('-topmost', False)
            temp_root.destroy()
        except Exception as e:
            logging.error(f"Error writing output file '{output_filename}': {e}")
            temp_root = tk.Tk()
            temp_root.withdraw()
            temp_root.attributes('-topmost', True)
            messagebox.showerror("Error", f"Error writing output file '{output_filename}':\n{e}", parent=temp_root)
            temp_root.attributes('-topmost', False)
            temp_root.destroy()
    else:
        logging.info("No data was transformed. Output file not created.")
        temp_root = tk.Tk()
        temp_root.withdraw()
        temp_root.attributes('-topmost', True)
        messagebox.showwarning("No Data", "No data was transformed. Output file will not be generated.", parent=temp_root)
        temp_root.attributes('-topmost', False)
        temp_root.destroy()

    logging.info("Transformation script finished.")

if __name__ == "__main__":
    main()
