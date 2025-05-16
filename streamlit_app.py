import streamlit as st
import pandas as pd
import logging
from io import StringIO, BytesIO 
from datetime import datetime
import traceback 
import config
from collections import defaultdict
import excel_transformer
from PIL import Image

# --- Page Configuration (Must be the FIRST Streamlit command) ---
st.set_page_config(
    page_title="CEJ Master Spec Transformer",
    layout="wide"
)

# --- Display Header Image (Immediately after page config) ---
# Load and resize image to 70% of original width, then center using columns
header_img = Image.open("header.png")
new_width = int(header_img.width * 0.8)  # 20% smaller
col_left, col_center, col_right = st.columns([1, 4, 1])
with col_center:
    st.image(header_img, width=new_width)

# --- Configuration (Copied and adapted from excel_transformer.py) ---
INPUT_SHEET_NAME = 'Tracker (Dual Lang)'
OUTPUT_FILE_BASENAME = 'transformed_CEJ_master_specsheet'

# Platform names to search for (case-insensitive)
PLATFORM_NAMES = ["YOUTUBE", "META", "TIKTOK", "PROGRAMMATIC", "AUDIO", "GAMING", "AMAZON"]

# Expected main headers
MAIN_HEADER_ASPECT_RATIO_GROUP_PRIMARY = "Aspect Ratio"
MAIN_HEADER_ASPECT_RATIO_GROUP_SECONDARY = "Format" 
MAIN_HEADER_LANGUAGES_GROUP = "Languages"
MAIN_HEADER_TOTAL_COL = "TOTAL"
CORE_MAIN_HEADERS = ["Funnel Stage", "Format", "Duration", MAIN_HEADER_LANGUAGES_GROUP, MAIN_HEADER_TOTAL_COL]

START_ROW_SEARCH_FOR_PLATFORM = 7
PLATFORM_TITLE_ROW_OFFSET = 0
MAIN_HEADER_ROW_OFFSET = 2
SUB_HEADER_ROW_OFFSET = 3
DATA_START_ROW_OFFSET = 4

# --- Helper Functions (Copied from excel_transformer.py) ---
# Error grouping class to track and summarize similar errors
class ErrorCollector:
    def __init__(self):
        # Track errors by category
        self.error_categories = defaultdict(list)
        # Count of warnings/errors by type
        self.error_counts = defaultdict(int)
        # Special tracking for row-specific errors (e.g., invalid TOTAL values)
        self.row_specific_errors = defaultdict(list)
    
    def add_error(self, category, message, row=None, platform=None):
        """Add an error to tracking with optional row and platform info"""
        # Store full details for debugging
        self.error_categories[category].append((message, row, platform))
        self.error_counts[category] += 1
        
        # If row specific, track separately for better summaries
        if row is not None:
            key = f"{category}_{platform if platform else 'unknown'}"
            self.row_specific_errors[key].append(row)
    
    def get_summary(self):
        """Generate a human-friendly summary of collected errors"""
        summary_lines = []
        
        # Process row-specific errors for compact representation
        for key, rows in self.row_specific_errors.items():
            category, platform = key.split('_', 1)
            if len(rows) > 0:
                # Sort rows for cleaner output
                rows.sort()
                # Compress consecutive numbers into ranges
                ranges = []
                start = rows[0]
                end = rows[0]
                
                for i in range(1, len(rows)):
                    if rows[i] == end + 1:
                        end = rows[i]
                    else:
                        if start == end:
                            ranges.append(str(start))
                        else:
                            ranges.append(f"{start}-{end}")
                        start = end = rows[i]
                
                # Add the last range
                if start == end:
                    ranges.append(str(start))
                else:
                    ranges.append(f"{start}-{end}")
                
                # Format the error message
                if category == "invalid_total":
                    summary_lines.append(f"Platform '{platform}': Found {len(rows)} cell(s) with non-numeric TOTAL values (rows: {', '.join(ranges)}). These were treated as 0.")
                elif category == "empty_ar_format":
                    summary_lines.append(f"Platform '{platform}': {len(rows)} row(s) had no Aspect Ratio/Format values selected (rows: {', '.join(ranges)}).")
                elif category == "total_mismatch":
                    summary_lines.append(f"Platform '{platform}': The calculated combinations didn't match the TOTAL value for {len(rows)} row(s) (rows: {', '.join(ranges)}).")
                else:
                    # Generic row-specific error
                    summary_lines.append(f"Platform '{platform}': {category.replace('_', ' ').title()} - {len(rows)} occurrence(s) in rows: {', '.join(ranges)}")
        
        # Process platform-level errors
        platform_errors = defaultdict(list)
        for category, errors in self.error_categories.items():
            if not any(e[1] is not None for e in errors):  # No row numbers, so platform-level
                # Group by platform for clean summary
                for _, _, platform in errors:
                    if platform:
                        platform_errors[category].append(platform)
        
        for category, platforms in platform_errors.items():
            if platforms:
                unique_platforms = sorted(set(platforms))
                if category == "missing_ar_header":
                    summary_lines.append(f"Could not find Aspect Ratio or Format group headers for {len(unique_platforms)} platform(s): {', '.join(unique_platforms)}")
                elif category == "missing_columns":
                    summary_lines.append(f"Essential columns missing in {len(unique_platforms)} platform(s): {', '.join(unique_platforms)}")
                else:
                    # Generic platform-level error
                    summary_lines.append(f"{category.replace('_', ' ').title()} in {len(unique_platforms)} platform(s): {', '.join(unique_platforms)}")
        
        # If no categorized errors were added, return a generic message
        if not summary_lines:
            if sum(self.error_counts.values()) > 0:
                summary_lines.append(f"Found {sum(self.error_counts.values())} issues during processing. Check the detailed log for more information.")
        
        return summary_lines


def setup_streamlit_logging():
    log_stream = StringIO()
    handler = logging.StreamHandler(log_stream)
    handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
    
    logger = logging.getLogger(__name__) 
    logger.setLevel(logging.INFO) 
    
    if logger.hasHandlers():
        logger.handlers.clear()
    logger.addHandler(handler)
    return log_stream, logger

# --- Core Excel Processing Logic (Adapted from excel_transformer.py) ---
def find_platform_tables_and_transform(df_full_sheet, logger, error_collector):
    all_transformed_data = []
    platform_data_frames = {}
    platform_creative_counts = {}

    platform_column_header = config.PLATFORM_COLUMN_HEADER
    funnel_stage_header = config.FUNNEL_STAGE_HEADER
    format_header = config.FORMAT_HEADER
    duration_header = config.DURATION_HEADER
    total_header = config.TOTAL_HEADER
    ar_group_header_primary = config.MAIN_HEADER_ASPECT_RATIO_GROUP_PRIMARY
    ar_group_header_secondary = config.MAIN_HEADER_ASPECT_RATIO_GROUP_SECONDARY
    language_header = config.LANGUAGE_HEADER 

    platform_header_indices = df_full_sheet[df_full_sheet.iloc[:, 0].astype(str).str.strip().str.fullmatch(platform_column_header, case=False, na=False)].index
    logger.info(f"Found {len(platform_header_indices)} potential platform blocks based on '{platform_column_header}'.")

    for platform_block_start_index in platform_header_indices:
        platform_name_row = df_full_sheet.iloc[platform_block_start_index]
        platform_name = None
        for cell_value in platform_name_row:
            if pd.notna(cell_value) and str(cell_value).strip() != platform_column_header and str(cell_value).strip() != "":
                platform_name = str(cell_value).strip()
                break
        
        if platform_name is None:
            logger.warning(f"Could not determine platform name at sheet row {platform_block_start_index + 1}. Skipping this block.")
            continue

        logger.info(f"Processing platform: '{platform_name}' (identified at sheet row {platform_block_start_index + 1}).")
        platform_title_row_index = platform_block_start_index

        platform_main_header_row_index = None
        platform_data_start_row_index = None

        for i in range(platform_title_row_index + 1, min(platform_title_row_index + 5, len(df_full_sheet))):
            row_values = df_full_sheet.iloc[i].astype(str).str.strip().tolist()
            main_headers_count = sum([
                any(funnel_stage_header.lower() in str(cell).lower() for cell in row_values),
                any(format_header.lower() == str(cell).lower().strip() for cell in row_values),
                any(duration_header.lower() in str(cell).lower() for cell in row_values),
                any(total_header.lower() in str(cell).lower() for cell in row_values)
            ])
            primary_format_present = any(format_header.lower() == str(cell).lower().strip() for cell in row_values)
            ar_group_header_present = (
                any(ar_group_header_primary.lower() in str(cell).lower() for cell in row_values) or
                any(ar_group_header_secondary.lower() in str(cell).lower() for cell in row_values)
            )
            if main_headers_count >= 3 and primary_format_present and ar_group_header_present:
                platform_main_header_row_index = i
                platform_data_start_row_index = i + 2
                logger.info(f"Platform '{platform_name}': Main headers identified on sheet row {i+1}.") 
                break

        if platform_main_header_row_index is None:
            logger.warning(f"Platform '{platform_name}': Could not find the main header row. Skipping this platform.")
            continue

        platform_main_header_row = df_full_sheet.iloc[platform_main_header_row_index].astype(str).str.strip().tolist()

        funnel_stage_col_idx = next(i for i, header in enumerate(platform_main_header_row) if funnel_stage_header.lower() in header.lower())
        format_col_idx = next(i for i, header in enumerate(platform_main_header_row) if format_header.lower() == header.lower()) 
        duration_col_idx = next(i for i, header in enumerate(platform_main_header_row) if duration_header.lower() in header.lower())
        total_col_idx = next(i for i, header in enumerate(platform_main_header_row) if total_header.lower() in header.lower())

        logger.info(f"Platform '{platform_name}': Key column indices -> Funnel Stage: {funnel_stage_col_idx + 1}, Format: {format_col_idx + 1}, Duration: {duration_col_idx + 1}, Total: {total_col_idx + 1}.") 

        ar_cols_start_idx = -1
        ar_group_header_found = None
        try:
            ar_cols_start_idx = next(i for i, header in enumerate(platform_main_header_row) if ar_group_header_primary.lower() in header.lower())
            ar_group_header_found = ar_group_header_primary
            logger.info(f"Platform '{platform_name}': Found primary AR Group Header '{ar_group_header_found}' at column {ar_cols_start_idx + 1}.")
        except StopIteration:
            logger.info(f"Platform '{platform_name}': Primary AR Group Header '{ar_group_header_primary}' not found. Looking for secondary '{ar_group_header_secondary}'.")
            try:
                format_indices = [i for i, header in enumerate(platform_main_header_row) if ar_group_header_secondary.lower() == header.lower()]
                if not format_indices:
                    raise StopIteration 
                
                if len(format_indices) == 1:
                    if format_indices[0] != format_col_idx:
                        ar_cols_start_idx = format_indices[0]
                    else:
                        logger.warning(f"Platform '{platform_name}': Only one '{ar_group_header_secondary}' column found, and it's the primary format column. Cannot use for AR group.")
                elif len(format_indices) > 1:
                    if format_col_idx in format_indices:
                        primary_format_occurrence_index_in_list = format_indices.index(format_col_idx)
                        if primary_format_occurrence_index_in_list + 1 < len(format_indices):
                            ar_cols_start_idx = format_indices[primary_format_occurrence_index_in_list + 1]
                        else:
                            logger.warning(f"Platform '{platform_name}': Primary format column is the LAST '{ar_group_header_secondary}'. No subsequent one for AR group.")
                    else:
                        ar_cols_start_idx = format_indices[0] 
                        logger.warning(f"Platform '{platform_name}': Primary format column (idx {format_col_idx}) not among listed '{ar_group_header_secondary}' columns ({format_indices}). Defaulting AR group to first found '{ar_group_header_secondary}' at index {ar_cols_start_idx+1}.")
                
                if ar_cols_start_idx != -1:
                     ar_group_header_found = ar_group_header_secondary
                     logger.info(f"Platform '{platform_name}': Found secondary AR Group Header '{ar_group_header_found}' at column {ar_cols_start_idx + 1}.")
                else:
                    logger.warning(f"Platform '{platform_name}': Secondary AR Group Header '{ar_group_header_secondary}' also not suitably found or configured for AR Group. Combinations might be limited.")
            except StopIteration:
                logger.warning(f"Platform '{platform_name}': Neither primary ('{ar_group_header_primary}') nor secondary ('{ar_group_header_secondary}') AR Group Header found. Aspect ratio/format combinations will not be processed.")

        if ar_cols_start_idx == -1:
            logger.warning(f"Platform '{platform_name}': Aspect Ratio/Format group header not identified. Proceeding without these combinations.")
        
        current_platform_data = []
        for r in range(platform_data_start_row_index, len(df_full_sheet)):
            row_data = df_full_sheet.iloc[r]
            primary_cell_value = row_data.iloc[format_col_idx]

            if pd.isna(primary_cell_value) or str(primary_cell_value).strip() == "":
                logger.info(f"Platform '{platform_name}': Empty primary cell at sheet row {r + 1}. End of data for this platform.")
                break
            
            funnel_stage = str(row_data.iloc[funnel_stage_col_idx]).strip()
            format_val = str(primary_cell_value).strip()
            duration = str(row_data.iloc[duration_col_idx]).strip()
            total_val = row_data.iloc[total_col_idx]

            try:
                num_total = int(pd.to_numeric(total_val, errors='coerce'))
                if pd.isna(num_total):
                    error_collector.add_error("invalid_total", f"TOTAL value '{total_val}' is not a valid number", r+1, platform_name)
                    logger.debug(f"Platform '{platform_name}', Sheet Row {r+1}: TOTAL value '{total_val}' is not a valid number. Setting to 0.")
                    num_total = 0
            except ValueError:
                error_collector.add_error("invalid_total", f"TOTAL value '{total_val}' could not be converted to a number", r+1, platform_name)
                logger.debug(f"Platform '{platform_name}', Sheet Row {r+1}: TOTAL value '{total_val}' could not be converted to a number. Setting to 0.")
                num_total = 0

            base_combination = [platform_name, funnel_stage, format_val, duration]

            if ar_cols_start_idx != -1 and ar_group_header_found:
                ar_cols_end_idx = total_col_idx
                for c_ar in range(ar_cols_start_idx, total_col_idx):
                    if pd.isna(platform_main_header_row[c_ar]) or platform_main_header_row[c_ar] == "":
                        ar_cols_end_idx = c_ar
                        break

                ar_format_values = [str(val).strip() for val in row_data.iloc[ar_cols_start_idx:ar_cols_end_idx].values if pd.notna(val) and str(val).strip() != ""]

                if ar_format_values:
                    for ar_format in ar_format_values:
                        if num_total > 0:
                            current_platform_data.append(base_combination + [ar_format, num_total])
                        elif num_total == 0 and config.INCLUDE_ZERO_TOTAL_COMBINATIONS:
                            current_platform_data.append(base_combination + [ar_format, num_total])
                else:
                    error_collector.add_error("empty_ar_format", "No Aspect Ratio/Format values found", r+1, platform_name)
                    logger.debug(f"Platform '{platform_name}', Sheet Row {r+1}: No Aspect Ratio/Format values found although AR columns were identified.")
                    if num_total > 0:
                        current_platform_data.append(base_combination + ["N/A", num_total])
                    elif num_total == 0 and config.INCLUDE_ZERO_TOTAL_COMBINATIONS:
                        current_platform_data.append(base_combination + ["N/A", num_total])

            else:
                if ar_cols_start_idx == -1:
                    error_collector.add_error("missing_ar_header", "No Aspect Ratio/Format group header identified", None, platform_name)
                if num_total > 0:
                    current_platform_data.append(base_combination + ["N/A", num_total])
                elif num_total == 0 and config.INCLUDE_ZERO_TOTAL_COMBINATIONS:
                    current_platform_data.append(base_combination + ["N/A", num_total])

        if not current_platform_data:
            logger.warning(f"Platform '{platform_name}': No creative combinations generated. This might be due to missing data, all TOTALS being zero (if INCLUDE_ZERO_TOTAL_COMBINATIONS is False), or header mismatch.")
        else:
            logger.info(f"Platform '{platform_name}': Generated {len(current_platform_data)} combinations for this platform.")

        all_transformed_data.extend(current_platform_data)
        if current_platform_data:
            df_platform = pd.DataFrame(current_platform_data, columns=['Platform', 'Funnel Stage', 'Format', 'Duration', 'Aspect Ratio / Format', 'TOTAL'])
            platform_data_frames[platform_name] = df_platform
            platform_creative_counts[platform_name] = len(current_platform_data)

    if not all_transformed_data:
        logger.error("No data transformed across all platforms. Please check the input file structure and platform markers.")
        raise ValueError("No data was transformed. The Excel sheet might be empty or incorrectly formatted.")

    df_combined = pd.DataFrame(all_transformed_data, columns=['Platform', 'Funnel Stage', 'Format', 'Duration', 'Aspect Ratio / Format', 'TOTAL'])
    return df_combined, platform_data_frames, platform_creative_counts

def process_uploaded_file(uploaded_file, logger):
    logger.info(f"Reading sheet: {INPUT_SHEET_NAME} from uploaded file: {uploaded_file.name}")
    try:
        # Use the thoroughly tested logic from excel_transformer to ensure parity
        df_full_sheet = pd.read_excel(uploaded_file, sheet_name=INPUT_SHEET_NAME, header=None)
        logger.info("Successfully read the Excel sheet.")

        transformed_rows = excel_transformer.find_platform_tables_and_transform(df_full_sheet)

        if not transformed_rows:
            logger.warning("No data was transformed. The resulting dataset is empty.")
            return pd.DataFrame(), {}, {}

        # Build DataFrame with the same column order as excel_transformer output
        df_transformed = pd.DataFrame(transformed_rows, columns=['Platform', 'Funnel Stage', 'Format', 'Duration', 'Aspect Ratio / Format', 'Languages'])

        # Prepare per-platform DataFrames and counts for UI display/download
        platform_dfs = {platform: df_transformed[df_transformed['Platform'] == platform].reset_index(drop=True) for platform in df_transformed['Platform'].unique()}
        platform_counts = {platform: len(pdf) for platform, pdf in platform_dfs.items()}

        logger.info(
            "Successfully transformed data. Generated %d unique creative combinations.",
            len(df_transformed)
        )

        return df_transformed, platform_dfs, platform_counts
     
    except ValueError as ve: 
        logger.error(f"A known error occurred during processing: {str(ve)}")
        st.error(f"Process stopped: {str(ve)}") 
        return pd.DataFrame(), {}, {}
    except Exception as e:
        logger.error(f"An unexpected error occurred: {str(e)}\n{traceback.format_exc()}")
        st.error("An unexpected error occurred while processing the file. Please check the details below or contact support if the issue persists.")
        with st.expander("Technical Error Details"):
            st.code(traceback.format_exc())
        return pd.DataFrame(), {}, {}

# --- Streamlit App UI ---
def run_streamlit_app():
    setup_streamlit_logging()
    st.title("CEJ Master Spec Sheet Transformer")

    st.markdown("Upload the 'Haleon CEJ Master Spec Sheet' Excel file to transform the 'Tracker (Dual Lang)' sheet.")

    uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")
    log_stream, logger_instance = setup_streamlit_logging()

    if uploaded_file is not None:
        if st.button("Transform Excel Data"):
            with st.spinner('Processing your Excel file... Please wait.'):
                df_final_transformed, platform_specific_dfs, platform_specific_counts = process_uploaded_file(uploaded_file, logger_instance)

                log_contents = log_stream.getvalue()
                st.session_state['log_contents'] = log_contents
                st.session_state['df_transformed'] = df_final_transformed
                st.session_state['platform_dfs'] = platform_specific_dfs
                st.session_state['platform_counts'] = platform_specific_counts
                st.session_state['file_processed'] = True

    if st.session_state.get('file_processed', False):
        st.subheader("Processing Log")
        log_placeholder = st.empty()
        log_placeholder.text_area("Log", st.session_state.get('log_contents', ''), height=200)

        df_transformed_display = st.session_state.get('df_transformed')
        platform_dfs_display = st.session_state.get('platform_dfs')
        platform_counts_display = st.session_state.get('platform_counts')

        if df_transformed_display is not None and not df_transformed_display.empty:
            st.subheader("Transformed Data Overview")
            st.write(f"Total unique creative combinations generated: {len(df_transformed_display.drop_duplicates())}")
            st.dataframe(df_transformed_display.head(20)) 

            output_excel = BytesIO()
            with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
                df_transformed_display.to_excel(writer, index=False, sheet_name='All_Platforms_Combined')
            output_excel.seek(0)
            st.download_button(
                label="Download Combined Transformed Data (Excel)",
                data=output_excel,
                file_name=f"{OUTPUT_FILE_BASENAME}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            st.subheader("Platform-Specific Data & Downloads")
            for platform_name, df_platform in platform_dfs_display.items():
                count = platform_counts_display.get(platform_name, 0)
                st.markdown(f"**{platform_name}**: {count} combinations generated.")
                if not df_platform.empty:
                    with st.expander(f"View and Download Data for {platform_name}"):
                        st.dataframe(df_platform.head(10))
                        platform_excel = BytesIO()
                        with pd.ExcelWriter(platform_excel, engine='openpyxl') as writer:
                            df_platform.to_excel(writer, index=False, sheet_name=platform_name)
                        platform_excel.seek(0)
                        st.download_button(
                            label=f"Download {platform_name} Data (Excel)",
                            data=platform_excel,
                            file_name=f"{platform_name}_transformed_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key=f"download_{platform_name}" 
                        )
        elif df_transformed_display is not None and df_transformed_display.empty and not st.session_state.get('log_contents', '').strip() == "":
            pass 
        else:
            st.info("No data to display. This could be due to an empty input or an issue during processing not caught as an error.")

# --- Main Execution ---
if __name__ == "__main__":
    run_streamlit_app()
