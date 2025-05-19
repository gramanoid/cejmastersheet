import streamlit as st
import pandas as pd
import logging
from io import StringIO, BytesIO 
from datetime import datetime
import traceback 
import config
import excel_transformer
from PIL import Image
import base64
import os
import tempfile


# --- Page Configuration (Must be the FIRST Streamlit command) ---
st.set_page_config(
    page_title="CEJ Master Spec Transformer",
    layout="wide"
)

# --- Display Header Image (Immediately after page config) ---
header_img = Image.open("header.png")
buffered = BytesIO()
header_img.save(buffered, format="PNG")
img_b64 = base64.b64encode(buffered.getvalue()).decode()

st.markdown(
    f"""
    <div style='text-align:center;'>
        <img src='data:image/png;base64,{img_b64}' style='width:40%; height:auto;' />
    </div>
    """,
    unsafe_allow_html=True
)


# setup_streamlit_logging remains to capture logs for display in the UI
def setup_streamlit_logging():
    logger = logging.getLogger('streamlit_app_logger') 
    logger.setLevel(logging.INFO)
    for handler in logger.handlers[:]:
        logger.removeHandler(handler)
    
    log_stream = StringIO()
    stream_handler = logging.StreamHandler(log_stream)
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    stream_handler.setFormatter(formatter)
    logger.addHandler(stream_handler)
    
    # Deduplicate root logger handlers that point to a StringIO (prevents log duplication on repeated runs)
    root_logger = logging.getLogger()
    for h in root_logger.handlers[:]:
        if isinstance(h, logging.StreamHandler) and isinstance(h.stream, StringIO):
            root_logger.removeHandler(h)
    root_logger.addHandler(stream_handler)
    root_logger.setLevel(logging.INFO)

    return log_stream, logger


# --- Streamlit App UI ---
def run_streamlit_app():
    # Removed initial call to setup_streamlit_logging to avoid duplicate handlers. Handlers are created when the user clicks the button.
    st.title("CEJ Master Spec Sheet Transformer")

    st.markdown("Upload the 'Haleon CEJ Master Spec Sheet' Excel file to transform the 'Tracker (Dual Lang)' and 'Tracker (Single Lang)' sheets.")

    uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")
    if uploaded_file is not None:
        if st.button("Transform Excel Data"):
            with st.spinner('Processing your Excel file... Please wait.'):
                with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
                    tmp_file.write(uploaded_file.getvalue())
                    tmp_file_path = tmp_file.name
                
                try:
                    results_by_sheet_type = excel_transformer.process_excel_file_for_streamlit(tmp_file_path)
                    st.session_state['results_by_sheet_type'] = results_by_sheet_type
                    st.session_state['file_processed'] = True
                except Exception as e:
                    log_stream, logger_instance = setup_streamlit_logging()
                    logger_instance.error(f"An error occurred during processing: {e}")
                    logger_instance.error(traceback.format_exc())
                    st.error(f"An error occurred: {e}") 
                    st.session_state['file_processed'] = False 
                finally:
                    os.remove(tmp_file_path) 

                if 'log_stream' in locals():
                    log_contents = log_stream.getvalue()
                    st.session_state['log_contents'] = log_contents
                
                st.session_state.pop('df_transformed', None)
                st.session_state.pop('platform_dfs', None)
                st.session_state.pop('platform_counts', None)

    if st.session_state.get('file_processed', False):
        st.subheader("Processing Log")
        log_placeholder = st.empty()
        log_placeholder.text_area("Log", st.session_state.get('log_contents', ''), height=200)

        results_data = st.session_state.get('results_by_sheet_type')
        any_data_processed = False

        if results_data:
            # Define the order and titles for display
            sheet_processing_configs = [
                {
                    "key": config.DUAL_LANG_INPUT_SHEET_NAME,
                    "title": "Dual Language Sheet Results ('Tracker (Dual Lang)')",
                    "output_sheet_name": config.OUTPUT_SHEET_NAME_DUAL_LANG
                },
                {
                    "key": config.SINGLE_LANG_INPUT_SHEET_NAME,
                    "title": "Single Language Sheet Results ('Tracker (Single Lang)')",
                    "output_sheet_name": config.OUTPUT_SHEET_NAME_SINGLE_LANG
                }
            ]

            for sp_config in sheet_processing_configs:
                sheet_key = sp_config["key"]
                sheet_title = sp_config["title"]
                df_current_sheet = results_data.get(sheet_key)

                st.subheader(sheet_title)
                if df_current_sheet is not None and not df_current_sheet.empty:
                    any_data_processed = True
                    st.write(f"Total unique creative combinations generated: {len(df_current_sheet)}")
                    st.dataframe(df_current_sheet.head(10)) # Show fewer rows for overview
                    
                    st.markdown("#### Platform-Specific Breakdowns & Downloads")
                    platforms_in_sheet = df_current_sheet['Platform'].unique()
                    if len(platforms_in_sheet) > 0:
                        for platform_name in sorted(list(platforms_in_sheet)):
                            df_platform_specific = df_current_sheet[df_current_sheet['Platform'] == platform_name]
                            count = len(df_platform_specific)
                            # Use a more robust key for expander and button to avoid conflicts
                            expander_key = f"expander_{sheet_key}_{platform_name}".replace(" ", "_")
                            button_key = f"button_dl_{sheet_key}_{platform_name}".replace(" ", "_")

                            with st.expander(f"{platform_name}: {count} combinations", key=expander_key):
                                st.dataframe(df_platform_specific.head(10))
                                platform_excel_bytes = BytesIO()
                                with pd.ExcelWriter(platform_excel_bytes, engine='openpyxl') as writer_platform:
                                    df_platform_specific.to_excel(writer_platform, index=False, sheet_name=platform_name[:30]) # Excel sheet names <= 31 chars
                                platform_excel_bytes.seek(0)
                                st.download_button(
                                    label=f"Download {platform_name} Data (Excel)",
                                    data=platform_excel_bytes,
                                    file_name=f"{platform_name}_transformed_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    key=button_key
                                )
                    else:
                        st.info("No platform data found within this sheet.")
                elif df_current_sheet is not None: # Empty DataFrame
                    st.info(f"No data was transformed for the '{sheet_key}' sheet, although the sheet might exist.")
                else: # df_current_sheet is None (sheet likely not found)
                    st.info(f"The sheet '{sheet_key}' was not found or not processed.")
                st.markdown("---") # Separator

            if any_data_processed:
                st.subheader("Download All Processed Data (Combined Excel)")
                output_excel_combined = BytesIO()
                with pd.ExcelWriter(output_excel_combined, engine='openpyxl') as writer_combined:
                    df_dual_lang_output = results_data.get(config.DUAL_LANG_INPUT_SHEET_NAME)
                    if df_dual_lang_output is not None and not df_dual_lang_output.empty:
                        df_dual_lang_output.to_excel(writer_combined, sheet_name=config.OUTPUT_SHEET_NAME_DUAL_LANG, index=False)
                    
                    df_single_lang_output = results_data.get(config.SINGLE_LANG_INPUT_SHEET_NAME)
                    if df_single_lang_output is not None and not df_single_lang_output.empty:
                        df_single_lang_output.to_excel(writer_combined, sheet_name=config.OUTPUT_SHEET_NAME_SINGLE_LANG, index=False)
                
                output_excel_combined.seek(0)
                # Check if any sheet was actually written to the combined file
                if output_excel_combined.getbuffer().nbytes > 0: # A more robust check might be needed if ExcelWriter writes headers for empty sheets
                    st.download_button(
                        label="Download Combined Data (Excel)",
                        data=output_excel_combined,
                        file_name=f"{config.OUTPUT_FILE_BASENAME}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="download_combined_all"
                    )
                else:
                    st.info("No data from any sheet was available to include in the combined download.")
            else:
                st.info("No data was transformed from any sheet. Nothing to download.")
        
        elif st.session_state.get('log_contents', ''): # If file_processed is true, but results_data is empty/None, but there are logs
            st.info("Processing was attempted, but no data was returned. Check logs for details.")
        # else: # No file processed or no results
            # No explicit message needed here as the UI will just not show data sections
            # st.info("Upload a file and click 'Transform Excel Data' to see results.") # This would show before processing

# --- Main Execution ---
if __name__ == "__main__":
    run_streamlit_app()
