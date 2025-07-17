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
try:
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
except FileNotFoundError:
    # If header image is missing, just show the title without image
    logging.warning("header.png not found. Proceeding without header image.")
except Exception as e:
    logging.error(f"Error loading header image: {e}")
    # Continue without header image


# setup_streamlit_logging remains to capture logs for display in the UI
def setup_streamlit_logging():
    app_logger = logging.getLogger('streamlit_app_logger') 
    app_logger.setLevel(logging.INFO)
    
    log_stream_for_ui = StringIO()
    ui_handler = logging.StreamHandler(log_stream_for_ui)
    # Added %(name)s to formatter to see the logger's name (e.g., root, excel_transformer)
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(name)s - %(message)s') 
    ui_handler.setFormatter(formatter)
    
    root_logger = logging.getLogger()
    for h in root_logger.handlers[:]:
        if isinstance(h, logging.StreamHandler) and isinstance(h.stream, StringIO):
            root_logger.removeHandler(h)
            h.close()
            
    root_logger.addHandler(ui_handler)
    root_logger.setLevel(logging.INFO)

    # --- Streamlit App Diagnostics for excel_transformer logger ---
    # These prints will go to the console where Streamlit is running:
    # et_logger = logging.getLogger('excel_transformer') # Get the instance of the other module's logger
    # print(f"CONSOLE_SA: excel_transformer.logger name (from SA): {et_logger.name}")
    # print(f"CONSOLE_SA: excel_transformer.logger level (from SA): {et_logger.level} (Effective: {et_logger.getEffectiveLevel()})")
    # print(f"CONSOLE_SA: excel_transformer.logger propagate (from SA): {et_logger.propagate}")
    # print(f"CONSOLE_SA: excel_transformer.logger handlers (from SA): {et_logger.handlers}")
    # print(f"CONSOLE_SA: root_logger handlers (from SA after setup): {root_logger.handlers}")
    # --- End Diagnostics ---

    return log_stream_for_ui, app_logger


# --- Streamlit App UI ---
def run_streamlit_app():
    # Initial call to setup_streamlit_logging was removed to ensure it's only called on button press.
    st.title(f"CEJ Master Spec Sheet Transformer v{config.VERSION}")

    st.markdown("Upload the 'Haleon CEJ Master Spec Sheet' Excel file to transform the 'Tracker (Dual Lang)' and 'Tracker (Single Lang)' sheets.")
    
    st.info("✅ **v2.2.0 Features**: Improved platform detection using 'Funnel Stage' markers • Support for all 7 platforms (YouTube, META, TikTok, Programmatic, Audio, Gaming, Amazon) • Enhanced dual/single language processing")

    uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")
    if uploaded_file is not None:
        if st.button("Transform Excel Data"):
            log_stream, logger_instance = setup_streamlit_logging()
            
            # Diagnostic log from streamlit_app's root logger
            logging.info("STREAMLIT_APP: 'Transform Excel Data' button clicked. Logging configured.")
            logger_instance.info("STREAMLIT_APP: Test message from 'streamlit_app_logger'.") # Should also propagate to root

            with st.spinner('Processing your Excel file... Please wait.'):
                with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
                    tmp_file.write(uploaded_file.getvalue())
                    tmp_file_path = tmp_file.name
                
                try:
                    # excel_transformer logs to root logger, which now has our stream_handler
                    results_by_sheet_type = excel_transformer.process_excel_file_for_streamlit(tmp_file_path)
                    st.session_state['results_by_sheet_type'] = results_by_sheet_type
                    st.session_state['file_processed'] = True
                except Exception as e:
                    # Use the named logger_instance for app-specific error messages
                    logger_instance.error(f"An error occurred during processing: {e}")
                    logger_instance.error(traceback.format_exc())
                    st.error(f"An error occurred: {e}") 
                finally:
                    if os.path.exists(tmp_file_path):
                        os.remove(tmp_file_path)
                    
                    # log_stream is now guaranteed to be defined here
                    log_contents = log_stream.getvalue()
                    st.session_state['log_contents'] = log_contents
                
                # Clean up old session state keys if they exist from previous versions
                st.session_state.pop('df_transformed', None)
                st.session_state.pop('platform_dfs', None)
                st.session_state.pop('platform_counts', None)

    if st.session_state.get('file_processed', False):
        # Display logs if available
        if 'log_contents' in st.session_state:
            st.subheader("Processing Log")
            log_text_to_display = st.session_state['log_contents']
            if log_text_to_display is None: # Should not happen with getvalue()
                log_text_to_display = "Log content is None. (Unexpected)"
            elif not log_text_to_display.strip(): # Checks for empty string or only whitespace
                log_text_to_display = "Log is empty. No messages were captured at INFO level or above."
            # Use a new key for text_area to force re-render if necessary
            st.text_area("Log Details", log_text_to_display, height=200, key="log_display_area_diagnostics") 

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
                            expander_key = f"expander_{sheet_key}_{platform_name.replace(' ', '_')}"
                            button_key = f"button_dl_{sheet_key}_{platform_name.replace(' ', '_')}"

                            platform_name_display = platform_name.upper()

                            with st.expander(f"{platform_name_display}: {count} combinations"):
                                st.dataframe(df_platform_specific.head(10))
                                platform_excel_bytes = BytesIO()
                                with pd.ExcelWriter(platform_excel_bytes, engine='openpyxl') as writer_platform:
                                    df_platform_specific.to_excel(writer_platform, index=False, sheet_name=platform_name[:30]) # Excel sheet names <= 31 chars
                                platform_excel_bytes.seek(0)
                                st.download_button(
                                    label=f"Download {platform_name_display} Data (Excel)",
                                    data=platform_excel_bytes,
                                    file_name=f"{platform_name}_{sheet_key.replace(' ', '_').lower()}_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
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
