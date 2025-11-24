import base64
import logging
import os
import tempfile
import traceback
from datetime import datetime
from io import BytesIO, StringIO
from pathlib import Path

import pandas as pd
import streamlit as st
from PIL import Image

from cej_transformer import config
from cej_transformer.logging_utils import configure_logging
from cej_transformer.transformer import process_workbook


# --- Page Configuration (Must be the FIRST Streamlit command) ---
st.set_page_config(
    page_title="CEJ Master Spec Transformer",
    layout="wide",
)


def _load_header_image() -> None:
    """Load and render the header image from common locations."""
    candidates = [Path("header.png"), Path("assets/header.png")]
    header_path = next((p for p in candidates if p.exists()), None)
    if not header_path:
        return

    try:
        header_img = Image.open(header_path)
        buffered = BytesIO()
        header_img.save(buffered, format="PNG")
        img_b64 = base64.b64encode(buffered.getvalue()).decode()

        st.markdown(
            f"""
            <div style='text-align:center;'>
                <img src='data:image/png;base64,{img_b64}' style='width:40%; height:auto;' />
            </div>
            """,
            unsafe_allow_html=True,
        )
    except Exception as e:  # pragma: no cover - defensive rendering
        st.warning(f"Header image could not be loaded ({header_path}): {e}")


_load_header_image()


def setup_streamlit_logging():
    """Configure logging and capture output for display in the UI."""
    configure_logging()
    app_logger = logging.getLogger("streamlit_app_logger")
    app_logger.setLevel(logging.INFO)

    log_stream_for_ui = StringIO()
    ui_handler = logging.StreamHandler(log_stream_for_ui)
    formatter = logging.Formatter("%(asctime)s - %(levelname)s - %(name)s - %(message)s")
    ui_handler.setFormatter(formatter)

    root_logger = logging.getLogger()
    for handler in root_logger.handlers[:]:
        if isinstance(handler, logging.StreamHandler) and isinstance(handler.stream, StringIO):
            root_logger.removeHandler(handler)
            handler.close()

    root_logger.addHandler(ui_handler)
    root_logger.setLevel(logging.INFO)

    return log_stream_for_ui, app_logger


def run_streamlit_app():
    st.title(f"CEJ Master Spec Sheet Transformer v{config.VERSION}")

    st.markdown(
        "Upload the 'Haleon CEJ Master Spec Sheet' Excel file to transform the 'Tracker (Dual Lang)' and 'Tracker (Single Lang)' sheets."
    )

    st.info(
        "✅ **v2.4.0 Features**: Improved platform detection using 'Funnel Stage' markers • "
        "Support for all 8 platforms (YouTube, META, TikTok, LinkedIn, Programmatic, Audio, Gaming, Amazon) • "
        "Enhanced dual/single language processing"
    )

    if config.EXPAND_ALL_TO_ACP:
        st.warning("ALL stage expansion is enabled: rows with Funnel Stage 'ALL' will emit Awareness, Consideration, and Purchase.")

    uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")
    if uploaded_file is not None:
        if st.button("Transform Excel Data"):
            log_stream, logger_instance = setup_streamlit_logging()

            logging.info("STREAMLIT_APP: 'Transform Excel Data' button clicked. Logging configured.")
            logger_instance.info("STREAMLIT_APP: Processing started.")

            with st.spinner("Processing your Excel file... Please wait."):
                with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_file:
                    tmp_file.write(uploaded_file.getvalue())
                    tmp_file_path = tmp_file.name

                try:
                    results_by_sheet_type = process_workbook(tmp_file_path)
                    st.session_state["results_by_sheet_type"] = results_by_sheet_type
                    st.session_state["file_processed"] = True
                except Exception as exc:
                    logger_instance.error("An error occurred during processing: %s", exc)
                    logger_instance.error(traceback.format_exc())
                    st.error(f"An error occurred: {exc}")
                finally:
                    if os.path.exists(tmp_file_path):
                        os.remove(tmp_file_path)

                    log_contents = log_stream.getvalue()
                    st.session_state["log_contents"] = log_contents

                st.session_state.pop("df_transformed", None)
                st.session_state.pop("platform_dfs", None)
                st.session_state.pop("platform_counts", None)

    if st.session_state.get("file_processed", False):
        if "log_contents" in st.session_state:
            st.subheader("Processing Log")
            log_text_to_display = st.session_state["log_contents"]
            if log_text_to_display is None:
                log_text_to_display = "Log content is None. (Unexpected)"
            elif not log_text_to_display.strip():
                log_text_to_display = "Log is empty. No messages were captured at INFO level or above."
            st.text_area("Log Details", log_text_to_display, height=200, key="log_display_area_diagnostics")

        results_data = st.session_state.get("results_by_sheet_type")
        any_data_processed = False

        if results_data:
            sheet_processing_configs = [
                {
                    "key": config.DUAL_LANG_INPUT_SHEET_NAME,
                    "title": "Dual Language Sheet Results ('Tracker (Dual Lang)')",
                    "output_sheet_name": config.OUTPUT_SHEET_NAME_DUAL_LANG,
                },
                {
                    "key": config.SINGLE_LANG_INPUT_SHEET_NAME,
                    "title": "Single Language Sheet Results ('Tracker (Single Lang)')",
                    "output_sheet_name": config.OUTPUT_SHEET_NAME_SINGLE_LANG,
                },
            ]

            for sp_config in sheet_processing_configs:
                sheet_key = sp_config["key"]
                sheet_title = sp_config["title"]
                df_current_sheet = results_data.get(sheet_key)

                st.subheader(sheet_title)
                if df_current_sheet is not None and not df_current_sheet.empty:
                    any_data_processed = True
                    st.write(f"Total unique creative combinations generated: {len(df_current_sheet)}")
                    st.dataframe(df_current_sheet.head(10))

                    st.markdown("#### Platform-Specific Breakdowns & Downloads")
                    platforms_in_sheet = df_current_sheet["Platform"].unique()
                    if len(platforms_in_sheet) > 0:
                        for platform_name in sorted(list(platforms_in_sheet)):
                            df_platform_specific = df_current_sheet[df_current_sheet["Platform"] == platform_name]
                            count = len(df_platform_specific)
                            expander_key = f"expander_{sheet_key}_{platform_name.replace(' ', '_')}"
                            button_key = f"button_dl_{sheet_key}_{platform_name.replace(' ', '_')}"

                            platform_name_display = platform_name.upper()

                            with st.expander(f"{platform_name_display}: {count} combinations"):
                                st.dataframe(df_platform_specific.head(10))
                                platform_excel_bytes = BytesIO()
                                with pd.ExcelWriter(platform_excel_bytes, engine="openpyxl") as writer_platform:
                                    df_platform_specific.to_excel(
                                        writer_platform, index=False, sheet_name=platform_name[:30]
                                    )
                                platform_excel_bytes.seek(0)
                                st.download_button(
                                    label=f"Download {platform_name_display} Data (Excel)",
                                    data=platform_excel_bytes,
                                    file_name=f"{platform_name}_{sheet_key.replace(' ', '_').lower()}_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    key=button_key,
                                )
                    else:
                        st.info("No platform data found within this sheet.")
                elif df_current_sheet is not None:
                    st.info(f"No data was transformed for the '{sheet_key}' sheet, although the sheet might exist.")
                else:
                    st.info(f"The sheet '{sheet_key}' was not found or not processed.")
                st.markdown("---")

            if any_data_processed:
                st.subheader("Download All Processed Data (Combined Excel)")
                output_excel_combined = BytesIO()
                with pd.ExcelWriter(output_excel_combined, engine="openpyxl") as writer_combined:
                    df_dual_lang_output = results_data.get(config.DUAL_LANG_INPUT_SHEET_NAME)
                    if df_dual_lang_output is not None and not df_dual_lang_output.empty:
                        df_dual_lang_output.to_excel(
                            writer_combined, sheet_name=config.OUTPUT_SHEET_NAME_DUAL_LANG, index=False
                        )

                    df_single_lang_output = results_data.get(config.SINGLE_LANG_INPUT_SHEET_NAME)
                    if df_single_lang_output is not None and not df_single_lang_output.empty:
                        df_single_lang_output.to_excel(
                            writer_combined, sheet_name=config.OUTPUT_SHEET_NAME_SINGLE_LANG, index=False
                        )

                output_excel_combined.seek(0)
                if output_excel_combined.getbuffer().nbytes > 0:
                    st.download_button(
                        label="Download Combined Data (Excel)",
                        data=output_excel_combined,
                        file_name=f"{config.OUTPUT_FILE_BASENAME}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="download_combined_all",
                    )
                else:
                    st.info("No data from any sheet was available to include in the combined download.")
            else:
                st.info("No data was transformed from any sheet. Nothing to download.")

        elif st.session_state.get("log_contents", ""):
            st.info("Processing was attempted, but no data was returned. Check logs for details.")


if __name__ == "__main__":
    run_streamlit_app()
