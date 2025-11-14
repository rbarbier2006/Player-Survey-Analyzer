import tempfile

import pandas as pd
import streamlit as st

from survey_processor import process_workbook

st.set_page_config(page_title="Player Survey Analyzer", layout="centered")

st.title("Player Survey Analyzer")

st.write(
    "Upload an Excel file with the player survey responses. "
    "The app will split the data by Team+Category (column G), "
    "add summary tables to each sheet, and return a processed Excel file."
)

uploaded = st.file_uploader("Upload Excel file", type=["xlsx", "xls"])

if uploaded is not None:
    # Save uploaded file to a temporary path
    suffix = ".xlsx"
    if uploaded.name.lower().endswith(".xls"):
        suffix = ".xls"

    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp_in:
        tmp_in.write(uploaded.getvalue())
        tmp_in_path = tmp_in.name

    with st.spinner("Processing workbook..."):
        # Use your existing function to build the processed workbook
        out_path = process_workbook(tmp_in_path)

    # Read processed file back into memory
    with open(out_path, "rb") as f:
        processed_bytes = f.read()

    st.success("Processing complete.")

    # Optional preview of the All_Data sheet (first 10 rows)
    try:
        df_preview = pd.read_excel(out_path, sheet_name="All_Data")
        st.subheader("Preview of All_Data sheet (first 10 rows)")
        st.dataframe(df_preview.head(10))
    except Exception:
        st.info("Could not load preview, but the file is ready to download.")

    # Download button
    st.download_button(
        label="Download processed Excel file",
        data=processed_bytes,
        file_name="player_survey_processed.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
else:
    st.info("Select an Excel file to begin.")
