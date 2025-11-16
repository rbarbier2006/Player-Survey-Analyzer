# streamlit_app.py

import os
import tempfile

import streamlit as st

from survey_processor import process_workbook, create_pdf_from_original


def main():
    st.title("Player Survey Analyzer")

    st.markdown(
        """
Upload the **original survey Excel file** from the players.

This app will:
1. Create a processed Excel file (split by team, with all tables and charts).
2. Create a combined PDF report (one page per team, with charts).
"""
    )

    cycle_label = st.text_input(
        "Cycle label (used in the PDF title)",
        value="Cycle X",
    )

    uploaded_file = st.file_uploader(
        "Upload original survey Excel (.xlsx)",
        type=["xlsx", "xls"],
    )

    if uploaded_file is None:
        st.info("Select an Excel file to begin.")
        return

    st.write(f"Uploaded file: **{uploaded_file.name}**")

    if st.button("Run analysis"):
        # Save uploaded file to a temporary path
        with st.spinner("Saving uploaded file..."):
            with tempfile.NamedTemporaryFile(
                delete=False, suffix=".xlsx"
            ) as tmp:
                tmp.write(uploaded_file.getbuffer())
                original_path = tmp.name

        # 1) Processed Excel
        with st.spinner("Creating processed Excel workbook..."):
            processed_path = process_workbook(original_path)

        # 2) Combined PDF
        with st.spinner("Creating PDF report..."):
            pdf_path = create_pdf_from_original(
                original_path,
                cycle_label=cycle_label,
            )

        st.success("Done! Download your files below.")

        # Download: processed Excel
        st.subheader("Processed Excel")
        with open(processed_path, "rb") as f:
            st.download_button(
                label="Download processed Excel",
                data=f,
                file_name=os.path.basename(processed_path),
                mime=(
                    "application/"
                    "vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                ),
            )

        # Download: combined PDF
        st.subheader("PDF report")
        with open(pdf_path, "rb") as fpdf:
            st.download_button(
                label="Download PDF report",
                data=fpdf,
                file_name=os.path.basename(pdf_path),
                mime="application/pdf",
            )


if __name__ == "__main__":
    main()
