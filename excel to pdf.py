from excel_to_pdf import excel_to_pdf

# after you’ve saved the uploaded processed Excel to, say, processed_path
pdf_path = excel_to_pdf(processed_path)

with open(pdf_path, "rb") as f:
    st.download_button(
        label="Download combined PDF (all sheets)",
        data=f,
        file_name=os.path.basename(pdf_path),
        mime="application/pdf",
    )

"""
excel_to_pdf.py

Convert an Excel workbook (.xlsx, .xlsm, .xls) into a single PDF,
with each worksheet rendered as a table on its own page.

Pure Python: works on Linux / Streamlit Cloud (no Excel needed).

Requirements (add to requirements.txt if needed):
    pandas
    matplotlib
    openpyxl
"""

import os
import sys

import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages


def excel_to_pdf(input_path: str, output_path: str | None = None) -> str:
    """
    Convert an Excel workbook to a single PDF, one page per sheet.

    Parameters
    ----------
    input_path : str
        Path to the Excel file.
    output_path : str | None
        Path for the output PDF. If None, uses the same
        base name as the Excel file with a .pdf extension.

    Returns
    -------
    str
        The path to the created PDF file.
    """
    input_path = os.path.abspath(input_path)
    if not os.path.isfile(input_path):
        raise FileNotFoundError(f"Input file not found: {input_path}")

    if output_path is None:
        base, _ = os.path.splitext(input_path)
        output_path = base + ".pdf"
    else:
        output_path = os.path.abspath(output_path)

    # Read workbook
    xls = pd.ExcelFile(input_path)

    # Create PDF
    with PdfPages(output_path) as pdf:
        for sheet_name in xls.sheet_names:
            df = xls.parse(sheet_name)

            # Turn NaN into empty strings and cast to str for nice display
            table_df = df.fillna("").astype(str)

            fig, ax = plt.subplots(figsize=(8.5, 11))  # US Letter portrait
            ax.axis("off")
            ax.set_title(sheet_name, fontsize=14, pad=12)

            # Build table
            table = ax.table(
                cellText=table_df.values,
                colLabels=table_df.columns,
                cellLoc="left",
                loc="upper left",
            )
            table.auto_set_font_size(False)
            # You can adjust these if things look too small/big
            table.set_fontsize(6)
            table.scale(1, 1.2)

            pdf.savefig(fig, bbox_inches="tight")
            plt.close(fig)

    return output_path


def main(argv: list[str]) -> None:
    if len(argv) < 2 or len(argv) > 3:
        print(
            "Usage:\n"
            "  python excel_to_pdf.py input.xlsx [output.pdf]\n\n"
            "If output.pdf is omitted, the PDF will be created next to the\n"
            "Excel file with the same base name."
        )
        return

    input_path = argv[1]
    output_path = argv[2] if len(argv) == 3 else None

    pdf_path = excel_to_pdf(input_path, output_path)
    print(f"PDF created at: {pdf_path}")


if __name__ == "__main__":
    main(sys.argv)
