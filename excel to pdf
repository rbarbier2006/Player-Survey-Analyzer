"""
excel_to_pdf.py

Convert an Excel workbook (.xlsx, .xlsm, .xls) into a single PDF,
with each worksheet exported as pages in the PDF.

Requirements:
- Windows
- Microsoft Excel installed
- pip install pywin32
"""

import os
import sys
from pathlib import Path

import win32com.client as win32


def convert_excel_to_pdf(input_path: str, output_path: str | None = None) -> str:
    """
    Convert an Excel workbook to a single PDF.

    Each sheet in the workbook becomes page(s) in the PDF,
    with charts and formatting preserved exactly as in Excel.

    Parameters
    ----------
    input_path : str
        Path to the Excel file (.xlsx, .xls, .xlsm).
    output_path : str | None
        Path for the output PDF. If None, uses the same
        base name as the Excel file with .pdf extension.

    Returns
    -------
    str
        The path to the created PDF file.
    """
    input_path = os.path.abspath(input_path)

    if not os.path.isfile(input_path):
        raise FileNotFoundError(f"Input file not found: {input_path}")

    if output_path is None:
        base = os.path.splitext(input_path)[0]
        output_path = base + ".pdf"
    else:
        output_path = os.path.abspath(output_path)

    # Make sure the output folder exists
    Path(os.path.dirname(output_path) or ".").mkdir(parents=True, exist_ok=True)

    # Launch Excel via COM
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    try:
        workbook = excel.Workbooks.Open(input_path)

        # Type 0 = PDF (see Excel's ExportAsFixedFormat docs)
        workbook.ExportAsFixedFormat(
            Type=0,
            Filename=output_path,
            Quality=0,                # 0 = standard, 1 = minimum
            IncludeDocProperties=True,
            IgnorePrintAreas=False,   # respect print areas if you set them
            OpenAfterPublish=False,
        )

        workbook.Close(SaveChanges=False)
    finally:
        excel.Quit()

    return output_path


def main(argv: list[str]) -> None:
    if len(argv) < 2 or len(argv) > 3:
        print(
            "Usage:\n"
            "  python excel_to_pdf.py input_excel.xlsx [output.pdf]\n\n"
            "If output.pdf is omitted, the PDF will be created next to the\n"
            "Excel file with the same base name."
        )
        return

    input_path = argv[1]
    output_path = argv[2] if len(argv) == 3 else None

    pdf_path = convert_excel_to_pdf(input_path, output_path)
    print(f"PDF created at: {pdf_path}")


if __name__ == "__main__":
    main(sys.argv)
