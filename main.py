# main.py

import argparse
from survey_processor import process_workbook


def main():
    parser = argparse.ArgumentParser(
        description="Split survey Excel by team/category and add summary tables."
    )
    parser.add_argument(
        "input",
        help="Path to input Excel file (for example survey_raw.xlsx)",
    )
    parser.add_argument(
        "-o",
        "--output",
        help="Path to output Excel file (default: input name with _processed)",
    )

    args = parser.parse_args()
    output_path = process_workbook(args.input, args.output)

    print(f"Done. Processed workbook saved to: {output_path}")


if __name__ == "__main__":
    main()
