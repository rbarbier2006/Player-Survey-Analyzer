# survey_processor.py

import os
import re
import numpy as np
import pandas as pd

# Column layout based on your description
# F: player name (not used directly here)
# G: team + category (grouping key)
# H, I, J, K, L, P, Q: 1-5 star rating questions
# M, N: Yes/No questions

RATING_COL_LETTERS = ["H", "I", "J", "K", "L", "P", "Q"]
YESNO_COL_LETTERS = ["M", "N"]


def col_letter_to_index(col_letter: str) -> int:
    """
    Convert an Excel column letter (for example 'A', 'G', 'AA')
    to a zero-based column index (A=0, B=1, ...).
    """
    col_letter = col_letter.strip().upper()
    if not col_letter:
        raise ValueError("Empty column letter")

    index = 0
    for ch in col_letter:
        if not ("A" <= ch <= "Z"):
            raise ValueError(f"Invalid column letter: {col_letter}")
        index = index * 26 + (ord(ch) - ord("A") + 1)
    return index - 1


# Precompute the indices we need
RATING_COL_INDICES = [col_letter_to_index(c) for c in RATING_COL_LETTERS]
YESNO_COL_INDICES = [col_letter_to_index(c) for c in YESNO_COL_LETTERS]
GROUP_COL_INDEX = col_letter_to_index("G")


def make_safe_sheet_name(raw_name: str, used_names=None) -> str:
    """
    Clean a raw sheet name so it is valid for Excel:
    - remove invalid characters \ / * ? : [ ]
    - trim to 31 characters
    - avoid duplicates by adding _1, _2, etc
    """
    if used_names is None:
        used_names = set()

    cleaned = re.sub(r"[\\/*?:\[\]]", "_", str(raw_name))
    cleaned = cleaned.strip()
    if not cleaned:
        cleaned = "Group"

    base = cleaned[:31]
    name = base
    counter = 1

    while name in used_names:
        suffix = f"_{counter}"
        max_base_len = 31 - len(suffix)
        if max_base_len < 1:
            raise ValueError(
                "Cannot create unique sheet name within Excel 31-char limit."
            )
        name = base[:max_base_len] + suffix
        counter += 1

    used_names.add(name)
    return name


def append_summary_tables(df: pd.DataFrame,
                          writer: pd.ExcelWriter,
                          sheet_name: str,
                          rating_indices,
                          yesno_indices) -> None:
    """
    For the given DataFrame and sheet, append:
    - ratings table (1..5 counts + average) for all rating columns
    - yes/no counts for all yes/no columns

    All tables start 3 blank rows after the last data row.
    """
    # Number of player rows
    n_rows = df.shape[0]

    # 3 blank rows after header + data
    # header is at row 0, data at rows 1..n_rows
    # so startrow = n_rows + 4 gives 3 blank rows
    startrow = n_rows + 4

    cols = list(df.columns)
    rating_cols = [cols[i] for i in rating_indices if i < len(cols)]
    yesno_cols = [cols[i] for i in yesno_indices if i < len(cols)]

    # Case 1: rating questions (1-5 stars plus average)
    if rating_cols:
        scores = list(range(1, 6))
        index = scores + ["Average"]
        rating_summary = pd.DataFrame(index=index,
                                      columns=rating_cols,
                                      dtype=float)

        for col in rating_cols:
            series = df[col]
            numeric_series = pd.to_numeric(series, errors="coerce")

            for s in scores:
                rating_summary.loc[s, col] = (numeric_series == s).sum()

            if numeric_series.notna().sum() > 0:
                rating_summary.loc["Average", col] = numeric_series.mean()
            else:
                rating_summary.loc["Average", col] = np.nan

        # Write the rating table
        rating_summary.to_excel(
            writer,
            sheet_name=sheet_name,
            startrow=startrow,
            startcol=0,
            index=True,
            index_label="Score",
        )

        # Move startrow below this table plus one blank row
        table_height = rating_summary.shape[0] + 1  # data rows + header
        startrow = startrow + table_height + 1

    # Case 2: Yes/No questions
    if yesno_cols:
        yesno_index = ["YES", "NO"]
        yesno_summary = pd.DataFrame(index=yesno_index,
                                     columns=yesno_cols,
                                     dtype=float)

        for col in yesno_cols:
            series = df[col].astype(str).str.strip().str.upper()
            yesno_summary.loc["YES", col] = (series == "YES").sum()
            yesno_summary.loc["NO", col] = (series == "NO").sum()

        yesno_summary.to_excel(
            writer,
            sheet_name=sheet_name,
            startrow=startrow,
            startcol=0,
            index=True,
            index_label="Response",
        )


def process_workbook(input_path: str, output_path: str = None) -> str:
    """
    Main function you will call.

    - Reads the first sheet of the input Excel file.
    - Groups rows by column G (team + category).
    - Creates an output workbook with:
      - Sheet 1: All data
      - One sheet per group (team + category).
    - On every sheet, appends the rating and yes/no summary tables
      starting 3 blank rows after the last data row.

    Returns the path to the output workbook.
    """
    if output_path is None:
        base, ext = os.path.splitext(input_path)
        if not ext:
            ext = ".xlsx"
        output_path = base + "_processed" + ext

    # Read the first sheet
    df = pd.read_excel(input_path, sheet_name=0)

    if GROUP_COL_INDEX >= len(df.columns):
        raise ValueError(
            "Group column G is outside the available columns in the sheet."
        )

    group_col_name = df.columns[GROUP_COL_INDEX]

    # Replace missing group names, so every row belongs to some group
    df[group_col_name] = df[group_col_name].fillna("UNASSIGNED")

    with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
        used_sheet_names = set()

        # Sheet 1: full data
        all_sheet_name = make_safe_sheet_name("All_Data", used_sheet_names)
        df.to_excel(writer, sheet_name=all_sheet_name, index=False)
        append_summary_tables(df,
                              writer,
                              all_sheet_name,
                              RATING_COL_INDICES,
                              YESNO_COL_INDICES)

        # One sheet per group
        groups = df.groupby(group_col_name, sort=True)
        for group_value, group_df in groups:
            sheet_name = make_safe_sheet_name(str(group_value), used_sheet_names)
            group_df.to_excel(writer, sheet_name=sheet_name, index=False)
            append_summary_tables(group_df,
                                  writer,
                                  sheet_name,
                                  RATING_COL_INDICES,
                                  YESNO_COL_INDICES)

    return output_path
