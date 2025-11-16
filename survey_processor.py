# survey_processor.py

import os
import re
import numpy as np
import pandas as pd

# Column layout based on your description
# F: player name
# G: team + category (grouping key)
# H, I, J, K, L, P, Q: 1-5 star rating questions
# M, N: Yes/No questions
# O: single-choice "values" question

PLAYER_NAME_COL_LETTER = "F"
RATING_COL_LETTERS = ["H", "I", "J", "K", "L", "P", "Q"]
YESNO_COL_LETTERS = ["M", "N"]
CHOICE_COL_LETTER = "O"


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


# Precompute indices we need
PLAYER_NAME_INDEX = col_letter_to_index(PLAYER_NAME_COL_LETTER)
RATING_COL_INDICES = [col_letter_to_index(c) for c in RATING_COL_LETTERS]
YESNO_COL_INDICES = [col_letter_to_index(c) for c in YESNO_COL_LETTERS]
CHOICE_COL_INDEX = col_letter_to_index(CHOICE_COL_LETTER)
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


def append_summary_tables(
    df: pd.DataFrame,
    writer: pd.ExcelWriter,
    sheet_name: str,
    rating_indices,
    yesno_indices,
):
    """
    Append the "score table" summaries to a sheet:

    - ratings table (1..5 counts + average) for all rating columns
    - yes/no counts for all yes/no columns

    All tables start 3 blank rows after the last data row.

    Returns:
        next_startrow: first empty row AFTER all summary tables.
        rating_info: metadata needed to build rating charts.
        yesno_info: metadata needed to build yes/no charts.
    """
    n_rows = df.shape[0]

    # 3 blank rows after header + data
    startrow = n_rows + 4

    cols = list(df.columns)
    rating_cols = [cols[i] for i in rating_indices if i < len(cols)]
    yesno_cols = [cols[i] for i in yesno_indices if i < len(cols)]

    rating_info = None
    yesno_info = None

    # ----- Rating questions (1-5 stars + average) -----
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

        rating_startrow = startrow

        rating_summary.to_excel(
            writer,
            sheet_name=sheet_name,
            startrow=rating_startrow,
            startcol=0,
            index=True,
            index_label="Score",
        )

        questions_meta = []
        for j, col_name in enumerate(rating_cols):
            avg_val = rating_summary.loc["Average", col_name]
            try:
                avg_val = float(avg_val)
            except (TypeError, ValueError):
                avg_val = None
            questions_meta.append(
                {
                    "name": col_name,
                    "col_index": 1 + j,  # Excel column index for values
                    "average": avg_val,
                }
            )

        rating_info = {
            "startrow": rating_startrow,
            "questions": questions_meta,
        }

        table_height = rating_summary.shape[0] + 1  # data rows + header
        startrow = rating_startrow + table_height + 1

    # ----- Yes/No questions -----
    if yesno_cols:
        yesno_index = ["YES", "NO"]
        yesno_summary = pd.DataFrame(index=yesno_index,
                                     columns=yesno_cols,
                                     dtype=float)

        for col in yesno_cols:
            series = df[col].astype(str).str.strip().str.upper()
            yesno_summary.loc["YES", col] = (series == "YES").sum()
            yesno_summary.loc["NO", col] = (series == "NO").sum()

        yesno_startrow = startrow

        yesno_summary.to_excel(
            writer,
            sheet_name=sheet_name,
            startrow=yesno_startrow,
            startcol=0,
            index=True,
            index_label="Response",
        )

        questions_meta = []
        for j, col_name in enumerate(yesno_cols):
            questions_meta.append(
                {
                    "name": col_name,
                    "col_index": 1 + j,  # column for values
                }
            )

        yesno_info = {
            "startrow": yesno_startrow,
            "questions": questions_meta,
        }

        table_height = yesno_summary.shape[0] + 1
        startrow = yesno_startrow + table_height + 1

    return startrow, rating_info, yesno_info


def build_low_ratings_table(
    df: pd.DataFrame,
    rating_indices,
    player_index,
):
    """
    Build a table listing all 1, 2, and 3 star answers.

    Each cell is "Player Name, (X★)" for the corresponding question.
    Columns = rating questions; rows = entries, padded with empty strings.
    """
    cols = list(df.columns)
    if player_index >= len(cols):
        return None

    player_col = cols[player_index]
    low_lists = {}

    for idx in rating_indices:
        if idx >= len(cols):
            continue
        col = cols[idx]
        entries = []

        for _, row in df.iterrows():
            value = row.iloc[idx]
            if pd.isna(value):
                continue
            try:
                rating = int(value)
            except (ValueError, TypeError):
                continue
            if rating in (1, 2, 3):
                name = str(row[player_col])
                entries.append(f"{name}, ({rating}★)")

        low_lists[col] = entries

    if not low_lists:
        return None

    max_len = max((len(v) for v in low_lists.values()), default=0)
    if max_len == 0:
        return None

    data = {}
    for question, vals in low_lists.items():
        padded = vals + [""] * (max_len - len(vals))
        data[question] = padded

    low_df = pd.DataFrame(data)
    return low_df


def build_no_answers_table(
    df: pd.DataFrame,
    yesno_indices,
    player_index,
) -> pd.DataFrame | None:
    """
    Build a table listing all 'NO' answers for Yes/No questions.

    Each cell is "Player Name, (NO)" for the corresponding question.
    """
    cols = list(df.columns)
    if player_index >= len(cols):
        return None

    player_col = cols[player_index]
    no_lists: dict[str, list[str]] = {}

    for idx in yesno_indices:
        if idx >= len(cols):
            continue
        col = cols[idx]
        entries: list[str] = []

        for _, row in df.iterrows():
            value = row.iloc[idx]
            if pd.isna(value):
                continue
            value_str = str(value).strip().upper()
            if value_str == "NO":
                name = str(row[player_col])
                entries.append(f"{name}, (NO)")

        no_lists[col] = entries

    if not no_lists:
        return None

    max_len = max((len(v) for v in no_lists.values()), default=0)
    if max_len == 0:
        return None

    data = {}
    for question, vals in no_lists.items():
        padded = vals + [""] * (max_len - len(vals))
        data[question] = padded

    no_df = pd.DataFrame(data)
    return no_df


def build_choice_counts(
    df: pd.DataFrame,
    choice_index: int,
) -> pd.DataFrame | None:
    """
    Build a frequency table for the single-choice question (column O).

    Returns a DataFrame with columns ["Choice", "Count"] or None if empty.
    """
    if choice_index >= len(df.columns):
        return None

    series = df.iloc[:, choice_index].dropna()
    if series.empty:
        return None

    counts = series.value_counts().sort_index()
    table_df = pd.DataFrame(
        {"Choice": counts.index.tolist(), "Count": counts.values.tolist()}
    )
    return table_df


def append_detail_tables(
    df: pd.DataFrame,
    writer: pd.ExcelWriter,
    sheet_name: str,
    startrow: int,
    rating_indices,
    yesno_indices,
    player_index,
) -> int:
    """
    Part 2:

    Under the score tables, append:

    - A "1-3 Star Reviews" section listing all players who gave 1, 2, or 3 stars.
    - A "NO Replies" section listing all players who answered "NO" to Yes/No questions.
    """
    worksheet = writer.sheets[sheet_name]

    # 1-3 Star Reviews
    low_df = build_low_ratings_table(df, rating_indices, player_index)
    if low_df is not None:
        n_rows = len(low_df)
        low_df.insert(0, "1-3 Star Reviews", list(range(1, n_rows + 1)))

        low_df.to_excel(
            writer,
            sheet_name=sheet_name,
            startrow=startrow,
            startcol=0,
            index=False,
        )
        startrow = startrow + n_rows + 1 + 2
    else:
        worksheet.write(startrow, 0, "0 1-3 star reviews")
        startrow = startrow + 2

    # NO replies for Yes/No questions
    no_df = build_no_answers_table(df, yesno_indices, player_index)
    if no_df is not None:
        n_rows = len(no_df)
        no_df.insert(0, "NO Replies", list(range(1, n_rows + 1)))

        no_df.to_excel(
            writer,
            sheet_name=sheet_name,
            startrow=startrow,
            startcol=0,
            index=False,
        )
        startrow = startrow + n_rows + 1 + 2
    else:
        worksheet.write(startrow, 0, "0 no answers")
        startrow = startrow + 2

    return startrow


def append_charts(
    writer: pd.ExcelWriter,
    sheet_name: str,
    rating_info,
    yesno_info,
    startrow_bottom: int,
    df: pd.DataFrame,
) -> None:
    """
    Part 3:

    At the very bottom of the sheet, create:
    - Column charts (histograms) for each rating question.
    - Pie charts for each Yes/No question, showing "percent, count".
    - A pie chart for the single-choice question in column O, also showing "percent, count".
    """
    workbook = writer.book
    worksheet = writer.sheets[sheet_name]

    row = startrow_bottom
    col = 0

    # ----- Rating charts (column) -----
    if rating_info is not None and rating_info.get("questions"):
        r_start = rating_info["startrow"]
        cat_first_row = r_start + 1
        cat_last_row = r_start + 5
        cat_col = 0  # "Score" column

        for q in rating_info["questions"]:
            chart = workbook.add_chart({"type": "column"})
            val_col = q["col_index"]

            chart.add_series(
                {
                    "name": q["name"],
                    "categories": [sheet_name, cat_first_row, cat_col,
                                   cat_last_row, cat_col],
                    "values": [sheet_name, cat_first_row, val_col,
                               cat_last_row, val_col],
                }
            )

            avg = q.get("average")
            if avg is not None and not np.isnan(avg):
                title_text = f"{q['name']} (Avg = {avg:.2f})"
            else:
                title_text = q["name"]

            chart.set_title(
                {
                    "name": title_text,
                    "name_font": {"size": 12, "bold": True},
                }
            )
            chart.set_legend({"none": True})
            chart.set_x_axis({"name": "# of Stars"})
            chart.set_y_axis({"name": "Player Count"})

            worksheet.insert_chart(row, col, chart)

            col += 8
            if col > 16:
                col = 0
                row += 18

        row += 18
        startrow_bottom = row

    # ----- Yes/No pie charts -----
    if yesno_info is not None and yesno_info.get("questions"):
        y_start = yesno_info["startrow"]
        cat_first_row = y_start + 1
        cat_last_row = y_start + 2  # YES, NO
        cat_col = 0  # "Response"

        row = startrow_bottom
        col = 0

        for q in yesno_info["questions"]:
            chart = workbook.add_chart({"type": "pie"})
            val_col = q["col_index"]

            chart.add_series(
                {
                    "name": q["name"],
                    "categories": [sheet_name, cat_first_row, cat_col,
                                   cat_last_row, cat_col],
                    "values": [sheet_name, cat_first_row, val_col,
                               cat_last_row, val_col],
                    "data_labels": {
                        "percentage": True,
                        "value": True,
                        "separator": ", ",
                    },
                }
            )

            chart.set_title({"name": q["name"]})
            worksheet.insert_chart(row, col, chart)

            col += 8
            if col > 16:
                col = 0
                row += 18

        row += 18

    # ----- Pie chart for 5-value choice in column O -----
    choice_df = build_choice_counts(df, CHOICE_COL_INDEX)
    if choice_df is not None:
        table_startrow = row
        table_startcol = 0

        choice_df.to_excel(
            writer,
            sheet_name=sheet_name,
            startrow=table_startrow,
            startcol=table_startcol,
            index=False,
        )

        chart = workbook.add_chart({"type": "pie"})
        choice_col_name = df.columns[CHOICE_COL_INDEX]

        chart.add_series(
            {
                "name": choice_col_name,
                "categories": [
                    sheet_name,
                    table_startrow + 1,
                    table_startcol,
                    table_startrow + len(choice_df),
                    table_startcol,
                ],
                "values": [
                    sheet_name,
                    table_startrow + 1,
                    table_startcol + 1,
                    table_startrow + len(choice_df),
                    table_startcol + 1,
                ],
                "data_labels": {
                    "percentage": True,
                    "value": True,
                    "separator": ", ",
                },
            }
        )

        chart.set_title({"name": choice_col_name})
        worksheet.insert_chart(table_startrow, table_startcol + 4, chart)


def process_workbook(input_path: str, output_path: str = None) -> str:
    """
    Main function you will call.
    """
    if output_path is None:
        base, ext = os.path.splitext(input_path)
        if not ext:
            ext = ".xlsx"
        output_path = base + "_processed" + ext

    df = pd.read_excel(input_path, sheet_name=0)

    if GROUP_COL_INDEX >= len(df.columns):
        raise ValueError(
            "Group column G is outside the available columns in the sheet."
        )

    group_col_name = df.columns[GROUP_COL_INDEX]
    df[group_col_name] = df[group_col_name].fillna("UNASSIGNED")

    with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
        used_sheet_names = set()

        # All-data sheet
        all_sheet_name = make_safe_sheet_name("All_Data", used_sheet_names)
        df.to_excel(writer, sheet_name=all_sheet_name, index=False)
        next_row, rating_info, yesno_info = append_summary_tables(
            df,
            writer,
            all_sheet_name,
            RATING_COL_INDICES,
            YESNO_COL_INDICES,
        )
        next_row = append_detail_tables(
            df,
            writer,
            all_sheet_name,
            next_row,
            RATING_COL_INDICES,
            YESNO_COL_INDICES,
            PLAYER_NAME_INDEX,
        )
        append_charts(
            writer,
            all_sheet_name,
            rating_info,
            yesno_info,
            next_row,
            df,
        )

        # One sheet per group
        for group_value, group_df in df.groupby(group_col_name, sort=True):
            sheet_name = make_safe_sheet_name(str(group_value), used_sheet_names)
            group_df.to_excel(writer, sheet_name=sheet_name, index=False)

            next_row, rating_info, yesno_info = append_summary_tables(
                group_df,
                writer,
                sheet_name,
                RATING_COL_INDICES,
                YESNO_COL_INDICES,
            )
            next_row = append_detail_tables(
                group_df,
                writer,
                sheet_name,
                next_row,
                RATING_COL_INDICES,
                YESNO_COL_INDICES,
                PLAYER_NAME_INDEX,
            )
            append_charts(
                writer,
                sheet_name,
                rating_info,
                yesno_info,
                next_row,
                group_df,
            )

    return output_path
