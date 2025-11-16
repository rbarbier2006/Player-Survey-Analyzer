# survey_processor.py

import os
import re
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages

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
        rating_summary = pd.DataFrame(
            index=index,
            columns=rating_cols,
            dtype=float,
        )

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
        yesno_summary = pd.DataFrame(
            index=yesno_index,
            columns=yesno_cols,
            dtype=float,
        )

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
    - Pie charts for each Yes/No question, with labels "XX%, N".
    - A pie chart for the single-choice question (column O), also with "XX%, N".
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
        row = startrow_bottom
        col = 0

        for q in yesno_info["questions"]:
            col_name = q["name"]
            if col_name not in df.columns:
                continue

            series = df[col_name].astype(str).str.strip().str.upper()
            yes_count = int((series == "YES").sum())
            no_count = int((series == "NO").sum())
            total = yes_count + no_count

            if total == 0:
                # Nothing to chart
                continue

            # Build custom labels "XX%, N"
            yes_pct = 100.0 * yes_count / total
            no_pct = 100.0 * no_count / total
            custom_labels = [
                {"value": f"{yes_pct:.0f}%, {yes_count}"},
                {"value": f"{no_pct:.0f}%, {no_count}"},
            ]

            # Data for the pie chart comes from the YES/NO summary table
            y_start = yesno_info["startrow"]
            cat_first_row = y_start + 1
            cat_last_row = y_start + 2  # YES, NO
            cat_col = 0  # "Response"
            val_col = q["col_index"]

            chart = workbook.add_chart({"type": "pie"})
            chart.add_series(
                {
                    "name": col_name,
                    "categories": [sheet_name, cat_first_row, cat_col,
                                   cat_last_row, cat_col],
                    "values": [sheet_name, cat_first_row, val_col,
                               cat_last_row, val_col],
                    "data_labels": {
                        "value": True,
                        "custom": custom_labels,
                    },
                }
            )

            chart.set_title({"name": col_name})
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

        # Write the Choice/Count table
        choice_df.to_excel(
            writer,
            sheet_name=sheet_name,
            startrow=table_startrow,
            startcol=table_startcol,
            index=False,
        )

        counts = choice_df["Count"].astype(int).tolist()
        total = sum(counts)
        if total > 0:
            custom_labels = []
            for c in counts:
                pct = 100.0 * c / total
                custom_labels.append({"value": f"{pct:.0f}%, {int(c)}"})
        else:
            custom_labels = [{"value": "0%, 0"} for _ in counts]

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
                    "value": True,
                    "custom": custom_labels,
                },
            }
        )

        chart.set_title({"name": choice_col_name})
        worksheet.insert_chart(table_startrow, table_startcol + 4, chart)


def process_workbook(input_path: str, output_path: str = None) -> str:
    """
    Main function you will call.

    - Reads the first sheet of the input Excel file.
    - Groups rows by column G (team + category).
    - Creates an output workbook with:
      - Sheet 1: All data
      - One sheet per group (team + category).
    - On every sheet, appends:
      - rating and yes/no summary tables
      - "1-3 star reviews" names
      - "NO replies" names
      - column and pie charts at the very bottom
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


# -------------------------------------------------------------------
# PDF REPORT GENERATION (from processed workbook)
# -------------------------------------------------------------------

def safe_filename(name: str) -> str:
    """Make a reasonably safe filename from a team name."""
    bad_chars = r'\/:*?"<>|'
    fname = "".join("_" if c in bad_chars else c for c in str(name))
    return fname.strip() or "Report"


def extract_data_block(df_full: pd.DataFrame) -> pd.DataFrame:
    """
    In the processed workbook, each sheet has:
      - original data rows at the top
      - 3 blank rows
      - summary tables, detail tables, etc.

    When we read it with pandas, everything comes as one big DataFrame.
    We want ONLY the original data block, so we:
      - look for the first row where *all* columns are NaN
      - keep everything *above* that row.
    """
    mask_all_nan = df_full.isna().all(axis=1)
    nan_rows = df_full[mask_all_nan]

    if nan_rows.empty:
        return df_full.reset_index(drop=True)

    first_blank_idx = nan_rows.index[0]
    data_df = df_full.loc[: first_blank_idx - 1].reset_index(drop=True)
    return data_df


def compute_counts_and_low(
    df_team: pd.DataFrame,
    question_index: int,
):
    """
    For one question column in one team:

    - counts: Series indexed 1..5 with counts
    - avg: mean rating (or None if no data)
    - low_entries: list of "Player Name, (X★)" for X in {1,2,3}
    - question_name: column header text
    """
    cols = list(df_team.columns)
    if question_index >= len(cols):
        return pd.Series(dtype=float), None, [], ""

    question_name = cols[question_index]

    ratings = pd.to_numeric(df_team.iloc[:, question_index], errors="coerce")
    valid = ratings.dropna()

    counts = valid.value_counts().reindex([1, 2, 3, 4, 5], fill_value=0)
    avg = float(valid.mean()) if not valid.empty else None

    low_entries: list[str] = []
    if PLAYER_NAME_INDEX < len(cols):
        player_col = cols[PLAYER_NAME_INDEX]
        for _, row in df_team.iterrows():
            value = row.iloc[question_index]
            if pd.isna(value):
                continue
            try:
                rating_int = int(value)
            except (ValueError, TypeError):
                continue
            if rating_int in (1, 2, 3):
                name = str(row[player_col])
                low_entries.append(f"{name}, ({rating_int}★)")

    return counts, avg, low_entries, question_name


def add_question_page(
    pdf: PdfPages,
    df_team: pd.DataFrame,
    question_index: int,
    team_label: str,
    cycle_label: str,
) -> None:
    """
    Create ONE PDF page for ONE star question for ONE team.

    Layout: chart on top, table of 1–3 star reviews below it.
    """
    counts, avg, low_entries, question_name = compute_counts_and_low(
        df_team, question_index
    )

    if question_name == "":
        return

    fig = plt.figure(figsize=(8.5, 11))  # US Letter
    gs = fig.add_gridspec(nrows=3, ncols=1, height_ratios=[3, 0.3, 4])

    ax_chart = fig.add_subplot(gs[0])
    ax_table = fig.add_subplot(gs[2])
    ax_table.axis("off")

    x_vals = [1, 2, 3, 4, 5]
    y_vals = [counts.get(x, 0) for x in x_vals]

    ax_chart.bar(x_vals, y_vals)
    ax_chart.set_xlabel("# of Stars")
    ax_chart.set_ylabel("Player Count")

    if avg is not None and not np.isnan(avg):
        avg_text = f"{avg:.2f}"
    else:
        avg_text = "N/A"

    chart_title = f"{question_name}  (Avg = {avg_text})"
    ax_chart.set_title(chart_title, fontsize=12, pad=12)
    ax_chart.set_xticks(x_vals)
    ax_chart.set_ylim(0, max(y_vals + [1]) * 1.2)

    ax_table.set_title("1–3 Star Reviews", fontsize=11, loc="left", pad=6)

    if low_entries:
        cell_text = [[entry] for entry in low_entries]
        table = ax_table.table(
            cellText=cell_text,
            colLabels=["Player, (Rating)"],
            loc="upper left",
            cellLoc="left",
        )
        table.auto_set_font_size(False)
        table.set_fontsize(9)
        table.scale(1, 1.1)
    else:
        ax_table.text(
            0.0,
            0.9,
            "No 1–3 star reviews for this question.",
            fontsize=10,
            va="top",
        )

    fig.suptitle(
        f"{team_label} – {cycle_label}",
        fontsize=14,
        y=0.98,
    )

    fig.tight_layout(rect=[0, 0, 1, 0.96])
    pdf.savefig(fig)
    plt.close(fig)


def build_team_report(
    df_team: pd.DataFrame,
    output_pdf_path: str,
    team_label: str,
    cycle_label: str,
) -> None:
    """
    Build a full PDF for ONE team (or 'All Teams'), one page per star question.
    """
    rating_indices = [
        idx for idx in RATING_COL_INDICES if idx < len(df_team.columns)
    ]
    if not rating_indices:
        return

    with PdfPages(output_pdf_path) as pdf:
        for q_idx in rating_indices:
            add_question_page(pdf, df_team, q_idx, team_label, cycle_label)


def create_pdf_reports_from_processed(
    processed_excel_path: str,
    cycle_label: str,
    output_dir: str = "pdf_reports",
) -> None:
    """
    Use the *processed* workbook (output of process_workbook) and generate:
      - 'All Teams - Cycle X.pdf'
      - '<Team Name> - Cycle X.pdf' for each team in column G.
    """
    os.makedirs(output_dir, exist_ok=True)

    xls = pd.ExcelFile(processed_excel_path)
    if "All_Data" in xls.sheet_names:
        df_full = xls.parse("All_Data")
    else:
        df_full = xls.parse(0)

    df_data = extract_data_block(df_full)

    if GROUP_COL_INDEX >= len(df_data.columns):
        raise ValueError(
            "Group column index is outside available columns in the sheet."
        )

    group_col_name = df_data.columns[GROUP_COL_INDEX]
    df_data[group_col_name] = df_data[group_col_name].fillna("UNASSIGNED")

    # Overall "All Teams" report
    all_label = "All Teams"
    all_fname = f"{safe_filename(all_label)} - {cycle_label}.pdf"
    all_path = os.path.join(output_dir, all_fname)
    build_team_report(df_data, all_path, all_label, cycle_label)

    # One report per team
    for team_value, df_team in df_data.groupby(group_col_name, sort=True):
        team_label = str(team_value)
        fname = f"{safe_filename(team_label)} - {cycle_label}.pdf"
        pdf_path = os.path.join(output_dir, fname)
        build_team_report(df_team, pdf_path, team_label, cycle_label)
