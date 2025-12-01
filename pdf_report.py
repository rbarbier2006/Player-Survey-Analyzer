# pdf_report.py

import os
import textwrap
from typing import Optional, List, Dict, Any

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import re
from matplotlib.backends.backend_pdf import PdfPages

# Mapping from team name -> coach name in the PDF titles :contentReference[oaicite:0]{index=0}
TEAM_COACH_MAP = {
    "MLS HG (T1) U19": "Jorge",
    "MLS HG (T1) U17": "Chris",
    "MLS HG (T1) U16": "David K",
    "MLS HG (T1) U15": "Jorge",
    "MLS HG (T1) U14": "David K",
    "MLS HG (T1) U13": "Chris M",

    "MLS AD (T2) U19": "Michael",
    "MLS AD (T2) U17": "Michael",
    "MLS AD (T2) U16": "Miguel",
    "MLS AD (T2) U15": "Miguel",
    "MLS AD (T2) U14": "?",          # unknown coach
    "MLS AD (T2) U13": "Miguel",

    "TX2 U19": "Jesus",
    "TX2 U17": "Fernando",
    "TX2 U16": "Jesus",
    "TX2 U15": "Claudia",
    "TX2 U14": "Rene/Claudia",
    "TX2 U13": "Claudia/Rene",
    "TX2 U12": "Armando",
    "TX2 U11": "Armando",

    "Athenians U16": "Rumen",
    "Athenians U13": "Keeley",
    "Athenians WDDOA U12": "Keeley",
    "Athenians WDDOA U11": "Robert",
    "Athenians PDF U10": "Robert",
    "Athenians PDF U9": "Robert",

    "WDDOA U12": "Adam",
    "WDDOA U11": "Adam",

    "PDF White U10": "Steven",
    "PDF White U9": "Steven",
    "PDF Red U10": "Pablo",
    "PDF Red U9": "Pablo",
}


# Mapping from chart number -> pretty label used in PDF tables
CHART_LABELS = {
    1: "(1)Safety and Support",
    2: "(2)Improvement",
    3: "(3)Instructions and Feedback",
    4: "(4)Coaches Listening",
    5: "(5)Effort and Discipline",
    6: "(6)SC Value Alignment",
    7: "(7)Overall Experience",
    8: "(8)Team Belonging",
    9: "(9)Cycle Enjoyment",
}

def _compose_group_title(title_label: str, cycle_label: str) -> str:
    """
    Build 'Team - Coach - Cycle X' for team pages and 'All Teams - Cycle X'
    for the combined page.

    If title_label already contains a coach (e.g. 'Team - Miguel'),
    we just append ' - Cycle X' to avoid duplicating.
    """
    base = str(title_label).strip()

    # All Teams page: no coach
    if base == "All Teams":
        return f"All Teams - {cycle_label}"

    # If caller already supplied 'Team - Coach', don't add again
    if " - " in base:
        return f"{base} - {cycle_label}"

    # Otherwise look up coach from the map, fallback to '?'
    coach = TEAM_COACH_MAP.get(base, "?")
    return f"{base} - {coach} - {cycle_label}"


from excel_processor import (
    PLAYER_NAME_INDEX,
    RATING_COL_INDICES,
    YESNO_COL_INDICES,
    CHOICE_COL_INDEX,
    GROUP_COL_INDEX,
    build_low_ratings_table,
    build_no_answers_table,
    build_choice_counts,
)


def _build_plot_metadata(df_group: pd.DataFrame) -> List[Dict[str, Any]]:
    """
    For a given group (All teams or a single team), build metadata
    for every plot we want to draw, and assign it a number
    starting at 1 for THIS group.
    """
    cols = list(df_group.columns)

    rating_indices = [idx for idx in RATING_COL_INDICES if idx < len(cols)]
    yesno_indices = [idx for idx in YESNO_COL_INDICES if idx < len(cols)]
    has_choice = CHOICE_COL_INDEX < len(cols)

    meta: List[Dict[str, Any]] = []
    number = 1  # restart numbering for each group

    for idx in rating_indices:
        meta.append(
            {
                "ptype": "rating",
                "idx": idx,
                "col_name": cols[idx],
                "number": number,
            }
        )
        number += 1

    for idx in yesno_indices:
        meta.append(
            {
                "ptype": "yesno",
                "idx": idx,
                "col_name": cols[idx],
                "number": number,
            }
        )
        number += 1

    if has_choice:
        meta.append(
            {
                "ptype": "choice",
                "idx": CHOICE_COL_INDEX,
                "col_name": cols[CHOICE_COL_INDEX],
                "number": number,
            }
        )

    return meta


def _add_group_charts_page_to_pdf(
    pdf: PdfPages,
    df_group: pd.DataFrame,
    title_label: str,
    cycle_label: str,
    plots_meta: List[Dict[str, Any]],
) -> None:
    """
    Page 1 for a group: all charts laid out in a grid, each labelled
    with a big number in the top-left (1,2,3,...) that resets per group.
    """
    if not plots_meta:
        return

    n_plots = len(plots_meta)
    ncols = 3
    nrows = int(np.ceil(n_plots / ncols))

    fig, axes = plt.subplots(nrows=nrows, ncols=ncols, figsize=(8.5, 11))

    # Normalize axes to a flat list
    if nrows == 1 and ncols == 1:
        axes = np.array([[axes]])
    elif nrows == 1:
        axes = np.array([axes])
    axes_flat = axes.flatten()

    # Turn off any unused axes
    for ax in axes_flat[n_plots:]:
        ax.axis("off")

    for ax, meta in zip(axes_flat, plots_meta):
        ptype = meta["ptype"]
        idx = meta["idx"]
        col_name = meta["col_name"]
        number = meta["number"]

        # Big chart number in top-left
        ax.text(
            0.02,
            0.98,
            str(number),
            transform=ax.transAxes,
            ha="left",
            va="top",
            fontsize=10,
            fontweight="bold",
        )

        # Wrap question text so titles do not overlap
        wrapped_name = textwrap.fill(col_name, width=40)

        if ptype == "rating":
            series = pd.to_numeric(df_group.iloc[:, idx], errors="coerce").dropna()
            counts = series.value_counts().reindex([1, 2, 3, 4, 5], fill_value=0)
            ax.bar(range(1, 6), counts.values)

            avg = series.mean() if not series.empty else None
            if avg is not None and not np.isnan(avg):
                title = f"{wrapped_name}\n(Avg = {avg:.2f})"
            else:
                title = wrapped_name

            ax.set_title(title, fontsize=8)
            ax.set_xlabel("# of Stars", fontsize=7)
            ax.set_ylabel("Player Count", fontsize=7)
            ax.tick_params(labelsize=7)
            ax.set_ylim(0, max(counts.values.tolist() + [1]) * 1.2)

        elif ptype == "yesno":
            series = df_group.iloc[:, idx].astype(str).str.strip().str.upper()
            yes_count = int((series == "YES").sum())
            no_count = int((series == "NO").sum())
            data = [yes_count, no_count]
            labels = ["YES", "NO"]

            if yes_count + no_count == 0:
                ax.text(
                    0.5,
                    0.5,
                    "No data",
                    ha="center",
                    va="center",
                    fontsize=8,
                )
                ax.axis("off")
            else:

                def make_label(pct, allvals=data):
                    total = sum(allvals)
                    if total == 0:
                        count = 0
                    else:
                        count = int(round(pct * total / 100.0))
                    return f"{pct:.0f}%, {count}"

                ax.pie(
                    data,
                    labels=labels,
                    autopct=make_label,
                    textprops={"fontsize": 7},
                )
                ax.set_title(wrapped_name, fontsize=8)

        elif ptype == "choice":
            series = df_group.iloc[:, idx].dropna()
            counts = series.value_counts()
            data = counts.values
            labels = counts.index.tolist()

            if len(data) == 0:
                ax.text(
                    0.5,
                    0.5,
                    "No data",
                    ha="center",
                    va="center",
                    fontsize=8,
                )
                ax.axis("off")
            else:

                def make_label(pct, allvals=data):
                    total = sum(allvals)
                    if total == 0:
                        count = 0
                    else:
                        count = int(round(pct * total / 100.0))
                    return f"{pct:.0f}%, {count}"

                ax.pie(
                    data,
                    labels=labels,
                    autopct=make_label,
                    textprops={"fontsize": 7},
                )
                ax.set_title(wrapped_name, fontsize=8)

    fig.suptitle(_compose_group_title(title_label, cycle_label), fontsize=12)
    fig.tight_layout(rect=[0, 0, 1, 0.94])
    pdf.savefig(fig)
    plt.close(fig)


def _build_all_players_grid(
    df_group: pd.DataFrame,
    max_cols: int = 6,
) -> Optional[pd.DataFrame]:
    """
    Build a compact grid with all players who completed the survey
    for this group.

    - Uses up to max_cols narrow columns (default 6).
    - Number of rows is just enough to hold all players.
    """
    if PLAYER_NAME_INDEX >= len(df_group.columns):
        return None

    names = (
        df_group.iloc[:, PLAYER_NAME_INDEX]
        .dropna()
        .astype(str)
        .str.strip()
    )
    names = names[names != ""].drop_duplicates()

    if names.empty:
        return None

    n = len(names)
    ncols = min(max_cols, n)
    nrows = int(np.ceil(n / ncols))

    grid = [["" for _ in range(ncols)] for _ in range(nrows)]
    i = 0
    for c in range(ncols):
        for r in range(nrows):
            if i >= n:
                break
            grid[r][c] = names.iloc[i]
            i += 1

    col_labels = [f"Players {i+1}" for i in range(ncols)]
    players_df = pd.DataFrame(grid, columns=col_labels)

    players_df = players_df[(players_df != "").any(axis=1)]
    return players_df


def _build_comments_table(df_group: pd.DataFrame) -> Optional[pd.DataFrame]:
    """
    Build a table with all free-text comments/suggestions for this group.

    It looks for any column whose header contains 'comment' or 'suggest'
    (case-insensitive). For each non-empty cell in those columns, we
    create a row: Player | Comment / Suggestion.
    """
    if PLAYER_NAME_INDEX >= len(df_group.columns):
        return None

    cols = list(df_group.columns)

    comment_indices: List[int] = []
    for i, name in enumerate(cols):
        name_lower = str(name).lower()
        if "comment" in name_lower or "suggest" in name_lower:
            comment_indices.append(i)

    if not comment_indices:
        return None

    rows: List[List[str]] = []
    for _, row in df_group.iterrows():
        player = str(row.iloc[PLAYER_NAME_INDEX]).strip()
        if not player:
            continue

        for idx in comment_indices:
            val = row.iloc[idx]
            if pd.isna(val):
                continue
            text = str(val).strip()
            if not text:
                continue

            col_label = cols[idx]
            if len(comment_indices) > 1:
                text_final = f"[{col_label}] {text}"
            else:
                text_final = text

            rows.append([player, text_final])

    if not rows:
        return None

    comments_df = pd.DataFrame(rows, columns=["Player", "Comment / Suggestion"])

    comments_df["Comment / Suggestion"] = comments_df[
        "Comment / Suggestion"
    ].apply(lambda s: textwrap.fill(str(s), width=70))

    return comments_df

def _filter_low_df_by_max_star(low_df: pd.DataFrame, max_star: int = 2) -> pd.DataFrame:
    """
    Given the low-ratings table whose cells look like 'Name, (3★)',
    keep only rows with rating <= max_star (per column).
    Columns are preserved; we just shorten / pad the lists.
    """
    pattern = re.compile(r"\((\d)★\)")
    new_cols: Dict[str, List[str]] = {}
    max_len = 0

    for col in low_df.columns:
        filtered: List[str] = []
        for val in low_df[col]:
            s = str(val).strip()
            if not s:
                continue
            m = pattern.search(s)
            if m:
                rating = int(m.group(1))
                if rating <= max_star:
                    filtered.append(s)
        new_cols[col] = filtered
        max_len = max(max_len, len(filtered))

    # If nothing left at all, just return one blank row so the header still shows
    if max_len == 0:
        for col in new_cols:
            new_cols[col] = [""]
        return pd.DataFrame(new_cols)

    # Pad each column to the same length
    for col, vals in new_cols.items():
        new_cols[col] = vals + [""] * (max_len - len(vals))

    return pd.DataFrame(new_cols)


def _add_group_tables_page_to_pdf(
    pdf: PdfPages,
    df_group: pd.DataFrame,
    title_label: str,
    cycle_label: str,
    plots_meta: List[Dict[str, Any]],
    is_all_teams: bool,
) -> None:
    """
    Page 2 for a group.

    For ALL TEAMS:
      - 1–2 star reviews (columns = chart numbers, but headers use CHART_LABELS)
      - "NO" replies
      - Completion summary

    For INDIVIDUAL TEAMS:
      - 1–3 star reviews
      - "NO" replies
      - Players who completed this survey
      - Comments / suggestions
    """
    rating_indices = [m["idx"] for m in plots_meta if m["ptype"] == "rating"]
    yesno_indices = [m["idx"] for m in plots_meta if m["ptype"] == "yesno"]

    rating_number_by_name = {
        m["col_name"]: m["number"] for m in plots_meta if m["ptype"] == "rating"
    }
    yesno_number_by_name = {
        m["col_name"]: m["number"] for m in plots_meta if m["ptype"] == "yesno"
    }

    # ----- 1–3 (or 1–2) star reviews table + labels -----
    low_df: Optional[pd.DataFrame] = None
    low_labels: Optional[List[str]] = None

    if rating_indices:
        low_df = build_low_ratings_table(df_group, rating_indices, PLAYER_NAME_INDEX)

        if low_df is not None:
            # For ALL TEAMS only, keep 1–2★ results (teams keep 1–3★)
            if is_all_teams:
                low_df = _filter_low_df_by_max_star(low_df, max_star=2)

            low_labels = []
            for col in low_df.columns:
                num = rating_number_by_name.get(col)
                if num is None:
                    try:
                        num = int(str(col))
                    except ValueError:
                        num = None

                if num is not None and num in CHART_LABELS:
                    low_labels.append(CHART_LABELS[num])
                else:
                    low_labels.append(str(col))

    # ----- "NO" replies table + labels -----
    no_df: Optional[pd.DataFrame] = None
    no_labels: Optional[List[str]] = None

    if yesno_indices:
        no_df = build_no_answers_table(df_group, yesno_indices, PLAYER_NAME_INDEX)
        if no_df is not None:
            no_labels = []
            for col in no_df.columns:
                num = yesno_number_by_name.get(col)
                if num is None:
                    try:
                        num = int(str(col))
                    except ValueError:
                        num = None

                if num is not None and num in CHART_LABELS:
                    no_labels.append(CHART_LABELS[num])
                else:
                    no_labels.append(str(col))

    # ----- Players / completion / comments -----
    players_df: Optional[pd.DataFrame] = None
    completion_df: Optional[pd.DataFrame] = None
    comments_df: Optional[pd.DataFrame] = None

    if is_all_teams:
        if PLAYER_NAME_INDEX < len(df_group.columns):
            names = (
                df_group.iloc[:, PLAYER_NAME_INDEX]
                .dropna()
                .astype(str)
                .str.strip()
            )
            names = names[names != ""]
            total = int(names.nunique())
            completion_df = pd.DataFrame(
                {"Metric": ["Players who completed this survey"], "Value": [total]}
            )
    else:
        players_df = _build_all_players_grid(df_group, max_cols=6)
        comments_df = _build_comments_table(df_group)

    # If literally nothing to show, skip page
    if (
        low_df is None
        and no_df is None
        and completion_df is None
        and players_df is None
        and comments_df is None
    ):
        return

    # ----- Decide sections and relative heights -----
    sections: List[str] = []
    if low_df is not None:
        sections.append("low")
    if no_df is not None:
        sections.append("no")
    if is_all_teams:
        if completion_df is not None:
            sections.append("completion")
    else:
        if players_df is not None:
            sections.append("players")
        if comments_df is not None:
            sections.append("comments")

    nrows = len(sections)

    height_ratios: List[float] = []
    for s in sections:
        if s == "low":
            height_ratios.append(1.1)
        elif s == "no":
            height_ratios.append(0.9)
        elif s == "completion":
            height_ratios.append(0.8)
        elif s == "players":
            height_ratios.append(1.3)
        elif s == "comments":
            height_ratios.append(1.6)

    fig, axes = plt.subplots(
        nrows=nrows,
        ncols=1,
        figsize=(11, 8.5),  # landscape
        gridspec_kw={"height_ratios": height_ratios},
    )
    if nrows == 1:
        axes = [axes]

    row_idx = 0

    # --------- 1–X Star Reviews ----------
    if low_df is not None:
        ax = axes[row_idx]
        ax.axis("off")

        table = ax.table(
            cellText=low_df.values,
            colLabels=low_labels if low_labels is not None else low_df.columns,
            loc="upper left",
        )
        table.auto_set_font_size(False)
        table.set_fontsize(7)

        ncols_low = len(low_df.columns)
        if ncols_low <= 8:
            width_scale = 1.0
        elif ncols_low <= 12:
            width_scale = 0.85
        else:
            width_scale = 0.7
        table.scale(width_scale, 1.15)

        # Title text still says "1–3 Star Reviews" to match the charts;
        # you can change to "1–2" for All Teams if you want:
        title_text = "1-3 Star Reviews (columns = chart numbers)"
        if is_all_teams:
            title_text = "1-2 Star Reviews (columns = chart numbers)"

        ax.set_title(title_text, fontsize=10, pad=6)
        row_idx += 1

    # --------- "NO" Replies ----------
    if no_df is not None:
        ax = axes[row_idx]
        ax.axis("off")

        table = ax.table(
            cellText=no_df.values,
            colLabels=no_labels if no_labels is not None else no_df.columns,
            loc="upper left",
        )
        table.auto_set_font_size(False)
        table.set_fontsize(7)

        ncols_no = len(no_df.columns)
        if ncols_no <= 6:
            width_scale = 1.0
        elif ncols_no <= 10:
            width_scale = 0.9
        else:
            width_scale = 0.75
        table.scale(width_scale, 1.15)

        ax.set_title(
            '"NO" Replies (columns = chart numbers)',
            fontsize=10,
            pad=6,
        )
        row_idx += 1

    # --------- Completion summary (All Teams only) ----------
    if is_all_teams and completion_df is not None:
        ax = axes[row_idx]
        ax.axis("off")

        table = ax.table(
            cellText=completion_df.values,
            colLabels=completion_df.columns,
            loc="center",
        )
        table.auto_set_font_size(False)
        table.set_fontsize(11)
        table.scale(1.2, 1.5)

        ax.set_title("Survey completion summary", fontsize=10, pad=4)
        row_idx += 1

    # --------- Players (team pages) ----------
    if not is_all_teams and players_df is not None:
        ax = axes[row_idx]
        ax.axis("off")

        table = ax.table(
            cellText=players_df.values,
            colLabels=players_df.columns,
            loc="upper left",
        )
        table.auto_set_font_size(False)
        table.set_fontsize(8)
        table.scale(1.0, 1.5)

        ax.set_title("Players who completed this survey", fontsize=10, pad=4)
        row_idx += 1

    # --------- Comments / Suggestions (team pages) ----------
    if not is_all_teams and comments_df is not None:
        ax = axes[row_idx]
        ax.axis("off")

        table = ax.table(
            cellText=comments_df.values,
            colLabels=comments_df.columns,
            loc="upper left",
            colWidths=[0.15, 0.85],  # 15% name, 85% comment
        )
        table.auto_set_font_size(False)
        table.set_fontsize(8)
        table.scale(1.0, 2.0)

        for (r, c), cell in table.get_celld().items():
            if r == 0:
                cell.set_text_props(ha="center", va="center")
                continue
            if c == 0:
                cell.set_text_props(ha="center", va="center")
            elif c == 1:
                txt = cell.get_text()
                txt.set_ha("left")
                txt.set_va("center")
                txt.set_wrap(True)
                cell.PAD = 0.02

        ax.set_title("Comments and Suggestions", fontsize=10, pad=6)

    # ----- Final layout -----
    fig.suptitle(
        _compose_group_title(title_label, cycle_label) + " (Details)",
        fontsize=12,
    )
    fig.tight_layout(rect=[0, 0.03, 1, 0.9])
    fig.subplots_adjust(hspace=0.55)

    pdf.savefig(fig)
    plt.close(fig)

def _add_cycle_summary_page(
    pdf: PdfPages,
    df: pd.DataFrame,
    group_col_name: str,
    cycle_label: str,
) -> None:
    """
    First page of the PDF.

    Ranks teams by:
      1) how many players completed the survey (descending)
      2) Overall Experience (question 7) average (descending) as tiebreaker

    Lines look like:
      1. Team - Coach: 3 players; Overall Experience: 4.33
    """
    summary_rows: List[Dict[str, Any]] = []

    for group_value, group_df in df.groupby(group_col_name, sort=True):
        team_name = str(group_value).strip()
        if team_name == "UNASSIGNED":
            # Skip fake / missing teams in the summary
            continue

        # ---------- number of players ----------
        if PLAYER_NAME_INDEX < len(group_df.columns):
            names = (
                group_df.iloc[:, PLAYER_NAME_INDEX]
                .dropna()
                .astype(str)
                .str.strip()
            )
            names = names[names != ""]
            n_players = int(names.nunique())
        else:
            n_players = 0

        # ---------- find Overall Experience (chart #7) ----------
        meta = _build_plot_metadata(group_df)
        overall_idx: Optional[int] = None
        for m in meta:
            if m["ptype"] == "rating" and m["number"] == 7:
                overall_idx = m["idx"]
                break

        if overall_idx is not None and overall_idx < len(group_df.columns):
            series = pd.to_numeric(
                group_df.iloc[:, overall_idx], errors="coerce"
            ).dropna()
            avg_q7 = float(series.mean()) if not series.empty else np.nan
        else:
            avg_q7 = np.nan

        # ---------- coach lookup ----------
        coach_name = TEAM_COACH_MAP.get(team_name, "?")

        summary_rows.append(
            {
                "Team": team_name,
                "Coach": coach_name,
                "Players": n_players,
                "OverallExp": avg_q7,
            }
        )

    if not summary_rows:
        return

    summary_df = pd.DataFrame(summary_rows)

    # Rank: more players first, then higher Overall Experience
    # Use a filled column so NaNs sort last.
    summary_df["OverallExp_sort"] = summary_df["OverallExp"].fillna(-1.0)
    summary_df = summary_df.sort_values(
        by=["Players", "OverallExp_sort", "Team"],
        ascending=[False, False, True],
        ignore_index=True,
    ).drop(columns=["OverallExp_sort"])

    # ---------- draw the page ----------
    fig, ax = plt.subplots(figsize=(8.5, 11))  # portrait
    ax.axis("off")

    fig.suptitle(f"{cycle_label} Summary", fontsize=14, fontweight="bold")

    y = 0.9
    line_height = 0.035

    for i, row in summary_df.iterrows():
        rank = i + 1
        team = row["Team"]
        coach = row["Coach"]
        players = int(row["Players"])

        if pd.isna(row["OverallExp"]):
            oe_str = "N/A"
        else:
            oe_str = f"{row['OverallExp']:.2f}"

        players_word = "player" if players == 1 else "players"

        text = (
            f"{rank}. {team} - {coach}: "
            f"{players} {players_word}; Overall Experience: {oe_str}"
        )

        ax.text(
            0.04,
            y,
            text,
            fontsize=9,
            ha="left",
            va="top",
            transform=ax.transAxes,
        )

        y -= line_height
        if y < 0.05:
            break  # don't run off the page

    fig.tight_layout(rect=[0, 0, 1, 0.95])
    pdf.savefig(fig)
    plt.close(fig)


def create_pdf_from_original(
    input_path: str,
    cycle_label: str = "Cycle",
    output_path: Optional[str] = None,
) -> str:
    """
    Use the ORIGINAL survey Excel file and create a multi-page PDF:

    Page 1: Cycle summary (ranked teams)
    Then for "All Teams" and for each individual team:
      - Page 1: all charts for that group, numbered 1,2,3,...
      - Page 2: tables with 1–2 or 1–3 star reviews and "NO" replies,
                plus a list of players and comments for team pages.
    """
    if output_path is None:
        base, _ = os.path.splitext(input_path)
        output_path = base + "_report.pdf"

    df = pd.read_excel(input_path, sheet_name=0)

    if GROUP_COL_INDEX >= len(df.columns):
        raise ValueError(
            "Group column G is outside the available columns in the sheet."
        )

    group_col_name = df.columns[GROUP_COL_INDEX]
    df[group_col_name] = df[group_col_name].fillna("UNASSIGNED")

    with PdfPages(output_path) as pdf:
        # --------- NEW: global summary page ---------
        _add_cycle_summary_page(pdf, df, group_col_name, cycle_label)

        # --------- All teams combined ---------
        all_meta = _build_plot_metadata(df)
        _add_group_charts_page_to_pdf(pdf, df, "All Teams", cycle_label, all_meta)
        _add_group_tables_page_to_pdf(
            pdf, df, "All Teams", cycle_label, all_meta, is_all_teams=True
        )

        # --------- Each team (two pages each) ---------
        for group_value, group_df in df.groupby(group_col_name, sort=True):
            team_name = str(group_value).strip()

            coach_name = TEAM_COACH_MAP.get(team_name, "?")
            title_label = f"{team_name} - {coach_name}"

            meta = _build_plot_metadata(group_df)
            _add_group_charts_page_to_pdf(
                pdf, group_df, title_label, cycle_label, meta
            )
            _add_group_tables_page_to_pdf(
                pdf, group_df, title_label, cycle_label, meta, is_all_teams=False
            )

    return output_path


