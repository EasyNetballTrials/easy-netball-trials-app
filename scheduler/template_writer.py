# template_writer.py
import os
import shutil
from typing import Iterable

import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

# Columns your VBA expects, in this exact order.
EXPECTED_ORDER: Iterable[str] = [
    "First Name", "Last Name", "Active", "Number",
    "PreferredPos1", "PreferredPos2", "PreferredPos3", "Seed",
]


def _normalize_players_df(df: pd.DataFrame) -> pd.DataFrame:
    """
    Ensure all expected columns exist and are in the exact order the macros expect.
    Handles BOM/whitespace in headers and keeps values as strings.
    """
    df = df.copy()

    # Clean headers: remove BOM (\ufeff), trim spaces
    cleaned_headers = [str(c).replace("\ufeff", "").strip() for c in df.columns]
    df.columns = cleaned_headers

    # Create any missing columns as blanks
    for col in EXPECTED_ORDER:
        if col not in df.columns:
            df[col] = ""

    # Force exact order (Number is 4th; Seed last)
    df = df.loc[:, list(EXPECTED_ORDER)]

    # Normalize text columns
    text_cols = ["First Name", "Last Name", "Active",
                 "PreferredPos1", "PreferredPos2", "PreferredPos3"]
    for col in text_cols:
        df[col] = df[col].fillna("").astype(str).str.strip()

    # Keep as text; VBA handles numbering/seed logic
    df["Seed"] = df["Seed"].fillna("").astype(str).str.strip()
    df["Number"] = df["Number"].fillna("").astype(str).str.strip()

    return df


def _write_dataframe_to_sheet(ws, df: pd.DataFrame, start_row: int = 1, start_col: int = 1):
    """
    Write a DataFrame (with header) to an openpyxl worksheet at (start_row, start_col).
    """
    # header
    for j, col in enumerate(df.columns, start=start_col):
        ws.cell(row=start_row, column=j, value=col)

    # data
    for i, (_, row) in enumerate(df.iterrows(), start=start_row + 1):
        for j, col in enumerate(df.columns, start=start_col):
            val = row[col]
            ws.cell(row=i, column=j, value=("" if pd.isna(val) else val))


def _clear_sheet(ws):
    """Clear all cell values in a sheet (values only)."""
    ws.delete_rows(1, ws.max_row)


def write_into_template(
    template_path: str,
    out_path: str,
    df_players: pd.DataFrame,
    num_courts: int = 2,
    num_rounds: int = 10,
    numbering: bool = True,
    show_prefs: bool = False,
):
    """
    Copy the macro-enabled template and write the Players table + Control flags.

    - Preserves VBA (keep_vba=True)
    - Normalizes Players columns (inserts any missing; forces exact order)
    - Sets Control!B2 (Rounds), B3 (Courts), B4 (Numbering YES/NO), B6 (Prefs YES/NO)
    - Sets Control!Z1 = "PENDING" so the workbook finalizes on first open
    - Clears Schedule, CourtBoards, GameTally so users get a clean build
    """
    if not os.path.exists(template_path):
        raise FileNotFoundError(f"Template not found: {template_path}")

    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    shutil.copyfile(template_path, out_path)

    # Normalize incoming data to exact schema
    df_norm = _normalize_players_df(df_players)

    # Open workbook (preserve macros)
    wb = load_workbook(out_path, keep_vba=True)

    # Ensure required sheets exist
    ws_players = wb["Players"] if "Players" in wb.sheetnames else wb.create_sheet("Players")
    ws_control = wb["Control"] if "Control" in wb.sheetnames else wb.create_sheet("Control")

    # Clear runtime sheets so the user gets a clean build
    for nm in ("Schedule", "CourtBoards", "GameTally"):
        if nm in wb.sheetnames:
            _clear_sheet(wb[nm])

    # Write Players
    _clear_sheet(ws_players)
    _write_dataframe_to_sheet(ws_players, df_norm, start_row=1, start_col=1)

    # Light "autofit" (approximate) for Players columns
    for j, col in enumerate(df_norm.columns, start=1):
        header_len = len(col)
        sample_len = max([len(str(v)) for v in df_norm[col][:50].astype(str).tolist()] + [header_len])
        ws_players.column_dimensions[get_column_letter(j)].width = min(max(sample_len + 2, 10), 30)

    # Control flags for finalize on first open
    ws_control["B2"].value = int(num_rounds)                 # Rounds
    ws_control["B3"].value = int(num_courts)                 # Courts
    ws_control["B4"].value = "YES" if numbering else "NO"    # Numbering switch
    ws_control["B6"].value = "YES" if show_prefs else "NO"   # Show preferred positions
    ws_control["Z1"].value = "PENDING"                       # Trigger finalize in Workbook_Open

    wb.save(out_path)
