# csv_ingest.py
import pandas as pd
from typing import Iterable, Dict

# Columns your app expects downstream (same as template_writer)
EXPECTED_ORDER: Iterable[str] = [
    "First Name", "Last Name", "Active", "Number",
    "PreferredPos1", "PreferredPos2", "PreferredPos3", "Seed",
]

# Common header variants we normalize
HEADER_ALIASES: Dict[str, str] = {
    "firstname": "First Name",
    "first name": "First Name",
    "givenname": "First Name",
    "forename": "First Name",

    "lastname": "Last Name",
    "surname": "Last Name",
    "familyname": "Last Name",
    "last name": "Last Name",

    "active": "Active",
    "status": "Active",

    "number": "Number",
    "player number": "Number",
    "bib": "Number",
    "bib number": "Number",
    "no": "Number",
    "#": "Number",
    "jersey": "Number",
    "jersey number": "Number",

    "preferredpos1": "PreferredPos1",
    "preferred pos1": "PreferredPos1",
    "preferred pos 1": "PreferredPos1",
    "p1": "PreferredPos1",

    "preferredpos2": "PreferredPos2",
    "preferred pos2": "PreferredPos2",
    "preferred pos 2": "PreferredPos2",
    "p2": "PreferredPos2",

    "preferredpos3": "PreferredPos3",
    "preferred pos3": "PreferredPos3",
    "preferred pos 3": "PreferredPos3",
    "p3": "PreferredPos3",

    "seed": "Seed",
}


def _normalize_headers(df: pd.DataFrame) -> pd.DataFrame:
    cols = []
    for c in df.columns:
        key = str(c).strip().lower()
        cols.append(HEADER_ALIASES.get(key, str(c).strip()))
    df = df.copy()
    df.columns = cols
    return df


def read_players_csv_to_df(path: str) -> pd.DataFrame:
    """
    Read a CSV of players and normalize it to the expected schema.
    - Accepts arbitrary column order and common header variants.
    - If 'Number' is missing, leaves it blank (template_writer will insert and order it).
    - Returns a DataFrame; any missing expected columns are created as blanks.
    """
    df = pd.read_csv(
        path,
        engine="python",
        dtype=str,
        keep_default_na=False,
        na_values=[],
    )

    df = _normalize_headers(df)

    # Ensure all expected columns exist
    for col in EXPECTED_ORDER:
        if col not in df.columns:
            df[col] = ""

    # Keep only expected columns (in any order here; writer will enforce final order)
    df = df[[c for c in df.columns if c in set(EXPECTED_ORDER)]]

    # Clean whitespace
    for col in df.columns:
        df[col] = df[col].astype(str).str.strip()

    return df

