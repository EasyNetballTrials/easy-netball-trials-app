# csv_ingest.py
import pandas as pd
from typing import Iterable, Dict

EXPECTED_ORDER: Iterable[str] = [
    "First Name", "Last Name", "Active", "Number",
    "PreferredPos1", "PreferredPos2", "PreferredPos3", "Seed",
]

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

def _clean_header(h: str) -> str:
    # remove BOM, collapse spaces, lower
    return str(h).replace("\ufeff", "").strip().lower()

def _normalize_headers(df: pd.DataFrame) -> pd.DataFrame:
    cols = []
    for c in df.columns:
        key = _clean_header(c)
        cols.append(HEADER_ALIASES.get(key, str(c).strip()))
    df = df.copy()
    df.columns = cols
    return df

def read_players_csv_to_df(path: str) -> pd.DataFrame:
    """
    Read the uploaded CSV/TSV and return a DataFrame with the exact schema our writer expects.
    - Delimiter auto-detected (comma or tab).
    - Headers normalized (BOM/spacing/aliases).
    - Any missing expected columns created as blank.
    - Final column order returned IS NOT enforced here (writer also enforces), but we
      build the full set so downstream is stable.
    """
    df = pd.read_csv(
        path,
        sep=None,              # auto-detect comma vs tab
        engine="python",
        dtype=str,
        keep_default_na=False,
        na_values=[],
    )

    df = _normalize_headers(df)

    # Ensure all expected columns exist (create blanks where missing)
    for col in EXPECTED_ORDER:
        if col not in df.columns:
            df[col] = ""

    # Return ONLY expected columns, in the ORIGINAL upload order + our blanks;
    # template_writer will force the final exact EXPECTED_ORDER.
    df = df[[c for c in df.columns if c in set(EXPECTED_ORDER)]].copy()

    # Clean whitespace in all cells
    for col in df.columns:
        df[col] = df[col].astype(str).str.strip()

    return df
