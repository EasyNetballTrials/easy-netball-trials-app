# capacity.py
from typing import Dict
import pandas as pd


def capacity_report(df: pd.DataFrame, num_courts: int, num_rounds: int) -> Dict:
    """
    Check whether the chosen courts/rounds provide enough slots
    for each player to get at least 3 games.

    Returns:
        {
            "ok": bool,
            "capacity": int,
            "required": int,
            "players": int,
            "message": str
        }
    """
    # How many players are marked active (or assume all if no Active column)
    if "Active" in df.columns:
        active_df = df[(df["Active"].astype(str).str.upper().isin(["YES", ""]))]  # YES or blank = active
    else:
        active_df = df

    players = len(active_df)
    required = players * 3
    capacity = num_courts * num_rounds * 14  # 14 positions per court per round

    if capacity < required:
        return {
            "ok": False,
            "capacity": capacity,
            "required": required,
            "players": players,
            "message": (
                f"Capacity warning: only {capacity} slots available "
                f"({num_courts} courts × {num_rounds} rounds × 14 slots), "
                f"but {required} slots are required to give {players} players 3 games each."
            ),
        }
    else:
        return {
            "ok": True,
            "capacity": capacity,
            "required": required,
            "players": players,
            "message": (
                f"Capacity OK: {capacity} slots available for {players} players "
                f"({num_courts} courts × {num_rounds} rounds × 14 slots)."
            ),
        }
