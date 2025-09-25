# capacity.py
from typing import Dict
import pandas as pd

def _count_active(df: pd.DataFrame) -> int:
    """
    Count active players. We treat Active 'YES' (case-insensitive) or blank as active.
    Anything else (e.g., 'NO') is inactive.
    """
    if "Active" not in df.columns:
        return len(df)
    series = df["Active"].astype(str).str.strip().str.upper()
    return int((series.eq("YES") | series.eq("")).sum())


def capacity_report(df: pd.DataFrame, num_courts: int, num_rounds: int) -> Dict:
    """
    Simple capacity guard used by app.py.
    - Each court hosts 14 players per round (7 positions x 2 teams).
    - Target is 3 games per active player.
    """
    active = _count_active(df)
    capacity = int(num_courts) * 14 * int(num_rounds)
    required = active * 3

    if active == 0:
        return {
            "ok": False,
            "message": "No active players detected. Set Active=YES (or leave blank) for players you want to include.",
            "active": 0,
            "capacity": capacity,
            "required": required,
        }

    if capacity < required:
        short = required - capacity
        msg = (
            f"Capacity warning: {capacity} slots but {required} are needed to give "
            f"{active} players 3 games each. You are short by {short}. "
            "Add rounds and/or a second court, or proceed and adjust manually."
        )
        return {
            "ok": False,
            "message": msg,
            "active": active,
            "capacity": capacity,
            "required": required,
        }

    return {
        "ok": True,
        "message": f"Capacity OK: {capacity} slots for {active} active players.",
        "active": active,
        "capacity": capacity,
        "required": required,
    }

