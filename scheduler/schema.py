import csv, io

REQUIRED_HEADERS = [
    "First Name","Last Name","Active","Number",
    "PreferredPos1","PreferredPos2","PreferredPos3","Seed"
]
VALID_POS = {"GS","GA","WA","C","WD","GD","GK",""}

def _norm_row(d: dict) -> dict:
    out = {h: (d.get(h, "") or "").strip() for h in REQUIRED_HEADERS}

    a = out["Active"].upper()
    out["Active"] = "" if a in {"", "YES", "Y"} else "NO"

    for k in ("PreferredPos1","PreferredPos2","PreferredPos3"):
        val = out[k].upper().replace(" ", "")
        out[k] = val if val in VALID_POS else ""

    s = out["Seed"].strip()
    out["Seed"] = s if s in {"1","2","3"} else "3"
    return out

def validate_and_normalize_csv(csv_bytes: io.BytesIO) -> list[dict]:
    csv_bytes.seek(0)
    text = io.TextIOWrapper(csv_bytes, encoding="utf-8-sig", newline="")
    rdr = csv.DictReader(text)
    if rdr.fieldnames is None:
        raise ValueError("CSV has no header row")
    got = [h.strip() for h in rdr.fieldnames]
    missing = [h for h in REQUIRED_HEADERS if h not in got]
    if missing:
        raise ValueError(f"Missing CSV columns: {missing}")
    rows = [_norm_row(row) for row in rdr]
    if not rows:
        raise ValueError("CSV contains no data rows")
    return rows
