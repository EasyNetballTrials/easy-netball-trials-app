# scheduler/schema.py
import csv
import io

VALID_POS = {"GS","GA","WA","C","WD","GD","GK"}

def _norm_pos(x: str) -> str:
    x = (x or "").strip().upper()
    return x if x in VALID_POS else ""

def validate_and_normalize_csv(fp: io.BytesIO):
    """
    Return list of dicts:
    first_name, last_name, active, number, p1, p2, p3, seed
    """
    fp.seek(0)
    text = fp.read().decode("utf-8-sig", errors="ignore")
    rdr = csv.DictReader(io.StringIO(text))

    out = []
    for row in rdr:
        fn = (row.get("First Name","") or "").strip()
        ln = (row.get("Last Name","") or "").strip()
        if not fn and not ln:
            continue

        active = (row.get("Active","") or "").strip().upper()
        number = (row.get("Number","") or "").strip()
        p1 = _norm_pos(row.get("PreferredPos1",""))
        p2 = _norm_pos(row.get("PreferredPos2",""))
        p3 = _norm_pos(row.get("PreferredPos3",""))
        seed_raw = (row.get("Seed","") or "").strip()
        try:
            seed = int(seed_raw) if seed_raw else ""
        except:
            seed = ""

        out.append({
            "first_name": fn,
            "last_name": ln,
            "active": active,    # blank handled in writer (→ YES)
            "number": number,
            "p1": p1, "p2": p2, "p3": p3,
            "seed": seed,
        })
    return out
