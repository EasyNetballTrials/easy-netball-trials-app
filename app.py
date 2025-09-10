import os
from flask import Flask
from flask_cors import CORS

app = Flask(__name__)

# folders
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
app.config["UPLOAD_FOLDER"] = os.path.join(BASE_DIR, "uploads")
app.config["RESULTS_FOLDER"] = os.path.join(BASE_DIR, "results")

# ensure folders exist
os.makedirs(app.config["UPLOAD_FOLDER"], exist_ok=True)
os.makedirs(app.config["RESULTS_FOLDER"], exist_ok=True)

# safety: 5 MB upload cap (CSV is tiny anyway)
app.config["MAX_CONTENT_LENGTH"] = 5 * 1024 * 1024

# CORS: allow your Wix domain (edit to your domain)
CORS(app, resources={
    r"/*": {"origins": [
        "https://www.YOURDOMAIN.com", 
        "https://YOUR-WIX-SUBDOMAIN.wixsite.com"
    ]}
})

# add this if you see 404s for favicon (optional)
@app.route('/favicon.ico')
def favicon():
    return ('', 204)

import os
from datetime import datetime
from io import BytesIO

from flask import Flask, render_template, request, send_from_directory, redirect, url_for, flash
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.worksheet.pagebreak import Break



app = Flask(__name__)
app.secret_key = "change-me-to-a-random-string"

ALLOWED_EXT = {".csv"}
POSITIONS = ["GS", "GA", "WA", "C", "WD", "GD", "GK"]

# ---------- Style helpers for Excel ----------
thin = Side(style="thin", color="000000")
border_all = Border(left=thin, right=thin, top=thin, bottom=thin)
fill_header = PatternFill("solid", fgColor="C8C8C8")
fill_title  = PatternFill("solid", fgColor="B4C8E6")
fill_teamA  = PatternFill("solid", fgColor="DCFBE6")
fill_teamB  = PatternFill("solid", fgColor="FFFFC8")
# Extra colours for the "Match-ups" sheet
fill_title_match  = PatternFill("solid", fgColor="A3D2FF")  # lighter blue title
fill_teamA_match  = PatternFill("solid", fgColor="E6F7FF")  # pale blue A
fill_teamB_match  = PatternFill("solid", fgColor="FFF0E6")  # pale orange B


# ---------- CSV loading & normalisation ----------
REQUIRED = ["First Name", "Last Name", "PreferredPos1", "PreferredPos2", "PreferredPos3"]
OPTIONAL = ["Active", "Seed"]

def normalize_headers(df: pd.DataFrame) -> pd.DataFrame:
    mapping = {}
    for c in df.columns:
        cc = c.strip()
        up = cc.replace(" ", "").lower()
        if up in ("firstname","first_name","first"):
            mapping[c] = "First Name"
        elif up in ("lastname","last_name","surname","last"):
            mapping[c] = "Last Name"
        elif up in ("preferredpos1","preferredposition1","pos1","preferredposi","preferredpos"):
            mapping[c] = "PreferredPos1"
        elif up in ("preferredpos2","preferredposition2","pos2"):
            mapping[c] = "PreferredPos2"
        elif up in ("preferredpos3","preferredposition3","pos3"):
            mapping[c] = "PreferredPos3"
        elif up in ("active","attendance","attend","present"):
            mapping[c] = "Active"
        elif up in ("seed","seeding","rank"):
            mapping[c] = "Seed"
        else:
            mapping[c] = cc
    df = df.rename(columns=mapping)
    for col in REQUIRED:
        if col not in df.columns:
            df[col] = ""
    for col in OPTIONAL:
        if col not in df.columns:
            df[col] = ""
    return df

def load_players(file_storage) -> pd.DataFrame:
    df = pd.read_csv(file_storage)
    df = normalize_headers(df)
    # defaults
    df["Active"] = df["Active"].fillna("").astype(str).str.strip().str.upper()
    df.loc[df["Active"]=="", "Active"] = "YES"
    df["Seed"] = pd.to_numeric(df["Seed"], errors="coerce").fillna(3).astype(int)
    for p in ["PreferredPos1","PreferredPos2","PreferredPos3"]:
        df[p] = df[p].fillna("").astype(str).str.strip().str.upper()
    df["Full Name"] = (df["First Name"].fillna("").astype(str).str.strip() + " " +
                       df["Last Name"].fillna("").astype(str).str.strip()).str.strip()
    return df

# ---------- Scheduling core ----------
def capacity_ok(num_active, num_courts, rounds, min_games=3):
    capacity = num_courts * 14 * rounds  # 7 pos * 2 teams * courts * rounds
    return capacity >= num_active * min_games, capacity

def build_player_pool(df_active):
    pool = []
    for _, row in df_active.iterrows():
        pool.append({
            "full": row["Full Name"],
            "p1": row["PreferredPos1"],
            "p2": row["PreferredPos2"],
            "p3": row["PreferredPos3"],
            "seed": int(row["Seed"]),
        })
    return pool

def pref_tag_for(full: str, pos: str, prefs_by_full: dict) -> str | None:
    """Return 'P1'/'P2'/'P3' for this player's pos, or None if not in prefs."""
    rec = prefs_by_full.get(full, {})
    pos = (pos or "").upper()
    if not rec:
        return None
    if pos == rec.get("p1"):
        return "P1"
    if pos == rec.get("p2"):
        return "P2"
    if pos == rec.get("p3"):
        return "P3"
    return None

def pick_next(pool, pos, games_played, court_hist, used_this_round,
              min_games_target, prefer_hist, prefs_by_full):
    """
    pool: list of dicts per player with keys: full, p1, p2, p3, seed, court
    prefer_hist: {full: {'P1':count,'P2':count,'P3':count}}
    prefs_by_full: {full: {'p1':..., 'p2':..., 'p3':...}}
    """
    pos = pos.upper()
    pref_weight = {"P1": 40, "P2": 25, "P3": 8}
    target_share = {"P1": 0.45, "P2": 0.35, "P3": 0.20}  # soft target mix

    # Two passes: try to help players under the min-games target first
    for pass_num in (1, 2):
        best_score = -10**9
        best = None

        for p in pool:
            full = p["full"]
            if full in used_this_round:
                continue
            prefs = (p["p1"], p["p2"], p["p3"])
            if pos not in prefs:
                continue

            gp = int(games_played.get(full, 0))
            if pass_num == 1 and gp >= min_games_target:
                continue

            # Base score: fewer games better, lower seed better
            score = 1000 - gp * 20 - int(p["seed"]) * 2

            # Court-spread bonus
            if court_hist.get(full, "").find(str(p["court"])) == -1:
                score += 10

            # Preference-aware scoring
            tag = pref_tag_for(full, pos, prefs_by_full)  # 'P1'/'P2'/'P3'
            if tag:
                # Strong preference bonus + diminishing returns if already many of that tag
                count_same = prefer_hist[full].get(tag, 0)
                score += pref_weight.get(tag, 0) - 12 * count_same

                # Under-representation nudge toward target share (soft)
                if gp > 0:
                    want = target_share[tag] * gp
                    diff = want - count_same  # >0 if under target
                    score += 8 * diff

            if score > best_score:
                best_score = score
                best = full

        if best:
            return best

    return ""

def schedule_rounds(df_active, rounds, courts, min_games_target=3):
    # Track total games and court usage
    games_played = {row["Full Name"]: 0 for _, row in df_active.iterrows()}
    court_hist = {row["Full Name"]: "" for _, row in df_active.iterrows()}

    # Build pool + a quick lookup of each player's prefs
    pool_base = []
    prefs_by_full = {}
    for _, row in df_active.iterrows():
        full = row["Full Name"]
        p1 = row["PreferredPos1"]; p2 = row["PreferredPos2"]; p3 = row["PreferredPos3"]
        pool_base.append({"full": full, "p1": p1, "p2": p2, "p3": p3, "seed": int(row["Seed"])})
        prefs_by_full[full] = {"p1": p1, "p2": p2, "p3": p3}

    # Track how many times each player has been assigned to P1/P2/P3 to spread usage
    prefer_hist = {full: {"P1": 0, "P2": 0, "P3": 0} for full in prefs_by_full.keys()}

    def bump_hist(full: str, pos: str):
        tag = pref_tag_for(full, pos, prefs_by_full)
        if tag:
            prefer_hist[full][tag] += 1

    rows = []  # (round, court, team, pos, player)

    for rnd in range(1, rounds + 1):
        used = set()  # no duplicates across both courts within the round

        for court in range(1, courts + 1):
            # Tag court to influence court-spread bonus
            pool = [dict(p, court=court) for p in pool_base]

            teamA, teamB = [], []
            for pos in POSITIONS:
                a = pick_next(pool, pos, games_played, court_hist, used,
                              min_games_target, prefer_hist, prefs_by_full)
                if a:
                    used.add(a)
                    games_played[a] += 1
                    court_hist[a] += f"{court},"
                    bump_hist(a, pos)
                teamA.append(a or "")

                b = pick_next(pool, pos, games_played, court_hist, used,
                              min_games_target, prefer_hist, prefs_by_full)
                if b:
                    used.add(b)
                    games_played[b] += 1
                    court_hist[b] += f"{court},"
                    bump_hist(b, pos)
                teamB.append(b or "")

            for i, pos in enumerate(POSITIONS):
                rows.append((rnd, court, "A", pos, teamA[i]))
                rows.append((rnd, court, "B", pos, teamB[i]))

    sched = pd.DataFrame(rows, columns=["Round", "Court", "Team", "Position", "Player"])
    return sched, games_played

# ---------- P1 ≥ target swaps ----------
def build_pref_arrays(df_active):
    prefs, primary = {}, {}
    for _, r in df_active.iterrows():
        full = (r["Full Name"] or "").strip().upper()
        if not full:
            continue
        allowed = [r["PreferredPos1"], r["PreferredPos2"], r["PreferredPos3"]]
        prefs[full] = {p for p in allowed if p}
        primary[full] = (r["PreferredPos1"] or "").upper()
    return prefs, primary

def enforce_p1_at_least(schedule_df, df_active, min_p1=2):
    prefs, primary = build_pref_arrays(df_active)
    sched = schedule_df.copy()

    pos_counts, rounds_by_player = {}, {}
    for _, row in sched.iterrows():
        full = (row["Player"] or "").strip().upper()
        if not full:
            continue
        rnd = int(row["Round"]); pos = row["Position"].upper()
        pos_counts.setdefault(full, {}).setdefault(pos, 0)
        pos_counts[full][pos] += 1
        rounds_by_player.setdefault(full, set()).add(rnd)

    notes = {}
    for full in list(rounds_by_player.keys()):
        p1 = primary.get(full, "")
        if not p1:
            continue
        have = pos_counts.get(full, {}).get(p1, 0)
        if have >= min_p1:
            continue
        for rnd in sorted(rounds_by_player[full]):
            if have >= min_p1:
                break
            idx_player = sched.index[(sched["Round"]==rnd) & (sched["Player"].str.upper()==full)]
            if len(idx_player)==0:
                continue
            ip = idx_player[0]
            pos_cur = sched.at[ip, "Position"].upper()
            if pos_cur == p1:
                have += 1
                continue
            idx_targets = sched.index[(sched["Round"]==rnd) & (sched["Position"].str.upper()==p1)]
            for it in idx_targets:
                other = (sched.at[it,"Player"] or "").strip().upper()
                if not other or other == full:
                    continue
                if other in prefs and pos_cur in prefs[other] and full in prefs and p1 in prefs[full]:
                    # swap
                    sched.at[it,"Player"], sched.at[ip,"Player"] = sched.at[ip,"Player"], sched.at[it,"Player"]
                    have += 1
                    break
        if have < min_p1:
            notes[full] = f"P1 {have}/{min_p1}"
    return sched, notes

# ---------- Numbering ----------
def build_numbers(df_active_all):
    tmp = df_active_all[df_active_all["Active"].str.upper()=="YES"].copy()
    tmp["LastLower"] = tmp["Last Name"].astype(str).str.lower()
    tmp["FirstLower"] = tmp["First Name"].astype(str).str.lower()
    tmp = tmp.sort_values(["LastLower","FirstLower"], kind="mergesort")
    tmp["Number"] = range(1, len(tmp)+1)
    num_map = dict(zip(tmp["Full Name"], tmp["Number"]))
    players_sheet = df_active_all.copy()
    players_sheet["Number"] = players_sheet["Full Name"].map(num_map).fillna("")
    players_sheet = players_sheet.drop(columns=["LastLower","FirstLower"], errors="ignore")
    return num_map, players_sheet

# ---------- Excel writers ----------
def write_schedule_ws(wb, schedule_df):
    ws = wb.create_sheet("Schedule")
    ws.append(["Round","Court","Team","Position","Player"])
    for _, r in schedule_df.iterrows():
        ws.append([int(r["Round"]), int(r["Court"]), r["Team"], r["Position"], r["Player"] or ""])
    return ws

def write_gametally_ws(wb, games_played, notes, df_active):
    ws = wb.create_sheet("GameTally")
    ws.append(["First Name","Last Name","Full Name","Games Played","Notes"])
    tmp = df_active[df_active["Active"].str.upper()=="YES"].copy()
    tmp["Games Played"] = tmp["Full Name"].map(games_played).fillna(0).astype(int)
    tmp["Notes"] = tmp["Full Name"].str.upper().map(notes).fillna("")
    tmp = tmp.sort_values(["Last Name","First Name"], kind="mergesort")
    for _, r in tmp.iterrows():
        ws.append([r["First Name"], r["Last Name"], r["Full Name"], int(r["Games Played"]), r["Notes"]])
    for col in "ABCDE":
        ws.column_dimensions[col].auto_size = True
    return ws

def write_players_ws(wb, players_df):
    ws = wb.create_sheet("Players")
    ws.append(list(players_df.columns))
    for _, r in players_df.iterrows():
        ws.append([r.get(c, "") for c in players_df.columns])
    for i, _ in enumerate(players_df.columns, start=1):
        col_letter = chr(64+i) if i <= 26 else "Z"
        ws.column_dimensions[col_letter].auto_size = True
    return ws

def write_courtboards_ws(wb, schedule_df, numbers_map=None, sheet_name="CourtBoards", matchups=False):
    """
    Writes a CourtBoards sheet.
    - When matchups=True, RHS position is the opposite (GS<->GK, GA<->GD, WA<->WD, C=C),
      so each row shows a direct head-to-head.
    - Inserts a manual page break BEFORE each round header (so each round starts a page)
      and keeps two courts per page.
    - Uses alternate colours and shows a small legend when matchups=True.
    """
    inverse_map = {"GS":"GK","GA":"GD","WA":"WD","C":"C","WD":"WA","GD":"GA","GK":"GS"}

    ws = wb.create_sheet(sheet_name)
    ws["A1"] = "Court Boards" if not matchups else "Court Boards (Match-ups)"
    ws["A1"].font = Font(bold=True, size=14)
    row_out = 3

    rounds = sorted(schedule_df["Round"].unique())

    for idx_round, rnd in enumerate(rounds, start=1):
        # We haven't written anything for this round yet; this will be the header row
        round_header_row = row_out

        # Page break BEFORE the round header, so header appears at top of new page
        if idx_round > 1:
            break_row = max(2, round_header_row - 1)
            try:
                ws.row_breaks.append(Break(id=break_row))
            except Exception:
                pass

        courts_in_round = sorted(schedule_df.loc[schedule_df["Round"] == rnd, "Court"].unique())

        for court in courts_in_round:
            # Header row (merged)
            ws.merge_cells(start_row=row_out, start_column=1, end_row=row_out, end_column=6)
            cell = ws.cell(row=row_out, column=1)
            title = f"Round {rnd} - Court {court}"
            cell.value = title if not matchups else f"{title} (Match-ups)"
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center")
            # Colour the title row
            for c in range(1, 6+1):
                ws.cell(row=row_out, column=c).fill = (fill_title if not matchups else fill_title_match)
                ws.cell(row=row_out, column=c).border = border_all
            row_out += 1


            # Column headers
            headers = ["Notes", "Pos", "Team A", "Team B", "Pos", "Notes"]
            for j, h in enumerate(headers, start=1):
                ws.cell(row=row_out, column=j, value=h)
                ws.cell(row=row_out, column=j).font = Font(bold=True)
                ws.cell(row=row_out, column=j).fill = fill_header
                ws.cell(row=row_out, column=j).border = border_all
            row_out += 1

            # Body rows
            block = schedule_df[(schedule_df["Round"] == rnd) & (schedule_df["Court"] == court)]
            block = block.set_index(["Team", "Position"]).sort_index()

            for pos in POSITIONS:
                # Left side = Team A at 'pos'
                a = block.at[("A", pos), "Player"] if ("A", pos) in block.index else ""

                # Right side pos depends on mode
                rhs_pos = inverse_map.get(pos, pos) if matchups else pos
                b = block.at[("B", rhs_pos), "Player"] if ("B", rhs_pos) in block.index else ""

                dispA, dispB = a, b
                if numbers_map:
                    if a and a in numbers_map: dispA = f"{a} {numbers_map[a]}"
                    if b and b in numbers_map: dispB = f"{b} {numbers_map[b]}"

                row_vals = ["", pos, dispA, dispB, rhs_pos, ""]
                for j, val in enumerate(row_vals, start=1):
                    ws.cell(row=row_out, column=j, value=val)
                    ws.cell(row=row_out, column=j).border = border_all
                    # Row shading: different palettes for normal vs match-ups
                    if j == 3:
                        ws.cell(row=row_out, column=j).fill = (fill_teamA if not matchups else fill_teamA_match)
                    if j == 4:
                        ws.cell(row=row_out, column=j).fill = (fill_teamB if not matchups else fill_teamB_match)
                ws.cell(row=row_out, column=2).font = Font(bold=True)
                ws.cell(row=row_out, column=5).font = Font(bold=True)
                row_out += 1

            # Spacer between courts
            row_out += 1

    # Widths + print setup
            # Widths
    ws.column_dimensions["A"].width = 20
    ws.column_dimensions["B"].width = 6
    ws.column_dimensions["C"].width = 22
    ws.column_dimensions["D"].width = 22
    ws.column_dimensions["E"].width = 6
    ws.column_dimensions["F"].width = 20

    # Print setup – landscape, fit to 1 page wide
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0

    # Margins – leave room for footer
    ws.page_margins.left = 0.5
    ws.page_margins.right = 0.5
    ws.page_margins.top = 0.5
    ws.page_margins.bottom = 0.75
    ws.page_margins.header = 0.3
    ws.page_margins.footer = 0.3

    # Footer: Page X of Y (Excel 2010 compatible, safe API)
    try:
        hf = ws.header_footer
        hf.center_footer = "Page &P of &N"
        hf.left_footer = ""
        hf.right_footer = ""
        hf.differentFirst = False
        hf.differentOddEven = False
    except Exception:
        pass

    return ws


# ---------- Routes ----------
@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")

@app.route("/generate", methods=["POST"])
def generate():
    f = request.files.get("file")
    if not f or os.path.splitext(f.filename)[1].lower() not in ALLOWED_EXT:
        flash("Please upload a .csv file.")
        return redirect(url_for("index"))

    try:
        df = load_players(f)
    except Exception as e:
        flash(f"Could not read CSV: {e}")
        return redirect(url_for("index"))

    # form values
    try:
        rounds = int(request.form.get("rounds", "10"))
    except:
        rounds = 10
    try:
        courts = int(request.form.get("courts", "2"))
    except:
        courts = 2
    number_players = (request.form.get("number_players","NO").upper() == "YES")
    try:
        min_p1 = int(request.form.get("min_p1", "2"))
        if min_p1 < 1: min_p1 = 1
    except:
        min_p1 = 2

    df_active = df[df["Active"].str.upper()=="YES"].copy()
    num_active = len(df_active)

    ok, cap = capacity_ok(num_active, courts, rounds, min_games=3)
    if not ok:
        flash(f"Warning: capacity {cap} is less than {num_active*3} needed to guarantee 3 games each. "
              f"(Courts={courts}, Rounds={rounds})")

    schedule_df, games_played = schedule_rounds(df_active, rounds, courts, min_games_target=3)
    schedule_df, notes = enforce_p1_at_least(schedule_df, df_active, min_p1=min_p1)

    numbers_map = None
    players_sheet_df = df.copy()
    if number_players:
        numbers_map, players_sheet_df = build_numbers(df)

    # Build workbook
    wb = Workbook()
    wb.remove(wb.active)

    write_schedule_ws(wb, schedule_df)
    write_courtboards_ws(wb, schedule_df, numbers_map=numbers_map, sheet_name="CourtBoards", matchups=False)
    write_courtboards_ws(wb, schedule_df, numbers_map=numbers_map, sheet_name="CourtBoards (Match-ups)", matchups=True)
    write_gametally_ws(wb, games_played, notes, df_active)
    write_players_ws(wb, players_sheet_df)


    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)

    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    filename = f"netball_courtboards_{stamp}.xlsx"

    # Save to disk in absolute /generated folder
    out_dir = os.path.join(app.root_path, "generated")
    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, filename)
    with open(out_path, "wb") as f_out:
        f_out.write(bio.getbuffer())
    print("Saved workbook to:", out_path)

    return render_template("result.html",
                           filename=filename,
                           timestamp=datetime.now().strftime("%Y-%m-%d %H:%M"),
                           rounds=rounds, courts=courts, min_p1=min_p1, number_players="YES" if number_players else "NO",
                           players=num_active)

@app.route("/download/<path:filename>")
def download_file(filename):
    directory = os.path.join(app.root_path, "generated")
    return send_from_directory(directory, filename, as_attachment=True)

if __name__ == "__main__":
    app.run(debug=True)
