import io, os
from datetime import datetime
from flask import (
    Flask, request, render_template, send_file,
    redirect, url_for, flash
)
from werkzeug.utils import secure_filename

# our modules
from scheduler.schema import validate_and_normalize_csv
from scheduler.excel_writer import inject_players_csv

app = Flask(__name__)
app.secret_key = "dev-only-secret"   # needed for flash()
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100MB

HERE = os.path.dirname(os.path.abspath(__file__))
ASSETS_DIR = os.path.join(HERE, "template")
TEMPLATE_PATH = os.path.join(ASSETS_DIR, "Netball_Trials_Template_Web.xlsm")

# --- helpers -----------------------------------------------------

def _pick_upload_file():
    """
    Accept 'file' (current form) or 'players_csv' (older form) and
    return a werkzeug FileStorage or None.
    """
    # Log what we received (helpful while testing)
    app.logger.info("request.files keys: %s", list(request.files.keys()))
    app.logger.info("request.form keys: %s", list(request.form.keys()))

    fs = request.files.get("file")
    if fs is None or fs.filename == "":
        fs = request.files.get("players_csv")  # backward-compat
        if fs is None or fs.filename == "":
            return None
    return fs

def _looks_like_csv(filename: str) -> bool:
    return filename.lower().endswith(".csv")

# --- routes ------------------------------------------------------

@app.get("/")
def index():
    return render_template("index.html")

# If someone browses directly to /build, send them back to the form.
@app.get("/build")
def build_get():
    return redirect(url_for("index"))

@app.post("/build")
def build_post():
    fs = _pick_upload_file()
    if fs is None:
        flash("No file provided. Please choose your Players CSV and try again.", "error")
        return redirect(url_for("index"))

    if not _looks_like_csv(fs.filename):
        flash("That doesn’t look like a CSV file. Please upload a .csv.", "error")
        return redirect(url_for("index"))

    # Read/normalize CSV
    rows = validate_and_normalize_csv(io.BytesIO(fs.read()))

    # Write into a copy of the template (Players tab only; macros preserved)
    out_stream = io.BytesIO()
    inject_players_csv(TEMPLATE_PATH, rows, out_stream, shell_mode=True)

    out_stream.seek(0)
    ts = datetime.now().strftime("%Y%m%d_%H%M")
    out_name = f"Netball_Trials_{ts}.xlsm"
    return send_file(
        out_stream,
        mimetype="application/vnd.ms-excel.sheet.macroEnabled.12",
        as_attachment=True,
        download_name=secure_filename(out_name),
    )

# local dev
if __name__ == "__main__":
    # visit http://127.0.0.1:10000/
    app.run(host="0.0.0.0", port=10000, debug=True)
