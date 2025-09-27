# app.py (only the parts below need changing/adding)
import io, os, traceback
from datetime import datetime
from flask import Flask, request, render_template, send_file, abort
from werkzeug.utils import secure_filename
from scheduler.schema import validate_and_normalize_csv
from scheduler.excel_writer import inject_players_csv

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100MB

HERE = os.path.dirname(os.path.abspath(__file__))
ASSETS_DIR = os.path.join(HERE, "template")
TEMPLATE_PATH = os.path.join(ASSETS_DIR, "Netball_Trials_Template_Web.xlsm")

@app.get("/")
def index():
    return render_template("index.html")

@app.post("/build")
def build():
    try:
        if "file" not in request.files:
            abort(400, "No file provided")
        f = request.files["file"]
        if f.filename == "":
            abort(400, "No selected file")

        # Validate + normalise the CSV
        rows = validate_and_normalize_csv(io.BytesIO(f.read()))

        # Write into a copy of the template (shell-only except Players)
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
    except Exception as e:
        # Log full traceback to server logs (shows up in Render → Logs)
        traceback.print_exc()
        # Show a nice page to the user
        return render_template("invalid.html", message=str(e)), 400

# Optional: a simple 500 handler (catches anything not handled above)
@app.errorhandler(500)
def server_error(e):
    traceback.print_exc()
    return render_template("invalid.html",
                           message="Something went wrong on the server. Please try again."), 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000, debug=True)
