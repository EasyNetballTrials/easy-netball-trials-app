import io, os
from datetime import datetime
from flask import Flask, request, render_template, send_file, abort
from werkzeug.utils import secure_filename
from scheduler.schema import validate_and_normalize_csv
from scheduler.template_writer import inject_players_csv

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100MB

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PATH = os.path.join(BASE_DIR, "template", "Netball_Trials_Template_Web.xlsm")

@app.get("/")
def index():
    # Uses templates/index.html
    return render_template("index.html")

@app.post("/build")
def build():
    if "file" not in request.files:
        abort(400, "No file provided")
    f = request.files["file"]
    if f.filename == "":
        abort(400, "No selected file")

    raw = f.read()
    csv_bytes = io.BytesIO(raw)

    # Validate + normalize CSV -> list[dict]
    rows = validate_and_normalize_csv(csv_bytes)

    # Inject into template (preserve VBA) -> shell workbook
    out_stream = io.BytesIO()
    timestamp = datetime.now().strftime("%Y%m%d_%H%M")
    out_name = f"Netball_Trials_{timestamp}.xlsm"

    inject_players_csv(TEMPLATE_PATH, rows, out_stream, shell_mode=True)

    out_stream.seek(0)
    return send_file(
        out_stream,
        mimetype="application/vnd.ms-excel.sheet.macroEnabled.12",
        as_attachment=True,
        download_name=secure_filename(out_name),
    )

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000, debug=True)
