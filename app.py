import os
import uuid
from flask import Flask, render_template, request, redirect, url_for, send_file, flash
from werkzeug.utils import secure_filename

# Helper modules (in C:\netball-app\scheduler\)
from scheduler.csv_ingest import read_players_csv_to_df
from scheduler.capacity import capacity_report
from scheduler.template_writer import write_into_template
from scheduler.webbridge import sanity_log_players_header   # ✅ debug helper

APP_ROOT = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = os.path.join(APP_ROOT, 'uploads')
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# ---- Auto-locate your macro template (.xlsm) in one of these folders ----
TEMPLATE_CANDIDATES = [
    os.path.join(APP_ROOT, 'app_data',  'Netball_Trials_Scheduler Download for Web.xlsm'),
    os.path.join(APP_ROOT, 'template',  'Netball_Trials_Scheduler Download for Web.xlsm'),
    os.path.join(APP_ROOT, 'templates', 'Netball_Trials_Scheduler Download for Web.xlsm'),
]
TEMPLATE_XLSM = next((p for p in TEMPLATE_CANDIDATES if os.path.exists(p)), None)
if not TEMPLATE_XLSM:
    raise FileNotFoundError(
        "Template .xlsm not found. Place "
        "'Netball_Trials_Scheduler Download for Web.xlsm' in app_data/ OR template/ OR templates/."
    )

ALLOWED_EXTENSIONS = {'.csv'}

app = Flask(__name__)
app.secret_key = os.environ.get('FLASK_SECRET', 'dev-secret')
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB


def allowed_file(filename: str) -> bool:
    return os.path.splitext(filename)[1].lower() in ALLOWED_EXTENSIONS


@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        # ---- Read form params ----
        try:
            num_courts = int(request.form.get('num_courts', '2'))
            num_rounds = int(request.form.get('num_rounds', '10'))
            numbering  = request.form.get('numbering', 'yes').lower() == 'yes'
            show_prefs = request.form.get('show_prefs', 'no').lower() == 'yes'
        except Exception:
            flash('Invalid parameters. Please check number of courts/rounds.', 'error')
            return redirect(url_for('index'))

        # ---- File upload ----
        f = request.files.get('players_csv')
        if not f or f.filename.strip() == '':
            flash('Please choose a CSV file.', 'error')
            return redirect(url_for('index'))
        if not allowed_file(f.filename):
            flash('Only .csv files are accepted.', 'error')
            return redirect(url_for('index'))

        token = uuid.uuid4().hex[:8]
        csv_path = os.path.join(UPLOAD_FOLDER, f'players_{token}.csv')
        out_name = f'netball_{token}.xlsm'
        out_path = os.path.join(UPLOAD_FOLDER, out_name)
        f.save(csv_path)

        # ---- Read CSV robustly (any header order; Number optional) ----
        try:
            df = read_players_csv_to_df(csv_path)
        except Exception as e:
            flash(f'Could not read CSV: {e}', 'error')
            return redirect(url_for('index'))

        # ---- Capacity check BEFORE writing template ----
        rep = capacity_report(df, num_courts=num_courts, num_rounds=num_rounds)
        if not rep['ok']:
            return render_template('result.html', report=rep, download_url=None)

        # ---- Write into your macro template (preserves VBA/hidden sheets) ----
        try:
            write_into_template(
                template_path=TEMPLATE_XLSM,
                out_path=out_path,
                df_players=df,
                num_courts=num_courts,
                num_rounds=num_rounds,
                numbering=numbering,
                show_prefs=show_prefs,  # writes YES/NO to Control
            )
            # ✅ Debug: print out Players header row that was written
            print(sanity_log_players_header(out_path))

        except Exception as e:
            flash(f'Failed to write template: {e}', 'error')
            return redirect(url_for('index'))

        # Ready to download
        dl_url = url_for('download_file', fname=out_name)
        rep['message'] = 'Workbook generated successfully.'
        return render_template('result.html', report=rep, download_url=dl_url)

    # GET
    return render_template('index.html')


@app.route('/download/<fname>')
def download_file(fname):
    path = os.path.join(UPLOAD_FOLDER, fname)
    if not os.path.exists(path):
        flash('File no longer exists. Please re-generate.', 'error')
        return redirect(url_for('index'))
    return send_file(path, as_attachment=True, download_name=fname)


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5000)), debug=True)
