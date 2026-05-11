import os
import subprocess
import tempfile
import threading
import webbrowser
from pathlib import Path
from uuid import uuid4

import pandas as pd
from dotenv import load_dotenv
from flask import Flask, flash, redirect, render_template, request, send_from_directory, session, url_for
from werkzeug.utils import secure_filename

from modules.fleet.registry import DEFAULT_FLEET_KEY, get_fleet_profile, list_fleet_profiles

load_dotenv()

BASE_DIR = Path(__file__).resolve().parent
UPLOAD_DIR = Path(tempfile.gettempdir()) / "insightflow_uploads"
ALLOWED_EXTENSIONS = {".xlsx", ".xls", ".xlsm"}
STARTUP_TOKEN = uuid4().hex

app = Flask(__name__)
base_secret = os.getenv("FLASK_SECRET_KEY", "insightflow-local-web-ui")
app.secret_key = f"{base_secret}:{STARTUP_TOKEN}"


def _resolve_folder(path_value, default_name):
    raw_value = str(path_value or default_name).strip() or default_name
    folder = Path(raw_value)
    if not folder.is_absolute():
        folder = BASE_DIR / folder
    folder.mkdir(parents=True, exist_ok=True)
    return folder


def _output_folder():
    return _resolve_folder(os.getenv("OUTPUT_FOLDER", "outputs"), "outputs")


def _open_url(url):
    opera_candidates = []

    opera_path = str(os.getenv("OPERA_PATH", "")).strip()
    if opera_path:
        opera_candidates.append(Path(opera_path))

    local_app_data = os.getenv("LOCALAPPDATA")
    program_files = os.getenv("PROGRAMFILES")
    program_files_x86 = os.getenv("PROGRAMFILES(X86)")

    if local_app_data:
        opera_candidates.append(Path(local_app_data) / "Programs" / "Opera" / "opera.exe")
    if program_files:
        opera_candidates.append(Path(program_files) / "Opera" / "launcher.exe")
    if program_files_x86:
        opera_candidates.append(Path(program_files_x86) / "Opera" / "launcher.exe")

    for candidate in opera_candidates:
        if candidate.exists():
            try:
                subprocess.Popen([str(candidate), url])
                return
            except OSError:
                pass

    webbrowser.open(url)


def _reset_runtime_state():
    if not UPLOAD_DIR.exists():
        return

    for file_path in UPLOAD_DIR.rglob("*"):
        if file_path.is_file():
            try:
                file_path.unlink()
            except OSError:
                pass


def _ensure_output_structure():
    root = _output_folder()
    for profile in list_fleet_profiles():
        profile_folder = _resolve_folder(root / profile.output_subfolder, profile.output_subfolder)
        _resolve_folder(profile_folder / "mensal", "mensal")


def _active_profile():
    return get_fleet_profile(session.get("fleet_profile", DEFAULT_FLEET_KEY))


def _upload_folder(profile=None):
    profile = profile or _active_profile()
    folder = UPLOAD_DIR / profile.upload_subfolder
    folder.mkdir(parents=True, exist_ok=True)
    return folder


def _profile_output_folder(profile=None):
    profile = profile or _active_profile()
    return _resolve_folder(_output_folder() / profile.output_subfolder, profile.output_subfolder)


def _allowed_file(filename):
    return Path(filename).suffix.lower() in ALLOWED_EXTENSIONS


def _load_sheet_names(excel_path):
    return pd.ExcelFile(excel_path).sheet_names


def _session_excel_path():
    excel_path = session.get("excel_path", "")
    if excel_path and os.path.exists(excel_path):
        return excel_path
    return ""


def _resolve_sheet_name(sheet_value, sheet_names):
    raw_value = str(sheet_value or "")
    if raw_value in sheet_names:
        return raw_value

    normalized = raw_value.strip()
    if not normalized:
        return ""

    for sheet_name in sheet_names:
        if str(sheet_name).strip() == normalized:
            return sheet_name
    return ""


def _valid_weekly_options(sheet_names, observation_sheet):
    return [sheet for sheet in sheet_names if sheet != observation_sheet]


def _page_context(results=None):
    profile = _active_profile()
    excel_path = _session_excel_path()
    if excel_path:
        try:
            sheet_names = _load_sheet_names(excel_path)
        except Exception:
            sheet_names = []
    else:
        sheet_names = []

    observation_sheet = _resolve_sheet_name(session.get("observation_sheet", ""), sheet_names)
    if not observation_sheet:
        observation_sheet = profile.processor.detect_observation_sheet(sheet_names)

    weekly_options = _valid_weekly_options(sheet_names, observation_sheet)
    weekly_sheet = _resolve_sheet_name(session.get("weekly_sheet", ""), weekly_options)

    return {
        "uploaded_name": session.get("uploaded_name", ""),
        "fleet_profiles": list_fleet_profiles(),
        "selected_fleet_key": profile.key,
        "selected_fleet_label": profile.label,
        "sheet_names": sheet_names,
        "weekly_options": weekly_options,
        "weekly_sheet": weekly_sheet,
        "observation_sheet": observation_sheet,
        "generate_all_month_sheets": session.get("generate_all_month_sheets", False),
        "results": results or [],
        "output_folder": str(_output_folder()),
        "asset_version": STARTUP_TOKEN,
    }


@app.get("/")
def index():
    return render_template("index.html", **_page_context())


@app.post("/load-sheets")
def load_sheets():
    profile = get_fleet_profile(request.form.get("fleet_profile"))
    session["fleet_profile"] = profile.key
    uploaded_file = request.files.get("excel_file")
    if not uploaded_file or not uploaded_file.filename:
        flash("Selecione um arquivo Excel para continuar.", "error")
        return render_template("index.html", **_page_context())

    if not _allowed_file(uploaded_file.filename):
        flash("Formato inválido. Use arquivos .xlsx, .xls ou .xlsm.", "error")
        return render_template("index.html", **_page_context())

    upload_folder = _upload_folder(profile)

    previous_path = _session_excel_path()
    if previous_path:
        try:
            os.remove(previous_path)
        except OSError:
            pass

    original_name = secure_filename(uploaded_file.filename) or "planilha.xlsx"
    saved_name = f"{uuid4().hex}_{original_name}"
    saved_path = upload_folder / saved_name
    uploaded_file.save(saved_path)

    try:
        sheet_names = _load_sheet_names(saved_path)
    except Exception as exc:
        saved_path.unlink(missing_ok=True)
        flash(f"Nao foi possivel ler a planilha: {exc}", "error")
        return render_template("index.html", **_page_context())

    session["excel_path"] = str(saved_path)
    session["uploaded_name"] = uploaded_file.filename
    session["weekly_sheet"] = ""
    session["observation_sheet"] = profile.processor.detect_observation_sheet(sheet_names)
    session["generate_all_month_sheets"] = False

    flash(
        f"Planilha carregada com sucesso em {profile.label}: "
        f"{uploaded_file.filename} ({len(sheet_names)} aba(s) encontrada(s)).",
        "success",
    )
    return redirect(url_for("index"))


@app.post("/generate")
def generate_reports():
    profile = _active_profile()
    excel_path = _session_excel_path()
    if not excel_path:
        flash("Envie a planilha Excel antes de gerar os relatórios.", "error")
        return render_template("index.html", **_page_context())

    generate_weekly = request.form.get("generate_weekly") == "on"
    generate_observation = request.form.get("generate_observation") == "on"
    generate_monthly = request.form.get("generate_monthly") == "on"
    generate_all_month_sheets = request.form.get("generate_all_month_sheets") == "on"

    if not any([generate_weekly, generate_observation, generate_monthly, generate_all_month_sheets]):
        flash("Selecione pelo menos um relatório para gerar.", "error")
        return render_template("index.html", **_page_context())

    sheet_names = _load_sheet_names(excel_path)
    observation_sheet = _resolve_sheet_name(request.form.get("observation_sheet", ""), sheet_names)

    weekly_options = _valid_weekly_options(sheet_names, observation_sheet)
    weekly_sheet = _resolve_sheet_name(request.form.get("weekly_sheet", ""), weekly_options)

    session["weekly_sheet"] = weekly_sheet
    session["observation_sheet"] = observation_sheet
    session["generate_all_month_sheets"] = generate_all_month_sheets

    output_root = _output_folder()
    output_folder = _profile_output_folder(profile)
    results = []
    generated_files = set()

    def add_result(label, file_path):
        filename = os.path.basename(file_path)
        relative_path = os.path.relpath(file_path, output_root).replace("\\", "/")
        dedupe_key = relative_path.lower()
        if dedupe_key in generated_files:
            return

        generated_files.add(dedupe_key)
        results.append(
            {
                "label": label,
                "filename": filename,
                "download_path": relative_path,
            }
        )

    if generate_weekly:
        if not weekly_sheet:
            flash("Escolha a aba semanal para gerar o relatório semanal.", "error")
            return render_template("index.html", **_page_context())

        weekly_pdf = profile.processor.create_weekly_report(excel_path, weekly_sheet, str(output_folder))
        if not weekly_pdf:
            flash("Nenhum dado foi encontrado para a aba semanal selecionada.", "error")
            return render_template("index.html", **_page_context())

        add_result(f"Relatorio semanal - {profile.label}", weekly_pdf)

    if generate_all_month_sheets:
        if not weekly_options:
            flash("Nenhuma aba semanal foi encontrada para gerar os relatórios do mês.", "error")
            return render_template("index.html", **_page_context())

        for sheet_name in weekly_options:
            weekly_pdf = profile.processor.create_weekly_report(excel_path, sheet_name, str(output_folder))
            if weekly_pdf:
                add_result(f"Relatorio semanal - {profile.label} - {sheet_name}", weekly_pdf)

    if generate_observation:
        if not observation_sheet:
            flash("Escolha a aba de observacoes para gerar esse relatório.", "error")
            return render_template("index.html", **_page_context())

        observation_pdf = profile.processor.create_observation_report(excel_path, observation_sheet, str(output_folder))
        if observation_pdf:
            add_result(f"Relatorio de observacoes - {profile.label}", observation_pdf)
        else:
            flash("Nao foi possivel gerar o relatório de observacoes.", "error")
            return render_template("index.html", **_page_context())

    if generate_monthly:
        if not weekly_options:
            flash("Nenhuma aba semanal foi encontrada para gerar o relatório mensal.", "error")
            return render_template("index.html", **_page_context())

        monthly_pdf = profile.processor.create_monthly_report(
            excel_path,
            observation_sheet or None,
            str(output_folder),
            selected_sheets=weekly_options,
        )
        if monthly_pdf:
            add_result(f"Relatorio mensal - {profile.label}", monthly_pdf)
        else:
            flash("Nao foi possivel gerar o relatório mensal.", "error")
            return render_template("index.html", **_page_context())

    flash("Relatorios gerados com sucesso.", "success")
    return render_template("index.html", **_page_context(results=results))


@app.get("/downloads/<path:filename>")
def download_file(filename):
    return send_from_directory(_output_folder(), filename, as_attachment=False)


@app.after_request
def add_no_cache_headers(response):
    if request.path == "/" or response.mimetype in {"text/html", "text/css", "application/javascript"}:
        response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
        response.headers["Pragma"] = "no-cache"
        response.headers["Expires"] = "0"
    return response


def run_app(host="127.0.0.1", port=5000, open_browser=True):
    _reset_runtime_state()
    _ensure_output_structure()

    if open_browser:
        threading.Timer(1.0, lambda: _open_url(f"http://{host}:{port}")).start()

    print("\n>>> SYSTEM STARTED: InsightFlow Web")
    print(f"[*] Interface disponível em: http://{host}:{port}")
    print(f"[*] PDFs gerados em: {_output_folder()}\n")
    app.run(host=host, port=port, debug=False)


if __name__ == "__main__":
    run_app()
