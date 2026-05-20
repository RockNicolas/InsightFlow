import atexit
import os
import subprocess
import tempfile
import threading
import webbrowser
from pathlib import Path
from uuid import uuid4

import pandas as pd
from dotenv import load_dotenv
from flask import Flask, jsonify, request, send_from_directory, session
from flask_cors import CORS
from werkzeug.utils import secure_filename

from modules.fleet.registry import DEFAULT_FLEET_KEY, get_fleet_profile, list_fleet_profiles
from modules.maintenance.db import disconnect_db
from modules.maintenance.store import (
    create_equipment,
    delete_equipment,
    get_fleet_payload,
    update_equipment,
    update_last_maintenance,
)

load_dotenv()

atexit.register(disconnect_db)

BASE_DIR = Path(__file__).resolve().parent
FRONTEND_DIST = BASE_DIR / "frontend" / "dist"
ASSETS_DIR = BASE_DIR / "assets"
UPLOAD_DIR = Path(tempfile.gettempdir()) / "insightflow_uploads"
ALLOWED_EXTENSIONS = {".xlsx", ".xls", ".xlsm"}
STARTUP_TOKEN = uuid4().hex
VITE_DEV_PORT = 5173
_vite_process = None

app = Flask(__name__, static_folder=None)
base_secret = os.getenv("FLASK_SECRET_KEY", "insightflow-local-web-ui")
app.secret_key = f"{base_secret}:{STARTUP_TOKEN}"

CORS(
    app,
    supports_credentials=True,
    origins=[
        "http://127.0.0.1:5000",
        "http://localhost:5000",
        "http://127.0.0.1:5173",
        "http://localhost:5173",
    ],
)


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


def _fleet_profiles_payload():
    return [{"key": profile.key, "label": profile.label} for profile in list_fleet_profiles()]


def _session_payload(results=None):
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
        "fleet_profiles": _fleet_profiles_payload(),
        "selected_fleet_key": profile.key,
        "selected_fleet_label": profile.label,
        "sheet_names": sheet_names,
        "weekly_options": weekly_options,
        "weekly_sheet": weekly_sheet,
        "observation_sheet": observation_sheet,
        "generate_all_month_sheets": session.get("generate_all_month_sheets", False),
        "results": results or [],
        "output_folder": str(_output_folder()),
    }


def _api_ok(data=None, message="", status=200):
    payload = {"ok": True, "message": message, "data": data or {}}
    return jsonify(payload), status


def _api_error(message, status=400):
    return jsonify({"ok": False, "message": message, "data": {}}), status


@app.get("/api/session")
def api_session():
    return _api_ok(_session_payload())


@app.get("/api/maintenance/fleet")
def api_maintenance_fleet():
    try:
        return _api_ok({"fleet": get_fleet_payload()})
    except Exception as exc:
        return _api_error(f"Banco de dados: {exc}", 503)


@app.post("/api/maintenance/equipment")
def api_create_equipment():
    body = request.get_json(silent=True) or {}
    try:
        fleet = create_equipment(
            body.get("category", ""),
            body.get("type", ""),
            body.get("code", ""),
            body.get("note", ""),
            bool(body.get("alert")),
        )
    except ValueError as exc:
        return _api_error(str(exc))

    return _api_ok({"fleet": fleet}, "Equipamento cadastrado com sucesso.")


@app.put("/api/maintenance/equipment")
def api_update_equipment():
    body = request.get_json(silent=True) or {}
    try:
        fleet = update_equipment(
            body.get("category", ""),
            body.get("equipment_id", ""),
            body.get("type", ""),
            body.get("code", ""),
            body.get("note", ""),
            bool(body.get("alert")),
        )
    except ValueError as exc:
        return _api_error(str(exc))

    return _api_ok({"fleet": fleet}, "Equipamento atualizado com sucesso.")


@app.delete("/api/maintenance/equipment")
def api_delete_equipment():
    body = request.get_json(silent=True) or {}
    try:
        fleet = delete_equipment(
            body.get("category", ""),
            body.get("equipment_id", ""),
        )
    except ValueError as exc:
        return _api_error(str(exc))

    return _api_ok({"fleet": fleet}, "Equipamento excluído com sucesso.")


@app.put("/api/maintenance/record")
def api_maintenance_record():
    body = request.get_json(silent=True) or {}
    category = body.get("category", "")
    equipment_id = body.get("equipment_id", "")
    last_maintenance = body.get("last_maintenance", "")

    try:
        fleet = update_last_maintenance(category, equipment_id, last_maintenance)
    except ValueError as exc:
        return _api_error(str(exc))

    return _api_ok({"fleet": fleet}, "Manutenção salva com sucesso.")


@app.post("/api/load-sheets")
def api_load_sheets():
    profile = get_fleet_profile(request.form.get("fleet_profile"))
    session["fleet_profile"] = profile.key
    uploaded_file = request.files.get("excel_file")

    if not uploaded_file or not uploaded_file.filename:
        return _api_error("Selecione um arquivo Excel para continuar.")

    if not _allowed_file(uploaded_file.filename):
        return _api_error("Formato inválido. Use arquivos .xlsx, .xls ou .xlsm.")

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
        return _api_error(f"Nao foi possivel ler a planilha: {exc}")

    session["excel_path"] = str(saved_path)
    session["uploaded_name"] = uploaded_file.filename
    session["weekly_sheet"] = ""
    session["observation_sheet"] = profile.processor.detect_observation_sheet(sheet_names)
    session["generate_all_month_sheets"] = False

    message = (
        f"Planilha carregada com sucesso em {profile.label}: "
        f"{uploaded_file.filename} ({len(sheet_names)} aba(s) encontrada(s))."
    )
    return _api_ok(_session_payload(), message)


@app.post("/api/generate")
def api_generate_reports():
    profile = _active_profile()
    excel_path = _session_excel_path()
    if not excel_path:
        return _api_error("Envie a planilha Excel antes de gerar os relatórios.")

    body = request.get_json(silent=True) or {}
    generate_weekly = bool(body.get("generate_weekly"))
    generate_observation = bool(body.get("generate_observation"))
    generate_monthly = bool(body.get("generate_monthly"))
    generate_all_month_sheets = bool(body.get("generate_all_month_sheets"))

    if not any([generate_weekly, generate_observation, generate_monthly, generate_all_month_sheets]):
        return _api_error("Selecione pelo menos um relatório para gerar.")

    sheet_names = _load_sheet_names(excel_path)
    observation_sheet = _resolve_sheet_name(body.get("observation_sheet", ""), sheet_names)

    weekly_options = _valid_weekly_options(sheet_names, observation_sheet)
    weekly_sheet = _resolve_sheet_name(body.get("weekly_sheet", ""), weekly_options)

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
            return _api_error("Escolha a aba semanal para gerar o relatório semanal.")

        weekly_pdf = profile.processor.create_weekly_report(excel_path, weekly_sheet, str(output_folder))
        if not weekly_pdf:
            return _api_error("Nenhum dado foi encontrado para a aba semanal selecionada.")

        add_result(f"Relatorio semanal - {profile.label}", weekly_pdf)

    if generate_all_month_sheets:
        if not weekly_options:
            return _api_error("Nenhuma aba semanal foi encontrada para gerar os relatórios do mês.")

        for sheet_name in weekly_options:
            weekly_pdf = profile.processor.create_weekly_report(excel_path, sheet_name, str(output_folder))
            if weekly_pdf:
                add_result(f"Relatorio semanal - {profile.label} - {sheet_name}", weekly_pdf)

    if generate_observation:
        if not observation_sheet:
            return _api_error("Escolha a aba de observacoes para gerar esse relatório.")

        observation_pdf = profile.processor.create_observation_report(
            excel_path, observation_sheet, str(output_folder)
        )
        if observation_pdf:
            add_result(f"Relatorio de observacoes - {profile.label}", observation_pdf)
        else:
            return _api_error("Nao foi possivel gerar o relatório de observacoes.")

    if generate_monthly:
        if not weekly_options:
            return _api_error("Nenhuma aba semanal foi encontrada para gerar o relatório mensal.")

        monthly_pdf = profile.processor.create_monthly_report(
            excel_path,
            observation_sheet or None,
            str(output_folder),
            selected_sheets=weekly_options,
        )
        if monthly_pdf:
            add_result(f"Relatorio mensal - {profile.label}", monthly_pdf)
        else:
            return _api_error("Nao foi possivel gerar o relatório mensal.")

    payload = _session_payload(results=results)
    return _api_ok(payload, "Relatorios gerados com sucesso.")


@app.get("/downloads/<path:filename>")
def download_file(filename):
    return send_from_directory(_output_folder(), filename, as_attachment=False)


@app.get("/assets/<path:filename>")
def serve_assets(filename):
    """Logos e imagens usados nos PDFs (pasta assets/ na raiz do projeto)."""
    if not ASSETS_DIR.is_dir():
        return _api_error("Pasta assets/ nao encontrada.", 404)
    return send_from_directory(ASSETS_DIR, filename)


@app.get("/", defaults={"path": ""})
@app.get("/<path:path>")
def serve_spa(path):
    # Evita que rotas /api/* sem handler POST caiam aqui e devolvam 405 confuso
    if path == "api" or path.startswith("api/"):
        return _api_error(
            "Rota da API não encontrada. Pare o servidor (Ctrl+C) e execute python main.py de novo.",
            404,
        )

    if not FRONTEND_DIST.exists():
        return (
            jsonify(
                {
                    "ok": False,
                    "message": (
                        "Frontend não compilado. Execute: cd frontend && npm install && npm run build"
                    ),
                }
            ),
            503,
        )

    requested = FRONTEND_DIST / path
    if path and requested.is_file():
        return send_from_directory(FRONTEND_DIST, path)

    return send_from_directory(FRONTEND_DIST, "index.html")


@app.after_request
def add_no_cache_headers(response):
    if request.path.startswith("/api") or response.mimetype in {
        "text/html",
        "text/css",
        "application/javascript",
    }:
        response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
        response.headers["Pragma"] = "no-cache"
        response.headers["Expires"] = "0"
    return response


def _npm_cmd():
    return "npm.cmd" if os.name == "nt" else "npm"


def _frontend_dir():
    return BASE_DIR / "frontend"


def _frontend_needs_build():
    dist_index = FRONTEND_DIST / "index.html"
    if not dist_index.exists():
        return True

    src_dir = _frontend_dir() / "src"
    if not src_dir.exists():
        return False

    dist_mtime = dist_index.stat().st_mtime
    for path in src_dir.rglob("*"):
        if path.is_file() and path.stat().st_mtime > dist_mtime:
            return True
    return False


def _npm_install(frontend_dir):
    npm = _npm_cmd()
    print("[*] Instalando dependencias do frontend (npm install)...")
    subprocess.run([npm, "install"], cwd=frontend_dir, check=True)


def _npm_build(frontend_dir):
    npm = _npm_cmd()
    print("[*] Compilando frontend React (npm run build)...")
    subprocess.run([npm, "run", "build"], cwd=frontend_dir, check=True)


def _ensure_frontend_built(force=False):
    frontend_dir = _frontend_dir()
    package_json = frontend_dir / "package.json"
    if not package_json.exists():
        print("[!] Pasta frontend/ nao encontrada.")
        return False

    npm = _npm_cmd()
    try:
        subprocess.run([npm, "--version"], cwd=frontend_dir, check=True, capture_output=True)
    except (OSError, subprocess.CalledProcessError):
        print("[!] Node.js/npm nao encontrado. Instale Node.js: https://nodejs.org/")
        return False

    try:
        if not (frontend_dir / "node_modules").exists():
            _npm_install(frontend_dir)
        elif force:
            _npm_install(frontend_dir)

        if force or _frontend_needs_build():
            _npm_build(frontend_dir)

        ready = (FRONTEND_DIST / "index.html").exists()
        if ready:
            print("[*] Frontend pronto em frontend/dist")
        return ready
    except (OSError, subprocess.CalledProcessError) as exc:
        print(f"[!] Nao foi possivel preparar o frontend: {exc}")
        return False


def _stop_vite_dev():
    global _vite_process
    if _vite_process and _vite_process.poll() is None:
        _vite_process.terminate()
        try:
            _vite_process.wait(timeout=5)
        except subprocess.TimeoutExpired:
            _vite_process.kill()
    _vite_process = None


def _start_vite_dev(host="127.0.0.1"):
    global _vite_process
    frontend_dir = _frontend_dir()
    if not (frontend_dir / "package.json").exists():
        print("[!] Pasta frontend/ nao encontrada.")
        return False

    if not (frontend_dir / "node_modules").exists():
        _npm_install(frontend_dir)

    npm = _npm_cmd()
    print(f"[*] Iniciando Vite (hot reload) em http://{host}:{VITE_DEV_PORT} ...")
    _vite_process = subprocess.Popen(
        [npm, "run", "dev", "--", "--host", host, "--port", str(VITE_DEV_PORT)],
        cwd=frontend_dir,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )
    atexit.register(_stop_vite_dev)
    return True


def run_app(host="127.0.0.1", port=5000, open_browser=True, dev_mode=False, rebuild_frontend=False):
    _reset_runtime_state()
    _ensure_output_structure()

    ui_url = f"http://{host}:{port}"

    if dev_mode:
        if not _start_vite_dev(host):
            print("[!] Modo dev indisponivel; usando build estatico.")
            _ensure_frontend_built(force=rebuild_frontend)
        else:
            ui_url = f"http://{host}:{VITE_DEV_PORT}"
    else:
        _ensure_frontend_built(force=rebuild_frontend)

    if open_browser:
        threading.Timer(1.5, lambda: _open_url(ui_url)).start()

    print("\n>>> SYSTEM STARTED: InsightFlow Web")
    print(f"[*] Interface: {ui_url}")
    print(f"[*] API Flask: http://{host}:{port}")
    print(f"[*] PDFs gerados em: {_output_folder()}")
    if dev_mode and _vite_process:
        print("[*] Modo dev: edite frontend/src e salve — a pagina atualiza sozinha")
    elif (FRONTEND_DIST / "index.html").exists():
        print("[*] Frontend React integrado (build servido pelo Flask)")
    else:
        print("[!] Frontend ausente — instale Node.js e rode python main.py novamente")
    if not dev_mode:
        print("[*] Hot reload: python main.py --dev\n")
    else:
        print()

    try:
        app.run(host=host, port=port, debug=False)
    finally:
        _stop_vite_dev()


if __name__ == "__main__":
    run_app()
