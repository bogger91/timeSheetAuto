"""
Flask-приложение для автоматизации отчёта по списанию часов.

Маршруты:
  GET  /              → редирект на /dashboard или /login
  GET  /login         → форма ввода кредов (AD + SMTP)
  POST /login         → сохранить в сессию, редирект на /dashboard
  GET  /logout        → очистить сессию, редирект на /login

  GET  /dashboard     → сводная таблица + форма загрузки Excel
  POST /upload        → загрузить Excel, разобрать, сохранить в сессию

  GET  /recipients    → список адресатов
  POST /recipients/fetch  → загрузить из AD
  POST /recipients/save   → сохранить отредактированный список

  POST /send          → отправить рассылку, вернуть JSON {email: status}
"""
import io
import os
import tempfile
import functools
import logging
import traceback

import pandas as pd
from flask import (Flask, render_template, request, redirect,
                   url_for, session, jsonify, flash)
from flask_session import Session

import config
import parser as report_parser
import ad_fetcher
import mailer

logging.basicConfig(
    level=logging.DEBUG,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler("timesheetauto.log", encoding="utf-8"),
    ],
)
log = logging.getLogger("timesheetauto")

# ---------------------------------------------------------------------------
# Приложение
# ---------------------------------------------------------------------------

app = Flask(__name__)
app.secret_key = os.urandom(32)

# Серверные сессии (хранятся на диске, не в cookie — данные могут быть большими)
app.config["SESSION_TYPE"] = "filesystem"
app.config["SESSION_FILE_DIR"] = os.path.join(tempfile.gettempdir(), "timesheetauto_sessions")
app.config["SESSION_PERMANENT"] = False
Session(app)


# ---------------------------------------------------------------------------
# Декоратор авторизации
# ---------------------------------------------------------------------------

def login_required(f):
    @functools.wraps(f)
    def wrapper(*args, **kwargs):
        if not session.get("logged_in"):
            return redirect(url_for("login"))
        return f(*args, **kwargs)
    return wrapper


# ---------------------------------------------------------------------------
# Авторизация
# ---------------------------------------------------------------------------

@app.route("/")
def index():
    if session.get("logged_in"):
        return redirect(url_for("dashboard"))
    return redirect(url_for("login"))


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        ad_user     = request.form.get("ad_user", "").strip()
        ad_password = request.form.get("ad_password", "")
        session["logged_in"]     = True
        session["ad_user"]       = ad_user
        session["ad_password"]   = ad_password
        session["smtp_host"]     = config.SMTP_HOST
        session["smtp_port"]     = config.SMTP_PORT
        session["smtp_from"]     = config.SMTP_FROM or ad_user
        session["smtp_password"] = ad_password
        return redirect(url_for("dashboard"))

    return render_template("login.html", ad_server=config.AD_SERVER)


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


# ---------------------------------------------------------------------------
# Dashboard
# ---------------------------------------------------------------------------

@app.route("/dashboard")
@login_required
def dashboard():
    pivot = None
    filename = session.get("upload_filename", "")
    error = session.pop("upload_error", None)

    log.debug("dashboard(): session keys=%r, has pivot_json=%r, filename=%r, error=%r",
              list(session.keys()), "pivot_json" in session, filename, error)

    if "pivot_json" in session:
        pivot_json = session["pivot_json"]
        log.debug("pivot_json length=%d chars", len(pivot_json))
        try:
            pivot = pd.read_json(io.StringIO(pivot_json), dtype=False)
            log.debug("pivot restored OK, shape=%s", pivot.shape)
        except Exception as e:
            log.error("Failed to restore pivot from session: %s\n%s", e, traceback.format_exc())
            pivot = None

    # Список уникальных управлений (без итоговой строки) для фильтра
    groups = []
    if pivot is not None:
        groups = [
            d for d in pivot.loc[pivot["row_type"] == "group", "Подразделение"].tolist()
            if str(d).upper() != "ИТОГО"
        ]

    # Словарь отдел → {name, email} для отображения начальников
    teamleads_raw = session.get("teamleads", [])
    dept_to_lead = {tl["department"]: tl for tl in teamleads_raw if tl.get("department")}

    return render_template(
        "dashboard.html",
        pivot=pivot,
        groups=groups,
        dept_to_lead=dept_to_lead,
        filename=filename,
        error=error,
    )


@app.route("/upload", methods=["POST"])
@login_required
def upload():
    file = request.files.get("excel_file")
    log.debug("upload() called, filename=%r", file.filename if file else None)

    if not file or file.filename == "":
        session["upload_error"] = "Файл не выбран."
        return redirect(url_for("dashboard"))

    if not file.filename.lower().endswith((".xlsx", ".xls")):
        session["upload_error"] = "Поддерживаются только файлы .xlsx / .xls."
        return redirect(url_for("dashboard"))

    tmp_path = None
    try:
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
            tmp_path = tmp.name
            file.save(tmp_path)
        file_size = os.path.getsize(tmp_path)
        log.debug("Saved to tmp=%r, size=%d bytes", tmp_path, file_size)

        log.debug("Config: COL_GROUP=%r, COL_CAPACITY=%r, COL_SPENT=%r",
                  config.COL_GROUP, config.COL_CAPACITY, config.COL_SPENT)

        import openpyxl
        wb = openpyxl.load_workbook(tmp_path, read_only=True)
        ws = wb.active
        headers = [cell.value for cell in next(ws.iter_rows(max_row=1))]
        wb.close()
        log.debug("Excel headers: %r", headers)

        pivot = report_parser.load_pivot(tmp_path)
        log.debug("pivot loaded OK, shape=%s, columns=%r", pivot.shape, list(pivot.columns))
        session["pivot_json"] = pivot.to_json(force_ascii=False)
        session["upload_filename"] = file.filename
        log.info("Upload success: %r", file.filename)
    except Exception as e:
        log.error("Upload failed: %s\n%s", e, traceback.format_exc())
        session["upload_error"] = f"Ошибка при чтении файла: {e}"
    finally:
        if tmp_path and os.path.exists(tmp_path):
            os.unlink(tmp_path)

    return redirect(url_for("dashboard"))


# ---------------------------------------------------------------------------
# Адресаты
# ---------------------------------------------------------------------------

@app.route("/recipients")
@login_required
def recipients():
    return render_template(
        "recipients.html",
        recipients=session.get("recipients", []),
        ad_error=session.pop("ad_error", None),
        ad_server=config.AD_SERVER,
    )


@app.route("/recipients/fetch", methods=["POST"])
@login_required
def fetch_recipients():
    try:
        teamleads = ad_fetcher.get_teamleads(
            user=session.get("ad_user"),
            password=session.get("ad_password"),
        )
        session["teamleads"]  = teamleads
        session["recipients"] = [tl["email"] for tl in teamleads]
    except Exception as e:
        session["ad_error"] = str(e)
    return redirect(url_for("recipients"))


@app.route("/recipients/save", methods=["POST"])
@login_required
def save_recipients():
    raw = request.form.get("recipients_text", "")
    session["recipients"] = [
        e.strip() for e in raw.splitlines() if e.strip()
    ]
    return redirect(url_for("recipients"))


# ---------------------------------------------------------------------------
# Отправка рассылки
# ---------------------------------------------------------------------------

@app.route("/send", methods=["POST"])
@login_required
def send():
    if "pivot_json" not in session:
        return jsonify({"error": "Нет данных. Загрузите Excel на главной странице."}), 400

    recipients_list = session.get("recipients", [])
    if not recipients_list:
        return jsonify({"error": "Список адресатов пуст."}), 400

    pivot = pd.read_json(io.StringIO(session["pivot_json"]), dtype=False)
    table_html = report_parser.pivot_to_html(pivot)

    smtp_host = session.get("smtp_host", "")
    smtp_port = session.get("smtp_port", 587)
    smtp_user = session.get("ad_user", "")
    smtp_password = session.get("smtp_password", "")
    from_addr = session.get("smtp_from", "")

    if not smtp_host:
        return jsonify({"error": "SMTP-сервер не задан. Проверьте настройки в форме входа."}), 400

    results = {}
    for email in recipients_list:
        results[email] = mailer.send_smtp(
            table_html=table_html,
            recipient=email,
            smtp_host=smtp_host,
            smtp_port=smtp_port,
            smtp_user=smtp_user,
            smtp_password=smtp_password,
            from_addr=from_addr,
        )

    return jsonify(results)


# ---------------------------------------------------------------------------
# Запуск
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    os.makedirs(app.config["SESSION_FILE_DIR"], exist_ok=True)
    app.run(host="127.0.0.1", port=5000, debug=False)
