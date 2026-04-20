# app.py — patched version for timeSheetAuto
# Changes vs original:
#  1. /upload auto-fetches recipients from AD after successful parse.
#     Builds session["recipients_meta"] — one entry per dept with status ok/manual/missing.
#     On AD failure: session["ad_autofetch_error"] is set; upload still succeeds.
#  2. /recipients (GET) redirects to /dashboard (wizard now includes step 3 inline).
#  3. /recipients/confirm (new, POST JSON) — updates enabled/email per row.
#  4. /preview returns JSON (body_html + subject + template fields) for the wizard iframe.
#  5. /email-template: GET returns JSON, POST accepts JSON body (null = reset to default).
#  6. /send accepts JSON {subject, recipients} and returns per-email status dict.
#  7. /reset_upload — clears pivot + recipients_meta, returns to step 1.
#  8. dashboard() builds `initial_state` for the wizard JS.
#
# Drop-in compatible with existing parser.py, ad_fetcher.py, mailer.py, config.py.

import io
import os
import tempfile
import json
import logging
from flask import (
    Flask, render_template, request, redirect, url_for,
    session, flash, jsonify,
)
from flask_session import Session
import pandas as pd

import config
from parser import load_pivot, load_period, pivot_to_html
import ad_fetcher
import mailer


logging.basicConfig(
    level=logging.DEBUG,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    handlers=[logging.StreamHandler()],
)

app = Flask(__name__)
app.secret_key = config.SECRET_KEY
app.config["MAX_CONTENT_LENGTH"] = 16 * 1024 * 1024  # 16 MB

# Серверные сессии — pivot_json может быть >4KB, в cookie не влезает
app.config["SESSION_TYPE"] = "filesystem"
app.config["SESSION_FILE_DIR"] = os.path.join(tempfile.gettempdir(), "timesheetauto_sessions")
app.config["SESSION_PERMANENT"] = False
Session(app)


# ───────────────────────── helpers ─────────────────────────

def _require_login():
    return session.get("logged_in") is True


def _get_pivot_df():
    raw = session.get("pivot_json")
    if not raw:
        return None
    return pd.read_json(io.StringIO(raw), orient="split")


def _build_recipients_meta(pivot_df, teamleads):
    """
    Build per-dept recipient rows by matching dept rows in pivot
    with AD lead records by lead name.

    teamleads: [{"department": ..., "name": ..., "email": ...}]
    returns:   [{"dept","lead","email","enabled","status"}]
    """
    # Map by dept name (string, case-insensitive)
    leads_by_dept = {}
    for tl in teamleads or []:
        dept = (tl.get("department") or "").strip().lower()
        if dept:
            leads_by_dept[dept] = tl

    meta = []
    if pivot_df is None or pivot_df.empty:
        return meta

    # Include both group (управления) and dept rows — exclude only the total row.
    if "row_type" in pivot_df.columns:
        dept_rows = pivot_df[pivot_df["row_type"].isin(["group", "dept"])]
    else:
        dept_rows = pivot_df

    seen = set()
    for _, row in dept_rows.iterrows():
        dept_name = str(row.get("Подразделение") or row.get("Отдел") or "").strip()
        if not dept_name or dept_name in seen:
            continue
        seen.add(dept_name)
        pct = float(row.get("% списания") or 0)
        above_threshold = pct >= config.EXCLUDE_PCT
        tl = leads_by_dept.get(dept_name.lower())
        if tl and tl.get("email"):
            meta.append({
                "dept": dept_name,
                "lead": tl.get("name") or "",
                "email": tl.get("email"),
                "enabled": not above_threshold,
                "status": "ok",
            })
        elif tl:
            meta.append({
                "dept": dept_name,
                "lead": tl.get("name") or "",
                "email": "",
                "enabled": False,
                "status": "missing",
            })
        else:
            meta.append({
                "dept": dept_name,
                "lead": "",
                "email": "",
                "enabled": False,
                "status": "missing",
            })
    return meta


def _recipients_from_meta():
    meta = session.get("recipients_meta") or []
    return [r["email"] for r in meta if r.get("enabled") and r.get("email")]


def _dept_to_lead_map():
    """For rendering the pivot — show lead name next to each dept."""
    out = {}
    for tl in session.get("teamleads", []) or []:
        dept = (tl.get("department") or "").strip()
        if dept:
            out[dept] = {"name": tl.get("name", ""), "email": tl.get("email", "")}
    return out


def _groups_list(pivot_df):
    if pivot_df is None or pivot_df.empty or "Управление" not in pivot_df.columns:
        return []
    vals = (
        pivot_df.loc[pivot_df.get("row_type", "") != "total", "Управление"]
        .dropna().astype(str).unique().tolist()
    )
    return sorted(vals)


# ───────────────────────── auth ─────────────────────────

@app.route("/", methods=["GET"])
def index():
    if _require_login():
        return redirect(url_for("dashboard"))
    return redirect(url_for("login"))


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        user = request.form.get("ad_user", "").strip()
        pwd = request.form.get("ad_password", "")
        if not user or not pwd:
            flash("Введите логин и пароль.", "danger")
            return render_template("login.html")
        # AD auth is performed lazily when needed; here we just stash credentials.
        session.clear()
        session["logged_in"] = True
        session["ad_user"] = user
        session["ad_password"] = pwd
        return redirect(url_for("dashboard"))
    return render_template("login.html")


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


# ───────────────────────── dashboard (wizard) ─────────────────────────

@app.route("/dashboard", methods=["GET"])
def dashboard():
    if not _require_login():
        return redirect(url_for("login"))

    pivot_df = _get_pivot_df()
    recipients_meta = session.get("recipients_meta") or []
    has_pivot = pivot_df is not None and not pivot_df.empty
    has_recipients = any(r.get("enabled") and r.get("email") for r in recipients_meta)

    # Decide which step to land on
    if not has_pivot:
        initial_step = 1
        completed = []
    elif not recipients_meta:
        initial_step = 2
        completed = [1]
    else:
        # Land on step 2 (pivot review) — user must click Next to proceed
        initial_step = 2
        completed = [1]

    initial_state = {
        "urls": {
            "upload":         url_for("upload"),
            "reset_upload":   url_for("reset_upload"),
            "fetch":          url_for("fetch_recipients"),
            "save_recipients": url_for("save_recipients"),
            "confirm":        url_for("confirm_recipients"),
            "preview":        url_for("preview"),
            "email_template": url_for("email_template"),
            "send":           url_for("send"),
        },
        "from_addr": session.get("ad_user", ""),
        "has_pivot": bool(has_pivot),
        "has_recipients": bool(has_recipients),
        "initial_step": initial_step,
        "completed_steps": completed,
    }

    return render_template(
        "dashboard.html",
        pivot=pivot_df,
        period=session.get("period"),
        filename=session.get("upload_filename"),
        groups=_groups_list(pivot_df),
        dept_to_lead=_dept_to_lead_map(),
        recipients_meta=recipients_meta,
        ad_autofetch_error=session.get("ad_autofetch_error"),
        ad_server=getattr(config, "AD_SERVER", ""),
        cc=mailer.load_template().get("cc", ""),
        initial_state=initial_state,
    )


# ───────────────────────── upload ─────────────────────────

@app.route("/upload", methods=["POST"])
def upload():
    if not _require_login():
        return redirect(url_for("login"))

    f = request.files.get("excel_file")
    if not f or not f.filename:
        flash("Не выбран файл.", "danger")
        return redirect(url_for("dashboard"))

    try:
        pivot_df = load_pivot(f)
        period = load_period(f)
    except Exception as e:
        flash(f"Не удалось разобрать Excel: {e}", "danger")
        return redirect(url_for("dashboard"))

    session["pivot_json"] = pivot_df.to_json(orient="split")
    session["period"] = period
    session["upload_filename"] = f.filename
    session.pop("ad_autofetch_error", None)

    # Auto-fetch recipients from AD
    try:
        teamleads = ad_fetcher.get_teamleads(
            user=session.get("ad_user", ""), password=session.get("ad_password", "")
        )
        session["teamleads"] = teamleads
        session["recipients_meta"] = _build_recipients_meta(pivot_df, teamleads)
        session["recipients"] = _recipients_from_meta()
    except Exception as e:
        session["ad_autofetch_error"] = str(e)
        session["teamleads"] = []
        # Still build meta rows so user can fill emails manually
        session["recipients_meta"] = _build_recipients_meta(pivot_df, [])
        session["recipients"] = []

    return redirect(url_for("dashboard"))


@app.route("/reset_upload", methods=["POST"])
def reset_upload():
    if not _require_login():
        return redirect(url_for("login"))
    for k in ("pivot_json", "period", "upload_filename",
              "teamleads", "recipients", "recipients_meta",
              "ad_autofetch_error"):
        session.pop(k, None)
    return redirect(url_for("dashboard"))


# ───────────────────────── recipients ─────────────────────────

@app.route("/recipients", methods=["GET"])
def recipients_page():
    # Adresaty now live on the dashboard as step 3
    return redirect(url_for("dashboard"))


@app.route("/recipients/fetch", methods=["POST"])
def fetch_recipients():
    if not _require_login():
        return redirect(url_for("login"))
    try:
        teamleads = ad_fetcher.get_teamleads(
            user=session.get("ad_user", ""), password=session.get("ad_password", "")
        )
        session["teamleads"] = teamleads
        pivot_df = _get_pivot_df()
        session["recipients_meta"] = _build_recipients_meta(pivot_df, teamleads)
        session["recipients"] = _recipients_from_meta()
        session.pop("ad_autofetch_error", None)
        flash(f"Загружено руководителей из AD: {len(teamleads)}.", "success")
    except Exception as e:
        session["ad_autofetch_error"] = str(e)
        flash(f"Ошибка AD: {e}", "danger")
    return redirect(url_for("dashboard"))


@app.route("/recipients/save", methods=["POST"])
def save_recipients():
    """Legacy endpoint — kept for backward compatibility.
    Accepts form field `recipients` as newline-separated emails."""
    if not _require_login():
        return redirect(url_for("login"))
    raw = request.form.get("recipients", "")
    emails = [ln.strip() for ln in raw.splitlines() if ln.strip()]
    session["recipients"] = emails
    flash(f"Сохранено получателей: {len(emails)}.", "success")
    return redirect(url_for("dashboard"))


@app.route("/recipients/confirm", methods=["POST"])
def confirm_recipients():
    """New endpoint — merges {idx: {enabled, email}} edits into recipients_meta."""
    if not _require_login():
        return jsonify(error="not logged in"), 401
    data = request.get_json(silent=True) or {}
    rows = data.get("rows") or []
    meta = session.get("recipients_meta") or []
    for patch in rows:
        i = patch.get("idx")
        if not isinstance(i, int) or i < 0 or i >= len(meta):
            continue
        if "enabled" in patch:
            meta[i]["enabled"] = bool(patch["enabled"])
        if "email" in patch:
            email = (patch.get("email") or "").strip()
            meta[i]["email"] = email
            if email and meta[i].get("status") == "missing":
                meta[i]["status"] = "manual"
    session["recipients_meta"] = meta
    session["recipients"] = _recipients_from_meta()
    return jsonify(ok=True, active=len(session["recipients"]))


# ───────────────────────── preview + template ─────────────────────────

@app.route("/preview", methods=["GET"])
def preview():
    if not _require_login():
        return jsonify(error="not logged in"), 401
    pivot_df = _get_pivot_df()
    if pivot_df is None or pivot_df.empty:
        return jsonify(error="Сначала загрузите Excel.")

    period = session.get("period") or ""
    tpl = mailer.load_template()
    greeting = session.get("tpl_greeting") or tpl["greeting"]
    intro    = session.get("tpl_intro")    or tpl["intro"]
    footer   = session.get("tpl_footer")   or tpl["footer"]

    table_html = pivot_to_html(pivot_df)
    body_html = mailer.build_html_body(
        greeting=greeting,
        intro=intro.format(period_str=period) if "{period_str}" in intro else intro,
        table_html=table_html,
        footer=footer,
    )
    subject = getattr(config, "MAIL_SUBJECT", "Отчёт по списанию часов")
    if "{period_str}" in subject:
        subject = subject.format(period_str=period)

    return jsonify(
        subject=subject,
        from_addr=session.get("ad_user", ""),
        body_html=body_html,
        tpl_greeting=greeting,
        tpl_intro=intro,
        tpl_footer=footer,
    )


@app.route("/email-template", methods=["GET", "POST"])
def email_template():
    if not _require_login():
        return jsonify(error="not logged in"), 401

    tpl = mailer.load_template()
    if request.method == "GET":
        return jsonify(
            greeting=session.get("tpl_greeting") or tpl["greeting"],
            intro=session.get("tpl_intro")        or tpl["intro"],
            footer=session.get("tpl_footer")      or tpl["footer"],
            cc=tpl["cc"],
        )

    data = request.get_json(silent=True) or {}
    save_kwargs = {}
    for key, sess_key in (
        ("greeting", "tpl_greeting"),
        ("intro",    "tpl_intro"),
        ("footer",   "tpl_footer"),
    ):
        if key in data:
            val = data[key]
            if val is None:
                session.pop(sess_key, None)
                save_kwargs[key] = getattr(mailer, f"DEFAULT_{key.upper()}")
            else:
                session[sess_key] = str(val)
                save_kwargs[key] = str(val)
    # CC сохраняется сразу в файл — персистентно без сессии
    if "cc" in data:
        save_kwargs["cc"] = (data["cc"] or "").strip()
    if save_kwargs:
        mailer.save_template(**save_kwargs)
    return jsonify(ok=True)


# ───────────────────────── send ─────────────────────────

@app.route("/send", methods=["POST"])
def send():
    if not _require_login():
        return jsonify(error="not logged in"), 401
    pivot_df = _get_pivot_df()
    if pivot_df is None or pivot_df.empty:
        return jsonify(error="Сначала загрузите Excel.")

    data = request.get_json(silent=True) or {}
    recipients = data.get("recipients") or _recipients_from_meta()
    if not recipients:
        return jsonify(error="Нет получателей.")
    subject = (data.get("subject") or "").strip() or getattr(
        config, "MAIL_SUBJECT", "Отчёт по списанию часов"
    )

    period = session.get("period") or ""
    tpl = mailer.load_template()
    greeting = session.get("tpl_greeting") or tpl["greeting"]
    intro    = session.get("tpl_intro")    or tpl["intro"]
    footer   = session.get("tpl_footer")   or tpl["footer"]
    if "{period_str}" in subject:
        subject = subject.format(period_str=period)

    table_html = pivot_to_html(pivot_df)
    body_html = mailer.build_html_body(
        greeting=greeting,
        intro=intro.format(period_str=period) if "{period_str}" in intro else intro,
        table_html=table_html,
        footer=footer,
    )

    cc = mailer.load_template().get("cc", "")

    results = {}
    for email in recipients:
        results[email] = mailer.send_outlook(
            mail_from=session.get("ad_user", ""),
            mail_to=email,
            subject=subject,
            html_body=body_html,
            cc=cc,
        )
    return jsonify(results)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
