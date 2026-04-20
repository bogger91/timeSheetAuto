"""
Отправка email-отчётов.

Два режима:
  - SMTP (Flask-версия): send_smtp() — через smtplib, не нужен Outlook.
  - Outlook COM (CLI/GUI): create_draft(), send() — через win32com, только Windows.
"""
import os
import json
import smtplib
import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import config

_TEMPLATE_FILE = os.path.join(os.path.dirname(__file__), "template.json")


def load_template() -> dict:
    """Загружает шаблон из template.json. Отсутствующие ключи заменяются дефолтами."""
    try:
        with open(_TEMPLATE_FILE, encoding="utf-8") as f:
            data = json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        data = {}
    return {
        "greeting": data.get("greeting", DEFAULT_GREETING),
        "intro":    data.get("intro",    DEFAULT_INTRO),
        "footer":   data.get("footer",   DEFAULT_FOOTER),
        "cc":       data.get("cc",       ""),
    }


def save_template(greeting: str | None = None,
                  intro: str | None = None,
                  footer: str | None = None,
                  cc: str | None = None) -> None:
    """Сохраняет шаблон в template.json. None — не менять поле."""
    try:
        with open(_TEMPLATE_FILE, encoding="utf-8") as f:
            data = json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        data = {}
    if greeting is not None:
        data["greeting"] = greeting
    if intro is not None:
        data["intro"] = intro
    if footer is not None:
        data["footer"] = footer
    if cc is not None:
        data["cc"] = cc
    with open(_TEMPLATE_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def _get_outlook():
    try:
        import win32com.client
    except ImportError:
        raise ImportError(
            "Пакет pywin32 не установлен. Выполните: pip install pywin32"
        )
    return win32com.client.Dispatch("Outlook.Application")


DEFAULT_GREETING = "Добрый день,"
DEFAULT_INTRO = "Направляю еженедельный отчёт по проценту списания рабочих часов по подразделениям разработки{period_str}."
DEFAULT_FOOTER = "Письмо сформировано автоматически. Данные из FineBI / Jira."


def build_html_body(table_html: str, period: str | None = None,
                    greeting: str | None = None,
                    intro: str | None = None,
                    footer: str | None = None) -> str:
    today = datetime.date.today().strftime("%d.%m.%Y")
    period_str = f" за период <b>{period}</b>" if period else f" по состоянию на <b>{today}</b>"

    def nl2br(text: str) -> str:
        return text.replace("\r\n", "<br>").replace("\n", "<br>")

    greeting_text = nl2br(greeting if greeting is not None else DEFAULT_GREETING)
    intro_text = nl2br((intro if intro is not None else DEFAULT_INTRO).replace("{period_str}", period_str))
    footer_text = nl2br(footer if footer is not None else DEFAULT_FOOTER)

    return f"""
<html>
<body style="font-family:Calibri,Arial,sans-serif;font-size:11pt;color:#000;">
  <p>{greeting_text}</p>
  <p>{intro_text}</p>
  {table_html}
  <br>
  <p style="color:#666;font-size:9pt;">
    {footer_text}
  </p>
</body>
</html>
"""


def create_draft(table_html: str, save_as_msg: bool = False, msg_path: str = "draft.msg"):
    """
    Создаёт черновик письма в Outlook.

    Параметры
    ---------
    table_html   : HTML-таблица из parser.pivot_to_html()
    save_as_msg  : если True — дополнительно сохраняет .msg-файл рядом
    msg_path     : путь для сохранения .msg (используется при save_as_msg=True)
    """
    outlook = _get_outlook()
    mail = outlook.CreateItem(0)  # 0 = olMailItem

    mail.Subject = config.MAIL_SUBJECT
    mail.HTMLBody = build_html_body(table_html)

    for addr in config.MAIL_TO:
        mail.Recipients.Add(addr).Type = 1  # olTo

    for addr in config.MAIL_CC:
        recipient = mail.Recipients.Add(addr)
        recipient.Type = 2  # olCC

    mail.Recipients.ResolveAll()

    if save_as_msg:
        abs_path = os.path.abspath(msg_path)
        mail.SaveAs(abs_path, 3)  # 3 = olMSG
        print(f"Черновик сохранён как MSG: {abs_path}")

    mail.Save()  # сохраняет в папку «Черновики» в Outlook
    print("Черновик добавлен в папку «Черновики» Outlook.")
    return mail


def send(table_html: str):
    """Создаёт и сразу отправляет письмо."""
    mail = create_draft(table_html)
    mail.Send()
    print("Письмо отправлено.")


# ---------------------------------------------------------------------------
# SMTP — для Flask-версии (не требует Outlook)
# ---------------------------------------------------------------------------

def send_smtp(smtp_host: str, smtp_port: int,
              smtp_user: str, smtp_password: str,
              mail_from: str, mail_to: str,
              subject: str, html_body: str,
              cc: str = "") -> str:
    """
    Отправляет одно письмо через SMTP (STARTTLS).

    cc — строка с адресами через запятую или пустая строка.
    Возвращает 'ok' или 'error: <текст ошибки>'.
    """
    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject or config.MAIL_SUBJECT
    msg["From"] = mail_from
    msg["To"] = mail_to
    cc_list = [a.strip() for a in cc.split(",") if a.strip()] if cc else []
    if cc_list:
        msg["Cc"] = ", ".join(cc_list)
    msg.attach(MIMEText(html_body, "html", "utf-8"))

    all_rcpt = [mail_to] + cc_list
    try:
        with smtplib.SMTP(smtp_host, smtp_port, timeout=15) as srv:
            srv.ehlo()
            srv.starttls()
            srv.ehlo()
            srv.login(smtp_user, smtp_password)
            srv.sendmail(mail_from, all_rcpt, msg.as_string())
        return "ok"
    except Exception as e:
        return f"error: {e}"
