"""
Отправка email-отчётов.

Два режима:
  - SMTP (Flask-версия): send_smtp() — через smtplib, не нужен Outlook.
  - Outlook COM (CLI/GUI): create_draft(), send() — через win32com, только Windows.
"""
import os
import smtplib
import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import config


def _get_outlook():
    try:
        import win32com.client
    except ImportError:
        raise ImportError(
            "Пакет pywin32 не установлен. Выполните: pip install pywin32"
        )
    return win32com.client.Dispatch("Outlook.Application")


def build_html_body(table_html: str) -> str:
    today = datetime.date.today().strftime("%d.%m.%Y")
    return f"""
<html>
<body style="font-family:Calibri,Arial,sans-serif;font-size:11pt;color:#000;">
  <p>Добрый день,</p>
  <p>
    Направляю еженедельный отчёт по проценту списания рабочих часов
    по подразделениям разработки по состоянию на <b>{today}</b>.
  </p>
  {table_html}
  <br>
  <p style="color:#666;font-size:9pt;">
    Письмо сформировано автоматически. Данные из FineBI / Jira.
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

def send_smtp(table_html: str, recipient: str,
              smtp_host: str, smtp_port: int,
              smtp_user: str, smtp_password: str,
              from_addr: str, subject: str = None) -> str:
    """
    Отправляет одно письмо через SMTP (STARTTLS).

    Возвращает 'ok' или 'error: <текст ошибки>'.
    """
    subj = subject or config.MAIL_SUBJECT

    msg = MIMEMultipart("alternative")
    msg["Subject"] = subj
    msg["From"] = from_addr
    msg["To"] = recipient
    msg.attach(MIMEText(build_html_body(table_html), "html", "utf-8"))

    try:
        with smtplib.SMTP(smtp_host, smtp_port, timeout=15) as srv:
            srv.ehlo()
            srv.starttls()
            srv.ehlo()
            srv.login(smtp_user, smtp_password)
            srv.sendmail(from_addr, [recipient], msg.as_string())
        return "ok"
    except Exception as e:
        return f"error: {e}"
