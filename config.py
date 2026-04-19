import os
from dotenv import load_dotenv

load_dotenv("config.env")


def get(key: str, default: str = "") -> str:
    return os.getenv(key, default)


SECRET_KEY = get("SECRET_KEY", os.urandom(32).hex())

EXCEL_PATH = get("EXCEL_PATH", "")
COL_GROUP = get("COL_GROUP", "Управление")
COL_DEPT  = get("COL_DEPT",  "Отдел")
COL_CAPACITY = get("COL_CAPACITY", "Capacity, ч")
COL_SPENT = get("COL_SPENT", "Потрачено, ч")
COL_DATE = get("COL_DATE", "Дата")

MAIL_SUBJECT = get("MAIL_SUBJECT", "Отчёт по списанию часов за неделю")
MAIL_TO = [a.strip() for a in get("MAIL_TO").split(",") if a.strip()]
MAIL_CC = [a.strip() for a in get("MAIL_CC").split(",") if a.strip()]

# SMTP (для Flask-версии)
SMTP_HOST = get("SMTP_HOST", "")
SMTP_PORT = int(get("SMTP_PORT", "587"))
SMTP_FROM = get("SMTP_FROM", "")   # email отправителя; если пусто — используется AD-логин

# Active Directory
AD_SERVER = get("AD_SERVER")          # ldap://dc01.company.ru
AD_BASE_DN = get("AD_BASE_DN")        # DC=company,DC=ru
AD_USER = get("AD_USER")              # DOMAIN\svc_reports  или  svc_reports@company.ru
AD_PASSWORD = get("AD_PASSWORD")
AD_USE_NTLM = get("AD_USE_NTLM", "false").lower() == "true"

# Фильтр по департаменту (поле department в AD); пусто — без фильтра
AD_DEPARTMENT = get("AD_DEPARTMENT")

# Режим поиска: "title" (по должности) или "group" (по группе AD)
AD_SEARCH_BY = get("AD_SEARCH_BY", "title")
AD_TITLE_MASK = get("AD_TITLE_MASK", "*Lead*")   # маска для поиска по title
AD_GROUP_DN = get("AD_GROUP_DN")                  # DN группы при AD_SEARCH_BY=group

# Заглушка для локального тестирования без AD (AD_STUB=true)
AD_STUB = get("AD_STUB", "false").lower() == "true"
