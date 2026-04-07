import os
from dotenv import load_dotenv

load_dotenv("config.env")


def get(key: str, default: str = "") -> str:
    return os.getenv(key, default)


EXCEL_PATH = get("EXCEL_PATH", "data/report.xlsx")
COL_GROUP = get("COL_GROUP", "Управление")
COL_CAPACITY = get("COL_CAPACITY", "Capacity, ч")
COL_SPENT = get("COL_SPENT", "Потрачено, ч")

MAIL_SUBJECT = get("MAIL_SUBJECT", "Отчёт по списанию часов за неделю")
MAIL_TO = [a.strip() for a in get("MAIL_TO").split(",") if a.strip()]
MAIL_CC = [a.strip() for a in get("MAIL_CC").split(",") if a.strip()]
