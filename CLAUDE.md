# TimeSheet Auto — CLAUDE.md

## Назначение проекта

Инструмент для автоматизации еженедельных отчётов по списанию часов (FineBI/Jira → Excel → HTML-таблица → email). Целевая среда — корпоративная Windows с доменом Active Directory.

## Архитектура

Два режима работы — CLI и Flask-веб-интерфейс — используют общий слой бизнес-логики.

```
config.py        — настройки из config.env (dotenv)
parser.py        — парсинг Excel, построение сводной таблицы + HTML
ad_fetcher.py    — получение email тим-лидов из Active Directory (ldap3)
mailer.py        — отправка: SMTP (основной) или Outlook COM (запасной, только Windows)
app.py           — Flask-приложение (веб-интерфейс)
run.bat          — запуск Flask на Windows без терминала
```

### Flask-маршруты (app.py)

| Метод | URL | Назначение |
|---|---|---|
| GET | `/` | редирект на `/dashboard` или `/login` |
| GET/POST | `/login` | форма AD + SMTP кредов, сохраняет в сессию |
| GET | `/logout` | очистка сессии |
| GET | `/dashboard` | сводная таблица + загрузка Excel |
| POST | `/upload` | загрузить Excel, распарсить, сохранить в сессию |
| GET | `/recipients` | список адресатов |
| POST | `/recipients/fetch` | загрузить адресатов из AD |
| POST | `/recipients/save` | сохранить отредактированный список |
| POST | `/send` | отправить рассылку, вернуть JSON `{email: status}` |

### Шаблоны (Jinja2 + Bootstrap 5.3)

- `templates/base.html` — навбар, flash-сообщения
- `templates/login.html` — форма входа
- `templates/dashboard.html` — таблица, загрузка файла, кнопка отправки
- `templates/recipients.html` — список адресатов
- `static/css/style.css` — кастомные стили

## Конфигурация

Настройки хранятся в `config.env` (не в репозитории). Шаблон:

```env
# Прокси для pip (используется только run.bat при установке пакетов; пусто — без прокси)
PIP_PROXY=http://proxy.company.ru:3128

# Excel
EXCEL_PATH=report.xlsx
COL_GROUP=Управление
COL_DEPT=Отдел
COL_CAPACITY=Capacity, ч
COL_SPENT=Потрачено, ч

# Письмо
MAIL_SUBJECT=Отчёт по списанию часов за неделю
MAIL_TO=boss@company.ru
MAIL_CC=

# SMTP (Flask-режим)
SMTP_HOST=mail.company.ru
SMTP_PORT=587

# Active Directory
AD_SERVER=ldap://dc01.company.ru
AD_BASE_DN=DC=company,DC=ru
AD_USER=DOMAIN\svc_reports
AD_PASSWORD=
AD_USE_NTLM=false
AD_DEPARTMENT=

# Поиск тим-лидов
AD_SEARCH_BY=title          # title | group
AD_TITLE_MASK=*Lead*
AD_GROUP_DN=                # DN группы, если AD_SEARCH_BY=group
```

## Зависимости

```
pandas>=2.0.0
openpyxl>=3.1.0
python-dotenv>=1.0.0
pywin32>=306          # только Windows, нужен для Outlook COM (CLI-режим)
ldap3>=2.9.1
flask>=3.0.0
flask-session>=0.6.0
```

Установка: `pip install -r requirements.txt`

## Запуск

```bash
python app.py
# или через run.bat на Windows (устанавливает зависимости, открывает браузер)
```

Сервер стартует на `http://127.0.0.1:5000`.

## Active Directory — важные нюансы

- **NTLM + OpenSSL 3.0**: MD4 отключён по умолчанию → нужен LDAPS (`AD_SERVER=ldaps://...`, `AD_USE_NTLM=false`)
- **strongerAuthRequired**: сервер требует шифрования → переключиться на LDAPS
- **Самоподписанный сертификат**: `ad_fetcher.py` отключает проверку (`ssl.CERT_NONE`)
- Диагностика подключения: `python ad_whoami.py`

## Сводная таблица

`parser.load_pivot()` читает Excel, группирует по `COL_GROUP`, считает:
- `Capacity, ч` — плановые часы
- `Потрачено, ч` — фактические
- `% списания` = Потрачено / Capacity × 100

Строки FineBI могут содержать неразрывные пробелы (`\xa0`) и запятые как десятичный разделитель — обрабатываются в `parser.py`.

Цветовая индикация % в HTML-письме: <80% — красный, 80–95% — жёлтый, ≥95% — зелёный.

## Сессии Flask

Сессии серверные (`flask-session`, тип `filesystem`), хранятся во временной папке ОС. Учётные данные AD и SMTP хранятся в сессии на время работы — в cookie не попадают.

## История веток

- `claude/automate-hours-calculation-CY4qC` — базовая автоматизация (Outlook COM, Tkinter GUI)
- `claude/redesign-flask-ui-OVTrd` — замена Tkinter на Flask-веб-интерфейс
- `main` — итоговая ветка; Tkinter GUI, CLI и тестовые утилиты удалены
