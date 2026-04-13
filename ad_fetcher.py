"""
Получение email-адресов тим-лидов из Active Directory через LDAP.

Поддерживаемые режимы аутентификации:
  - NTLM (Windows integrated auth) — пароль не нужен, работает под доменным аккаунтом
  - Simple bind — логин + пароль из config.env

Поддерживаемые режимы поиска:
  - AD_SEARCH_BY=title  — по полю должности (title)
  - AD_SEARCH_BY=group  — по членству в группе AD
"""
import config
from ldap3 import Server, Connection, ALL, SUBTREE, NTLM


def _connect() -> Connection:
    server = Server(config.AD_SERVER, get_info=ALL, connect_timeout=10)

    if config.AD_USE_NTLM:
        conn = Connection(
            server,
            user=config.AD_USER,
            password=config.AD_PASSWORD,
            authentication=NTLM,
            auto_bind=True,
        )
    else:
        conn = Connection(
            server,
            user=config.AD_USER,
            password=config.AD_PASSWORD,
            auto_bind=True,
        )

    return conn


def get_teamlead_emails() -> list[str]:
    """
    Возвращает список email-адресов тим-лидов из AD.
    Поднимает исключение при ошибке подключения или поиска.
    """
    conn = _connect()

    search_filter = _build_filter()

    conn.search(
        search_base=config.AD_BASE_DN,
        search_filter=search_filter,
        search_scope=SUBTREE,
        attributes=["displayName", "mail", "department", "title"],
    )

    emails = []
    skipped = []

    for entry in conn.entries:
        email = str(entry.mail).strip()
        name = str(entry.displayName).strip()

        if not email or email.lower() in ("", "none"):
            skipped.append(name or str(entry.entry_dn))
            continue

        emails.append(email)
        print(f"  + {name} <{email}>")

    if skipped:
        print(f"  [!] Пропущены (нет email в AD): {', '.join(skipped)}")

    conn.unbind()
    return emails


def _build_filter() -> str:
    base = "(objectClass=user)(!(userAccountControl:1.2.840.113556.1.4.803:=2))"  # активные юзеры

    if config.AD_SEARCH_BY == "group":
        group_dn = config.AD_GROUP_DN
        if not group_dn:
            raise ValueError("AD_GROUP_DN не задан в config.env (нужен при AD_SEARCH_BY=group)")
        return f"(&{base}(memberOf={group_dn}))"

    # По умолчанию — поиск по должности (title)
    title_mask = config.AD_TITLE_MASK or "*"
    return f"(&{base}(title={title_mask}))"


def test_connection() -> str:
    """Проверка подключения к AD. Возвращает строку с результатом."""
    try:
        conn = _connect()
        who = conn.extend.standard.who_am_i()
        conn.unbind()
        return f"Подключение успешно. Аккаунт: {who}"
    except Exception as e:
        return f"Ошибка подключения: {e}"
