"""
Получение email-адресов тим-лидов из Active Directory через LDAP.

Поддерживаемые режимы аутентификации:
  - NTLM  (AD_USE_NTLM=true)  — шифрование встроено, пароль в config.env
  - LDAPS (AD_SERVER=ldaps://…) — TLS-соединение на порту 636
  - Simple bind — только если сервер явно разрешает (редко в корпоративной среде)

Поддерживаемые режимы поиска:
  - AD_SEARCH_BY=title  — по полю должности (title)
  - AD_SEARCH_BY=group  — по членству в группе AD
"""
import ssl
import config
from ldap3 import Server, Connection, ALL, SUBTREE, NTLM, Tls


def _connect() -> Connection:
    use_ldaps = config.AD_SERVER.lower().startswith("ldaps://")

    if use_ldaps:
        # LDAPS: TLS с отключённой проверкой сертификата (корпоративный самоподписанный)
        tls = Tls(validate=ssl.CERT_NONE, version=ssl.PROTOCOL_TLS_CLIENT)
        tls.check_hostname = False
        server = Server(config.AD_SERVER, use_ssl=True, tls=tls,
                        get_info=ALL, connect_timeout=10)
    else:
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
        hint = ""
        msg = str(e).lower()
        if "strongeraruthrequired" in msg or "strongerauthrequired" in msg:
            hint = (
                "\n\nСервер требует шифрования. Попробуйте одно из:\n"
                "  1. AD_USE_NTLM=true в config.env\n"
                "  2. Сменить AD_SERVER=ldaps://dc01.company.ru (порт 636)"
            )
        return f"Ошибка подключения: {e}{hint}"
