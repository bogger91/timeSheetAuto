"""
Получение email-адресов тим-лидов из Active Directory через LDAP.

Поддерживаемые режимы аутентификации:
  - NTLM  (AD_USE_NTLM=true)  — шифрование встроено, пароль в config.env
  - LDAPS (AD_SERVER=ldaps://…) — TLS-соединение на порту 636
  - Simple bind — только если сервер явно разрешает (редко в корпоративной среде)

Поддерживаемые режимы поиска:
  - AD_SEARCH_BY=title  — по полю должности (title)
  - AD_SEARCH_BY=group  — по членству в группе AD

Функции принимают необязательные параметры server/user/password,
которые перекрывают значения из config.env. Это позволяет Flask-приложению
передавать учётные данные из сессии без изменения конфигурационного файла.

При AD_STUB=true в config.env все функции делегируются заглушке ad_fetcher_stub
без обращения к реальному LDAP-серверу.
"""
import config

if config.AD_STUB:
    # Заглушка для локального тестирования — ldap3 не нужен
    from ad_fetcher_stub import get_teamleads, get_teamlead_emails, test_connection

else:
    import ssl
    from ldap3 import Server, Connection, ALL, SUBTREE, NTLM, Tls

    def _connect(server: str = None, user: str = None,
                 password: str = None, use_ntlm: bool = None) -> Connection:
        server_url = server or config.AD_SERVER
        ad_user    = user     or config.AD_USER
        ad_pass    = password or config.AD_PASSWORD
        ntlm_flag  = use_ntlm if use_ntlm is not None else config.AD_USE_NTLM

        use_ldaps = server_url.lower().startswith("ldaps://")

        if use_ldaps:
            # validate=CERT_NONE достаточно: ldap3 сам выставляет check_hostname=False
            # когда validate==CERT_NONE, без зависимости от версии ldap3.
            tls = Tls(validate=ssl.CERT_NONE)
            srv = Server(server_url, use_ssl=True, tls=tls,
                         get_info=ALL, connect_timeout=10)
        else:
            srv = Server(server_url, get_info=ALL, connect_timeout=10)

        if ntlm_flag:
            conn = Connection(
                srv,
                user=ad_user,
                password=ad_pass,
                authentication=NTLM,
                auto_bind=True,
            )
        else:
            conn = Connection(
                srv,
                user=ad_user,
                password=ad_pass,
                auto_bind=True,
            )

        return conn

    def get_teamleads(server: str = None, user: str = None,
                      password: str = None) -> list[dict]:
        """
        Возвращает список тим-лидов из AD в виде:
          [{"name": "Иванов Иван", "email": "ivanov@company.ru", "department": "Отдел ..."}]
        Параметры server/user/password перекрывают значения из config.env.
        Поднимает исключение при ошибке подключения или поиска.
        """
        conn = _connect(server=server, user=user, password=password)
        search_filter = _build_filter()

        conn.search(
            search_base=config.AD_BASE_DN,
            search_filter=search_filter,
            search_scope=SUBTREE,
            attributes=["displayName", "mail", "department", "employeeType", "title"],
        )

        leads = []
        skipped = []

        for entry in conn.entries:
            email = str(entry.mail).strip()
            name = str(entry.displayName).strip()

            if not email or email.lower() in ("", "none"):
                skipped.append(name or str(entry.entry_dn))
                continue

            # employeeType — подразделение/отдел для сопоставления с Excel
            emp_type = str(entry.employeeType).strip()
            if emp_type.lower() == "none":
                emp_type = ""

            leads.append({"name": name, "email": email, "department": emp_type})

        if skipped:
            pass  # тихо игнорируем записи без email

        conn.unbind()
        return leads

    def get_teamlead_emails(server: str = None, user: str = None,
                            password: str = None) -> list[str]:
        """
        Возвращает список email-адресов тим-лидов из AD.
        Обёртка над get_teamleads() для обратной совместимости.
        """
        return [tl["email"] for tl in get_teamleads(server=server, user=user, password=password)]

    def _build_filter() -> str:
        parts = [
            "(objectClass=user)",
            "(!(userAccountControl:1.2.840.113556.1.4.803:=2))",  # только активные
        ]

        if config.AD_DEPARTMENT:
            parts.append(f"(department={config.AD_DEPARTMENT})")

        if config.AD_SEARCH_BY == "group":
            group_dn = config.AD_GROUP_DN
            if not group_dn:
                raise ValueError("AD_GROUP_DN не задан в config.env (нужен при AD_SEARCH_BY=group)")
            parts.append(f"(memberOf={group_dn})")
        else:
            title_mask = config.AD_TITLE_MASK or "*"
            parts.append(f"(title={title_mask})")

        return "(&" + "".join(parts) + ")"

    def test_connection(server: str = None, user: str = None,
                       password: str = None) -> str:
        """Проверка подключения к AD. Возвращает строку с результатом."""
        try:
            conn = _connect(server=server, user=user, password=password)
            who = conn.extend.standard.who_am_i()
            conn.unbind()
            return f"Подключение успешно. Аккаунт: {who}"
        except Exception as e:
            hint = ""
            msg = str(e).lower()
            if "strongeraruthrequired" in msg or "strongerauthrequired" in msg:
                hint = (
                    "\n\nСервер требует шифрования. Попробуйте:\n"
                    "  AD_SERVER=ldaps://ваш-dc.company.ru\n"
                    "  AD_USE_NTLM=false"
                )
            elif "timed out" in msg or "handshake" in msg:
                hint = (
                    "\n\nSSL handshake завис. Возможные причины:\n"
                    "  1. Порт 636 (LDAPS) закрыт файрволом — проверьте доступность сервера.\n"
                    "  2. AD_SERVER указывает на ldaps://, но сервер принимает только ldap:// + NTLM.\n"
                    "     Попробуйте: AD_SERVER=ldap://ваш-dc.company.ru  AD_USE_NTLM=true"
                )
            elif "md4" in msg or "unsupported hash" in msg:
                hint = (
                    "\n\nNTLM не работает: OpenSSL 3.0 отключил MD4.\n"
                    "Переключитесь на LDAPS в config.env:\n"
                    "  AD_SERVER=ldaps://ваш-dc.company.ru\n"
                    "  AD_USE_NTLM=false"
                )
            return f"Ошибка подключения: {e}{hint}"
