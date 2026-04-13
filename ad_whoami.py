"""
Диагностика: выводит все атрибуты вашего пользователя из AD.
Помогает проверить подключение и узнать точные названия полей.

Запуск: python ad_whoami.py
"""
import config
import ssl
from ldap3 import Server, Connection, ALL, SUBTREE, NTLM, Tls


def connect():
    use_ldaps = config.AD_SERVER.lower().startswith("ldaps://")
    if use_ldaps:
        tls = Tls(validate=ssl.CERT_NONE, version=ssl.PROTOCOL_TLS_CLIENT)
        tls.check_hostname = False
        server = Server(config.AD_SERVER, use_ssl=True, tls=tls,
                        get_info=ALL, connect_timeout=10)
    else:
        server = Server(config.AD_SERVER, get_info=ALL, connect_timeout=10)

    if config.AD_USE_NTLM:
        conn = Connection(server, user=config.AD_USER, password=config.AD_PASSWORD,
                          authentication=NTLM, auto_bind=True)
    else:
        conn = Connection(server, user=config.AD_USER, password=config.AD_PASSWORD,
                          auto_bind=True)
    return conn


def extract_samaccountname(ad_user: str) -> str:
    """Извлекает sAMAccountName из DOMAIN\\user или user@domain."""
    if "\\" in ad_user:
        return ad_user.split("\\", 1)[1]
    if "@" in ad_user:
        return ad_user.split("@")[0]
    return ad_user


def main():
    print(f"Подключение к {config.AD_SERVER} ...")
    try:
        conn = connect()
    except Exception as e:
        print(f"[Ошибка] Не удалось подключиться: {e}")
        return

    username = extract_samaccountname(config.AD_USER)
    print(f"Поиск пользователя: {username}\n")

    conn.search(
        search_base=config.AD_BASE_DN,
        search_filter=f"(&(objectClass=user)(sAMAccountName={username}))",
        search_scope=SUBTREE,
        attributes=["*"],   # все атрибуты
    )

    if not conn.entries:
        print("[!] Пользователь не найден. Проверьте AD_BASE_DN и AD_USER в config.env.")
        conn.unbind()
        return

    entry = conn.entries[0]
    print("=" * 55)
    print(f"  DN: {entry.entry_dn}")
    print("=" * 55)

    # Интересные поля — показываем первыми
    priority = [
        "displayName", "sAMAccountName", "mail", "title",
        "department", "company", "manager", "telephoneNumber",
        "memberOf",
    ]

    printed = set()
    for attr in priority:
        try:
            val = entry[attr].value
            if val:
                label = f"{attr}:"
                if isinstance(val, list):
                    print(f"  {label}")
                    for v in val:
                        print(f"      {v}")
                else:
                    print(f"  {label:<25} {val}")
                printed.add(attr.lower())
        except Exception:
            pass

    # Остальные непустые атрибуты
    print("\n  --- Остальные атрибуты ---")
    for attr in sorted(entry.entry_attributes):
        if attr.lower() in printed:
            continue
        try:
            val = entry[attr].value
            if val is None or val == [] or val == "":
                continue
            label = f"{attr}:"
            if isinstance(val, list):
                print(f"  {label}")
                for v in val:
                    print(f"      {v}")
            else:
                print(f"  {label:<25} {val}")
        except Exception:
            pass

    conn.unbind()
    print("=" * 55)


if __name__ == "__main__":
    main()
