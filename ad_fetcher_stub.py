"""
Заглушка AD для локального тестирования.

Активируется через AD_STUB=true в config.env или переменной окружения.
Возвращает фиктивных тим-лидов без подключения к LDAP.

Чтобы добавить или изменить тестовых тим-лидов — редактируйте STUB_LEADS ниже.
"""

STUB_LEADS = [
    {
        "name": "Иванов Иван Иванович",
        "email": "i.ivanov@company.ru",
        "department": "Отдел разработки",
    },
    {
        "name": "Петрова Мария Сергеевна",
        "email": "m.petrova@company.ru",
        "department": "Отдел тестирования",
    },
    {
        "name": "Сидоров Алексей Николаевич",
        "email": "a.sidorov@company.ru",
        "department": "Отдел аналитики",
    },
    {
        "name": "Козлова Екатерина Дмитриевна",
        "email": "e.kozlova@company.ru",
        "department": "Отдел DevOps",
    },
    {
        "name": "Новиков Дмитрий Александрович",
        "email": "d.novikov@company.ru",
        "department": "Отдел безопасности",
    },
]


def get_teamleads(server=None, user=None, password=None) -> list[dict]:
    """Возвращает фиктивный список тим-лидов (без обращения к AD)."""
    return list(STUB_LEADS)


def get_teamlead_emails(server=None, user=None, password=None) -> list[str]:
    return [tl["email"] for tl in get_teamleads()]


def test_connection(server=None, user=None, password=None) -> str:
    return "AD_STUB=true: заглушка активна, реального подключения нет."
