"""
Точка входа.

Использование:
    python main.py              — создаёт черновик в Outlook
    python main.py --send       — сразу отправляет письмо
    python main.py --msg        — сохраняет draft.msg (+ черновик)
    python main.py --file other.xlsx  — использует конкретный файл

    python main.py --preview    — только выводит таблицу в консоль (без Outlook)
"""
import argparse
import sys

import parser as rpt
import mailer
import config


def main():
    ap = argparse.ArgumentParser(description="Отчёт по списанию часов")
    ap.add_argument("--send", action="store_true", help="Отправить письмо сразу")
    ap.add_argument("--msg", action="store_true", help="Сохранить draft.msg")
    ap.add_argument("--file", default=None, help="Путь к Excel-файлу (переопределяет config.env)")
    ap.add_argument("--preview", action="store_true", help="Показать таблицу в консоли, не трогать Outlook")
    args = ap.parse_args()

    # 1. Читаем данные и строим сводную таблицу
    try:
        pivot = rpt.load_pivot(args.file)
    except FileNotFoundError as e:
        print(f"[Ошибка] Файл не найден: {e}")
        sys.exit(1)
    except ValueError as e:
        print(f"[Ошибка] {e}")
        sys.exit(1)

    print("\n=== Сводная таблица ===")
    print(pivot.to_string(index=False))
    print()

    if args.preview:
        return

    # 2. Формируем HTML-таблицу
    table_html = rpt.pivot_to_html(pivot)

    # 3. Действие с Outlook
    try:
        if args.send:
            mailer.send(table_html)
        else:
            mailer.create_draft(table_html, save_as_msg=args.msg)
    except ImportError as e:
        print(f"[Ошибка] {e}")
        sys.exit(1)
    except Exception as e:
        print(f"[Ошибка Outlook] {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
