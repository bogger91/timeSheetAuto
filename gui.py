"""
Графический интерфейс для запуска отчёта по списанию часов.
Запуск: python gui.py  (или двойной клик по run.bat)
"""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
import sys
import io
import os

# Перехватываем print для вывода в окно
class TextRedirect(io.StringIO):
    def __init__(self, widget):
        super().__init__()
        self.widget = widget

    def write(self, text):
        self.widget.configure(state="normal")
        self.widget.insert(tk.END, text)
        self.widget.see(tk.END)
        self.widget.configure(state="disabled")

    def flush(self):
        pass


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Отчёт по списанию часов")
        self.resizable(False, False)
        self._build_ui()
        self._load_config()

    def _build_ui(self):
        pad = {"padx": 10, "pady": 5}

        # --- Файл ---
        file_frame = ttk.LabelFrame(self, text="Excel-файл из FineBI")
        file_frame.grid(row=0, column=0, columnspan=2, sticky="ew", **pad)

        self.file_var = tk.StringVar()
        ttk.Entry(file_frame, textvariable=self.file_var, width=55).grid(row=0, column=0, padx=5, pady=5)
        ttk.Button(file_frame, text="Обзор…", command=self._browse).grid(row=0, column=1, padx=5)

        # --- Получатели ---
        mail_frame = ttk.LabelFrame(self, text="Получатели (через запятую)")
        mail_frame.grid(row=1, column=0, columnspan=2, sticky="ew", **pad)

        self.mail_var = tk.StringVar()
        ttk.Entry(mail_frame, textvariable=self.mail_var, width=50).grid(row=0, column=0, padx=5, pady=5)
        ttk.Button(mail_frame, text="Загрузить из AD",
                   command=self._load_from_ad, width=18).grid(row=0, column=1, padx=5)
        ttk.Button(mail_frame, text="Проверить AD",
                   command=self._test_ad, width=14).grid(row=0, column=2, padx=5)

        # --- Кнопки ---
        btn_frame = ttk.Frame(self)
        btn_frame.grid(row=2, column=0, columnspan=2, **pad)

        ttk.Button(btn_frame, text="Предпросмотр таблицы",
                   command=self._run_preview, width=25).grid(row=0, column=0, padx=5)
        ttk.Button(btn_frame, text="Создать черновик в Outlook",
                   command=self._run_draft, width=25).grid(row=0, column=1, padx=5)
        ttk.Button(btn_frame, text="Отправить письмо",
                   command=self._run_send, width=20).grid(row=0, column=2, padx=5)

        # --- Лог ---
        log_frame = ttk.LabelFrame(self, text="Вывод")
        log_frame.grid(row=3, column=0, columnspan=2, sticky="nsew", **pad)

        self.log = scrolledtext.ScrolledText(log_frame, width=75, height=15,
                                             state="disabled", font=("Consolas", 9))
        self.log.pack(padx=5, pady=5)

        ttk.Button(self, text="Очистить лог",
                   command=self._clear_log).grid(row=4, column=1, sticky="e", **pad)

    def _load_config(self):
        """Подставить значения из config.env в поля."""
        try:
            from dotenv import dotenv_values
            cfg = dotenv_values("config.env")
            if cfg.get("EXCEL_PATH"):
                self.file_var.set(cfg["EXCEL_PATH"])
            if cfg.get("MAIL_TO"):
                self.mail_var.set(cfg["MAIL_TO"])
        except Exception:
            pass

    def _browse(self):
        path = filedialog.askopenfilename(
            title="Выбрать Excel-файл",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if path:
            self.file_var.set(path)

    def _clear_log(self):
        self.log.configure(state="normal")
        self.log.delete("1.0", tk.END)
        self.log.configure(state="disabled")

    def _save_config(self):
        """Сохранить текущий путь и почту обратно в config.env."""
        try:
            lines = []
            if os.path.exists("config.env"):
                with open("config.env", encoding="utf-8") as f:
                    lines = f.readlines()

            updated = {}
            new_lines = []
            for line in lines:
                if line.startswith("EXCEL_PATH="):
                    new_lines.append(f"EXCEL_PATH={self.file_var.get()}\n")
                    updated["EXCEL_PATH"] = True
                elif line.startswith("MAIL_TO="):
                    new_lines.append(f"MAIL_TO={self.mail_var.get()}\n")
                    updated["MAIL_TO"] = True
                else:
                    new_lines.append(line)

            if "EXCEL_PATH" not in updated:
                new_lines.append(f"EXCEL_PATH={self.file_var.get()}\n")
            if "MAIL_TO" not in updated:
                new_lines.append(f"MAIL_TO={self.mail_var.get()}\n")

            with open("config.env", "w", encoding="utf-8") as f:
                f.writelines(new_lines)
        except Exception:
            pass

    def _run(self, func):
        """Запускает func в отдельном потоке, перехватывая stdout."""
        redirector = TextRedirect(self.log)
        sys.stdout = redirector
        sys.stderr = redirector

        def task():
            try:
                func()
            except SystemExit:
                pass
            except Exception as e:
                print(f"[Ошибка] {e}")
            finally:
                sys.stdout = sys.__stdout__
                sys.stderr = sys.__stderr__

        threading.Thread(target=task, daemon=True).start()

    def _test_ad(self):
        def do():
            import importlib, config, ad_fetcher
            importlib.reload(config)
            importlib.reload(ad_fetcher)
            print("\n=== Проверка подключения к AD ===")
            result = ad_fetcher.test_connection()
            print(result)
            print()

        self._run(do)

    def _load_from_ad(self):
        def do():
            import importlib, config, ad_fetcher
            importlib.reload(config)
            importlib.reload(ad_fetcher)
            print("\n=== Загрузка тим-лидов из AD ===")
            emails = ad_fetcher.get_teamlead_emails()
            if not emails:
                print("[!] Никого не найдено. Проверьте AD_SEARCH_BY и маску в config.env.")
                return
            joined = ",".join(emails)
            self.mail_var.set(joined)
            self._save_config()
            print(f"\nНайдено: {len(emails)} адрес(ов). Поле получателей обновлено.")
            print()

        self._run(do)

    def _validate_file(self) -> bool:
        path = self.file_var.get().strip()
        if not path:
            messagebox.showwarning("Файл не выбран", "Выберите Excel-файл через кнопку «Обзор…»")
            return False
        if not os.path.exists(path):
            messagebox.showerror("Файл не найден", f"Файл не найден:\n{path}")
            return False
        return True

    def _run_preview(self):
        if not self._validate_file():
            return
        self._save_config()

        def do():
            import importlib, config, parser as rpt
            importlib.reload(config)
            importlib.reload(rpt)
            pivot = rpt.load_pivot(self.file_var.get())
            print("\n=== Сводная таблица ===")
            print(pivot.to_string(index=False))
            print()

        self._run(do)

    def _run_draft(self):
        if not self._validate_file():
            return
        self._save_config()

        def do():
            import importlib, config, parser as rpt, mailer
            importlib.reload(config)
            importlib.reload(rpt)
            importlib.reload(mailer)
            pivot = rpt.load_pivot(self.file_var.get())
            table_html = rpt.pivot_to_html(pivot)
            mailer.create_draft(table_html)

        self._run(do)

    def _run_send(self):
        if not self._validate_file():
            return
        if not messagebox.askyesno("Подтверждение", "Отправить письмо всем получателям?"):
            return
        self._save_config()

        def do():
            import importlib, config, parser as rpt, mailer
            importlib.reload(config)
            importlib.reload(rpt)
            importlib.reload(mailer)
            pivot = rpt.load_pivot(self.file_var.get())
            table_html = rpt.pivot_to_html(pivot)
            mailer.send(table_html)

        self._run(do)


if __name__ == "__main__":
    app = App()
    app.mainloop()
