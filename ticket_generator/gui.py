from __future__ import annotations

import logging
import tkinter as tk
from pathlib import Path
from tkinter import messagebox, ttk

from .core import DataError, ResultJournal, read_students, read_tickets


class TicketApp:
    def __init__(self, root: tk.Tk, data_dir: Path):
        self.root = root
        self.data_dir = data_dir
        self.groups = {}
        self.tickets = {}
        self.journal = ResultJournal(data_dir / "results.xlsx")
        root.title("Генератор экзаменационных билетов")
        root.geometry("680x500")
        root.minsize(580, 460)
        style = ttk.Style(root)
        if "clam" in style.theme_names():
            style.theme_use("clam")
        style.configure("TLabel", font=("Segoe UI", 11))
        style.configure("TButton", font=("Segoe UI", 11), padding=10)
        body = ttk.Frame(root, padding=30)
        body.pack(fill="both", expand=True)
        ttk.Label(body, text="Экзаменационные билеты", font=("Segoe UI", 22, "bold")).pack(anchor="w")
        ttk.Label(body, text="Выберите группу и студента, чтобы получить билет.").pack(anchor="w", pady=(8, 24))
        ttk.Label(body, text="Группа").pack(anchor="w")
        self.group = ttk.Combobox(body, state="readonly", font=("Segoe UI", 12))
        self.group.pack(fill="x", pady=(6, 16))
        self.group.bind("<<ComboboxSelected>>", self.select_group)
        ttk.Label(body, text="Студент").pack(anchor="w")
        self.student = ttk.Combobox(body, state="disabled", font=("Segoe UI", 12))
        self.student.pack(fill="x", pady=(6, 20))
        self.student.bind("<<ComboboxSelected>>", self.select_student)
        self.generate_button = ttk.Button(body, text="Сгенерировать билет", command=self.generate, state="disabled")
        self.generate_button.pack(fill="x")
        self.status = tk.StringVar(value="Загрузка данных…")
        ttk.Label(body, textvariable=self.status, wraplength=590).pack(anchor="w", pady=(16, 8))
        ttk.Button(body, text="Перезагрузить файлы", command=self.load_sources).pack(anchor="w")
        root.after(100, self.load_sources)

    def load_sources(self):
        self.groups, self.tickets = {}, {}
        self.group.set("")
        self.student.set("")
        self.group.configure(values=())
        self.student.configure(values=(), state="disabled")
        self.generate_button.state(["disabled"])
        try:
            while True:
                try:
                    self.journal.initialize()
                    break
                except PermissionError:
                    if not messagebox.askretrycancel("Журнал недоступен", "Закройте results.xlsx в Excel и нажмите «Повторить».", parent=self.root):
                        self.status.set("Журнал недоступен. После закрытия файла перезагрузите данные.")
                        return
            self.groups = read_students(self.data_dir / "students.xlsx")
            self.tickets = read_tickets(self.data_dir / "tickets.docx")
            self.group.configure(values=tuple(self.groups))
            self.status.set(f"Загружено групп: {len(self.groups)}. Билетов: {len(self.tickets)}.")
        except (DataError, OSError) as exc:
            logging.exception("Ошибка загрузки данных")
            self.status.set("Проверьте файлы и нажмите «Перезагрузить файлы».")
            messagebox.showerror("Не удалось загрузить данные", str(exc), parent=self.root)

    def select_group(self, _event=None):
        students = self.groups.get(self.group.get(), [])
        self.student.set("")
        self.student.configure(values=[person.display_name for person in students],
                               state="readonly" if students else "disabled")
        self.generate_button.state(["disabled"])
        self.status.set("Выберите студента." if students else "В этой группе нет студентов.")

    def select_student(self, _event=None):
        valid = self.group.get() in self.groups and self.student.current() >= 0 and bool(self.tickets)
        self.generate_button.state(["!disabled"] if valid else ["disabled"])

    def generate(self):
        index = self.student.current()
        students = self.groups.get(self.group.get(), [])
        if not 0 <= index < len(students) or not self.tickets:
            return
        student = students[index]
        while True:
            try:
                ticket, repeated = self.journal.assign(student, self.tickets)
                break
            except PermissionError:
                if not messagebox.askretrycancel("Не удалось сохранить билет", "Закройте results.xlsx в Excel и нажмите «Повторить». Билет будет показан после сохранения.", parent=self.root):
                    return
            except (DataError, OSError) as exc:
                logging.exception("Ошибка назначения билета")
                messagebox.showerror("Не удалось выдать билет", str(exc), parent=self.root)
                return
        self.show_result(student, ticket, repeated)

    def show_result(self, student, ticket, repeated):
        window = tk.Toplevel(self.root)
        window.title(f"Ваш билет № {ticket.number}")
        window.geometry("760x560")
        window.minsize(500, 360)
        window.transient(self.root)
        window.grab_set()
        body = ttk.Frame(window, padding=28)
        body.pack(fill="both", expand=True)
        ttk.Label(body, text=f"Ваш билет № {ticket.number}", font=("Segoe UI", 23, "bold")).pack(anchor="w")
        ttk.Label(body, text=f"{student.group} · {student.display_name}").pack(anchor="w", pady=(8, 18))
        text_frame = ttk.Frame(body)
        text_frame.pack(fill="both", expand=True)
        text = tk.Text(text_frame, wrap="word", font=("Segoe UI", 13), relief="flat", padx=12, pady=12)
        scrollbar = ttk.Scrollbar(text_frame, command=text.yview)
        text.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        text.pack(side="left", fill="both", expand=True)
        text.insert("1.0", "\n\n".join(f"{index}. {question}" for index, question in enumerate(ticket.questions, 1)))
        text.configure(state="disabled")
        ttk.Label(body, text="Результат сохранён в results.xlsx" + (" · Повторная выдача" if repeated else "")).pack(anchor="w", pady=(16, 12))

        def close(_event=None):
            window.grab_release()
            window.destroy()
            self.student.set("")
            self.generate_button.state(["disabled"])
            self.status.set("Выберите следующего студента.")
            self.student.focus_set()
            return "break"

        ttk.Button(body, text="К выбору студента · Esc", command=close).pack(fill="x")
        window.bind("<Escape>", close)
        window.protocol("WM_DELETE_WINDOW", close)
        window.focus_set()
