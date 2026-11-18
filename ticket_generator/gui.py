from __future__ import annotations

import logging
import os
import subprocess
import sys
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from openpyxl import load_workbook

from .core import DataError, ResultJournal, Student, Ticket, read_students, read_tickets

logger = logging.getLogger(__name__)

THEMES = {
    "dark": {
        "name": "dark",
        "bg": "#0b0f19",
        "header_bg1": "#1e1b4b",
        "header_bg2": "#312e81",
        "card_bg": "#111827",
        "card_border": "#1f2937",
        "card_alt": "#1e293b",
        "fg": "#f9fafb",
        "fg_muted": "#9ca3af",
        "fg_dim": "#6b7280",
        "accent": "#6366f1",
        "accent_hover": "#4f46e5",
        "accent_active": "#4338ca",
        "accent_fg": "#ffffff",
        "green": "#10b981",
        "green_bg": "#064e3b",
        "green_fg": "#6ee7b7",
        "amber": "#f59e0b",
        "amber_bg": "#78350f",
        "amber_fg": "#fde68a",
        "purple": "#8b5cf6",
        "purple_bg": "#4c1d95",
        "purple_fg": "#ddd6fe",
        "cyan": "#06b6d4",
        "cyan_bg": "#164e63",
        "cyan_fg": "#a5f3fc",
        "rose": "#f43f5e",
        "input_bg": "#1f2937",
        "input_fg": "#f9fafb",
        "input_border": "#374151",
        "tree_bg": "#111827",
        "tree_fg": "#f9fafb",
        "tree_heading_bg": "#1f2937",
        "tree_heading_fg": "#e5e7eb",
        "tree_alt": "#162032",
        "tree_sel_bg": "#4f46e5",
        "tree_sel_fg": "#ffffff",
    },
    "light": {
        "name": "light",
        "bg": "#f8fafc",
        "header_bg1": "#3730a3",
        "header_bg2": "#4f46e5",
        "card_bg": "#ffffff",
        "card_border": "#e2e8f0",
        "card_alt": "#f1f5f9",
        "fg": "#0f172a",
        "fg_muted": "#475569",
        "fg_dim": "#94a3b8",
        "accent": "#4f46e5",
        "accent_hover": "#4338ca",
        "accent_active": "#3730a3",
        "accent_fg": "#ffffff",
        "green": "#059669",
        "green_bg": "#ecfdf5",
        "green_fg": "#065f46",
        "amber": "#d97706",
        "amber_bg": "#fffbeb",
        "amber_fg": "#92400e",
        "purple": "#7c3aed",
        "purple_bg": "#f5f3ff",
        "purple_fg": "#5b21b6",
        "cyan": "#0284c7",
        "cyan_bg": "#f0f9ff",
        "cyan_fg": "#075985",
        "rose": "#e11d48",
        "input_bg": "#ffffff",
        "input_fg": "#0f172a",
        "input_border": "#cbd5e1",
        "tree_bg": "#ffffff",
        "tree_fg": "#0f172a",
        "tree_heading_bg": "#f1f5f9",
        "tree_heading_fg": "#1e293b",
        "tree_alt": "#f8fafc",
        "tree_sel_bg": "#c7d2fe",
        "tree_sel_fg": "#1e1b4b",
    },
}


def open_system_path(path: Path) -> None:
    try:
        if hasattr(os, "startfile"):
            os.startfile(path)
        elif sys.platform == "darwin":
            subprocess.run(["open", str(path)], check=False)
        else:
            subprocess.run(["xdg-open", str(path)], check=False)
    except Exception as exc:
        logger.warning("Не удалось открыть путь %s: %s", path, exc)


def load_journal_history(path: Path) -> list[dict]:
    if not path.is_file():
        return []
    try:
        book = load_workbook(path, read_only=True, data_only=True)
        try:
            sheet = book.active
            records = []
            for idx, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), 1):
                if not any(row):
                    continue
                group = str(row[0]) if len(row) > 0 and row[0] is not None else ""
                surname = str(row[1]) if len(row) > 1 and row[1] is not None else ""
                name = str(row[2]) if len(row) > 2 and row[2] is not None else ""
                ticket_num = row[3]
                dt = row[4]
                if isinstance(dt, datetime):
                    dt_str = dt.strftime("%d.%m.%Y %H:%M:%S")
                else:
                    dt_str = str(dt) if dt is not None else ""
                rep = str(row[5]).strip().lower() if len(row) > 5 and row[5] is not None else ""
                records.append({
                    "idx": idx,
                    "group": group,
                    "surname": surname,
                    "name": name,
                    "student": f"{surname} {name}".strip(),
                    "ticket": ticket_num,
                    "datetime": dt_str,
                    "repeated": rep == "да" or rep == "yes" or rep == "true",
                })
            return records
        finally:
            book.close()
    except Exception as exc:
        logger.warning("Не удалось прочитать историю журнала: %s", exc)
        return []


class TicketApp:
    def __init__(self, root: tk.Tk, data_dir: Path):
        self.root = root
        self.data_dir = data_dir
        self.groups: dict[str, list[Student]] = {}
        self.tickets: dict[int, Ticket] = {}
        self.journal = ResultJournal(data_dir / "results.xlsx")
        self.journal_records: list[dict] = []
        self.current_theme = "dark"

        root.title("Генератор экзаменационных билетов · Pro")
        root.geometry("1060x720")
        root.minsize(880, 600)

        self.style = ttk.Style(root)
        if "clam" in self.style.theme_names():
            self.style.theme_use("clam")

        self.status = tk.StringVar(value="Инициализация системы…")
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", self._on_search_change)

        self._build_ui()
        self.apply_theme(self.current_theme)

        root.after(100, self.load_sources)

    def _build_ui(self):
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(2, weight=1)

        # 1. Header Banner
        self.header_canvas = tk.Canvas(self.root, height=76, highlightthickness=0)
        self.header_canvas.grid(row=0, column=0, sticky="ew")
        self.header_canvas.bind("<Configure>", lambda e: self._draw_header())

        # 2. KPI / Metrics Bar
        self.kpi_frame = tk.Frame(self.root, padx=20, pady=12)
        self.kpi_frame.grid(row=1, column=0, sticky="ew")
        for i in range(4):
            self.kpi_frame.columnconfigure(i, weight=1, uniform="kpi")

        self.kpi_cards = []
        kpi_defs = [
            ("📁 ГРУПП В БАЗЕ", "0", "🏛️"),
            ("👥 ВСЕГО СТУДЕНТОВ", "0", "🎓"),
            ("🎟️ БИЛЕТОВ В ПУЛЕ", "0", "📄"),
            ("📝 ВЫДАНО БИЛЕТОВ", "0", "✅"),
        ]
        for idx, (title, default_val, icon) in enumerate(kpi_defs):
            card = tk.Frame(self.kpi_frame, padx=14, pady=10, relief="flat", highlightthickness=1)
            card.grid(row=0, column=idx, padx=6, sticky="nsew")
            card.columnconfigure(1, weight=1)

            icon_lbl = tk.Label(card, text=icon, font=("Segoe UI Emoji", 20))
            icon_lbl.grid(row=0, column=0, rowspan=2, padx=(0, 10), sticky="w")

            t_lbl = tk.Label(card, text=title, font=("Segoe UI", 8, "bold"))
            t_lbl.grid(row=0, column=1, sticky="w")

            v_lbl = tk.Label(card, text=default_val, font=("Segoe UI", 16, "bold"))
            v_lbl.grid(row=1, column=1, sticky="w")

            self.kpi_cards.append({
                "frame": card,
                "icon": icon_lbl,
                "title": t_lbl,
                "val": v_lbl,
            })

        # 3. Main Split Container (Left = Controls, Right = History Feed)
        self.main_container = tk.Frame(self.root, padx=20, pady=6)
        self.main_container.grid(row=2, column=0, sticky="nsew")
        self.main_container.columnconfigure(0, weight=5, uniform="col")
        self.main_container.columnconfigure(1, weight=6, uniform="col")
        self.main_container.rowconfigure(0, weight=1)

        # LEFT COLUMN (Controls & Profile)
        self.left_col = tk.Frame(self.main_container)
        self.left_col.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        self.left_col.rowconfigure(1, weight=1)
        self.left_col.columnconfigure(0, weight=1)

        # Card: Selection
        self.select_card = tk.Frame(self.left_col, padx=18, pady=16, highlightthickness=1)
        self.select_card.pack(fill="x", pady=(0, 12))
        self.select_card.columnconfigure(0, weight=1)

        self.step1_lbl = tk.Label(self.select_card, text="1. Выберите академическую группу", font=("Segoe UI", 10, "bold"))
        self.step1_lbl.pack(anchor="w")

        self.group = ttk.Combobox(self.select_card, state="readonly", font=("Segoe UI", 11))
        self.group.pack(fill="x", pady=(6, 4))
        self.group.bind("<<ComboboxSelected>>", self.select_group)

        self.group_info_lbl = tk.Label(self.select_card, text="Выберите группу из списка", font=("Segoe UI", 9))
        self.group_info_lbl.pack(anchor="w", pady=(0, 4))

        self.group_progress = ttk.Progressbar(self.select_card, orient="horizontal", mode="determinate")
        self.group_progress.pack(fill="x", pady=(0, 10))

        self.step2_lbl = tk.Label(self.select_card, text="2. Выберите студента", font=("Segoe UI", 10, "bold"))
        self.step2_lbl.pack(anchor="w")

        # Live search bar for students
        search_row = tk.Frame(self.select_card)
        search_row.pack(fill="x", pady=(4, 6))
        self.search_icon_lbl = tk.Label(search_row, text="🔍", font=("Segoe UI Emoji", 10))
        self.search_icon_lbl.pack(side="left", padx=(0, 4))
        self.search_entry = ttk.Entry(search_row, textvariable=self.search_var, font=("Segoe UI", 10))
        self.search_entry.pack(side="left", fill="x", expand=True)
        self.search_clear_btn = tk.Button(search_row, text="✕", font=("Segoe UI", 9),
                                          command=lambda: self.search_var.set(""),
                                          relief="flat", cursor="hand2", padx=6, bd=0)
        self.search_clear_btn.pack(side="right", padx=(2, 0))

        self.student = ttk.Combobox(self.select_card, state="disabled", font=("Segoe UI", 11))
        self.student.pack(fill="x", pady=(2, 6))
        self.student.bind("<<ComboboxSelected>>", self.select_student)

        # Card: Candidate Profile & Status
        self.profile_card = tk.Frame(self.left_col, padx=18, pady=14, highlightthickness=1)
        self.profile_card.pack(fill="both", expand=True, pady=(0, 12))

        self.profile_title = tk.Label(self.profile_card, text="Сведения о кандидате", font=("Segoe UI", 10, "bold"))
        self.profile_title.pack(anchor="w")

        self.profile_name_lbl = tk.Label(self.profile_card, text="Студент не выбран", font=("Segoe UI", 13, "bold"))
        self.profile_name_lbl.pack(anchor="w", pady=(6, 2))

        self.profile_badge_lbl = tk.Label(self.profile_card, text="Ожидание выбора студента", font=("Segoe UI", 9), padx=8, pady=4)
        self.profile_badge_lbl.pack(anchor="w", pady=(2, 6))

        self.profile_desc_lbl = tk.Label(self.profile_card, text="Выберите группу и студента для проверки статуса выдачи билета.",
                                         font=("Segoe UI", 9), wraplength=420, justify="left")
        self.profile_desc_lbl.pack(anchor="w", pady=(0, 6))

        # Card: Action Button & Status
        self.action_card = tk.Frame(self.left_col, padx=18, pady=16, highlightthickness=1)
        self.action_card.pack(fill="x")

        self.generate_button = ttk.Button(self.action_card, text="🎲  Сгенерировать экзаменационный билет",
                                          command=self.generate, state="disabled", style="Accent.TButton")
        self.generate_button.pack(fill="x", ipady=4)
        self.root.bind("<Return>", lambda e: self._on_enter())
        self.root.bind("<F5>", lambda e: self.load_sources())
        self.root.bind("<Control-r>", lambda e: self.load_sources())
        self.root.bind("<Control-R>", lambda e: self.load_sources())
        self.root.bind("<Control-t>", lambda e: self.toggle_theme())
        self.root.bind("<Control-T>", lambda e: self.toggle_theme())
        self.root.bind("<Control-o>", lambda e: self.open_excel_file())
        self.root.bind("<Control-O>", lambda e: self.open_excel_file())

        status_box = tk.Frame(self.action_card)
        status_box.pack(fill="x", pady=(10, 0))

        self.status_dot = tk.Label(status_box, text="●", font=("Segoe UI", 12))
        self.status_dot.pack(side="left", padx=(0, 6))

        self.status_lbl = tk.Label(status_box, textvariable=self.status, font=("Segoe UI", 9),
                                   wraplength=410, justify="left")
        self.status_lbl.pack(side="left", fill="x", expand=True)

        # RIGHT COLUMN (Live Activity Feed / Table)
        self.right_col = tk.Frame(self.main_container)
        self.right_col.grid(row=0, column=1, sticky="nsew", padx=(10, 0))
        self.right_col.rowconfigure(1, weight=1)
        self.right_col.columnconfigure(0, weight=1)

        self.history_card = tk.Frame(self.right_col, padx=16, pady=14, highlightthickness=1)
        self.history_card.pack(fill="both", expand=True)
        self.history_card.rowconfigure(2, weight=1)
        self.history_card.columnconfigure(0, weight=1)

        hist_header = tk.Frame(self.history_card)
        hist_header.grid(row=0, column=0, sticky="ew", pady=(0, 8))

        self.hist_title = tk.Label(hist_header, text="📋 Журнал выдачи (results.xlsx)", font=("Segoe UI", 11, "bold"))
        self.hist_title.pack(side="left")

        self.hist_badge = tk.Label(hist_header, text="0 записей", font=("Segoe UI", 9, "bold"), padx=8, pady=2)
        self.hist_badge.pack(side="right")

        # History live filter
        self.hist_filter_frame = tk.Frame(self.history_card)
        self.hist_filter_frame.grid(row=1, column=0, sticky="ew", pady=(0, 8))

        self.hist_filter_icon = tk.Label(self.hist_filter_frame, text="🔍", font=("Segoe UI Emoji", 9))
        self.hist_filter_icon.pack(side="left", padx=(0, 4))

        self.hist_search_var = tk.StringVar()
        self.hist_search_var.trace_add("write", lambda *_: self._refresh_history_table())
        self.hist_filter_entry = ttk.Entry(self.hist_filter_frame, textvariable=self.hist_search_var, font=("Segoe UI", 9))
        self.hist_filter_entry.pack(side="left", fill="x", expand=True)

        self.hist_clear_btn = tk.Button(self.hist_filter_frame, text="✕", font=("Segoe UI", 8),
                                        command=lambda: self.hist_search_var.set(""),
                                        relief="flat", cursor="hand2", padx=6, bd=0)
        self.hist_clear_btn.pack(side="right", padx=(2, 0))

        tree_frame = tk.Frame(self.history_card)
        tree_frame.grid(row=2, column=0, sticky="nsew")
        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)

        columns = ("id", "datetime", "group", "student", "ticket", "repeated")
        self.tree = ttk.Treeview(tree_frame, columns=columns, show="headings", selectmode="browse")
        self.tree.heading("id", text="№")
        self.tree.heading("datetime", text="Дата и время")
        self.tree.heading("group", text="Группа")
        self.tree.heading("student", text="Студент")
        self.tree.heading("ticket", text="Билет №")
        self.tree.heading("repeated", text="Повтор")

        self.tree.column("id", width=36, anchor="center")
        self.tree.column("datetime", width=135, anchor="center")
        self.tree.column("group", width=80, anchor="center")
        self.tree.column("student", width=140, anchor="w")
        self.tree.column("ticket", width=65, anchor="center")
        self.tree.column("repeated", width=65, anchor="center")

        tree_scroll = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=tree_scroll.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        tree_scroll.grid(row=0, column=1, sticky="ns")

        self.tree.bind("<Double-1>", self._on_tree_double_click)
        self.tree.bind("<<TreeviewSelect>>", self._on_tree_select)

        # Bottom Bar of History Card
        hist_footer = tk.Frame(self.history_card)
        hist_footer.grid(row=3, column=0, sticky="ew", pady=(10, 0))

        self.hist_summary = tk.Label(hist_footer, text="Двойной клик по строке открывает подробности", font=("Segoe UI", 8))
        self.hist_summary.pack(side="left")

        self.open_excel_btn = ttk.Button(hist_footer, text="📊 Открыть в Excel",
                                         command=self.open_excel_file, style="Secondary.TButton")
        self.open_excel_btn.pack(side="right")

        self.open_folder_btn = ttk.Button(hist_footer, text="📁 Папка данных",
                                          command=self.open_data_folder, style="Secondary.TButton")
        self.open_folder_btn.pack(side="right", padx=(0, 6))

    def _draw_header(self):
        theme = THEMES[self.current_theme]
        w = max(self.header_canvas.winfo_width(), 100)
        h = max(self.header_canvas.winfo_height(), 76)
        self.header_canvas.delete("all")

        # Draw smooth horizontal gradient
        r1, g1, b1 = self.header_canvas.winfo_rgb(theme["header_bg1"])
        r2, g2, b2 = self.header_canvas.winfo_rgb(theme["header_bg2"])
        steps = 60
        for i in range(steps):
            r = int(r1 + (r2 - r1) * (i / steps)) // 256
            g = int(g1 + (g2 - g1) * (i / steps)) // 256
            b = int(b1 + (b2 - b1) * (i / steps)) // 256
            col = f"#{r:02x}{g:02x}{b:02x}"
            x0 = int(w * (i / steps))
            x1 = int(w * ((i + 1) / steps)) + 1
            self.header_canvas.create_rectangle(x0, 0, x1, h, fill=col, outline=col)

        # Header Title and Subtitle
        self.header_canvas.create_text(26, 26, text="🎓  ЭКЗАМЕНАЦИОННЫЕ БИЛЕТЫ", anchor="w",
                                       font=("Segoe UI", 16, "bold"), fill="#ffffff")
        self.header_canvas.create_text(28, 50, text="Интеллектуальная система распределения и электронный журнал",
                                       anchor="w", font=("Segoe UI", 9), fill="#cbd5e1")

        # Header Control Buttons (Theme Toggle & Reload & Demo)
        theme_btn_text = "☀️ Светлая тема" if self.current_theme == "dark" else "🌙 Тёмная тема"
        if not hasattr(self, "_header_btns_created"):
            self._header_btns_created = True
            self.header_btns_frame = tk.Frame(self.header_canvas)
            self.demo_btn = ttk.Button(self.header_btns_frame, text="🧪 Демо-файлы",
                                       command=self.create_demo_files, style="Header.TButton")
            self.demo_btn.pack(side="left", padx=3)
            self.reload_btn = ttk.Button(self.header_btns_frame, text="🔄 Перезагрузить",
                                         command=self.load_sources, style="Header.TButton")
            self.reload_btn.pack(side="left", padx=3)
            self.theme_btn = ttk.Button(self.header_btns_frame, text=theme_btn_text,
                                        command=self.toggle_theme, style="Header.TButton")
            self.theme_btn.pack(side="left", padx=3)

        self.theme_btn.configure(text=theme_btn_text)
        self.header_btns_frame.configure(bg=theme["header_bg2"])
        self.header_canvas.create_window(w - 20, h // 2, window=self.header_btns_frame, anchor="e")

    def toggle_theme(self):
        self.current_theme = "light" if self.current_theme == "dark" else "dark"
        self.apply_theme(self.current_theme)

    def apply_theme(self, theme_key: str):
        theme = THEMES[theme_key]
        self.root.configure(bg=theme["bg"])

        # Configure TTK Styles
        self.style.configure(".", background=theme["bg"], foreground=theme["fg"])

        # Frames
        self.style.configure("TFrame", background=theme["bg"])

        # Buttons
        self.style.configure("TButton", font=("Segoe UI", 10), padding=6)
        self.style.configure("Accent.TButton", font=("Segoe UI", 11, "bold"),
                             background=theme["accent"], foreground=theme["accent_fg"],
                             padding=(12, 10), borderwidth=0)
        self.style.map("Accent.TButton",
                       background=[("disabled", theme["card_border"]),
                                   ("pressed", theme["accent_active"]),
                                   ("active", theme["accent_hover"])],
                       foreground=[("disabled", theme["fg_dim"]),
                                   ("pressed", theme["accent_fg"]),
                                   ("active", theme["accent_fg"])])

        self.style.configure("Secondary.TButton", font=("Segoe UI", 9),
                             background=theme["card_alt"], foreground=theme["fg"],
                             padding=(10, 5), borderwidth=1)
        self.style.map("Secondary.TButton",
                       background=[("active", theme["card_border"])],
                       foreground=[("active", theme["fg"])])

        self.style.configure("Header.TButton", font=("Segoe UI", 9, "bold"),
                             background=theme["accent"], foreground=theme["accent_fg"],
                             padding=(8, 4))
        self.style.map("Header.TButton",
                       background=[("active", theme["accent_hover"])])

        # Combobox
        self.style.configure("TCombobox",
                             fieldbackground=theme["input_bg"],
                             background=theme["card_alt"],
                             foreground=theme["input_fg"],
                             arrowcolor=theme["accent"],
                             font=("Segoe UI", 10))
        self.style.map("TCombobox",
                       fieldbackground=[("readonly", theme["input_bg"]), ("disabled", theme["card_alt"])],
                       foreground=[("readonly", theme["input_fg"]), ("disabled", theme["fg_dim"])])

        self.root.option_add("*TCombobox*Listbox.background", theme["input_bg"])
        self.root.option_add("*TCombobox*Listbox.foreground", theme["input_fg"])
        self.root.option_add("*TCombobox*Listbox.selectBackground", theme["accent"])
        self.root.option_add("*TCombobox*Listbox.selectForeground", theme["accent_fg"])
        self.root.option_add("*TCombobox*Listbox.font", ("Segoe UI", 10))

        # Entry
        self.style.configure("TEntry", fieldbackground=theme["input_bg"], foreground=theme["input_fg"])

        # Treeview
        self.style.configure("Treeview",
                             background=theme["tree_bg"],
                             foreground=theme["tree_fg"],
                             fieldbackground=theme["tree_bg"],
                             borderwidth=0,
                             font=("Segoe UI", 9),
                             rowheight=26)
        self.style.configure("Treeview.Heading",
                             background=theme["tree_heading_bg"],
                             foreground=theme["tree_heading_fg"],
                             font=("Segoe UI", 9, "bold"),
                             borderwidth=1)
        self.style.map("Treeview",
                       background=[("selected", theme["tree_sel_bg"])],
                       foreground=[("selected", theme["tree_sel_fg"])])

        # Configure tk widgets colors
        self.kpi_frame.configure(bg=theme["bg"])
        for card_data in self.kpi_cards:
            card_data["frame"].configure(bg=theme["card_bg"], highlightbackground=theme["card_border"])
            card_data["icon"].configure(bg=theme["card_bg"], fg=theme["accent"])
            card_data["title"].configure(bg=theme["card_bg"], fg=theme["fg_muted"])
            card_data["val"].configure(bg=theme["card_bg"], fg=theme["fg"])

        self.main_container.configure(bg=theme["bg"])
        self.left_col.configure(bg=theme["bg"])
        self.right_col.configure(bg=theme["bg"])

        # Left Cards
        for card in (self.select_card, self.profile_card, self.action_card, self.history_card):
            card.configure(bg=theme["card_bg"], highlightbackground=theme["card_border"])

        self.step1_lbl.configure(bg=theme["card_bg"], fg=theme["fg"])
        self.group_info_lbl.configure(bg=theme["card_bg"], fg=theme["fg_muted"])
        self.step2_lbl.configure(bg=theme["card_bg"], fg=theme["fg"])
        self.search_icon_lbl.configure(bg=theme["card_bg"], fg=theme["fg_muted"])

        self.profile_title.configure(bg=theme["card_bg"], fg=theme["fg_muted"])
        self.profile_name_lbl.configure(bg=theme["card_bg"], fg=theme["fg"])
        self.profile_badge_lbl.configure(bg=theme["card_alt"], fg=theme["fg_muted"])
        self.profile_desc_lbl.configure(bg=theme["card_bg"], fg=theme["fg_muted"])

        self.status_dot.configure(bg=theme["card_bg"], fg=theme["green"] if "готов" in self.status.get().lower() or "загружено" in self.status.get().lower() else theme["amber"])
        self.status_lbl.configure(bg=theme["card_bg"], fg=theme["fg_muted"])

        self.hist_title.configure(bg=theme["card_bg"], fg=theme["fg"])
        self.hist_badge.configure(bg=theme["purple_bg"], fg=theme["purple_fg"])
        self.hist_summary.configure(bg=theme["card_bg"], fg=theme["fg_dim"])

        if hasattr(self, "search_clear_btn"):
            self.search_clear_btn.configure(bg=theme["input_bg"], fg=theme["fg_muted"],
                                            activebackground=theme["card_border"], activeforeground=theme["fg"])
        if hasattr(self, "hist_filter_frame"):
            self.hist_filter_frame.configure(bg=theme["card_bg"])
            self.hist_filter_icon.configure(bg=theme["card_bg"], fg=theme["fg_muted"])
            self.hist_clear_btn.configure(bg=theme["input_bg"], fg=theme["fg_muted"],
                                          activebackground=theme["card_border"], activeforeground=theme["fg"])

        self._draw_header()
        self._update_candidate_preview()

    def _on_enter(self):
        if self.generate_button.instate(["!disabled"]):
            self.generate()

    def _on_search_change(self, *_args):
        query = self.search_var.get().strip().lower()
        students = self.groups.get(self.group.get(), [])
        if not students:
            return
        if not query:
            names = [person.display_name for person in students]
        else:
            names = [person.display_name for person in students if query in person.display_name.lower()]
        self.student.configure(values=names)
        if names:
            self.student.configure(state="readonly")
        else:
            self.student.configure(state="disabled")
        self.select_student()

    def load_sources(self):
        self.groups, self.tickets = {}, {}
        self.group.set("")
        self.student.set("")
        self.search_var.set("")
        self.group.configure(values=())
        self.student.configure(values=(), state="disabled")
        self.generate_button.state(["disabled"])

        theme = THEMES[self.current_theme]
        self.status_dot.configure(fg=theme["amber"])

        try:
            while True:
                try:
                    self.journal.initialize()
                    break
                except PermissionError:
                    if not messagebox.askretrycancel("Журнал недоступен",
                                                    "Закройте results.xlsx в Excel и нажмите «Повторить».",
                                                    parent=self.root):
                        self.status.set("Журнал недоступен. После закрытия файла перезагрузите данные.")
                        self.status_dot.configure(fg=theme["rose"])
                        return

            self.groups = read_students(self.data_dir / "students.xlsx")
            self.tickets = read_tickets(self.data_dir / "tickets.docx")
            self.group.configure(values=tuple(self.groups))

            # Read journal history for the table
            self.journal_records = load_journal_history(self.data_dir / "results.xlsx")
            self._refresh_history_table()
            self._update_kpi()

            total_students = sum(len(st_list) for st_list in self.groups.values())
            self.status.set(f"Готово: {len(self.groups)} групп ({total_students} студ.), {len(self.tickets)} билетов.")
            self.status_dot.configure(fg=theme["green"])
            self.group_info_lbl.configure(text=f"Всего групп: {len(self.groups)}. Выберите нужную группу.")
            self._update_candidate_preview()

        except (DataError, OSError) as exc:
            logger.exception("Ошибка загрузки данных")
            self.status.set("Проверьте файлы и нажмите «Перезагрузить файлы».")
            self.status_dot.configure(fg=theme["rose"])
            self._update_kpi()
            messagebox.showerror("Не удалось загрузить данные", str(exc), parent=self.root)

    def _update_kpi(self):
        total_students = sum(len(st_list) for st_list in self.groups.values())
        self.kpi_cards[0]["val"].configure(text=str(len(self.groups)))
        self.kpi_cards[1]["val"].configure(text=str(total_students))
        self.kpi_cards[2]["val"].configure(text=str(len(self.tickets)))
        self.kpi_cards[3]["val"].configure(text=str(len(self.journal_records)))
        self.hist_badge.configure(text=f"{len(self.journal_records)} записей")

    def _refresh_history_table(self):
        for item in self.tree.get_children():
            self.tree.delete(item)

        theme = THEMES[self.current_theme]
        query = self.hist_search_var.get().strip().lower() if hasattr(self, "hist_search_var") else ""

        count = 0
        # Show reverse chronological (newest first)
        for rec in reversed(self.journal_records):
            if query:
                haystack = f"{rec['student']} {rec['group']} {rec['ticket']} {rec['datetime']}".lower()
                if query not in haystack:
                    continue

            rep_text = "Да" if rec["repeated"] else "Нет"
            item_id = self.tree.insert("", "end", values=(
                rec["idx"],
                rec["datetime"],
                rec["group"],
                rec["student"],
                rec["ticket"],
                rep_text
            ))
            if rec["repeated"]:
                self.tree.item(item_id, tags=("repeated",))
            count += 1

        self.tree.tag_configure("repeated", foreground=theme["amber"])
        if hasattr(self, "hist_badge"):
            if query:
                self.hist_badge.configure(text=f"{count} из {len(self.journal_records)}")
            else:
                self.hist_badge.configure(text=f"{len(self.journal_records)} записей")

    def _on_tree_select(self, _event=None):
        selected = self.tree.selection()
        if not selected:
            return
        vals = self.tree.item(selected[0], "values")
        if vals:
            self.hist_summary.configure(
                text=f"Выбрана запись №{vals[0]}: {vals[3]} ({vals[2]}) → Билет №{vals[4]} [выдан {vals[1]}]"
            )

    def _on_tree_double_click(self, _event):
        selected = self.tree.selection()
        if not selected:
            return
        vals = self.tree.item(selected[0], "values")
        if not vals:
            return
        ticket_num_raw = vals[4]
        ticket_extra = ""
        try:
            tnum = int(ticket_num_raw)
            if tnum in self.tickets:
                t = self.tickets[tnum]
                q_text = "\n".join(f"  {i}. {q}" for i, q in enumerate(t.questions, 1))
                ticket_extra = f"\n\nВопросы билета № {tnum}:\n{q_text}"
        except Exception:
            pass

        msg = (f"Запись № {vals[0]}\n"
               f"Дата и время: {vals[1]}\n"
               f"Группа: {vals[2]}\n"
               f"Студент: {vals[3]}\n"
               f"Назначенный билет: № {vals[4]}\n"
               f"Повторная выдача: {vals[5]}"
               f"{ticket_extra}")
        messagebox.showinfo(f"Запись журнала № {vals[0]}", msg, parent=self.root)

    def select_group(self, _event=None):
        group_name = self.group.get()
        students = self.groups.get(group_name, [])
        self.student.set("")
        self.search_var.set("")
        self.student.configure(values=[person.display_name for person in students],
                               state="readonly" if students else "disabled")
        self.generate_button.state(["disabled"])

        theme = THEMES[self.current_theme]
        if students:
            assigned_names = {r["student"] for r in self.journal_records if r["group"] == group_name}
            completed = sum(1 for s in students if s.display_name in assigned_names)
            total = len(students)
            self.group_progress.configure(maximum=total, value=completed)
            pct = int((completed / total) * 100) if total else 0
            self.status.set(f"В группе «{group_name}» {total} студ. Выберите студента.")
            self.group_info_lbl.configure(text=f"Студентов: {total} (сдали: {completed}, осталось: {total - completed} · {pct}%)")
            self.status_dot.configure(fg=theme["accent"])
        else:
            self.group_progress.configure(maximum=1, value=0)
            self.status.set("В этой группе нет студентов.")
            self.group_info_lbl.configure(text="В этой группе нет студентов.")
            self.status_dot.configure(fg=theme["amber"])

        self._update_candidate_preview()

    def select_student(self, _event=None):
        selected_name = self.student.get().strip()
        students = self.groups.get(self.group.get(), [])
        is_valid_student = any(person.display_name == selected_name for person in students) or (self.student.current() >= 0 and self.group.get() in self.groups)

        valid = self.group.get() in self.groups and is_valid_student and bool(self.tickets)
        self.generate_button.state(["!disabled"] if valid else ["disabled"])

        theme = THEMES[self.current_theme]
        if valid:
            self.status.set("Студент выбран. Нажмите «Сгенерировать билет» (или Enter).")
            self.status_dot.configure(fg=theme["green"])
        elif self.group.get() and not selected_name:
            self.status.set("Выберите студента.")
            self.status_dot.configure(fg=theme["amber"])

        self._update_candidate_preview()

    def _find_prior_record(self, group: str, student_name: str) -> dict | None:
        for rec in self.journal_records:
            if rec["group"] == group and rec["student"] == student_name:
                return rec
        return None

    def _update_candidate_preview(self):
        theme = THEMES[self.current_theme]
        group = self.group.get().strip()
        name = self.student.get().strip()

        if not group or not name:
            self.profile_name_lbl.configure(text="Студент не выбран", fg=theme["fg"])
            self.profile_badge_lbl.configure(text="Ожидание выбора студента", bg=theme["card_alt"], fg=theme["fg_muted"])
            self.profile_desc_lbl.configure(text="Выберите группу и студента из списков выше для проверки статуса сдачи.", fg=theme["fg_muted"])
            return

        self.profile_name_lbl.configure(text=f"{name}  ({group})", fg=theme["fg"])
        prior = self._find_prior_record(group, name)

        if prior:
            self.profile_badge_lbl.configure(
                text="⚠️  Повторная выдача билета",
                bg=theme["amber_bg"],
                fg=theme["amber_fg"]
            )
            self.profile_desc_lbl.configure(
                text=f"Студент ранее уже получал билет № {prior['ticket']} ({prior['datetime']}). Согласно регламенту, будет сохранён тот же номер билета.",
                fg=theme["amber_fg"] if self.current_theme == "dark" else theme["amber"]
            )
        else:
            self.profile_badge_lbl.configure(
                text="✨  Первичное получение билета",
                bg=theme["green_bg"],
                fg=theme["green_fg"]
            )
            self.profile_desc_lbl.configure(
                text=f"Студент сдаёт впервые. Будет случайно определён билет из {len(self.tickets)} доступных в базе.",
                fg=theme["green_fg"] if self.current_theme == "dark" else theme["green"]
            )

    def generate(self):
        students = self.groups.get(self.group.get(), [])
        selected_name = self.student.get().strip()
        student = next((p for p in students if p.display_name == selected_name), None)

        if not student:
            idx = self.student.current()
            if 0 <= idx < len(students):
                student = students[idx]

        if not student or not self.tickets:
            return

        while True:
            try:
                ticket, repeated = self.journal.assign(student, self.tickets)
                break
            except PermissionError:
                if not messagebox.askretrycancel(
                    "Не удалось сохранить билет",
                    "Закройте results.xlsx в Excel и нажмите «Повторить». Билет будет показан после сохранения.",
                    parent=self.root
                ):
                    return
            except (DataError, OSError) as exc:
                logger.exception("Ошибка назначения билета")
                messagebox.showerror("Не удалось выдать билет", str(exc), parent=self.root)
                return

        # Reload journal records so history table and KPIs are fresh immediately
        self.journal_records = load_journal_history(self.data_dir / "results.xlsx")
        self._refresh_history_table()
        self._update_kpi()

        self.show_result(student, ticket, repeated)

    def show_result(self, student: Student, ticket: Ticket, repeated: bool):
        try:
            import winsound
            winsound.MessageBeep(winsound.MB_ICONASTERISK)
        except Exception:
            pass

        theme = THEMES[self.current_theme]

        window = tk.Toplevel(self.root)
        window.title(f"Ваш билет № {ticket.number}")
        window.geometry("820x640")
        window.minsize(680, 500)
        window.configure(bg=theme["bg"])
        window.transient(self.root)
        window.grab_set()

        # Center relative to root window
        try:
            rx, ry = self.root.winfo_x(), self.root.winfo_y()
            rw, rh = self.root.winfo_width(), self.root.winfo_height()
            x = max(rx + (rw - 820) // 2, 20)
            y = max(ry + (rh - 640) // 2, 20)
            window.geometry(f"820x640+{x}+{y}")
        except Exception:
            pass

        # Outer padding frame
        container = tk.Frame(window, bg=theme["bg"], padx=24, pady=20)
        container.pack(fill="both", expand=True)

        # Header Certificate Banner
        hdr_card = tk.Frame(container, bg=theme["card_bg"], highlightbackground=theme["card_border"],
                            highlightthickness=1, padx=20, pady=16)
        hdr_card.pack(fill="x", pady=(0, 14))

        hdr_top = tk.Frame(hdr_card, bg=theme["card_bg"])
        hdr_top.pack(fill="x")

        tk.Label(hdr_top, text=f"Ваш билет № {ticket.number}",
                 font=("Segoe UI", 21, "bold"), fg=theme["fg"], bg=theme["card_bg"]).pack(side="left")

        # Status badge
        badge_text = "⚠️  ПОВТОРНАЯ ВЫДАЧА" if repeated else "✨  ПЕРВИЧНАЯ ВЫДАЧА"
        badge_bg = theme["amber_bg"] if repeated else theme["green_bg"]
        badge_fg = theme["amber_fg"] if repeated else theme["green_fg"]

        badge = tk.Label(hdr_top, text=badge_text, font=("Segoe UI", 9, "bold"),
                         bg=badge_bg, fg=badge_fg, padx=10, pady=5)
        badge.pack(side="right")

        # Subtitle info bar
        info_str = f"{student.group} · {student.display_name}   •   {datetime.now().strftime('%d.%m.%Y %H:%M')}"
        tk.Label(hdr_card, text=info_str, font=("Segoe UI", 11),
                 fg=theme["fg_muted"], bg=theme["card_bg"]).pack(anchor="w", pady=(8, 0))

        # Questions Scrollable Area
        questions_canvas = tk.Canvas(container, bg=theme["bg"], highlightthickness=0)
        q_scroll = ttk.Scrollbar(container, orient="vertical", command=questions_canvas.yview)
        q_frame = tk.Frame(questions_canvas, bg=theme["bg"])

        q_window = questions_canvas.create_window((0, 0), window=q_frame, anchor="nw")
        questions_canvas.configure(yscrollcommand=q_scroll.set)

        def _on_qframe_config(_event):
            questions_canvas.configure(scrollregion=questions_canvas.bbox("all"))

        def _on_canvas_config(event):
            questions_canvas.itemconfig(q_window, width=event.width)

        q_frame.bind("<Configure>", _on_qframe_config)
        questions_canvas.bind("<Configure>", _on_canvas_config)

        # Mouse wheel support on questions canvas
        def _on_mousewheel(event):
            questions_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        questions_canvas.bind_all("<MouseWheel>", _on_mousewheel)

        questions_canvas.pack(side="top", fill="both", expand=True)
        q_scroll.pack(side="right", fill="y", before=questions_canvas)

        # Render each question in an eye-catching Card
        colors = [
            (theme["purple"], theme["purple_bg"], theme["purple_fg"]),
            (theme["accent"], theme["header_bg1"], theme["cyan_fg"]),
            (theme["cyan"], theme["cyan_bg"], theme["cyan_fg"]),
        ]

        full_ticket_text_lines = [
            f"ЭКЗАМЕНАЦИОННЫЙ БИЛЕТ № {ticket.number}",
            f"Группа: {student.group}",
            f"Студент: {student.display_name}",
            f"Дата выдачи: {datetime.now().strftime('%d.%m.%Y %H:%M:%S')}",
            f"Статус: {'Повторная выдача' if repeated else 'Первичная выдача'}",
            "",
            "ВОПРОСЫ:",
        ]

        for idx, question in enumerate(ticket.questions, 1):
            full_ticket_text_lines.append(f"{idx}. {question}")

            q_card = tk.Frame(q_frame, bg=theme["card_bg"], highlightbackground=theme["card_border"],
                              highlightthickness=1, padx=18, pady=14)
            q_card.pack(fill="x", pady=6)

            tag_accent, tag_bg, tag_fg = colors[(idx - 1) % len(colors)]

            q_tag = tk.Label(q_card, text=f"  ВОПРОС № {idx}  ", font=("Segoe UI", 9, "bold"),
                             bg=tag_bg, fg=tag_fg, padx=6, pady=3)
            q_tag.pack(anchor="w")

            q_text = tk.Label(q_card, text=question, font=("Segoe UI", 12),
                              fg=theme["fg"], bg=theme["card_bg"],
                              wraplength=720, justify="left")
            q_text.pack(anchor="w", pady=(8, 2))

        # Save notification
        save_status_text = "Результат сохранён в results.xlsx" + (" · Повторная выдача" if repeated else "")
        tk.Label(container, text=save_status_text, font=("Segoe UI", 9),
                 fg=theme["fg_muted"], bg=theme["bg"]).pack(anchor="w", pady=(6, 2))

        # Bottom Action Bar
        bottom_bar = tk.Frame(container, bg=theme["bg"], pady=10)
        bottom_bar.pack(fill="x", side="bottom")

        copy_notify_lbl = tk.Label(bottom_bar, text="", font=("Segoe UI", 9, "bold"),
                                   fg=theme["green"], bg=theme["bg"])
        copy_notify_lbl.pack(side="left")

        def copy_ticket():
            text_to_copy = "\n".join(full_ticket_text_lines)
            self.root.clipboard_clear()
            self.root.clipboard_append(text_to_copy)
            copy_notify_lbl.configure(text="✓ Билет скопирован в буфер обмена!")
            window.after(2500, lambda: copy_notify_lbl.configure(text=""))

        def export_ticket():
            text_to_save = "\n".join(full_ticket_text_lines)
            default_fname = f"Билет_{ticket.number}_{student.surname}_{student.name}.txt"
            fpath = filedialog.asksaveasfilename(
                parent=window,
                initialfile=default_fname,
                defaultextension=".txt",
                filetypes=[("Текстовые файлы", "*.txt"), ("Все файлы", "*.*")]
            )
            if fpath:
                try:
                    with open(fpath, "w", encoding="utf-8") as f:
                        f.write(text_to_save)
                    messagebox.showinfo("Сохранено", f"Билет сохранён в:\n{fpath}", parent=window)
                except Exception as err:
                    messagebox.showerror("Ошибка сохранения", str(err), parent=window)

        def close(_event=None):
            try:
                questions_canvas.unbind_all("<MouseWheel>")
            except Exception:
                pass
            window.grab_release()
            window.destroy()
            self.student.set("")
            self.search_var.set("")
            self.generate_button.state(["disabled"])
            self.status.set("Выберите следующего студента.")
            self.status_dot.configure(fg=theme["green"])
            self._update_candidate_preview()
            self.student.focus_set()
            return "break"

        copy_btn = ttk.Button(bottom_bar, text="📋 Скопировать билет", command=copy_ticket, style="Secondary.TButton")
        copy_btn.pack(side="left", padx=(10, 6))

        export_btn = ttk.Button(bottom_bar, text="💾 Сохранить в файл", command=export_ticket, style="Secondary.TButton")
        export_btn.pack(side="left")

        close_btn = ttk.Button(bottom_bar, text="К выбору студента · Esc", command=close, style="Accent.TButton")
        close_btn.pack(side="right")

        window.bind("<Escape>", close)
        window.protocol("WM_DELETE_WINDOW", close)
        window.focus_set()

    def open_data_folder(self):
        open_system_path(self.data_dir)

    def open_excel_file(self):
        path = self.data_dir / "results.xlsx"
        if path.exists():
            open_system_path(path)
        else:
            messagebox.showinfo("Журнал отсутствует", "Файл results.xlsx ещё не создан.", parent=self.root)

    def create_demo_files(self):
        try:
            from create_demo_data import create_demo
            create_demo(self.data_dir)
            self.load_sources()
            messagebox.showinfo("Учебные примеры", "Демонстрационные файлы проверены/созданы.", parent=self.root)
        except Exception as exc:
            messagebox.showerror("Ошибка", f"Не удалось создать примеры: {exc}", parent=self.root)
