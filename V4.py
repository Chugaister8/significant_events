from __future__ import annotations

from typing import Callable, Literal, TypeAlias

import csv
import json
import os
import re
from collections import Counter
from datetime import datetime

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

try:
    import pandas as pd
except ImportError:
    pd = None

try:
    import matplotlib

    matplotlib.use("TkAgg")
    import matplotlib.pyplot as plt
    from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
    from matplotlib.figure import Figure

    HAS_MPL = True
except ImportError:
    HAS_MPL = False


# =============================================================================
#  ГЛОБАЛЬНІ КОНСТАНТИ ТА ПАЛІТРА
# =============================================================================

DATA_FILE = "essential_events.json"

COLORS = {
    # --- Основні поверхні ---
    "bg_main": "#243640",        # Головний фон (глибокий сірий-синій)
    "bg_sidebar": "#1E2C33",     # Трохи темніший для контрасту
    "bg_header": "#1E2C33",      # Верхній бар
    "bg_surface": "#2E4450",     # Поверхні / картки
    "bg_surface_alt": "#344E5A", # Трохи світліша тінь
    "bg_input": "#1E2C33",       # Поля вводу в мінімалістичному стилі

    # --- Акценти ---
    "accent": "#4F46E5",         # Індиго — ідеально гармонує з #243640
    "accent_soft": "#6366F1",    # Hover на кнопках
    "accent_muted": "#818CF8",   # Легкий індиго для тексту
    "accent_success": "#10B981", # Зелений позитив
    "accent_danger": "#EF4444",  # Червоний попередження
    "accent_warning": "#F59E0B", # Жовтий попередження

    # --- Текст ---
    "text_primary": "#F3F4F6",   # Майже білий
    "text_muted": "#CBD5E1",     # Світло-сірий
    "text_subtle": "#94A3B8",    # Блідо-сірий (підтексти)

    # --- Кордони ---
    "border_soft": "#3B4F59",
    "border_strong": "#556871",

    # --- Ряди таблиці ---
    "row_even": "#2A3D47",
    "row_odd": "#243640",
    "row_select": "#3B82F6",     # Гарний синій для виділення
}

RISK_COLORS = {
    "Операцiйний": COLORS["accent_warning"],
    "Технiчний": COLORS["accent"],
    "Фiнансовий": COLORS["accent_danger"],
    "Репутацiйний": "#a855f7",
    "Екологiчний": COLORS["accent_success"],
    "Надзвичайна ситуацiя": "#f97316",
}

EVENT_TYPES = [
    "Вимушений простiй < 24 год",
    "Вимушений простiй >= 24 год",
    "Зупинка виробництва",
    "Аварiя обладнання",
    "Пошкодження майна",
    "Порушення дозволiв",
    "Крадiжка / диверсiя",
    "Iнше",
]

RISK_TYPES = [
    "Операцiйний",
    "Технiчний",
    "Фiнансовий",
    "Репутацiйний",
    "Екологiчний",
    "Надзвичайна ситуацiя",
]


# =============================================================================
#  ХЕЛПЕРИ ТА СТИЛЬ
# =============================================================================

def is_valid_date(s: str) -> bool:
    """Перевірка дати у форматі дд.мм.рррр."""
    if not s or s in ("дд.мм.рррр", ""):
        return True
    if not re.match(r"^\d{2}\.\d{2}\.\d{4}$", s):
        return False
    try:
        datetime.strptime(s, "%d.%m.%Y")
        return True
    except ValueError:
        return False


def apply_dark_style(root: tk.Misc) -> None:
    """Налаштовує темний, мінімалістичний стиль для ttk."""
    style = ttk.Style(root)
    style.theme_use("clam")
    C = COLORS

    # --- Базові налаштування ---
    style.configure(
        ".",
        background=C["bg_main"],
        foreground=C["text_primary"],
        fieldbackground=C["bg_input"],
        troughcolor=C["bg_surface"],
        bordercolor=C["border_soft"],
        darkcolor=C["bg_surface"],
        lightcolor=C["bg_surface"],
        insertcolor=C["text_primary"],
        selectbackground=C["row_select"],
        selectforeground=C["text_primary"],
        font=("Arial", 9),
    )

    style.configure("TFrame", background=C["bg_main"])
    style.configure("Surface.TFrame", background=C["bg_surface"])
    style.configure("Sidebar.TFrame", background=C["bg_sidebar"])
    style.configure("Header.TFrame", background=C["bg_header"])

    style.configure(
        "TLabel",
        background=C["bg_main"],
        foreground=C["text_primary"],
        font=("Arial", 9),
    )
    style.configure(
        "Muted.TLabel",
        background=C["bg_main"],
        foreground=C["text_muted"],
        font=("Arial", 8),
    )

    # --- Entry ---
    style.configure(
        "TEntry",
        fieldbackground=C["bg_input"],
        foreground=C["text_primary"],
        bordercolor=C["border_soft"],
        insertcolor=C["text_primary"],
    )
    style.map(
        "TEntry",
        fieldbackground=[("focus", C["bg_surface_alt"])],
        bordercolor=[("focus", C["accent"])],
    )

    # --- Combobox ---
    style.configure(
        "TCombobox",
        fieldbackground=C["bg_surface"],
        background=C["bg_surface"],
        foreground=C["text_primary"],
        bordercolor=C["border_soft"],
        arrowcolor=C["text_muted"],
    )
    style.map(
        "TCombobox",
        fieldbackground=[
            ("readonly", C["bg_surface"]),
            ("hover", C["bg_surface_alt"]),
            ("focus", C["bg_surface_alt"]),
        ],
        background=[
            ("readonly", C["bg_surface"]),
            ("hover", C["bg_surface_alt"]),
            ("focus", C["bg_surface_alt"]),
        ],
        foreground=[
            ("disabled", C["text_subtle"]),
        ],
        arrowcolor=[
            ("hover", C["text_primary"]),
            ("focus", C["accent"]),
        ],
    )

    # --- Notebook / Tabs ---
    style.configure(
        "TNotebook",
        background=C["bg_main"],
        bordercolor=C["border_soft"],
        tabmargins=[0, 0, 0, 0],
    )
    style.configure(
        "TNotebook.Tab",
        background=C["bg_sidebar"],
        foreground=C["text_muted"],
        padding=(14, 6),
        font=("Arial", 9),
    )
    style.map(
        "TNotebook.Tab",
        background=[
            ("selected", C["bg_surface"]),
            ("active", C["bg_surface_alt"]),
        ],
        foreground=[
            ("selected", C["text_primary"]),
            ("active", C["text_primary"]),
        ],
    )

    # --- Treeview, Scrollbars тощо (як було, можеш залишити) ---
    style.configure(
        "Treeview",
        background=C["row_odd"],
        foreground=C["text_primary"],
        fieldbackground=C["row_odd"],
        bordercolor=C["border_soft"],
        font=("Arial", 9),
        rowheight=24,
    )
    style.configure(
        "Treeview.Heading",
        background=C["bg_surface"],
        foreground=C["text_muted"],
        bordercolor=C["border_soft"],
        font=("Arial", 8, "bold"),
        relief="flat",
    )
    style.map(
        "Treeview",
        background=[("selected", C["row_select"])],
        foreground=[("selected", C["text_primary"])],
    )

    style.configure(
        "Vertical.TScrollbar",
        background=C["bg_surface"],
        troughcolor=C["bg_main"],
        arrowcolor=C["text_muted"],
        bordercolor=C["bg_main"],
    )
    style.configure(
        "Horizontal.TScrollbar",
        background=C["bg_surface"],
        troughcolor=C["bg_main"],
        arrowcolor=C["text_muted"],
        bordercolor=C["bg_main"],
    )


def make_dark_text(parent: tk.Misc, **kwargs) -> tk.Text:
    """Створює Text у стилі нового дизайну."""
    C = COLORS
    return tk.Text(
        parent,
        bg=C["bg_input"],
        fg=C["text_primary"],
        insertbackground=C["text_primary"],
        selectbackground=C["row_select"],
        selectforeground=C["text_primary"],
        relief="flat",
        bd=1,
        highlightthickness=1,
        highlightbackground=C["border_soft"],
        highlightcolor=C["accent"],
        font=("Arial", 9),
        **kwargs,
    )

def add_placeholder(entry: tk.Entry, text: str) -> None:
    """Додає плейсхолдер у Entry."""
    entry.insert(0, text)
    entry.configure(fg=COLORS["text_muted"])

    def on_in(_: object) -> None:
        if entry.get() == text:
            entry.delete(0, tk.END)
            entry.configure(fg=COLORS["text_primary"])

    def on_out(_: object) -> None:
        if not entry.get():
            entry.insert(0, text)
            entry.configure(fg=COLORS["text_muted"])

    entry.bind("<FocusIn>", on_in)
    entry.bind("<FocusOut>", on_out)


# =============================================================================
#  ДЕТАЛЬНЕ ВІКНО ЗАПИСУ
# =============================================================================

class EventDetailWindow:
    """Спливаюче вікно для перегляду та редагування збереженої події."""

    def __init__(
        self,
        parent_root: tk.Misc,
        record: tuple,
        all_records: list[tuple],
        save_callback: Callable[[str, tuple], None],
        delete_callback: Callable[[str], None],
        toast_callback: Callable[[str], None],
    ) -> None:
        self.parent_root = parent_root
        self.record = list(record)
        self.all_records = all_records
        self.save_callback = save_callback
        self.delete_callback = delete_callback
        self.toast_callback = toast_callback
        self.is_edit_mode = False

        self._build_window()

    def _build_window(self) -> None:
        C = COLORS
        self.win = tk.Toplevel(self.parent_root)
        self.win.title(f"Подiя #{self.record[0]}  —  {self.record[1]}")
        self.win.geometry("780x700")
        self.win.minsize(640, 500)
        self.win.configure(bg=C["bg_main"])
        self.win.grab_set()

        self.win.update_idletasks()
        rx = self.parent_root.winfo_x()
        ry = self.parent_root.winfo_y()
        rw = self.parent_root.winfo_width()
        rh = self.parent_root.winfo_height()
        ww, wh = 780, 700
        x = rx + (rw - ww) // 2
        y = ry + (rh - wh) // 2
        self.win.geometry(f"{ww}x{wh}+{x}+{y}")

        self.win.rowconfigure(1, weight=1)
        self.win.columnconfigure(0, weight=1)

        # Header
        header = tk.Frame(self.win, bg=C["bg_header"], height=58)
        header.grid(row=0, column=0, sticky="ew")
        header.columnconfigure(1, weight=1)
        header.grid_propagate(False)

        risk_color = RISK_COLORS.get(self.record[4], COLORS["accent"])
        tk.Frame(header, bg=risk_color, width=4).grid(row=0, column=0, sticky="ns")

        title_frame = tk.Frame(header, bg=C["bg_header"])
        title_frame.grid(row=0, column=1, sticky="ew", padx=16, pady=10)

        self.lbl_title = tk.Label(
            title_frame,
            text=f"Запис #{self.record[0]}",
            bg=C["bg_header"],
            fg=COLORS["accent_muted"],
            font=("Arial", 11, "bold"),
        )
        self.lbl_title.pack(anchor="w")

        self.lbl_subtitle = tk.Label(
            title_frame,
            text=self.record[1],
            bg=C["bg_header"],
            fg=COLORS["text_muted"],
            font=("Arial", 9),
        )
        self.lbl_subtitle.pack(anchor="w")

        status_val = self.record[10] if len(self.record) > 10 else "—"
        status_colors = {
            "Вiдкрито": COLORS["accent_danger"],
            "В обробцi": COLORS["accent_warning"],
            "Закрито": COLORS["text_muted"],
            "Вирiшено": COLORS["accent_success"],
        }
        sc = status_colors.get(status_val, COLORS["text_muted"])
        self.lbl_status_badge = tk.Label(
            header,
            text=f"  {status_val}  ",
            bg=sc,
            fg="white",
            font=("Arial", 8, "bold"),
            pady=3,
        )
        self.lbl_status_badge.grid(row=0, column=2, padx=16, pady=18)

        # Scroll content
        canvas = tk.Canvas(self.win, bg=C["bg_main"], highlightthickness=0)
        sb = ttk.Scrollbar(self.win, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=sb.set)
        canvas.grid(row=1, column=0, sticky="nsew")
        sb.grid(row=1, column=1, sticky="ns")

        self.content = tk.Frame(canvas, bg=C["bg_main"])
        self._cw = canvas.create_window((0, 0), window=self.content, anchor="nw")

        def on_conf(_: object) -> None:
            canvas.configure(scrollregion=canvas.bbox("all"))
            canvas.itemconfig(self._cw, width=canvas.winfo_width())

        self.content.bind("<Configure>", on_conf)
        canvas.bind("<Configure>", on_conf)
        canvas.bind_all(
            "<MouseWheel>",
            lambda e: canvas.yview_scroll(int(-1 * e.delta / 120), "units"),
        )

        self.content.columnconfigure(0, weight=1)
        self.content.columnconfigure(1, weight=1)

        self._build_view_content()

        # Bottom buttons
        btn_bar = tk.Frame(self.win, bg=C["bg_header"], height=46)
        btn_bar.grid(row=2, column=0, columnspan=2, sticky="ew")
        btn_bar.grid_propagate(False)
        btn_bar.columnconfigure(0, weight=1)

        left_btns = tk.Frame(btn_bar, bg=C["bg_header"])
        left_btns.pack(side="left", padx=12, pady=8)

        self.btn_edit = tk.Button(
            left_btns,
            text="Редагувати",
            bg=COLORS["accent_warning"],
            fg=COLORS["bg_main"],
            activebackground="#d97706",
            activeforeground="white",
            relief="flat",
            bd=0,
            cursor="hand2",
            font=("Arial", 9, "bold"),
            padx=14,
            pady=4,
            command=self._toggle_edit_mode,
        )
        self.btn_edit.pack(side="left", padx=(0, 8))

        self.btn_save = tk.Button(
            left_btns,
            text="Зберегти змiни",
            bg=COLORS["accent_success"],
            fg="white",
            activebackground="#16a34a",
            activeforeground="white",
            relief="flat",
            bd=0,
            cursor="hand2",
            font=("Arial", 9, "bold"),
            padx=14,
            pady=4,
            command=self._save_changes,
        )
        self.btn_save.pack_forget()

        self.btn_cancel_edit = tk.Button(
            left_btns,
            text="Скасувати",
            bg=COLORS["bg_surface"],
            fg=COLORS["text_muted"],
            activebackground=COLORS["bg_surface_alt"],
            activeforeground=COLORS["text_primary"],
            relief="flat",
            bd=0,
            cursor="hand2",
            font=("Arial", 9),
            padx=12,
            pady=4,
            command=self._cancel_edit,
        )
        self.btn_cancel_edit.pack_forget()

        right_btns = tk.Frame(btn_bar, bg=C["bg_header"])
        right_btns.pack(side="right", padx=12, pady=8)

        tk.Button(
            right_btns,
            text="Видалити",
            bg=COLORS["accent_danger"],
            fg="white",
            activebackground="#dc2626",
            activeforeground="white",
            relief="flat",
            bd=0,
            cursor="hand2",
            font=("Arial", 9, "bold"),
            padx=14,
            pady=4,
            command=self._delete_record,
        ).pack(side="right", padx=(8, 0))

        tk.Button(
            right_btns,
            text="Закрити",
            bg=COLORS["bg_surface"],
            fg=COLORS["text_primary"],
            activebackground=COLORS["bg_surface_alt"],
            activeforeground=COLORS["text_primary"],
            relief="flat",
            bd=0,
            cursor="hand2",
            font=("Arial", 9),
            padx=12,
            pady=4,
            command=self.win.destroy,
        ).pack(side="right")

    def _section_label(self, parent: tk.Misc, text: str, row: int) -> None:
        C = COLORS
        f = tk.Frame(parent, bg=C["bg_main"])
        f.grid(row=row, column=0, columnspan=2, sticky="ew", padx=8, pady=(14, 4))
        tk.Frame(f, bg=COLORS["accent_soft"], width=2, height=16).pack(side="left")
        tk.Label(
            f,
            text=text,
            bg=C["bg_main"],
            fg=COLORS["accent_muted"],
            font=("Arial", 9, "bold"),
        ).pack(side="left", padx=8)

    def _info_row(
        self, parent: tk.Misc, label: str, value: str, row: int, col: int = 0
    ) -> None:
        C = COLORS
        cell = tk.Frame(parent, bg=C["bg_surface"], padx=10, pady=6)
        cell.grid(
            row=row,
            column=col,
            sticky="nsew",
            padx=(8 if col == 0 else 4, 4 if col == 0 else 8),
            pady=3,
        )
        cell.columnconfigure(0, weight=1)
        tk.Label(
            cell,
            text=label,
            bg=C["bg_surface"],
            fg=C["text_subtle"],
            font=("Arial", 7, "bold"),
        ).grid(row=0, column=0, sticky="w")
        tk.Label(
            cell,
            text=value or "—",
            bg=C["bg_surface"],
            fg=C["text_primary"],
            font=("Arial", 9),
            wraplength=260,
            justify="left",
        ).grid(row=1, column=0, sticky="w", pady=(2, 0))

    def _text_block(self, parent: tk.Misc, label: str, value: str, row: int) -> None:
        C = COLORS
        cell = tk.Frame(parent, bg=C["bg_surface"], padx=10, pady=6)
        cell.grid(row=row, column=0, columnspan=2, sticky="nsew", padx=8, pady=3)
        cell.columnconfigure(0, weight=1)
        tk.Label(
            cell,
            text=label,
            bg=C["bg_surface"],
            fg=C["text_subtle"],
            font=("Arial", 7, "bold"),
        ).grid(row=0, column=0, sticky="w")
        t = make_dark_text(cell, height=3, wrap="word", state="normal")
        t.insert("1.0", value or "—")
        t.configure(state="disabled")
        t.grid(row=1, column=0, sticky="ew", pady=(4, 0))

    def _build_view_content(self) -> None:
        C = COLORS
        rec = self.record

        for w in self.content.winfo_children():
            w.destroy()

        r = rec
        row = 0

        self._section_label(self.content, "Iнформацiя про пiдприємство", row)
        row += 1

        self._info_row(
            self.content, "Пiдприємство", r[1] if len(r) > 1 else "—", row, 0
        )
        priority_val = r[9] if len(r) > 9 else "—"
        priority_colors = {
            "Критичний": COLORS["accent_danger"],
            "Високий": COLORS["accent_warning"],
            "Середнiй": COLORS["accent"],
            "Низький": COLORS["accent_success"],
        }
        p_color = priority_colors.get(priority_val, COLORS["text_primary"])
        cell_p = tk.Frame(self.content, bg=C["bg_surface"], padx=10, pady=6)
        cell_p.grid(row=row, column=1, sticky="nsew", padx=(4, 8), pady=3)
        cell_p.columnconfigure(0, weight=1)
        tk.Label(
            cell_p,
            text="Прiоритет",
            bg=C["bg_surface"],
            fg=C["text_subtle"],
            font=("Arial", 7, "bold"),
        ).grid(row=0, column=0, sticky="w")
        tk.Label(
            cell_p,
            text=priority_val,
            bg=C["bg_surface"],
            fg=p_color,
            font=("Arial", 10, "bold"),
        ).grid(row=1, column=0, sticky="w", pady=(2, 0))
        row += 1

        self._section_label(self.content, "Опис подiї / ризику", row)
        row += 1

        self._info_row(
            self.content, "Назва подiї", r[2] if len(r) > 2 else "—", row, 0
        )

        risk_val = r[4] if len(r) > 4 else "—"
        rc = RISK_COLORS.get(risk_val, C["text_primary"])
        cell_r = tk.Frame(self.content, bg=C["bg_surface"], padx=10, pady=6)
        cell_r.grid(row=row, column=1, sticky="nsew", padx=(4, 8), pady=3)
        cell_r.columnconfigure(0, weight=1)
        tk.Label(
            cell_r,
            text="Тип ризику",
            bg=C["bg_surface"],
            fg=C["text_subtle"],
            font=("Arial", 7, "bold"),
        ).grid(row=0, column=0, sticky="w")
        tk.Label(
            cell_r,
            text=risk_val,
            bg=C["bg_surface"],
            fg=rc,
            font=("Arial", 10, "bold"),
        ).grid(row=1, column=0, sticky="w", pady=(2, 0))
        row += 1

        self._info_row(
            self.content, "Статус", r[10] if len(r) > 10 else "—", row, 0
        )
        self._info_row(
            self.content,
            "Задiянi пiдроздiли / особи",
            r[3] if len(r) > 3 else "—",
            row,
            1,
        )
        row += 1

        self._section_label(self.content, "Дати", row)
        row += 1

        self._info_row(
            self.content, "Дата подiї", r[5] if len(r) > 5 else "—", row, 0
        )
        self._info_row(
            self.content, "Дата виявлення", r[8] if len(r) > 8 else "—", row, 1
        )
        row += 1

        self._section_label(self.content, "Деталi подiї", row)
        row += 1

        self._text_block(
            self.content,
            "Детальний опис подiї",
            r[6] if len(r) > 6 else "—",
            row,
        )
        row += 1

        self._text_block(
            self.content,
            "Вжитi заходи",
            r[7] if len(r) > 7 else "—",
            row,
        )
        row += 1

        tk.Frame(self.content, bg=C["bg_main"], height=12).grid(
            row=row, column=0, columnspan=2
        )

    def _build_edit_content(self) -> None:
        C = COLORS
        rec = self.record

        for w in self.content.winfo_children():
            w.destroy()

        self.content.columnconfigure(0, weight=1)

        def section(txt: str, row_idx: int) -> int:
            f = tk.Frame(self.content, bg=C["bg_main"])
            f.grid(row=row_idx, column=0, sticky="ew", padx=8, pady=(14, 4))
            tk.Frame(f, bg=COLORS["accent_warning"], width=2, height=16).pack(
                side="left"
            )
            tk.Label(
                f,
                text=txt,
                bg=C["bg_main"],
                fg=COLORS["accent_warning"],
                font=("Arial", 9, "bold"),
            ).pack(side="left", padx=8)
            return row_idx + 1

        def lbl(text: str, row_idx: int) -> None:
            tk.Label(
                self.content,
                text=text,
                bg=C["bg_main"],
                fg=C["text_subtle"],
                font=("Arial", 8),
            ).grid(row=row_idx, column=0, sticky="w", padx=10, pady=(6, 0))

        def make_entry(**kw: object) -> tk.Entry:
            return tk.Entry(
                self.content,
                bg=C["bg_input"],
                fg=C["text_primary"],
                insertbackground=C["text_primary"],
                relief="flat",
                bd=2,
                highlightthickness=1,
                highlightbackground=C["border_soft"],
                highlightcolor=COLORS["accent_warning"],
                font=("Arial", 9),
                **kw,
            )

        def make_combo(values: list[str] | None = None, **kw: object) -> ttk.Combobox:
            return ttk.Combobox(
                self.content,
                values=values or [],
                state="readonly",
                font=("Arial", 9),
                **kw,
            )

        row = 0
        row = section("Пiдприємство та подiя", row)

        lbl("Скорочена назва пiдприємства:", row)
        row += 1
        self.e_entity = make_entry()
        self.e_entity.grid(row=row, column=0, sticky="ew", padx=10, pady=(2, 0))
        self.e_entity.insert(0, rec[1] if len(rec) > 1 else "")
        row += 1

        lbl("Назва подiї:", row)
        row += 1
        self.e_event = make_combo(values=EVENT_TYPES)
        self.e_event.grid(row=row, column=0, sticky="ew", padx=10, pady=(2, 0))
        self.e_event.set(rec[2] if len(rec) > 2 else "")
        row += 1

        lbl("Тип ризику:", row)
        row += 1
        self.e_risk = make_combo(values=RISK_TYPES)
        self.e_risk.grid(row=row, column=0, sticky="ew", padx=10, pady=(2, 0))
        self.e_risk.set(rec[4] if (len(rec) > 4 and rec[4] != "—") else "")
        row += 1

        lbl("Задiянi пiдроздiли / особи:", row)
        row += 1
        self.e_involved = make_dark_text(self.content, height=2, wrap="word")
        self.e_involved.grid(
            row=row, column=0, sticky="ew", padx=10, pady=(2, 0)
        )
        if len(rec) > 3 and rec[3] and rec[3] != "—":
            self.e_involved.insert("1.0", rec[3])
        row += 1

        row = section("Дати", row)
        date_frame = tk.Frame(self.content, bg=C["bg_main"])
        date_frame.grid(row=row, column=0, sticky="w", padx=10, pady=4)
        row += 1

        for col_i, (lbl_t, attr, val_idx) in enumerate(
            [("Дата виявлення:", "e_detect", 8), ("Дата подiї:", "e_event_date", 5)]
        ):
            tk.Label(
                date_frame,
                text=lbl_t,
                bg=C["bg_main"],
                fg=C["text_subtle"],
                font=("Arial", 8),
            ).grid(
                row=0,
                column=col_i,
                padx=(0 if col_i == 0 else 20, 0),
                sticky="w",
            )
            e = tk.Entry(
                date_frame,
                bg=C["bg_input"],
                fg=C["text_primary"],
                insertbackground=C["text_primary"],
                relief="flat",
                bd=2,
                highlightthickness=1,
                highlightbackground=C["border_soft"],
                highlightcolor=COLORS["accent_warning"],
                font=("Arial", 9),
                width=14,
            )
            e.grid(
                row=1,
                column=col_i,
                padx=(0 if col_i == 0 else 20, 0),
                pady=2,
            )
            val = rec[val_idx] if len(rec) > val_idx else ""
            if val and val != "—":
                e.insert(0, val)
            else:
                add_placeholder(e, "дд.мм.рррр")
            setattr(self, attr, e)

        row = section("Прiоритет та статус", row)
        pri_stat_frame = tk.Frame(self.content, bg=C["bg_main"])
        pri_stat_frame.grid(row=row, column=0, sticky="w", padx=10, pady=4)
        row += 1

        tk.Label(
            pri_stat_frame,
            text="Прiоритет:",
            bg=C["bg_main"],
            fg=C["text_subtle"],
            font=("Arial", 8),
        ).grid(row=0, column=0, sticky="w")
        self.e_priority = ttk.Combobox(
            pri_stat_frame,
            values=["Критичний", "Високий", "Середнiй", "Низький"],
            state="readonly",
            font=("Arial", 9),
            width=14,
        )
        self.e_priority.grid(row=1, column=0, pady=(2, 0), padx=(0, 20))
        self.e_priority.set(rec[9] if len(rec) > 9 else "Середнiй")

        tk.Label(
            pri_stat_frame,
            text="Статус:",
            bg=C["bg_main"],
            fg=C["text_subtle"],
            font=("Arial", 8),
        ).grid(row=0, column=1, sticky="w")
        self.e_status = ttk.Combobox(
            pri_stat_frame,
            values=["Вiдкрито", "В обробцi", "Закрито", "Вирiшено"],
            state="readonly",
            font=("Arial", 9),
            width=14,
        )
        self.e_status.grid(row=1, column=1, pady=(2, 0))

        row = section("Деталi подiї", row)

        lbl("Детальний опис подiї:", row)
        row += 1
        self.e_description = make_dark_text(self.content, height=4, wrap="word")
        self.e_description.grid(
            row=row, column=0, sticky="ew", padx=10, pady=(2, 0)
        )
        if len(rec) > 6 and rec[6] and rec[6] != "—":
            self.e_description.insert("1.0", rec[6])
        row += 1

        lbl("Вжитi заходи:", row)
        row += 1
        self.e_measures = make_dark_text(self.content, height=3, wrap="word")
        self.e_measures.grid(
            row=row, column=0, sticky="ew", padx=10, pady=(2, 0)
        )
        if len(rec) > 7 and rec[7] and rec[7] != "—":
            self.e_measures.insert("1.0", rec[7])
        row += 1

        tk.Frame(self.content, bg=C["bg_main"], height=16).grid(
            row=row, column=0
        )

    def _toggle_edit_mode(self) -> None:
        self.is_edit_mode = True
        self._build_edit_content()

        self.btn_edit.pack_forget()
        self.btn_save.pack(side="left", padx=(0, 8))
        self.btn_cancel_edit.pack(side="left")

        self.lbl_title.configure(
            text=f"Редагування запису #{self.record[0]}",
            fg=COLORS["accent_warning"],
        )

    def _cancel_edit(self) -> None:
        self.is_edit_mode = False
        self._build_view_content()

        self.btn_save.pack_forget()
        self.btn_cancel_edit.pack_forget()
        self.btn_edit.pack(side="left", padx=(0, 8))

        self.lbl_title.configure(
            text=f"Запис #{self.record[0]}",
            fg=COLORS["accent_muted"],
        )

    def _save_changes(self) -> None:
        detect = self.e_detect.get().strip()
        event_d = self.e_event_date.get().strip()

        for val, label in [
            (detect, "дати виявлення"),
            (event_d, "дати подiї"),
        ]:
            if val in ("дд.мм.рррр", ""):
                continue
            if not is_valid_date(val):
                messagebox.showwarning(
                    "Помилка",
                    f"Неправильний формат {label} (очiкується дд.мм.рррр)",
                    parent=self.win,
                )
                return

        detect = "" if detect == "дд.мм.рррр" else detect
        event_d = "" if event_d == "дд.мм.рррр" else event_d

        entity = self.e_entity.get().strip()
        event = self.e_event.get().strip()

        if not entity or not event:
            messagebox.showwarning(
                "Обов'язковi поля",
                "Заповнiть назву пiдприємства та назву подiї",
                parent=self.win,
            )
            return

        old_id = self.record[0]

        new_record = (
            self.record[0],
            entity,
            event,
            self.e_involved.get("1.0", tk.END).strip(),
            self.e_risk.get().strip() or "—",
            event_d or "—",
            self.e_description.get("1.0", tk.END).strip(),
            self.e_measures.get("1.0", tk.END).strip(),
            detect or "—",
            self.e_priority.get().strip() or "Середнiй",
            self.e_status.get().strip() or "Вiдкрито",
        )

        self.record = list(new_record)
        self.save_callback(str(old_id), new_record)

        self.lbl_subtitle.configure(text=entity)
        status_val = new_record[10]
        status_colors = {
            "Вiдкрито": COLORS["accent_danger"],
            "В обробцi": COLORS["accent_warning"],
            "Закрито": COLORS["text_muted"],
            "Вирiшено": COLORS["accent_success"],
        }
        sc = status_colors.get(status_val, COLORS["text_muted"])
        self.lbl_status_badge.configure(text=f"  {status_val}  ", bg=sc)

        self._cancel_edit()
        self.toast_callback("Запис оновлено")

    def _delete_record(self) -> None:
        idx_str = self.record[0]
        if not messagebox.askyesno(
            "Пiдтвердження",
            f"Видалити запис #{idx_str}?\nЦю дiю не можна скасувати.",
            parent=self.win,
        ):
            return
        self.delete_callback(str(idx_str))
        self.toast_callback("Запис видалено")
        self.win.destroy()


# =============================================================================
#  ВКЛАДКА: РЕЄСТР
# =============================================================================

class RegistryTab:
    """Вкладка 'Реєстр суттєвих подій'."""

    def __init__(
        self,
        parent: tk.Misc,
        on_data_change: Callable[[list[tuple]], None] | None = None,
    ) -> None:
        self.parent = parent
        self.on_data_change = on_data_change
        self.frame = ttk.Frame(parent)
        self.all_records: list[tuple] = []

        self._build_ui()
        self._load_data()

    def _build_ui(self) -> None:
        C = COLORS
        self.frame.rowconfigure(1, weight=1)
        self.frame.columnconfigure(0, weight=1)

        # Header
        header = tk.Frame(self.frame, bg=C["bg_header"], height=56)
        header.grid(row=0, column=0, sticky="ew")
        header.columnconfigure(1, weight=1)
        header.grid_propagate(False)

        tk.Label(
            header,
            text="РЕЄСТР СУТТЄВИХ ПОДIЙ",
            bg=C["bg_header"],
            fg=COLORS["accent_muted"],
            font=("Arial", 13, "bold"),
        ).grid(row=0, column=0, padx=20, pady=14, sticky="w")

        search_frame = tk.Frame(header, bg=C["bg_header"])
        search_frame.grid(row=0, column=1, sticky="e", padx=20)

        tk.Label(
            search_frame,
            text="Пошук:",
            bg=C["bg_header"],
            fg=C["text_muted"],
            font=("Arial", 8),
        ).pack(side="left", padx=(0, 6))

        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *_: self._apply_filter())

        ent_search = tk.Entry(
            search_frame,
            textvariable=self.search_var,
            bg=C["bg_input"],
            fg=C["text_primary"],
            insertbackground=C["text_primary"],
            relief="flat",
            bd=2,
            font=("Arial", 9),
            width=34,
        )
        ent_search.pack(side="left", padx=(0, 8), ipady=2)

        tk.Button(
            search_frame,
            text="Скинути",
            bg=C["bg_surface"],
            fg=C["text_muted"],
            activebackground=C["bg_surface_alt"],
            activeforeground=C["text_primary"],
            relief="flat",
            bd=0,
            font=("Arial", 8),
            cursor="hand2",
            padx=8,
            pady=2,
            command=self._reset_filter,
        ).pack(side="left")

        # Основна зона
        paned = ttk.PanedWindow(self.frame, orient="horizontal")
        paned.grid(row=1, column=0, sticky="nsew")

        left_wrap = ttk.Frame(paned)
        right_wrap = ttk.Frame(paned)
        paned.add(left_wrap, weight=4)
        paned.add(right_wrap, weight=7)

        self._build_form(left_wrap)
        self._build_table(right_wrap)

    # --- Форма ----------------------------------------------------------------

    def _build_form(self, container: tk.Misc) -> None:
        C = COLORS
        container.columnconfigure(0, weight=1)
        container.rowconfigure(0, weight=1)

        canvas = tk.Canvas(container, bg=C["bg_main"], highlightthickness=0)
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")
        container.columnconfigure(0, weight=1)
        container.rowconfigure(0, weight=1)

        form = tk.Frame(canvas, bg=C["bg_main"])
        form_window = canvas.create_window((0, 0), window=form, anchor="nw")

        def on_configure(_: object) -> None:
            canvas.configure(scrollregion=canvas.bbox("all"))
            canvas.itemconfig(form_window, width=canvas.winfo_width())

        form.bind("<Configure>", on_configure)
        canvas.bind("<Configure>", on_configure)

        def on_mousewheel(e: tk.Event) -> None:  # type: ignore[override]
            canvas.yview_scroll(int(-1 * (e.delta / 120)), "units")

        canvas.bind_all("<MouseWheel>", on_mousewheel)

        form.columnconfigure(0, weight=1)

        def section(txt: str, row_idx: int) -> int:
            f = tk.Frame(form, bg=C["bg_main"])
            f.grid(row=row_idx, column=0, sticky="ew", padx=16, pady=(16, 4))
            tk.Frame(f, bg=COLORS["accent"], width=2, height=16).pack(side="left")
            tk.Label(
                f,
                text=txt,
                bg=C["bg_main"],
                fg=COLORS["accent"],
                font=("Arial", 9, "bold"),
            ).pack(side="left", padx=8)
            return row_idx + 1

        def field(
            lbl_text: str, row_idx: int, widget_factory: Callable[[tk.Misc], tk.Widget]
        ) -> tuple[tk.Widget, int]:
            tk.Label(
                form,
                text=lbl_text,
                bg=C["bg_main"],
                fg=C["text_subtle"],
                font=("Arial", 8),
            ).grid(row=row_idx, column=0, sticky="w", padx=16, pady=(4, 0))
            w = widget_factory(form)
            w.grid(row=row_idx + 1, column=0, sticky="ew", padx=16, pady=(2, 0))
            return w, row_idx + 2

        def make_entry(parent: tk.Misc, **kw: object) -> tk.Entry:
            return tk.Entry(
                parent,
                bg=C["bg_input"],
                fg=C["text_primary"],
                insertbackground=C["text_primary"],
                relief="flat",
                bd=2,
                highlightthickness=1,
                highlightbackground=C["border_soft"],
                highlightcolor=C["accent"],
                font=("Arial", 9),
                **kw,
            )

        def make_combo(
            parent: tk.Misc, values: list[str] | None = None, **kw: object
        ) -> ttk.Combobox:
            return ttk.Combobox(
                parent,
                values=values or [],
                state="readonly",
                font=("Arial", 9),
                **kw,
            )

        row = 0

        # badge "новий запис"
        new_badge = tk.Frame(form, bg=C["bg_main"])
        new_badge.grid(row=row, column=0, sticky="ew", padx=16, pady=(12, 0))
        tk.Label(
            new_badge,
            text="  + НОВИЙ ЗАПИС  ",
            bg=COLORS["accent"],
            fg="white",
            font=("Arial", 8, "bold"),
            pady=3,
        ).pack(anchor="w")
        row += 1

        row = section("Iнформацiя про пiдприємство та особу", row)

        self.ent_entity, row = field(
            "Скорочена назва пiдприємства:", row, make_entry
        )
        self.ent_position, row = field("Посада:", row, make_entry)
        self.ent_reporter, row = field("ПIБ особи, що звiтує:", row, make_entry)

        row = section("Опис подiї / ризику", row)

        self.cb_event, row = field(
            "Назва подiї:", row, lambda p: make_combo(p, values=EVENT_TYPES)
        )
        self.cb_risk, row = field(
            "Тип ризику:", row, lambda p: make_combo(p, values=RISK_TYPES)
        )

        tk.Label(
            form,
            text="Задiянi пiдроздiли / особи:",
            bg=C["bg_main"],
            fg=C["text_subtle"],
            font=("Arial", 8),
        ).grid(row=row, column=0, sticky="w", padx=16, pady=(4, 0))
        row += 1
        self.txt_involved = make_dark_text(form, height=2, wrap="word")
        self.txt_involved.grid(row=row, column=0, sticky="ew", padx=16, pady=(2, 0))
        row += 1

        row = section("Фiнансовий вплив (млн грн)", row)

        fin_frame = tk.Frame(form, bg=C["bg_main"])
        fin_frame.grid(row=row, column=0, sticky="ew", padx=16, pady=4)
        fin_frame.columnconfigure((0, 1, 2, 3), weight=1)
        row += 1

        for col, title in enumerate(
            ["Втрати", "Резерв", "Заплановані втрати", "Відшкодування"]
        ):
            tk.Label(
                fin_frame,
                text=title,
                bg=C["bg_main"],
                fg=C["text_muted"],
                font=("Arial", 7),
            ).grid(row=0, column=col, sticky="w", padx=4)

        self.ent_loss = make_entry(fin_frame, width=11)
        self.ent_reserve = make_entry(fin_frame, width=11)
        self.ent_planned = make_entry(fin_frame, width=11)
        self.ent_refund = make_entry(fin_frame, width=11)

        for col, e in enumerate(
            [self.ent_loss, self.ent_reserve, self.ent_planned, self.ent_refund]
        ):
            e.grid(row=1, column=col, padx=4, pady=2, sticky="ew")

        net_frame = tk.Frame(form, bg=C["bg_main"])
        net_frame.grid(row=row, column=0, sticky="ew", padx=16, pady=(2, 6))
        row += 1
        tk.Label(
            net_frame,
            text="Чистий вплив (млн грн):",
            bg=C["bg_main"],
            fg=C["text_muted"],
            font=("Arial", 8),
        ).pack(side="left")
        self.lbl_net = tk.Label(
            net_frame,
            text="0.00",
            bg=C["bg_main"],
            fg=COLORS["accent_success"],
            font=("Arial", 13, "bold"),
        )
        self.lbl_net.pack(side="left", padx=10)

        def update_net(_: object) -> None:
            try:
                loss = float(self.ent_loss.get().replace(",", ".") or 0)
                refund = float(self.ent_refund.get().replace(",", ".") or 0)
                net = loss - refund
                col = COLORS["accent_danger"] if net > 0 else COLORS["accent_success"]
                self.lbl_net.configure(text=f"{net:,.2f}", fg=col)
            except Exception:
                self.lbl_net.configure(text="—", fg=COLORS["text_muted"])

        for e in [self.ent_loss, self.ent_refund]:
            e.bind("<KeyRelease>", update_net)

        row = section("Деталi подiї", row)

        text_fields: list[tuple[str, str, int]] = [
            ("Вплив на iншi пiдприємства:", "txt_impact", 2),
            ("Нефiнансовий / якiсний вплив:", "txt_qualitative", 2),
            ("Детальний опис подiї:", "txt_description", 4),
            ("Вжитi заходи:", "txt_measures", 3),
        ]
        for lbl_text, attr, height in text_fields:
            tk.Label(
                form,
                text=lbl_text,
                bg=C["bg_main"],
                fg=C["text_subtle"],
                font=("Arial", 8),
            ).grid(row=row, column=0, sticky="w", padx=16, pady=(6, 0))
            row += 1
            t = make_dark_text(form, height=height, wrap="word")
            t.grid(row=row, column=0, sticky="ew", padx=16, pady=(2, 0))
            setattr(self, attr, t)
            row += 1

        row = section("Дати", row)
        date_frame = tk.Frame(form, bg=C["bg_main"])
        date_frame.grid(row=row, column=0, sticky="w", padx=16, pady=4)
        row += 1

        for col, (lbl_text, attr) in enumerate(
            [("Дата виявлення:", "ent_detect"), ("Дата подiї:", "ent_event_date")]
        ):
            tk.Label(
                date_frame,
                text=lbl_text,
                bg=C["bg_main"],
                fg=C["text_subtle"],
                font=("Arial", 8),
            ).grid(
                row=0,
                column=col,
                padx=(0 if col == 0 else 20, 0),
                sticky="w",
            )
            e = make_entry(date_frame, width=14)
            e.grid(
                row=1,
                column=col,
                padx=(0 if col == 0 else 20, 0),
                pady=2,
            )
            add_placeholder(e, "дд.мм.рррр")
            setattr(self, attr, e)

        tk.Label(
            form,
            text="Прiоритет:",
            bg=C["bg_main"],
            fg=C["text_subtle"],
            font=("Arial", 8),
        ).grid(row=row, column=0, sticky="w", padx=16, pady=(8, 0))
        row += 1
        self.cb_priority = make_combo(
            form, values=["Критичний", "Високий", "Середнiй", "Низький"]
        )
        self.cb_priority.grid(row=row, column=0, sticky="w", padx=16, pady=(2, 0))
        row += 1

        tk.Label(
            form,
            text="Статус:",
            bg=C["bg_main"],
            fg=C["text_subtle"],
            font=("Arial", 8),
        ).grid(row=row, column=0, sticky="w", padx=16, pady=(8, 0))
        row += 1
        self.cb_status = make_combo(
            form, values=["Вiдкрито", "В обробцi", "Закрито", "Вирiшено"]
        )
        self.cb_status.grid(row=row, column=0, sticky="w", padx=16, pady=(2, 0))
        row += 1

        btn_frame = tk.Frame(form, bg=C["bg_main"])
        btn_frame.grid(row=row, column=0, sticky="ew", padx=16, pady=(16, 20))
        row += 1
        btn_frame.columnconfigure((0, 1), weight=1)

        tk.Button(
            btn_frame,
            text="Очистити",
            bg=C["bg_surface"],
            fg=C["text_muted"],
            activebackground=C["bg_surface_alt"],
            activeforeground=C["text_primary"],
            relief="flat",
            bd=0,
            cursor="hand2",
            font=("Arial", 9),
            padx=14,
            pady=6,
            command=self._clear_form,
        ).grid(row=0, column=0, padx=4, sticky="ew")

        tk.Button(
            btn_frame,
            text="Додати запис",
            bg=COLORS["accent"],
            fg="white",
            activebackground=COLORS["accent_soft"],
            activeforeground="white",
            relief="flat",
            bd=0,
            cursor="hand2",
            font=("Arial", 9, "bold"),
            padx=14,
            pady=6,
            command=self._add_record,
        ).grid(row=0, column=1, padx=4, sticky="ew")

    # --- Таблиця -------------------------------------------------------------

    def _build_table(self, container: tk.Misc) -> None:
        C = COLORS
        container.rowconfigure(1, weight=1)
        container.columnconfigure(0, weight=1)

        toolbar = tk.Frame(container, bg=C["bg_surface"], height=40)
        toolbar.grid(row=0, column=0, sticky="ew")
        toolbar.grid_propagate(False)

        tk.Label(
            toolbar,
            text="Записи",
            bg=C["bg_surface"],
            fg=C["text_muted"],
            font=("Arial", 8),
        ).pack(side="left", padx=12, pady=8)

        self.lbl_count = tk.Label(
            toolbar,
            text="0",
            bg=C["bg_surface"],
            fg=COLORS["accent"],
            font=("Arial", 8, "bold"),
        )
        self.lbl_count.pack(side="left", pady=8)

        tk.Label(
            toolbar,
            text="  |  Ризик:",
            bg=C["bg_surface"],
            fg=C["text_muted"],
            font=("Arial", 8),
        ).pack(side="left", pady=8)
        self.filter_risk = ttk.Combobox(
            toolbar,
            width=16,
            state="readonly",
            values=["Всi"] + RISK_TYPES,
            font=("Arial", 8),
        )
        self.filter_risk.set("Всi")
        self.filter_risk.pack(side="left", padx=6, pady=8)
        self.filter_risk.bind("<<ComboboxSelected>>", lambda _: self._apply_filter())

        tk.Label(
            toolbar,
            text="Статус:",
            bg=C["bg_surface"],
            fg=C["text_muted"],
            font=("Arial", 8),
        ).pack(side="left", pady=8)
        self.filter_status = ttk.Combobox(
            toolbar,
            width=12,
            state="readonly",
            values=["Всi", "Вiдкрито", "В обробцi", "Закрито", "Вирiшено"],
            font=("Arial", 8),
        )
        self.filter_status.set("Всi")
        self.filter_status.pack(side="left", padx=6, pady=8)
        self.filter_status.bind(
            "<<ComboboxSelected>>", lambda _: self._apply_filter()
        )

        for txt, cmd, bg in [
            ("Переглянути", self._open_selected_detail, COLORS["accent"]),
            ("Дублювати", self._duplicate_record, C["bg_surface_alt"]),
            ("Видалити", self._delete_selected, COLORS["accent_danger"]),
        ]:
            tk.Button(
                toolbar,
                text=txt,
                bg=bg,
                fg="white" if bg != C["bg_surface_alt"] else C["text_primary"],
                activebackground=bg,
                activeforeground="white",
                relief="flat",
                bd=0,
                cursor="hand2",
                font=("Arial", 8),
                padx=10,
                pady=3,
                command=cmd,
            ).pack(side="right", padx=4, pady=6)

        tree_frame = ttk.Frame(container)
        tree_frame.grid(row=1, column=0, sticky="nsew")
        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)

        cols = (
            "id",
            "entity",
            "event",
            "risk",
            "priority",
            "status",
            "date",
            "involved",
            "desc",
            "measures",
        )
        self.tree = ttk.Treeview(
            tree_frame, columns=cols, show="headings", selectmode="browse"
        )
        self.tree.grid(row=0, column=0, sticky="nsew")

        headers = {
            "id": ("№", 46),
            "entity": ("Пiдприємство", 155),
            "event": ("Назва подiї", 185),
            "risk": ("Тип ризику", 105),
            "priority": ("Прiоритет", 90),
            "status": ("Статус", 90),
            "date": ("Дата подiї", 90),
            "involved": ("Задiянi", 130),
            "desc": ("Опис", 200),
            "measures": ("Заходи", 200),
        }
        for col, (txt, w) in headers.items():
            self.tree.heading(col, text=txt, command=lambda c=col: self._sort_tree(c))
            self.tree.column(col, width=w, anchor="w")

        sy = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        sy.grid(row=0, column=1, sticky="ns")
        sx = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        sx.grid(row=1, column=0, sticky="ew")
        self.tree.configure(yscrollcommand=sy.set, xscrollcommand=sx.set)

        self.tree.tag_configure("even", background=C["row_even"])
        self.tree.tag_configure("odd", background=C["row_odd"])

        for risk, color in RISK_COLORS.items():
            self.tree.tag_configure(f"risk_{risk}", foreground=color)

        self.tree.bind("<<TreeviewSelect>>", self._on_select)
        self.tree.bind("<Double-1>", self._on_double_click)

        hint_frame = tk.Frame(container, bg=C["bg_surface"])
        hint_frame.grid(row=2, column=0, sticky="ew")
        tk.Label(
            hint_frame,
            text="  Подвiйний клiк по рядку — переглянути / редагувати запис",
            bg=C["bg_surface"],
            fg=C["text_muted"],
            font=("Arial", 7, "italic"),
        ).pack(side="left", padx=8, pady=4)

        detail_frame = tk.Frame(container, bg=C["bg_surface"])
        detail_frame.grid(row=3, column=0, sticky="ew")
        detail_frame.columnconfigure((0, 1), weight=1)

        for col, (lbl, attr) in enumerate(
            [("Опис подiї", "det_desc"), ("Вжитi заходи", "det_measures")]
        ):
            sub = tk.Frame(detail_frame, bg=C["bg_surface"])
            sub.grid(
                row=0,
                column=col,
                sticky="nsew",
                padx=(12 if col == 0 else 4, 4),
                pady=8,
            )
            sub.columnconfigure(0, weight=1)

            tk.Label(
                sub,
                text=lbl,
                bg=C["bg_surface"],
                fg=C["text_muted"],
                font=("Arial", 7, "bold"),
            ).grid(row=0, column=0, sticky="w")
            t = make_dark_text(sub, height=4, wrap="word", state="disabled")
            t.grid(row=1, column=0, sticky="ew", pady=(2, 0))
            setattr(self, attr, t)

        exp_bar = tk.Frame(container, bg=C["bg_main"])
        exp_bar.grid(row=4, column=0, sticky="ew", padx=8, pady=6)

        tk.Button(
            exp_bar,
            text="Експорт CSV",
            bg=C["bg_surface"],
            fg=C["text_primary"],
            activebackground=C["bg_surface_alt"],
            activeforeground=C["text_primary"],
            relief="flat",
            bd=0,
            cursor="hand2",
            font=("Arial", 8),
            padx=12,
            pady=4,
            command=self._export_csv,
        ).pack(side="left", padx=(0, 6))

        if pd:
            tk.Button(
                exp_bar,
                text="Експорт Excel",
                bg=COLORS["accent_success"],
                fg="white",
                activebackground="#16a34a",
                activeforeground="white",
                relief="flat",
                bd=0,
                cursor="hand2",
                font=("Arial", 8),
                padx=12,
                pady=4,
                command=self._export_excel,
            ).pack(side="left", padx=(0, 6))

        tk.Button(
            exp_bar,
            text="Iмпорт JSON",
            bg=C["bg_surface"],
            fg=C["text_primary"],
            activebackground=C["bg_surface_alt"],
            activeforeground=C["text_primary"],
            relief="flat",
            bd=0,
            cursor="hand2",
            font=("Arial", 8),
            padx=12,
            pady=4,
            command=self._import_json,
        ).pack(side="left")

    # --- Подвійний клік / деталі --------------------------------------------

    def _on_double_click(self, event: tk.Event) -> None:  # type: ignore[override]
        sel = self.tree.selection()
        if not sel:
            return
        region = self.tree.identify("region", event.x, event.y)
        if region != "cell":
            return
        self._open_selected_detail()

    def _open_selected_detail(self) -> None:
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Перегляд", "Оберiть запис для перегляду")
            return
        iid = sel[0]
        idx_str = self.tree.set(iid, "id")
        rec = self._find_record(idx_str)
        if not rec:
            return

        EventDetailWindow(
            parent_root=self.frame.winfo_toplevel(),
            record=rec,
            all_records=self.all_records,
            save_callback=lambda old_id, new_rec: self._on_detail_save(
                iid, old_id, new_rec
            ),
            delete_callback=lambda idx_s: self._on_detail_delete(iid, idx_s),
            toast_callback=self._show_toast,
        )

    def _find_record(self, idx_str: str) -> tuple | None:
        for r in self.all_records:
            if str(r[0]) == idx_str or str(r[0]).lstrip("0") == str(idx_str).lstrip("0"):
                return r
        return None

    def _on_detail_save(self, iid: str, old_id: str, new_record: tuple) -> None:
        for i, r in enumerate(self.all_records):
            if str(r[0]) == str(old_id) or str(r[0]).lstrip("0") == str(old_id).lstrip(
                "0"
            ):
                self.all_records[i] = new_record
                break

        try:
            self.tree.item(
                iid,
                values=(
                    new_record[0],
                    new_record[1],
                    new_record[2],
                    new_record[4],
                    new_record[9],
                    new_record[10],
                    new_record[5],
                    new_record[3],
                    new_record[6],
                    new_record[7],
                ),
            )
        except tk.TclError:
            pass
        self._recolor_rows()
        self._save_data()
        if self.on_data_change:
            self.on_data_change(self.all_records)

    def _on_detail_delete(self, iid: str, idx_str: str) -> None:
        try:
            self.tree.delete(iid)
        except tk.TclError:
            pass
        self.all_records = [
            r for r in self.all_records if str(r[0]) != str(idx_str)
        ]
        self._recolor_rows()
        self._save_data()
        self._update_count()
        if self.on_data_change:
            self.on_data_change(self.all_records)

    # --- Сортування / перефарбовка -------------------------------------------

    def _sort_tree(self, col: str) -> None:
        data = [(self.tree.set(iid, col), iid) for iid in self.tree.get_children("")]
        try:
            data.sort(
                key=lambda x: float(x[0].replace(",", ""))
                if x[0] not in ("—", "")
                else 0
            )
        except ValueError:
            data.sort(key=lambda x: x[0].lower())
        for i, (_, iid) in enumerate(data):
            self.tree.move(iid, "", i)
        self._recolor_rows()

    def _recolor_rows(self) -> None:
        for i, iid in enumerate(self.tree.get_children()):
            risk = self.tree.set(iid, "risk")
            base_tag = "even" if i % 2 == 0 else "odd"
            tags = [base_tag]
            if risk in RISK_COLORS:
                tags.append(f"risk_{risk}")
            self.tree.item(iid, tags=tags)

    # --- Завантаження / збереження -------------------------------------------

    def _load_data(self) -> None:
        self.all_records.clear()
        self.tree.delete(*self.tree.get_children())
        if not os.path.exists(DATA_FILE):
            return
        try:
            with open(DATA_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            for row in data:
                if isinstance(row, (list, tuple)) and len(row) >= 8:
                    if len(row) == 8:
                        row = list(row) + ["—", "Середнiй", "Вiдкрито"]
                    self.all_records.append(tuple(row))
                    self._insert_tree_row(tuple(row))
        except Exception as e:
            messagebox.showerror("Помилка завантаження", str(e))
        self._update_count()
        if self.on_data_change:
            self.on_data_change(self.all_records)

    def _save_data(self) -> None:
        try:
            with open(DATA_FILE, "w", encoding="utf-8") as f:
                json.dump(self.all_records, f, ensure_ascii=False, indent=2)
        except Exception as e:
            messagebox.showerror("Помилка збереження", str(e))

    def _insert_tree_row(self, data: tuple) -> str:
        if len(data) < 11:
            data = tuple(list(data) + ["—"] * (11 - len(data)))
        iid = self.tree.insert(
            "",
            tk.END,
            values=(
                data[0],
                data[1],
                data[2],
                data[4],
                data[9],
                data[10],
                data[5],
                data[3],
                data[6],
                data[7],
            ),
        )
        self._recolor_rows()
        return iid

    # --- Форма ----------------------------------------------------------------

    def _get_form_data(self) -> tuple | None:
        detect = self.ent_detect.get().strip()
        event_d = self.ent_event_date.get().strip()

        for val, label in [
            (detect, "дати виявлення"),
            (event_d, "дати подiї"),
        ]:
            if val in ("дд.мм.рррр", ""):
                continue
            if not is_valid_date(val):
                messagebox.showwarning(
                    "Помилка",
                    f"Неправильний формат {label} (очiкується дд.мм.рррр)",
                )
                return None

        detect = "" if detect == "дд.мм.рррр" else detect
        event_d = "" if event_d == "дд.мм.рррр" else event_d

        idx = len(self.all_records) + 1
        return (
            f"{idx:03d}",
            self.ent_entity.get().strip(),
            self.cb_event.get().strip(),
            self.txt_involved.get("1.0", tk.END).strip(),
            self.cb_risk.get().strip() or "—",
            event_d or "—",
            self.txt_description.get("1.0", tk.END).strip(),
            self.txt_measures.get("1.0", tk.END).strip(),
            detect or "—",
            self.cb_priority.get().strip() or "Середнiй",
            self.cb_status.get().strip() or "Вiдкрито",
        )

    def _clear_form(self, silent: bool = False) -> None:  # noqa: ARG002
        for w in [
            self.ent_entity,
            self.ent_position,
            self.ent_reporter,
            self.ent_loss,
            self.ent_reserve,
            self.ent_planned,
            self.ent_refund,
        ]:
            w.delete(0, tk.END)
            w.configure(fg=COLORS["text_primary"])
        for w in [self.cb_event, self.cb_risk, self.cb_priority, self.cb_status]:
            w.set("")
        for w in [
            self.txt_involved,
            self.txt_impact,
            self.txt_qualitative,
            self.txt_description,
            self.txt_measures,
        ]:
            w.delete("1.0", tk.END)
        self.lbl_net.configure(text="0.00", fg=COLORS["accent_success"])

        for e, ph in [(self.ent_detect, "дд.мм.рррр"), (self.ent_event_date, "дд.мм.рррр")]:
            e.delete(0, tk.END)
            add_placeholder(e, ph)

    def _add_record(self) -> None:
        data = self._get_form_data()
        if not data:
            return
        if not data[1] or not data[2]:
            messagebox.showwarning(
                "Обов'язковi поля", "Заповнiть назву пiдприємства та назву подiї"
            )
            return

        self.all_records.append(data)
        self._insert_tree_row(data)
        self._clear_form()
        self._save_data()
        if self.on_data_change:
            self.on_data_change(self.all_records)
        self._update_count()
        self._show_toast("Запис додано")

    # --- Видалення / дублювання ----------------------------------------------

    def _delete_selected(self) -> None:
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Видалення", "Оберiть запис для видалення")
            return
        iid = sel[0]
        idx_str = self.tree.set(iid, "id")
        if not messagebox.askyesno(
            "Пiдтвердження",
            f"Видалити запис #{idx_str}?\nЦю дiю не можна скасувати.",
        ):
            return
        self.tree.delete(iid)
        self.all_records = [r for r in self.all_records if str(r[0]) != idx_str]
        self._recolor_rows()
        self._save_data()
        self._update_count()
        if self.on_data_change:
            self.on_data_change(self.all_records)
        self._show_toast("Запис видалено")

    def _duplicate_record(self) -> None:
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Дублювання", "Оберiть запис для дублювання")
            return
        idx_str = self.tree.set(sel[0], "id")
        rec = self._find_record(idx_str)
        if not rec:
            return
        new_idx = f"{len(self.all_records) + 1:03d}"
        new_rec = (new_idx,) + tuple(rec[1:])
        self.all_records.append(new_rec)
        self._insert_tree_row(new_rec)
        self._save_data()
        self._update_count()
        if self.on_data_change:
            self.on_data_change(self.all_records)
        self._show_toast("Запис продубльовано")

    # --- Фільтри / пошук -----------------------------------------------------

    def _apply_filter(self) -> None:
        q = self.search_var.get().strip().lower()
        risk = self.filter_risk.get()
        status = self.filter_status.get()

        self.tree.delete(*self.tree.get_children())
        for row in self.all_records:
            row_str = " ".join(str(v).lower() for v in row)
            if q and q not in row_str:
                continue
            if risk != "Всi" and row[4] != risk:
                continue
            if status != "Всi" and (len(row) <= 10 or row[10] != status):
                continue
            self._insert_tree_row(row)
        self._update_count()

    def _reset_filter(self) -> None:
        self.search_var.set("")
        self.filter_risk.set("Всi")
        self.filter_status.set("Всi")

        self.tree.delete(*self.tree.get_children())
        for row in self.all_records:
            self._insert_tree_row(row)
        self._update_count()

    # --- Вибір рядка / швидкий перегляд --------------------------------------

    def _on_select(self, _: object | None = None) -> None:
        sel = self.tree.selection()
        if not sel:
            return
        desc = self.tree.set(sel[0], "desc")
        measures = self.tree.set(sel[0], "measures")
        for widget, text in [(self.det_desc, desc), (self.det_measures, measures)]:
            widget.configure(state="normal")
            widget.delete("1.0", tk.END)
            widget.insert("1.0", text)
            widget.configure(state="disabled")

    # --- Лічильник -----------------------------------------------------------

    def _update_count(self) -> None:
        n = len(self.tree.get_children())
        self.lbl_count.configure(text=f" {n}")

    # --- Toast ----------------------------------------------------------------

    def _show_toast(self, msg: str) -> None:
        C = COLORS
        toast = tk.Toplevel(self.frame)
        toast.overrideredirect(True)
        toast.attributes("-topmost", True)
        toast.configure(bg=COLORS["accent_success"])

        tk.Label(
            toast,
            text=f"  {msg}  ",
            bg=COLORS["accent_success"],
            fg="white",
            font=("Arial", 9, "bold"),
            pady=6,
        ).pack()

        root = self.frame.winfo_toplevel()
        x = root.winfo_x() + root.winfo_width() - 220
        y = root.winfo_y() + root.winfo_height() - 80
        toast.geometry(f"+{x}+{y}")
        toast.after(2000, toast.destroy)

    # --- Експорт / імпорт -----------------------------------------------------

    def _export_csv(self) -> None:
        if not self.tree.get_children():
            messagebox.showinfo("Експорт", "Таблиця порожня")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV файли", "*.csv")],
            title="Зберегти як CSV",
        )
        if not path:
            return
        try:
            headers = [
                "ID",
                "Пiдприємство",
                "Назва подiї",
                "Задiянi особи",
                "Тип ризику",
                "Дата подiї",
                "Опис",
                "Заходи",
                "Дата виявлення",
                "Прiоритет",
                "Статус",
            ]
            with open(path, "w", newline="", encoding="utf-8-sig") as f:
                w = csv.writer(f)
                w.writerow(headers)
                for row in self.all_records:
                    w.writerow(row)
            self._show_toast("CSV збережено")
        except Exception as e:
            messagebox.showerror("Помилка", str(e))

    def _export_excel(self) -> None:
        if not pd:
            messagebox.showwarning("Excel", "Встановiть pandas та openpyxl")
            return
        if not self.all_records:
            messagebox.showinfo("Експорт", "Немає записiв")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel файли", "*.xlsx")],
            title="Зберегти як Excel",
        )
        if not path:
            return
        try:
            headers = [
                "ID",
                "Пiдприємство",
                "Назва подiї",
                "Задiянi особи",
                "Тип ризику",
                "Дата подiї",
                "Опис",
                "Заходи",
                "Дата виявлення",
                "Прiоритет",
                "Статус",
            ]
            data = [
                r if len(r) == 11 else list(r) + [""] * (11 - len(r))
                for r in self.all_records
            ]
            df = pd.DataFrame(data, columns=headers)  # type: ignore[call-arg]
            with pd.ExcelWriter(path, engine="openpyxl") as writer:  # type: ignore[call-arg]
                df.to_excel(writer, index=False, sheet_name="Реєстр")
                ws = writer.sheets["Реєстр"]
                for col_cells in ws.columns:
                    max_len = max(len(str(c.value or "")) for c in col_cells)
                    ws.column_dimensions[col_cells[0].column_letter].width = min(
                        max_len + 4, 60
                    )
            self._show_toast("Excel збережено")
        except Exception as e:
            messagebox.showerror("Помилка", str(e))

    def _import_json(self) -> None:
        path = filedialog.askopenfilename(
            filetypes=[("JSON файли", "*.json")],
            title="Iмпорт JSON",
        )
        if not path:
            return
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            if not isinstance(data, list):
                raise ValueError("Файл повинен мiстити список записiв")
            added = 0
            for row in data:
                if isinstance(row, (list, tuple)):
                    if len(row) == 8:
                        row = list(row) + ["—", "Середнiй", "Вiдкрито"]
                    row[0] = f"{len(self.all_records) + 1:03d}"
                    self.all_records.append(tuple(row))
                    self._insert_tree_row(tuple(row))
                    added += 1
            self._save_data()
            self._update_count()
            if self.on_data_change:
                self.on_data_change(self.all_records)
            self._show_toast(f"Iмпортовано: {added} записiв")
        except Exception as e:
            messagebox.showerror("Помилка iмпорту", str(e))

    def get_frame(self) -> ttk.Frame:
        return self.frame

# =============================================================================
#  ВКЛАДКА: АНАЛІТИКА
# =============================================================================

class AnalyticsTab:
    """Аналітика по записах реєстру (діаграми + таблиця)."""

    def __init__(self, parent: tk.Misc) -> None:
        self.frame = ttk.Frame(parent)
        self.records: list[tuple] = []
        self._build_ui()

    def _build_ui(self) -> None:
        C = COLORS
        self.frame.rowconfigure(1, weight=1)
        self.frame.columnconfigure(0, weight=1)

        header = tk.Frame(self.frame, bg=C["bg_header"], height=56)
        header.grid(row=0, column=0, sticky="ew")
        header.grid_propagate(False)

        tk.Label(
            header,
            text="АНАЛIТИКА ТА ЗВIТИ",
            bg=C["bg_header"],
            fg=COLORS["accent_muted"],
            font=("Arial", 13, "bold"),
        ).pack(side="left", padx=20, pady=14)

        tk.Button(
            header,
            text="Оновити",
            bg=COLORS["accent"],
            fg="white",
            activebackground=COLORS["accent_soft"],
            activeforeground="white",
            relief="flat",
            bd=0,
            cursor="hand2",
            font=("Arial", 9, "bold"),
            padx=12,
            pady=4,
            command=self.refresh,
        ).pack(side="right", padx=20, pady=12)

        canvas = tk.Canvas(self.frame, bg=C["bg_main"], highlightthickness=0)
        sb = ttk.Scrollbar(self.frame, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=sb.set)
        canvas.grid(row=1, column=0, sticky="nsew")
        sb.grid(row=1, column=1, sticky="ns")

        self.content = tk.Frame(canvas, bg=C["bg_main"])
        self._cw = canvas.create_window((0, 0), window=self.content, anchor="nw")

        def on_conf(_: object) -> None:
            canvas.configure(scrollregion=canvas.bbox("all"))
            canvas.itemconfig(self._cw, width=canvas.winfo_width())

        self.content.bind("<Configure>", on_conf)
        canvas.bind("<Configure>", on_conf)
        canvas.bind_all(
            "<MouseWheel>",
            lambda e: canvas.yview_scroll(int(-1 * e.delta / 120), "units"),
        )

        self.content.columnconfigure(0, weight=1)
        self.canvas_ref = canvas

        self._build_stat_cards()
        self._build_charts_and_table()

    def _build_stat_cards(self) -> None:
        C = COLORS
        cards_frame = tk.Frame(self.content, bg=C["bg_main"])
        cards_frame.grid(row=0, column=0, sticky="ew", padx=16, pady=(16, 0))
        for i in range(4):
            cards_frame.columnconfigure(i, weight=1)

        self.stat_cards: dict[str, tk.Label] = {}
        defs = [
            ("total", "Всього записiв", "0", C["accent"]),
            ("open", "Вiдкрито / в обробцi", "0", C["accent_danger"]),
            ("critical", "Критичних", "0", C["accent_warning"]),
            ("closed", "Закрито / вирiшено", "0", C["accent_success"]),
        ]
        for col, (key, title, val, color) in enumerate(defs):
            card = tk.Frame(cards_frame, bg=C["bg_surface"], padx=18, pady=12)
            card.grid(row=0, column=col, padx=6, sticky="nsew")
            tk.Frame(card, bg=color, height=3).pack(fill="x")
            tk.Label(
                card,
                text=title,
                bg=C["bg_surface"],
                fg=C["text_muted"],
                font=("Arial", 8),
            ).pack(anchor="w", pady=(8, 2))
            lbl = tk.Label(
                card,
                text=val,
                bg=C["bg_surface"],
                fg=color,
                font=("Arial", 22, "bold"),
            )
            lbl.pack(anchor="w")
            self.stat_cards[key] = lbl

    def _build_charts_and_table(self) -> None:
        C = COLORS
        if HAS_MPL:
            charts_row = tk.Frame(self.content, bg=C["bg_main"])
            charts_row.grid(row=1, column=0, sticky="ew", padx=16, pady=16)
            charts_row.columnconfigure((0, 1), weight=1)

            # Left (pie)
            self.fig_left = Figure(figsize=(5, 3.5), dpi=90, facecolor=C["bg_surface"])
            self.ax_left = self.fig_left.add_subplot(111)
            self._style_ax(self.ax_left)
            self.ax_left.set_title(
                "Розподiл за типом ризику", color=C["text_muted"], fontsize=9
            )

            frame_l = tk.Frame(charts_row, bg=C["bg_surface"], padx=8, pady=8)
            frame_l.grid(row=0, column=0, padx=(0, 8), sticky="nsew")
            self.canvas_left = FigureCanvasTkAgg(self.fig_left, master=frame_l)
            self.canvas_left.get_tk_widget().pack(fill="both", expand=True)

            # Right (bar)
            self.fig_right = Figure(
                figsize=(5, 3.5), dpi=90, facecolor=C["bg_surface"]
            )
            self.ax_right = self.fig_right.add_subplot(111)
            self._style_ax(self.ax_right)
            self.ax_right.set_title(
                "Записи за статусом", color=C["text_muted"], fontsize=9
            )

            frame_r = tk.Frame(charts_row, bg=C["bg_surface"], padx=8, pady=8)
            frame_r.grid(row=0, column=1, padx=(8, 0), sticky="nsew")
            self.canvas_right = FigureCanvasTkAgg(self.fig_right, master=frame_r)
            self.canvas_right.get_tk_widget().pack(fill="both", expand=True)

            # Bottom (barh)
            self.fig_bottom = Figure(
                figsize=(10, 3), dpi=90, facecolor=C["bg_surface"]
            )
            self.ax_bottom = self.fig_bottom.add_subplot(111)
            self._style_ax(self.ax_bottom)
            self.ax_bottom.set_title(
                "Топ-5 пiдприємств за кiлькiстю подiй",
                color=C["text_muted"],
                fontsize=9,
            )

            frame_b = tk.Frame(self.content, bg=C["bg_surface"], padx=8, pady=8)
            frame_b.grid(row=2, column=0, padx=16, pady=(0, 16), sticky="ew")
            self.canvas_bottom = FigureCanvasTkAgg(self.fig_bottom, master=frame_b)
            self.canvas_bottom.get_tk_widget().pack(fill="both", expand=True)
        else:
            tk.Label(
                self.content,
                text=(
                    "Встановiть matplotlib для вiдображення графiкiв:\n"
                    "  pip install matplotlib"
                ),
                bg=COLORS["bg_main"],
                fg=COLORS["text_muted"],
                font=("Arial", 10),
            ).grid(row=1, column=0, pady=40)

        # Таблиця статистики
        frame = tk.Frame(self.content, bg=COLORS["bg_surface"], padx=16, pady=12)
        frame.grid(row=3, column=0, padx=16, pady=(0, 16), sticky="ew")
        frame.columnconfigure(0, weight=1)

        tk.Label(
            frame,
            text="Деталiзована статистика за типом ризику",
            bg=COLORS["bg_surface"],
            fg=COLORS["text_muted"],
            font=("Arial", 9, "bold"),
        ).grid(row=0, column=0, sticky="w", pady=(0, 8))

        cols = ("risk", "count", "open", "closed")
        self.stats_tree = ttk.Treeview(
            frame, columns=cols, show="headings", height=7
        )
        for col, hdr, w in [
            ("risk", "Тип ризику", 200),
            ("count", "Всього", 80),
            ("open", "Вiдкрито", 80),
            ("closed", "Закрито", 80),
        ]:
            self.stats_tree.heading(col, text=hdr)
            self.stats_tree.column(col, width=w, anchor="w")
        self.stats_tree.grid(row=1, column=0, sticky="ew")

    def _style_ax(self, ax: plt.Axes) -> None:  # type: ignore[name-defined]
        C = COLORS
        ax.set_facecolor(C["bg_surface"])
        ax.tick_params(colors=C["text_muted"], labelsize=8)
        for spine in ax.spines.values():
            spine.set_edgecolor(C["border_soft"])

    def update_data(self, records: list[tuple]) -> None:
        self.records = records
        self.refresh()

    def refresh(self) -> None:
        if not self.records:
            # якщо немає даних — просто обнулити картки і все
            for key in ["total", "open", "critical", "closed"]:
                self.stat_cards[key].configure(text="0")
            if HAS_MPL:
                self._clear_charts()
            self.stats_tree.delete(*self.stats_tree.get_children())
            return

        C = COLORS
        records = self.records

        total = len(records)
        open_c = sum(
            1
            for r in records
            if len(r) > 10 and r[10] in ("Вiдкрито", "В обробцi")
        )
        critical = sum(
            1 for r in records if len(r) > 9 and r[9] == "Критичний"
        )
        closed = sum(
            1
            for r in records
            if len(r) > 10 and r[10] in ("Закрито", "Вирiшено")
        )

        self.stat_cards["total"].configure(text=str(total))
        self.stat_cards["open"].configure(text=str(open_c))
        self.stat_cards["critical"].configure(text=str(critical))
        self.stat_cards["closed"].configure(text=str(closed))

        if not HAS_MPL:
            return

        risk_counter = Counter(r[4] for r in records if len(r) > 4)
        status_counter = Counter(r[10] for r in records if len(r) > 10)
        entity_counter = Counter(r[1] for r in records if len(r) > 1 and r[1])

        # Pie по ризиках
        self.ax_left.clear()
        self._style_ax(self.ax_left)
        self.ax_left.set_title(
            "Розподiл за типом ризику", color=C["text_muted"], fontsize=9
        )
        if risk_counter:
            labels = list(risk_counter.keys())
            values = list(risk_counter.values())
            colors = [RISK_COLORS.get(l, C["text_muted"]) for l in labels]
            wedges, texts, autotexts = self.ax_left.pie(
                values,
                labels=labels,
                autopct="%1.0f%%",
                colors=colors,
                startangle=90,
                textprops={"color": C["text_muted"], "fontsize": 7},
            )
            for at in autotexts:
                at.set_fontsize(7)
                at.set_color("white")
        else:
            self.ax_left.text(
                0.5,
                0.5,
                "Немає даних",
                transform=self.ax_left.transAxes,
                ha="center",
                va="center",
                color=C["text_muted"],
            )
        self.canvas_left.draw()

        # Bars по статусах
        self.ax_right.clear()
        self._style_ax(self.ax_right)
        self.ax_right.set_title(
            "Записи за статусом", color=C["text_muted"], fontsize=9
        )
        if status_counter:
            s_labels = list(status_counter.keys())
            s_values = list(status_counter.values())
            s_colors = [
                C["accent_danger"],
                C["accent_warning"],
                C["accent_success"],
                C["accent_muted"],
            ][: len(s_labels)]
            bars = self.ax_right.bar(
                s_labels, s_values, color=s_colors, edgecolor="none"
            )
            for bar, val in zip(bars, s_values, strict=False):
                self.ax_right.text(
                    bar.get_x() + bar.get_width() / 2,
                    bar.get_height() + 0.1,
                    str(val),
                    ha="center",
                    va="bottom",
                    color=C["text_muted"],
                    fontsize=8,
                )
            self.ax_right.tick_params(axis="x", labelrotation=15, labelsize=7)
            self.ax_right.set_ylim(0, max(s_values) * 1.2 + 1)
        else:
            self.ax_right.text(
                0.5,
                0.5,
                "Немає даних",
                transform=self.ax_right.transAxes,
                ha="center",
                va="center",
                color=C["text_muted"],
            )
        self.canvas_right.draw()

        # Top-5 підприємств
        self.ax_bottom.clear()
        self._style_ax(self.ax_bottom)
        self.ax_bottom.set_title(
            "Топ-5 пiдприємств за кiлькiстю подiй",
            color=C["text_muted"],
            fontsize=9,
        )
        top5 = entity_counter.most_common(5)
        if top5:
            ent_labels = [e[0][:20] for e in top5]
            ent_values = [e[1] for e in top5]
            bars = self.ax_bottom.barh(
                ent_labels, ent_values, color=C["accent"], edgecolor="none"
            )
            for bar, val in zip(bars, ent_values, strict=False):
                self.ax_bottom.text(
                    bar.get_width() + 0.05,
                    bar.get_y() + bar.get_height() / 2,
                    str(val),
                    ha="left",
                    va="center",
                    color=C["text_muted"],
                    fontsize=8,
                )
            self.ax_bottom.tick_params(
                axis="y", labelsize=8, colors=C["text_primary"]
            )
        else:
            self.ax_bottom.text(
                0.5,
                0.5,
                "Немає даних",
                transform=self.ax_bottom.transAxes,
                ha="center",
                va="center",
                color=C["text_muted"],
            )
        self.canvas_bottom.draw()

        # Таблиця статистики
        self.stats_tree.delete(*self.stats_tree.get_children())
        all_risks = set(RISK_TYPES) | set(r[4] for r in records if len(r) > 4)
        for risk in sorted(all_risks):
            recs = [r for r in records if len(r) > 4 and r[4] == risk]
            cnt = len(recs)
            open_ = sum(
                1
                for r in recs
                if len(r) > 10 and r[10] in ("Вiдкрито", "В обробцi")
            )
            cl = sum(
                1
                for r in recs
                if len(r) > 10 and r[10] in ("Закрито", "Вирiшено")
            )
            if cnt:
                self.stats_tree.insert("", tk.END, values=(risk, cnt, open_, cl))

    def _clear_charts(self) -> None:
        self.ax_left.clear()
        self.ax_right.clear()
        self.ax_bottom.clear()
        self.canvas_left.draw()
        self.canvas_right.draw()
        self.canvas_bottom.draw()

    def get_frame(self) -> ttk.Frame:
        return self.frame


# =============================================================================
#  ВКЛАДКА: НАЛАШТУВАННЯ
# =============================================================================

class SettingsTab:
    """Вкладка налаштувань модуля 'Реєстр суттєвих подій'."""

    def __init__(self, parent: tk.Misc) -> None:
        self.frame = ttk.Frame(parent)
        self._build_ui()

    def _build_ui(self) -> None:
        C = COLORS
        self.frame.columnconfigure(0, weight=1)

        header = tk.Frame(self.frame, bg=C["bg_header"], height=56)
        header.grid(row=0, column=0, sticky="ew")
        header.grid_propagate(False)
        tk.Label(
            header,
            text="НАЛАШТУВАННЯ РЕЄСТРУ",
            bg=C["bg_header"],
            fg=C["text_muted"],
            font=("Arial", 13, "bold"),
        ).pack(side="left", padx=20, pady=14)

        content = tk.Frame(self.frame, bg=C["bg_main"])
        content.grid(row=1, column=0, sticky="nsew", padx=40, pady=30)
        content.columnconfigure(0, weight=1)

        self._row(content, 0, "Файл даних:", DATA_FILE, C)
        self._row(content, 1, "Версiя:", "2.1 — Детальне вiкно записiв", C)
        self._row(
            content,
            2,
            "matplotlib:",
            "встановлено" if HAS_MPL else "не встановлено",
            C,
        )
        self._row(
            content,
            3,
            "pandas:",
            "встановлено" if pd else "не встановлено",
            C,
        )

        tk.Label(
            content,
            text="Встановлення залежностей:",
            bg=C["bg_main"],
            fg=C["text_muted"],
            font=("Arial", 8, "bold"),
        ).grid(row=4, column=0, sticky="w", pady=(24, 6))
        tk.Label(
            content,
            text="  pip install matplotlib pandas openpyxl",
            bg=C["bg_surface"],
            fg=COLORS["accent_muted"],
            font=("Courier", 9),
            padx=12,
            pady=8,
        ).grid(row=5, column=0, sticky="w")

        tk.Label(
            content,
            text="Пiдказки:",
            bg=C["bg_main"],
            fg=C["text_muted"],
            font=("Arial", 8, "bold"),
        ).grid(row=6, column=0, sticky="w", pady=(24, 6))

        hints = [
            "Подвiйний клiк по рядку таблицi — вiдкрити детальне вiкно запису",
            "У детальному вiкнi доступне редагування та видалення запису",
            "Лiва панель призначена виключно для створення нових записiв",
            "Кнопка 'Переглянути' у тулбарi вiдкриває те саме детальне вiкно",
        ]
        for i, hint in enumerate(hints):
            f = tk.Frame(content, bg=C["bg_main"])
            f.grid(row=7 + i, column=0, sticky="w", pady=2)
            tk.Frame(f, bg=COLORS["accent_success"], width=4, height=4).pack(
                side="left", padx=(0, 8)
            )
            tk.Label(
                f,
                text=hint,
                bg=C["bg_main"],
                fg=C["text_subtle"],
                font=("Arial", 8),
            ).pack(side="left")

    def _row(self, parent: tk.Misc, row: int, label: str, value: str, C: dict) -> None:
        f = tk.Frame(parent, bg=C["bg_main"])
        f.grid(row=row, column=0, sticky="ew", pady=4)
        tk.Label(
            f,
            text=label,
            bg=C["bg_main"],
            fg=C["text_muted"],
            font=("Arial", 9),
            width=22,
            anchor="w",
        ).pack(side="left")
        tk.Label(
            f,
            text=value,
            bg=C["bg_main"],
            fg=C["text_primary"],
            font=("Arial", 9),
        ).pack(side="left")

    def get_frame(self) -> ttk.Frame:
        return self.frame


# =============================================================================
#  СТОРІНКА "РЕЄСТР СУТТЄВИХ ПОДІЙ" ДЛЯ ATLAS
# =============================================================================

class MaterialEventsPage(tk.Frame):
    """Сторінка 'Реєстр суттєвих подій' (вкладки + статусбар)."""

    def __init__(self, master: tk.Misc, **kwargs) -> None:
        super().__init__(master, bg=COLORS["bg_main"], **kwargs)

        self.grid_rowconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=0)
        self.grid_columnconfigure(0, weight=1)

        self.notebook = ttk.Notebook(self)
        self.notebook.grid(row=0, column=0, sticky="nsew")

        self.analytics_tab = AnalyticsTab(self.notebook)
        self.registry_tab = RegistryTab(
            self.notebook, on_data_change=self.analytics_tab.update_data
        )
        self.settings_tab = SettingsTab(self.notebook)

        self.notebook.add(self.registry_tab.get_frame(), text="  Реєстр подiй  ")
        self.notebook.add(self.analytics_tab.get_frame(), text="  Аналiтика  ")
        self.notebook.add(self.settings_tab.get_frame(), text="  Налаштування  ")

        statusbar = tk.Frame(self, bg=COLORS["bg_header"], height=22)
        statusbar.grid(row=1, column=0, sticky="ew")
        statusbar.grid_propagate(False)

        self._status_lbl = tk.Label(
            statusbar,
            text="Готово",
            bg=COLORS["bg_header"],
            fg=COLORS["text_muted"],
            font=("Arial", 7),
            padx=10,
        )
        self._status_lbl.pack(side="left", pady=3)

        self._time_lbl = tk.Label(
            statusbar,
            text="",
            bg=COLORS["bg_header"],
            fg=COLORS["text_muted"],
            font=("Arial", 7),
            padx=10,
        )
        self._time_lbl.pack(side="right", pady=3)

        self._start_clock()
        self._schedule_autosave()
        self.after(
            600, lambda: self.analytics_tab.update_data(self.registry_tab.all_records)
        )

    def _start_clock(self) -> None:
        self._time_lbl.configure(
            text=datetime.now().strftime("%d.%m.%Y  %H:%M:%S")
        )
        self.after(1000, self._start_clock)

    def _schedule_autosave(self) -> None:
        try:
            self.registry_tab._save_data()  # noqa: SLF001
            self._status_lbl.configure(
                text=f"Автозбережено о {datetime.now().strftime('%H:%M:%S')}"
            )
        except Exception:
            self._status_lbl.configure(text="Помилка автозбереження")
        self.after(30000, self._schedule_autosave)

    def save_before_exit(self) -> None:
        try:
            self.registry_tab._save_data()  # noqa: SLF001
        except Exception:
            pass


# =============================================================================
#  ATLAS ШЕЛЛ (TOP-BAR + SIDEBAR + CONTENT)
# =============================================================================

PageKey: TypeAlias = Literal[
    "risk_register",
    "material_events",
    "risk_appetite",
    "analytics",
    "reports",
    "risk_coordinators",
    "settings",
]

APP_TITLE = "ATLAS"
COPYRIGHT_TEXT = "© 2026 Chugaister8"


class AtlasApp(tk.Tk):
    """Головне вікно ATLAS."""

    def __init__(self, user_full_name: str) -> None:
        super().__init__()

        self.title("ATLAS | Risk Management System")
        self.geometry("1300x800")
        self.minsize(1100, 650)
        self.configure(bg=COLORS["bg_main"])

        apply_dark_style(self)
        self.option_add("*Font", "Arial 9")

        self._user_full_name = user_full_name
        self._material_events_page: MaterialEventsPage | None = None

        self._build_layout()

        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _build_layout(self) -> None:
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(1, weight=1)

        self._build_topbar()
        self._build_sidebar()
        self._build_content()

    def _build_topbar(self) -> None:
        C = COLORS

        topbar = tk.Frame(self, bg=C["bg_header"], height=56)
        topbar.grid(row=0, column=0, columnspan=2, sticky="ew")
        topbar.grid_propagate(False)
        topbar.columnconfigure(1, weight=1)

        # Назва ATLAS
        tk.Label(
            topbar,
            text=APP_TITLE,
            bg=C["bg_header"],
            fg=C["text_primary"],
            font=("Arial", 14, "bold"),
        ).grid(row=0, column=0, padx=20, pady=8, sticky="w")

        # ПІБ
        tk.Label(
            topbar,
            text=self._user_full_name,
            bg=C["bg_header"],
            fg=C["text_muted"],
            font=("Arial", 10),
        ).grid(row=0, column=1, padx=8, pady=8, sticky="e")

        # Кнопка нагадувань
        bell_btn = tk.Button(
            topbar,
            text="🔔",
            bg=C["bg_surface"],          # фон іконки
            fg=C["text_primary"],
            activebackground=C["bg_surface_alt"],
            activeforeground=C["text_primary"],
            relief="flat",
            bd=0,
            cursor="hand2",
            font=("Arial", 12),          # трохи більший розмір емодзі
            padx=8,                      # симетричний паддінг
            pady=4,
            command=self._handle_reminders_click,
        )
        bell_btn.grid(row=0, column=2, padx=20, pady=8, sticky="e")
        bell_btn.configure(width=2, height=1)


    def _build_sidebar(self) -> None:
        C = COLORS
        sidebar = tk.Frame(self, bg=C["bg_sidebar"], width=210)
        sidebar.grid(row=1, column=0, sticky="nsw")
        sidebar.grid_propagate(False)
        sidebar.rowconfigure(1, weight=1)

        tk.Label(
            sidebar,
            text="Навігація",
            bg=C["bg_sidebar"],
            fg=C["text_muted"],
            font=("Arial", 10, "bold"),
        ).grid(row=0, column=0, padx=16, pady=(5, 6), sticky="w")

        nav_frame = tk.Frame(sidebar, bg=C["bg_sidebar"])
        nav_frame.grid(row=1, column=0, sticky="nsew", padx=8)

        self._nav_buttons: dict[PageKey, tk.Button] = {}

        menu_items: list[tuple[PageKey, str]] = [
            ("risk_register", "Реєстр ризиків"),
            ("material_events", "Реєстр суттєвих подій"),
            ("risk_appetite", "Ризик апетит"),
            ("analytics", "Аналітика"),
            ("reports", "Звіти"),
            ("risk_coordinators", "Ризик координатори"),
            ("settings", "Налаштування"),
        ]

        for i, (key, label) in enumerate(menu_items):
            btn = tk.Button(
                nav_frame,
                text=label,
                bg=C["bg_sidebar"],
                fg=C["text_muted"],
                activebackground=C["bg_surface"],
                activeforeground=C["text_primary"],
                relief="flat",
                bd=0,
                cursor="hand2",
                anchor="w",
                font=("Arial", 9),
                padx=24,
                pady=6,
                command=lambda k=key: self._on_nav_click(k),
            )
            btn.grid(row=i, column=0, sticky="ew")
            self._nav_buttons[key] = btn

        bottom = tk.Frame(sidebar, bg=C["bg_sidebar"])
        bottom.grid(row=2, column=0, sticky="ew", pady=(6, 8))
        tk.Label(
            bottom,
            text=COPYRIGHT_TEXT,
            bg=C["bg_sidebar"],
            fg=C["text_subtle"],
            font=("Arial", 8),
        ).pack(side="left", padx=14)

    def _build_content(self) -> None:
        self.content_frame = tk.Frame(self, bg=COLORS["bg_main"])
        self.content_frame.grid(row=1, column=1, sticky="nsew")
        self.content_frame.grid_rowconfigure(0, weight=1)
        self.content_frame.grid_columnconfigure(0, weight=1)

        self._current_page: PageKey | None = None
        self._pages: dict[PageKey, tk.Frame] = {}

        self._on_nav_click("material_events")

    def _on_nav_click(self, page_key: PageKey) -> None:
        if self._current_page == page_key:
            return
        self._current_page = page_key

        for key, btn in self._nav_buttons.items():
            if key == page_key:
                btn.configure(bg=COLORS["bg_surface"], fg=COLORS["text_primary"])
            else:
                btn.configure(bg=COLORS["bg_sidebar"], fg=COLORS["text_muted"])

        for child in self.content_frame.winfo_children():
            child.grid_forget()

        if page_key == "material_events":
            page = self._pages.get(page_key)
            if page is None:
                page = MaterialEventsPage(self.content_frame)
                self._pages[page_key] = page
            page.grid(row=0, column=0, sticky="nsew")
            return

        page = self._pages.get(page_key)
        if page is None:
            page = self._create_placeholder_page(page_key)
            self._pages[page_key] = page
        page.grid(row=0, column=0, sticky="nsew")

    def _create_placeholder_page(self, page_key: PageKey) -> tk.Frame:
        C = COLORS
        page = tk.Frame(self.content_frame, bg=C["bg_main"])
        page.grid_rowconfigure(1, weight=1)
        page.grid_columnconfigure(0, weight=1)

        headers = {
            "risk_register": "Реєстр ризиків",
            "risk_appetite": "Ризик апетит",
            "analytics": "Аналітика",
            "reports": "Звіти",
            "risk_coordinators": "Ризик координатори",
            "settings": "Налаштування",
            "material_events": "Реєстр суттєвих подій",
        }
        desc = {
            "risk_register": "Тут буде модуль реєстру ризиків.",
            "risk_appetite": "Тут буде модуль ризик-апетиту.",
            "analytics": "Тут буде загальна аналітика по всіх модулях.",
            "reports": "Тут буде конструктор звітів.",
            "risk_coordinators": "Тут буде управління ризик-координаторами.",
            "settings": "Тут будуть глобальні налаштування ATLAS.",
            "material_events": "",
        }

        tk.Label(
            page,
            text=headers.get(page_key, "Сторінка"),
            bg=C["bg_main"],
            fg=C["text_primary"],
            font=("Arial", 18, "bold"),
        ).grid(row=0, column=0, sticky="w", padx=24, pady=(24, 8))

        tk.Label(
            page,
            text=desc.get(page_key, ""),
            bg=C["bg_main"],
            fg=C["text_muted"],
            font=("Arial", 10),
            justify="left",
        ).grid(row=1, column=0, sticky="nw", padx=24, pady=4)

        return page

    def _handle_reminders_click(self) -> None:
        messagebox.showinfo("Нагадування", "Нагадувань поки немає.")

    def _on_close(self) -> None:
        page = self._pages.get("material_events")
        if isinstance(page, MaterialEventsPage):
            page.save_before_exit()
        self.destroy()


# =============================================================================
#  MAIN
# =============================================================================

def main() -> int:
    current_user_full_name = "Оніщенко Андрій Сергійович"
    app = AtlasApp(user_full_name=current_user_full_name)
    app.mainloop()
    return 0

if __name__ == "__main__":
    raise SystemExit(main())
