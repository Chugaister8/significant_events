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
    import numpy as np
    from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
    from matplotlib.figure import Figure
    HAS_MPL = True
except ImportError:
    HAS_MPL = False
    np = None

# =============================================================================
#  ГЛОБАЛЬНІ КОНСТАНТИ ТА ПАЛІТРА
# =============================================================================

DATA_FILE      = "essential_events.json"
RISK_DATA_FILE = "risk_register.json"

COLORS = {
    "bg_main":        "#243640",
    "bg_sidebar":     "#1E2C33",
    "bg_header":      "#1E2C33",
    "bg_surface":     "#2E4450",
    "bg_surface_alt": "#344E5A",
    "bg_input":       "#1E2C33",

    "accent":         "#4F46E5",
    "accent_soft":    "#6366F1",
    "accent_muted":   "#818CF8",
    "accent_success": "#10B981",
    "accent_danger":  "#EF4444",
    "accent_warning": "#F59E0B",

    "text_primary":   "#F3F4F6",
    "text_muted":     "#CBD5E1",
    "text_subtle":    "#94A3B8",

    "border_soft":    "#3B4F59",
    "border_strong":  "#556871",

    "row_even":       "#2A3D47",
    "row_odd":        "#243640",
    "row_select":     "#3B82F6",
}

RISK_COLORS = {
    "Операцiйний":        COLORS["accent_warning"],
    "Технiчний":          COLORS["accent"],
    "Фiнансовий":         COLORS["accent_danger"],
    "Репутацiйний":       "#a855f7",
    "Екологiчний":        COLORS["accent_success"],
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

RISK_CATEGORIES = [
    "Стратегiчний",
    "Операцiйний",
    "Фiнансовий",
    "Комплаєнс",
    "Репутацiйний",
    "IТ / Кiбербезпека",
    "Кадровий",
    "Екологiчний",
    "Iнше",
]

PROBABILITY_LEVELS = [
    "1 — Мiнiмальна", "2 — Низька", "3 — Середня",
    "4 — Висока", "5 — Критична",
]
IMPACT_LEVELS = [
    "1 — Незначний", "2 — Малий", "3 — Помiрний",
    "4 — Суттєвий", "5 — Катастрофiчний",
]


def _score_color(score: int) -> str:
    if score <= 4:   return COLORS["accent_success"]
    elif score <= 9: return COLORS["accent_warning"]
    elif score <= 16: return "#f97316"
    else:            return COLORS["accent_danger"]


def _score_label(score: int) -> str:
    if score <= 4:    return "Низький"
    elif score <= 9:  return "Помiрний"
    elif score <= 16: return "Високий"
    else:             return "Критичний"


# =============================================================================
#  ХЕЛПЕРИ ТА СТИЛІ
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
    """Налаштовує темний мінімалістичний стиль для ttk."""
    style = ttk.Style(root)
    style.theme_use("clam")
    C = COLORS

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

    style.configure("TFrame",        background=C["bg_main"])
    style.configure("Surface.TFrame", background=C["bg_surface"])
    style.configure("Sidebar.TFrame", background=C["bg_sidebar"])
    style.configure("Header.TFrame",  background=C["bg_header"])

    style.configure("TLabel",
                    background=C["bg_main"],
                    foreground=C["text_primary"],
                    font=("Arial", 9))
    style.configure("Muted.TLabel",
                    background=C["bg_main"],
                    foreground=C["text_muted"],
                    font=("Arial", 8))

    style.configure("TEntry",
                    fieldbackground=C["bg_input"],
                    foreground=C["text_primary"],
                    bordercolor=C["border_soft"],
                    insertcolor=C["text_primary"])
    style.map("TEntry",
              fieldbackground=[("focus", C["bg_surface_alt"])],
              bordercolor=[("focus", C["accent"])])

    style.configure("TCombobox",
                    fieldbackground=C["bg_surface"],
                    background=C["bg_surface"],
                    foreground=C["text_primary"],
                    bordercolor=C["border_soft"],
                    arrowcolor=C["text_muted"])
    style.map("TCombobox",
              fieldbackground=[("readonly", C["bg_surface"]),
                               ("hover",    C["bg_surface_alt"]),
                               ("focus",    C["bg_surface_alt"])],
              background=[("readonly", C["bg_surface"]),
                          ("hover",    C["bg_surface_alt"]),
                          ("focus",    C["bg_surface_alt"])],
              foreground=[("disabled", C["text_subtle"])],
              arrowcolor=[("hover", C["text_primary"]),
                          ("focus", C["accent"])])

    style.configure("TNotebook",
                    background=C["bg_main"],
                    bordercolor=C["border_soft"],
                    tabmargins=[0, 0, 0, 0])
    style.configure("TNotebook.Tab",
                    background=C["bg_sidebar"],
                    foreground=C["text_muted"],
                    padding=(14, 6),
                    font=("Arial", 9))
    style.map("TNotebook.Tab",
              background=[("selected", C["bg_surface"]),
                          ("active",   C["bg_surface_alt"])],
              foreground=[("selected", C["text_primary"]),
                          ("active",   C["text_primary"])])

    style.configure("Treeview",
                    background=C["row_odd"],
                    foreground=C["text_primary"],
                    fieldbackground=C["row_odd"],
                    bordercolor=C["border_soft"],
                    font=("Arial", 9),
                    rowheight=24)
    style.configure("Treeview.Heading",
                    background=C["bg_surface"],
                    foreground=C["text_muted"],
                    bordercolor=C["border_soft"],
                    font=("Arial", 8, "bold"),
                    relief="flat")
    style.map("Treeview",
              background=[("selected", C["row_select"])],
              foreground=[("selected", C["text_primary"])])

    for orient in ("Vertical", "Horizontal"):
        style.configure(f"{orient}.TScrollbar",
                        background=C["bg_surface"],
                        troughcolor=C["bg_main"],
                        arrowcolor=C["text_muted"],
                        bordercolor=C["bg_main"])


def make_dark_text(parent: tk.Misc, **kwargs) -> tk.Text:
    """Створює Text-віджет у єдиному стилі програми."""
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


def make_dark_entry(parent: tk.Misc, accent: str | None = None, **kwargs) -> tk.Entry:
    """Створює Entry-віджет у єдиному стилі програми."""
    C = COLORS
    return tk.Entry(
        parent,
        bg=C["bg_input"],
        fg=C["text_primary"],
        insertbackground=C["text_primary"],
        relief="flat",
        bd=2,
        highlightthickness=1,
        highlightbackground=C["border_soft"],
        highlightcolor=accent or C["accent"],
        font=("Arial", 9),
        **kwargs,
    )


def make_dark_combo(parent: tk.Misc, values: list[str] | None = None,
                    **kwargs) -> ttk.Combobox:
    """Створює Combobox у єдиному стилі програми."""
    return ttk.Combobox(
        parent,
        values=values or [],
        state="readonly",
        font=("Arial", 9),
        **kwargs,
    )


def make_button(parent: tk.Misc, text: str, bg: str,
                fg: str = "white", **kwargs) -> tk.Button:
    """Уніфікована кнопка."""
    C = COLORS
    active_bg = kwargs.pop("activebackground", bg)
    active_fg = kwargs.pop("activeforeground", fg)
    # font береться з kwargs, або встановлюється дефолтний
    font = kwargs.pop("font", ("Arial", 9))
    return tk.Button(
        parent,
        text=text,
        bg=bg,
        fg=fg,
        activebackground=active_bg,
        activeforeground=active_fg,
        relief="flat",
        bd=0,
        cursor="hand2",
        font=font,
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

    entry.bind("<FocusIn>",  on_in)
    entry.bind("<FocusOut>", on_out)


def _extract_num(val: str) -> int:
    """Витягує цифру з рядку."""
    try:
        return int(str(val).split()[0])
    except (ValueError, IndexError, TypeError):
        try:
            return int(str(val))
        except ValueError:
            return 1


def _build_section_label(parent: tk.Misc, text: str, row: int,
                          accent: str | None = None) -> None:
    """Уніфікований заголовок секції."""
    C = COLORS
    color = accent or C["accent"]
    f = tk.Frame(parent, bg=C["bg_main"])
    f.grid(row=row, column=0, columnspan=2, sticky="ew",
           padx=8, pady=(14, 4))
    tk.Frame(f, bg=color, width=3, height=16).pack(side="left")
    tk.Label(f, text=text, bg=C["bg_main"], fg=color,
             font=("Arial", 9, "bold")).pack(side="left", padx=8)


def _build_info_cell(parent: tk.Misc, label: str, value: str,
                     row: int, col: int = 0,
                     value_color: str | None = None) -> None:
    """Уніфікована інформаційна комірка."""
    C = COLORS
    cell = tk.Frame(parent, bg=C["bg_surface"], padx=10, pady=6)
    cell.grid(
        row=row, column=col, sticky="nsew",
        padx=(8 if col == 0 else 4, 4 if col == 0 else 8), pady=3,
    )
    cell.columnconfigure(0, weight=1)
    tk.Label(cell, text=label, bg=C["bg_surface"], fg=C["text_subtle"],
             font=("Arial", 7, "bold")).grid(row=0, column=0, sticky="w")
    fg = value_color if value_color else C["text_primary"]
    fnt = ("Arial", 10, "bold") if value_color else ("Arial", 9)
    tk.Label(cell, text=value or "—", bg=C["bg_surface"], fg=fg,
             font=fnt, wraplength=260, justify="left").grid(
        row=1, column=0, sticky="w", pady=(2, 0))


def _build_text_block(parent: tk.Misc, label: str, value: str,
                      row: int) -> None:
    """Уніфікований текстовий блок (read-only)."""
    C = COLORS
    cell = tk.Frame(parent, bg=C["bg_surface"], padx=10, pady=6)
    cell.grid(row=row, column=0, columnspan=2, sticky="nsew",
              padx=8, pady=3)
    cell.columnconfigure(0, weight=1)
    tk.Label(cell, text=label, bg=C["bg_surface"], fg=C["text_subtle"],
             font=("Arial", 7, "bold")).grid(row=0, column=0, sticky="w")
    t = make_dark_text(cell, height=3, wrap="word", state="normal")
    t.insert("1.0", value or "—")
    t.configure(state="disabled")
    t.grid(row=1, column=0, sticky="ew", pady=(4, 0))


def _show_toast(frame: tk.Widget, msg: str,
                color: str | None = None) -> None:
    """Уніфікований toast-сповіщення."""
    bg = color or COLORS["accent_success"]
    toast = tk.Toplevel(frame)
    toast.overrideredirect(True)
    toast.attributes("-topmost", True)
    toast.configure(bg=bg)
    tk.Label(toast, text=f"  {msg}  ", bg=bg, fg="white",
             font=("Arial", 9, "bold"), pady=6).pack()
    root = frame.winfo_toplevel()
    x = root.winfo_x() + root.winfo_width() - 240
    y = root.winfo_y() + root.winfo_height() - 80
    toast.geometry(f"+{x}+{y}")
    toast.after(2200, toast.destroy)


def _scrollable_canvas(container: tk.Misc) -> tuple[tk.Canvas, tk.Frame]:
    """Повертає (canvas, inner_frame) з прокруткою."""
    C = COLORS
    container.columnconfigure(0, weight=1)
    container.rowconfigure(0, weight=1)

    canvas = tk.Canvas(container, bg=C["bg_main"], highlightthickness=0)
    sb = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
    canvas.configure(yscrollcommand=sb.set)
    canvas.grid(row=0, column=0, sticky="nsew")
    sb.grid(row=0, column=1, sticky="ns")

    inner = tk.Frame(canvas, bg=C["bg_main"])
    cw = canvas.create_window((0, 0), window=inner, anchor="nw")

    def _on_conf(_: object) -> None:
        canvas.configure(scrollregion=canvas.bbox("all"))
        canvas.itemconfig(cw, width=canvas.winfo_width())

    inner.bind("<Configure>", _on_conf)
    canvas.bind("<Configure>", _on_conf)
    canvas.bind_all(
        "<MouseWheel>",
        lambda e: canvas.yview_scroll(int(-1 * e.delta / 120), "units"),
    )
    return canvas, inner


# =============================================================================
#  ДЕТАЛЬНЕ ВІКНО ЗАПИСУ ПОДІЇ
# =============================================================================

class EventDetailWindow:
    """Спливаюче вікно для перегляду та редагування збереженої події."""

    def __init__(
        self,
        parent_root:     tk.Misc,
        record:          tuple,
        all_records:     list[tuple],
        save_callback:   Callable[[str, tuple], None],
        delete_callback: Callable[[str], None],
        toast_callback:  Callable[[str], None],
    ) -> None:
        self.parent_root     = parent_root
        self.record          = list(record)
        self.all_records     = all_records
        self.save_callback   = save_callback
        self.delete_callback = delete_callback
        self.toast_callback  = toast_callback
        self.is_edit_mode    = False
        self._build_window()

    # ------------------------------------------------------------------
    def _build_window(self) -> None:
        C = COLORS
        self.win = tk.Toplevel(self.parent_root)
        self.win.title(f"Подiя #{self.record[0]}  —  {self.record[1]}")
        self.win.geometry("780x700")
        self.win.minsize(640, 500)
        self.win.configure(bg=C["bg_main"])
        self.win.grab_set()

        self.win.update_idletasks()
        rx, ry = self.parent_root.winfo_x(), self.parent_root.winfo_y()
        rw, rh = self.parent_root.winfo_width(), self.parent_root.winfo_height()
        ww, wh = 780, 700
        self.win.geometry(f"{ww}x{wh}+{rx+(rw-ww)//2}+{ry+(rh-wh)//2}")

        self.win.rowconfigure(1, weight=1)
        self.win.columnconfigure(0, weight=1)

        # ── Header ──────────────────────────────────────────────────────
        header = tk.Frame(self.win, bg=C["bg_header"], height=58)
        header.grid(row=0, column=0, sticky="ew")
        header.columnconfigure(1, weight=1)
        header.grid_propagate(False)

        risk_color = RISK_COLORS.get(
            self.record[4] if len(self.record) > 4 else "", C["accent"])
        tk.Frame(header, bg=risk_color, width=4).grid(row=0, column=0, sticky="ns")

        title_frame = tk.Frame(header, bg=C["bg_header"])
        title_frame.grid(row=0, column=1, sticky="ew", padx=16, pady=10)

        self.lbl_title = tk.Label(
            title_frame, text=f"Запис #{self.record[0]}",
            bg=C["bg_header"], fg=C["accent_muted"],
            font=("Arial", 11, "bold"))
        self.lbl_title.pack(anchor="w")

        self.lbl_subtitle = tk.Label(
            title_frame,
            text=self.record[1] if len(self.record) > 1 else "",
            bg=C["bg_header"], fg=C["text_muted"], font=("Arial", 9))
        self.lbl_subtitle.pack(anchor="w")

        status_val = self.record[10] if len(self.record) > 10 else "—"
        sc = self._status_color(status_val)
        self.lbl_status_badge = tk.Label(
            header, text=f"  {status_val}  ",
            bg=sc, fg="white", font=("Arial", 8, "bold"), pady=3)
        self.lbl_status_badge.grid(row=0, column=2, padx=16, pady=18)

        # ── Scroll ──────────────────────────────────────────────────────
        _, self.content = _scrollable_canvas(
            tk.Frame(self.win, bg=C["bg_main"]))
        # замінюємо на власний canvas щоб мати grid
        canvas = tk.Canvas(self.win, bg=C["bg_main"], highlightthickness=0)
        sb = ttk.Scrollbar(self.win, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=sb.set)
        canvas.grid(row=1, column=0, sticky="nsew")
        sb.grid(row=1, column=1, sticky="ns")

        self.content = tk.Frame(canvas, bg=C["bg_main"])
        self._cw = canvas.create_window((0, 0), window=self.content, anchor="nw")

        def _on_conf(_: object) -> None:
            canvas.configure(scrollregion=canvas.bbox("all"))
            canvas.itemconfig(self._cw, width=canvas.winfo_width())

        self.content.bind("<Configure>", _on_conf)
        canvas.bind("<Configure>", _on_conf)
        canvas.bind_all("<MouseWheel>",
                        lambda e: canvas.yview_scroll(
                            int(-1 * e.delta / 120), "units"))

        self.content.columnconfigure(0, weight=1)
        self.content.columnconfigure(1, weight=1)
        self._build_view_content()

        # ── Bottom buttons ───────────────────────────────────────────────
        btn_bar = tk.Frame(self.win, bg=C["bg_header"], height=46)
        btn_bar.grid(row=2, column=0, columnspan=2, sticky="ew")
        btn_bar.grid_propagate(False)

        left_btns = tk.Frame(btn_bar, bg=C["bg_header"])
        left_btns.pack(side="left", padx=12, pady=8)

        self.btn_edit = make_button(
            left_btns, "Редагувати",
            bg=C["accent_warning"], fg=C["bg_main"],
            activebackground="#d97706", activeforeground="white",
            font=("Arial", 9, "bold"), padx=14, pady=4,
            command=self._toggle_edit_mode)
        self.btn_edit.pack(side="left", padx=(0, 8))

        self.btn_save = make_button(
            left_btns, "Зберегти змiни",
            bg=C["accent_success"],
            activebackground="#16a34a",
            font=("Arial", 9, "bold"), padx=14, pady=4,
            command=self._save_changes)
        self.btn_save.pack_forget()

        self.btn_cancel_edit = make_button(
            left_btns, "Скасувати",
            bg=C["bg_surface"], fg=C["text_muted"],
            activebackground=C["bg_surface_alt"],
            activeforeground=C["text_primary"],
            font=("Arial", 9), padx=12, pady=4,
            command=self._cancel_edit)
        self.btn_cancel_edit.pack_forget()

        right_btns = tk.Frame(btn_bar, bg=C["bg_header"])
        right_btns.pack(side="right", padx=12, pady=8)

        make_button(right_btns, "Видалити",
                    bg=C["accent_danger"],
                    activebackground="#dc2626",
                    font=("Arial", 9, "bold"), padx=14, pady=4,
                    command=self._delete_record
                    ).pack(side="right", padx=(8, 0))

        make_button(right_btns, "Закрити",
                    bg=C["bg_surface"], fg=C["text_primary"],
                    activebackground=C["bg_surface_alt"],
                    activeforeground=C["text_primary"],
                    font=("Arial", 9), padx=12, pady=4,
                    command=self.win.destroy
                    ).pack(side="right")

    # ------------------------------------------------------------------
    @staticmethod
    def _status_color(status: str) -> str:
        return {
            "Вiдкрито":  COLORS["accent_danger"],
            "В обробцi": COLORS["accent_warning"],
            "Закрито":   COLORS["text_muted"],
            "Вирiшено":  COLORS["accent_success"],
        }.get(status, COLORS["text_muted"])

    @staticmethod
    def _priority_color(priority: str) -> str:
        return {
            "Критичний": COLORS["accent_danger"],
            "Високий":   COLORS["accent_warning"],
            "Середнiй":  COLORS["accent"],
            "Низький":   COLORS["accent_success"],
        }.get(priority, COLORS["text_primary"])

    # ------------------------------------------------------------------
    def _build_view_content(self) -> None:
        for w in self.content.winfo_children():
            w.destroy()

        r   = self.record
        row = 0

        _build_section_label(self.content, "Iнформацiя про пiдприємство", row)
        row += 1

        _build_info_cell(self.content, "Пiдприємство",
                         r[1] if len(r) > 1 else "—", row, 0)

        p_val = r[9] if len(r) > 9 else "—"
        _build_info_cell(self.content, "Прiоритет", p_val, row, 1,
                         value_color=self._priority_color(p_val))
        row += 1

        _build_section_label(self.content, "Опис подiї / ризику", row)
        row += 1

        _build_info_cell(self.content, "Назва подiї",
                         r[2] if len(r) > 2 else "—", row, 0)

        risk_val = r[4] if len(r) > 4 else "—"
        _build_info_cell(self.content, "Тип ризику", risk_val, row, 1,
                         value_color=RISK_COLORS.get(risk_val))
        row += 1

        status_val = r[10] if len(r) > 10 else "—"
        _build_info_cell(self.content, "Статус", status_val, row, 0,
                         value_color=self._status_color(status_val))
        _build_info_cell(self.content, "Задiянi пiдроздiли / особи",
                         r[3] if len(r) > 3 else "—", row, 1)
        row += 1

        _build_section_label(self.content, "Дати", row)
        row += 1

        _build_info_cell(self.content, "Дата подiї",
                         r[5] if len(r) > 5 else "—", row, 0)
        _build_info_cell(self.content, "Дата виявлення",
                         r[8] if len(r) > 8 else "—", row, 1)
        row += 1

        _build_section_label(self.content, "Деталi подiї", row)
        row += 1

        _build_text_block(self.content, "Детальний опис подiї",
                          r[6] if len(r) > 6 else "—", row)
        row += 1
        _build_text_block(self.content, "Вжитi заходи",
                          r[7] if len(r) > 7 else "—", row)
        row += 1

        tk.Frame(self.content, bg=COLORS["bg_main"], height=12).grid(
            row=row, column=0, columnspan=2)

    # ------------------------------------------------------------------
    def _build_edit_content(self) -> None:
        C   = COLORS
        rec = self.record
        for w in self.content.winfo_children():
            w.destroy()
        self.content.columnconfigure(0, weight=1)

        def section(txt: str, r: int) -> int:
            f = tk.Frame(self.content, bg=C["bg_main"])
            f.grid(row=r, column=0, sticky="ew", padx=8, pady=(14, 4))
            tk.Frame(f, bg=C["accent_warning"], width=3, height=16).pack(side="left")
            tk.Label(f, text=txt, bg=C["bg_main"], fg=C["accent_warning"],
                     font=("Arial", 9, "bold")).pack(side="left", padx=8)
            return r + 1

        def lbl(text: str, r: int) -> None:
            tk.Label(self.content, text=text, bg=C["bg_main"],
                     fg=C["text_subtle"], font=("Arial", 8)).grid(
                row=r, column=0, sticky="w", padx=10, pady=(6, 0))

        def mk_entry(**kw: object) -> tk.Entry:
            return make_dark_entry(self.content,
                                   accent=C["accent_warning"], **kw)

        def mk_combo(values: list[str] | None = None,
                     **kw: object) -> ttk.Combobox:
            return make_dark_combo(self.content, values=values, **kw)

        row = 0
        row = section("Пiдприємство та подiя", row)

        lbl("Скорочена назва пiдприємства:", row); row += 1
        self.e_entity = mk_entry()
        self.e_entity.grid(row=row, column=0, sticky="ew", padx=10, pady=(2, 0))
        self.e_entity.insert(0, rec[1] if len(rec) > 1 else "")
        row += 1

        lbl("Назва подiї:", row); row += 1
        self.e_event = mk_combo(values=EVENT_TYPES)
        self.e_event.grid(row=row, column=0, sticky="ew", padx=10, pady=(2, 0))
        self.e_event.set(rec[2] if len(rec) > 2 else "")
        row += 1

        lbl("Тип ризику:", row); row += 1
        self.e_risk = mk_combo(values=RISK_TYPES)
        self.e_risk.grid(row=row, column=0, sticky="ew", padx=10, pady=(2, 0))
        self.e_risk.set(rec[4] if len(rec) > 4 and rec[4] != "—" else "")
        row += 1

        lbl("Задiянi пiдроздiли / особи:", row); row += 1
        self.e_involved = make_dark_text(self.content, height=2, wrap="word")
        self.e_involved.grid(row=row, column=0, sticky="ew",
                             padx=10, pady=(2, 0))
        if len(rec) > 3 and rec[3] and rec[3] != "—":
            self.e_involved.insert("1.0", rec[3])
        row += 1

        row = section("Дати", row)
        date_f = tk.Frame(self.content, bg=C["bg_main"])
        date_f.grid(row=row, column=0, sticky="w", padx=10, pady=4)
        row += 1

        for ci, (lbl_t, attr, vi) in enumerate(
            [("Дата виявлення:", "e_detect", 8),
             ("Дата подiї:",    "e_event_date", 5)]
        ):
            tk.Label(date_f, text=lbl_t, bg=C["bg_main"],
                     fg=C["text_subtle"], font=("Arial", 8)).grid(
                row=0, column=ci,
                padx=(0 if ci == 0 else 20, 0), sticky="w")
            e = make_dark_entry(date_f, accent=C["accent_warning"], width=14)
            e.grid(row=1, column=ci,
                   padx=(0 if ci == 0 else 20, 0), pady=2)
            val = rec[vi] if len(rec) > vi else ""
            e.insert(0, val) if (val and val != "—") else add_placeholder(e, "дд.мм.рррр")
            setattr(self, attr, e)

        row = section("Прiоритет та статус", row)
        ps_f = tk.Frame(self.content, bg=C["bg_main"])
        ps_f.grid(row=row, column=0, sticky="w", padx=10, pady=4)
        row += 1

        tk.Label(ps_f, text="Прiоритет:", bg=C["bg_main"],
                 fg=C["text_subtle"], font=("Arial", 8)).grid(
            row=0, column=0, sticky="w")
        self.e_priority = make_dark_combo(
            ps_f, values=["Критичний", "Високий", "Середнiй", "Низький"],
            width=14)
        self.e_priority.grid(row=1, column=0, pady=(2, 0), padx=(0, 20))
        self.e_priority.set(rec[9] if len(rec) > 9 else "Середнiй")

        tk.Label(ps_f, text="Статус:", bg=C["bg_main"],
                 fg=C["text_subtle"], font=("Arial", 8)).grid(
            row=0, column=1, sticky="w")
        self.e_status = make_dark_combo(
            ps_f,
            values=["Вiдкрито", "В обробцi", "Закрито", "Вирiшено"],
            width=14)
        self.e_status.grid(row=1, column=1, pady=(2, 0))
        self.e_status.set(rec[10] if len(rec) > 10 else "Вiдкрито")

        row = section("Деталi подiї", row)

        lbl("Детальний опис подiї:", row); row += 1
        self.e_description = make_dark_text(self.content, height=4, wrap="word")
        self.e_description.grid(row=row, column=0, sticky="ew",
                                padx=10, pady=(2, 0))
        if len(rec) > 6 and rec[6] and rec[6] != "—":
            self.e_description.insert("1.0", rec[6])
        row += 1

        lbl("Вжитi заходи:", row); row += 1
        self.e_measures = make_dark_text(self.content, height=3, wrap="word")
        self.e_measures.grid(row=row, column=0, sticky="ew",
                             padx=10, pady=(2, 0))
        if len(rec) > 7 and rec[7] and rec[7] != "—":
            self.e_measures.insert("1.0", rec[7])
        row += 1

        tk.Frame(self.content, bg=C["bg_main"], height=16).grid(
            row=row, column=0)

    # ------------------------------------------------------------------
    def _toggle_edit_mode(self) -> None:
        self.is_edit_mode = True
        self._build_edit_content()
        self.btn_edit.pack_forget()
        self.btn_save.pack(side="left", padx=(0, 8))
        self.btn_cancel_edit.pack(side="left")
        self.lbl_title.configure(
            text=f"Редагування запису #{self.record[0]}",
            fg=COLORS["accent_warning"])

    def _cancel_edit(self) -> None:
        self.is_edit_mode = False
        self._build_view_content()
        self.btn_save.pack_forget()
        self.btn_cancel_edit.pack_forget()
        self.btn_edit.pack(side="left", padx=(0, 8))
        self.lbl_title.configure(
            text=f"Запис #{self.record[0]}",
            fg=COLORS["accent_muted"])

    def _save_changes(self) -> None:
        detect  = self.e_detect.get().strip()
        event_d = self.e_event_date.get().strip()

        for val, label in [(detect, "дати виявлення"),
                           (event_d, "дати подiї")]:
            if val in ("дд.мм.рррр", ""):
                continue
            if not is_valid_date(val):
                messagebox.showwarning(
                    "Помилка",
                    f"Неправильний формат {label} (очiкується дд.мм.рррр)",
                    parent=self.win)
                return

        detect  = "" if detect  == "дд.мм.рррр" else detect
        event_d = "" if event_d == "дд.мм.рррр" else event_d

        entity = self.e_entity.get().strip()
        event  = self.e_event.get().strip()
        if not entity or not event:
            messagebox.showwarning(
                "Обов'язковi поля",
                "Заповнiть назву пiдприємства та назву подiї",
                parent=self.win)
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
            self.e_status.get().strip()   or "Вiдкрито",
        )
        self.record = list(new_record)
        self.save_callback(str(old_id), new_record)

        self.lbl_subtitle.configure(text=entity)
        sc = self._status_color(new_record[10])
        self.lbl_status_badge.configure(text=f"  {new_record[10]}  ", bg=sc)
        self._cancel_edit()
        self.toast_callback("Запис оновлено")

    def _delete_record(self) -> None:
        idx_str = self.record[0]
        if not messagebox.askyesno(
            "Пiдтвердження",
            f"Видалити запис #{idx_str}?\nЦю дiю не можна скасувати.",
            parent=self.win):
            return
        self.delete_callback(str(idx_str))
        self.toast_callback("Запис видалено")
        self.win.destroy()


# =============================================================================
#  ВКЛАДКА: РЕЄСТР СУТТЄВИХ ПОДІЙ
# =============================================================================

class RegistryTab:
    """Вкладка 'Реєстр суттєвих подiй'."""

    def __init__(
        self,
        parent: tk.Misc,
        on_data_change: Callable[[list[tuple]], None] | None = None,
    ) -> None:
        self.parent         = parent
        self.on_data_change = on_data_change
        self.frame          = ttk.Frame(parent)
        self.all_records:   list[tuple] = []

        self._build_ui()
        self._load_data()

    # ------------------------------------------------------------------
    def _build_ui(self) -> None:
        C = COLORS
        self.frame.rowconfigure(1, weight=1)
        self.frame.columnconfigure(0, weight=1)

        # ── Header ──────────────────────────────────────────────────────
        header = tk.Frame(self.frame, bg=C["bg_header"], height=56)
        header.grid(row=0, column=0, sticky="ew")
        header.columnconfigure(1, weight=1)
        header.grid_propagate(False)

        tk.Label(header, text="РЕЄСТР СУТТЄВИХ ПОДIЙ",
                 bg=C["bg_header"], fg=C["accent_muted"],
                 font=("Arial", 13, "bold")).grid(
            row=0, column=0, padx=20, pady=14, sticky="w")

        sf = tk.Frame(header, bg=C["bg_header"])
        sf.grid(row=0, column=1, sticky="e", padx=20)
        tk.Label(sf, text="Пошук:", bg=C["bg_header"],
                 fg=C["text_muted"], font=("Arial", 8)).pack(
            side="left", padx=(0, 6))

        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *_: self._apply_filter())
        tk.Entry(sf, textvariable=self.search_var,
                 bg=C["bg_input"], fg=C["text_primary"],
                 insertbackground=C["text_primary"],
                 relief="flat", bd=2, font=("Arial", 9),
                 width=34).pack(side="left", padx=(0, 8), ipady=2)
        make_button(sf, "Скинути",
                    bg=C["bg_surface"], fg=C["text_muted"],
                    activebackground=C["bg_surface_alt"],
                    activeforeground=C["text_primary"],
                    font=("Arial", 8), padx=8, pady=2,
                    command=self._reset_filter).pack(side="left")

        # ── Paned ───────────────────────────────────────────────────────
        paned = ttk.PanedWindow(self.frame, orient="horizontal")
        paned.grid(row=1, column=0, sticky="nsew")
        lw = ttk.Frame(paned)
        rw = ttk.Frame(paned)
        paned.add(lw, weight=4)
        paned.add(rw, weight=7)

        self._build_form(lw)
        self._build_table(rw)

    # ------------------------------------------------------------------
    def _build_form(self, container: tk.Misc) -> None:
        C = COLORS
        container.columnconfigure(0, weight=1)
        container.rowconfigure(0, weight=1)

        canvas = tk.Canvas(container, bg=C["bg_main"], highlightthickness=0)
        sb = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=sb.set)
        canvas.grid(row=0, column=0, sticky="nsew")
        sb.grid(row=0, column=1, sticky="ns")

        form = tk.Frame(canvas, bg=C["bg_main"])
        fw   = canvas.create_window((0, 0), window=form, anchor="nw")

        def _conf(_: object) -> None:
            canvas.configure(scrollregion=canvas.bbox("all"))
            canvas.itemconfig(fw, width=canvas.winfo_width())

        form.bind("<Configure>", _conf)
        canvas.bind("<Configure>", _conf)
        canvas.bind_all("<MouseWheel>",
                        lambda e: canvas.yview_scroll(
                            int(-1 * e.delta / 120), "units"))
        form.columnconfigure(0, weight=1)

        def section(txt: str, r: int) -> int:
            f = tk.Frame(form, bg=C["bg_main"])
            f.grid(row=r, column=0, sticky="ew", padx=16, pady=(16, 4))
            tk.Frame(f, bg=C["accent"], width=3, height=16).pack(side="left")
            tk.Label(f, text=txt, bg=C["bg_main"], fg=C["accent"],
                     font=("Arial", 9, "bold")).pack(side="left", padx=8)
            return r + 1

        def lbl(text: str, r: int) -> None:
            tk.Label(form, text=text, bg=C["bg_main"],
                     fg=C["text_subtle"], font=("Arial", 8)).grid(
                row=r, column=0, sticky="w", padx=16, pady=(4, 0))

        def mk_e(**kw: object) -> tk.Entry:
            return make_dark_entry(form, **kw)

        def mk_c(values: list[str] | None = None,
                 **kw: object) -> ttk.Combobox:
            return make_dark_combo(form, values=values, **kw)

        def field(lbl_txt: str, r: int,
                  factory: Callable) -> tuple[tk.Widget, int]:
            lbl(lbl_txt, r)
            w = factory()
            w.grid(row=r + 1, column=0, sticky="ew", padx=16, pady=(2, 0))
            return w, r + 2

        row = 0

        # Badge
        badge_f = tk.Frame(form, bg=C["bg_main"])
        badge_f.grid(row=row, column=0, sticky="ew", padx=16, pady=(12, 0))
        tk.Label(badge_f, text="  + НОВИЙ ЗАПИС  ",
                 bg=C["accent"], fg="white",
                 font=("Arial", 8, "bold"), pady=3).pack(anchor="w")
        row += 1

        row = section("Iнформацiя про пiдприємство та особу", row)
        self.ent_entity,   row = field("Скорочена назва пiдприємства:", row, mk_e)
        self.ent_position, row = field("Посада:", row, mk_e)
        self.ent_reporter, row = field("ПIБ особи, що звiтує:", row, mk_e)

        row = section("Опис подiї / ризику", row)
        self.cb_event, row = field(
            "Назва подiї:", row, lambda: mk_c(values=EVENT_TYPES))
        self.cb_risk, row = field(
            "Тип ризику:", row, lambda: mk_c(values=RISK_TYPES))

        lbl("Задiянi пiдроздiли / особи:", row); row += 1
        self.txt_involved = make_dark_text(form, height=2, wrap="word")
        self.txt_involved.grid(row=row, column=0, sticky="ew",
                               padx=16, pady=(2, 0))
        row += 1

        row = section("Фiнансовий вплив (млн грн)", row)
        fin_f = tk.Frame(form, bg=C["bg_main"])
        fin_f.grid(row=row, column=0, sticky="ew", padx=16, pady=4)
        fin_f.columnconfigure((0, 1, 2, 3), weight=1)
        row += 1

        for ci, title in enumerate(
            ["Втрати", "Резерв", "Запланованi втрати", "Вiдшкодування"]
        ):
            tk.Label(fin_f, text=title, bg=C["bg_main"],
                     fg=C["text_muted"], font=("Arial", 7)).grid(
                row=0, column=ci, sticky="w", padx=4)

        self.ent_loss    = make_dark_entry(fin_f, width=11)
        self.ent_reserve = make_dark_entry(fin_f, width=11)
        self.ent_planned = make_dark_entry(fin_f, width=11)
        self.ent_refund  = make_dark_entry(fin_f, width=11)
        for ci, e in enumerate(
            [self.ent_loss, self.ent_reserve, self.ent_planned, self.ent_refund]
        ):
            e.grid(row=1, column=ci, padx=4, pady=2, sticky="ew")

        net_f = tk.Frame(form, bg=C["bg_main"])
        net_f.grid(row=row, column=0, sticky="ew", padx=16, pady=(2, 6))
        row += 1
        tk.Label(net_f, text="Чистий вплив (млн грн):",
                 bg=C["bg_main"], fg=C["text_muted"],
                 font=("Arial", 8)).pack(side="left")
        self.lbl_net = tk.Label(net_f, text="0.00",
                                bg=C["bg_main"], fg=C["accent_success"],
                                font=("Arial", 13, "bold"))
        self.lbl_net.pack(side="left", padx=10)

        def _upd_net(_: object) -> None:
            try:
                loss   = float(self.ent_loss.get().replace(",", ".") or 0)
                refund = float(self.ent_refund.get().replace(",", ".") or 0)
                net    = loss - refund
                col    = C["accent_danger"] if net > 0 else C["accent_success"]
                self.lbl_net.configure(text=f"{net:,.2f}", fg=col)
            except Exception:
                self.lbl_net.configure(text="—", fg=C["text_muted"])

        self.ent_loss.bind("<KeyRelease>",   _upd_net)
        self.ent_refund.bind("<KeyRelease>", _upd_net)

        row = section("Деталi подiї", row)
        for lbl_txt, attr, h in [
            ("Вплив на iншi пiдприємства:",  "txt_impact",      2),
            ("Нефiнансовий / якiсний вплив:", "txt_qualitative", 2),
            ("Детальний опис подiї:",         "txt_description", 4),
            ("Вжитi заходи:",                 "txt_measures",    3),
        ]:
            lbl(lbl_txt, row); row += 1
            t = make_dark_text(form, height=h, wrap="word")
            t.grid(row=row, column=0, sticky="ew", padx=16, pady=(2, 0))
            setattr(self, attr, t)
            row += 1

        row = section("Дати", row)
        date_f = tk.Frame(form, bg=C["bg_main"])
        date_f.grid(row=row, column=0, sticky="w", padx=16, pady=4)
        row += 1
        for ci, (lbl_t, attr) in enumerate(
            [("Дата виявлення:", "ent_detect"),
             ("Дата подiї:",    "ent_event_date")]
        ):
            tk.Label(date_f, text=lbl_t, bg=C["bg_main"],
                     fg=C["text_subtle"], font=("Arial", 8)).grid(
                row=0, column=ci, padx=(0 if ci == 0 else 20, 0), sticky="w")
            e = make_dark_entry(date_f, width=14)
            e.grid(row=1, column=ci,
                   padx=(0 if ci == 0 else 20, 0), pady=2)
            add_placeholder(e, "дд.мм.рррр")
            setattr(self, attr, e)

        lbl("Прiоритет:", row); row += 1
        self.cb_priority = mk_c(
            values=["Критичний", "Високий", "Середнiй", "Низький"])
        self.cb_priority.grid(row=row, column=0, sticky="w",
                              padx=16, pady=(2, 0))
        row += 1

        lbl("Статус:", row); row += 1
        self.cb_status = mk_c(
            values=["Вiдкрито", "В обробцi", "Закрито", "Вирiшено"])
        self.cb_status.grid(row=row, column=0, sticky="w",
                            padx=16, pady=(2, 0))
        row += 1

        btn_f = tk.Frame(form, bg=C["bg_main"])
        btn_f.grid(row=row, column=0, sticky="ew", padx=16, pady=(16, 20))
        btn_f.columnconfigure((0, 1), weight=1)

        make_button(btn_f, "Очистити",
                    bg=C["bg_surface"], fg=C["text_muted"],
                    activebackground=C["bg_surface_alt"],
                    activeforeground=C["text_primary"],
                    padx=14, pady=6,
                    command=self._clear_form).grid(
            row=0, column=0, padx=4, sticky="ew")
        make_button(btn_f, "Додати запис",
                    bg=C["accent"],
                    activebackground=C["accent_soft"],
                    font=("Arial", 9, "bold"), padx=14, pady=6,
                    command=self._add_record).grid(
            row=0, column=1, padx=4, sticky="ew")

    # ------------------------------------------------------------------
    def _build_table(self, container: tk.Misc) -> None:
        C = COLORS
        container.rowconfigure(1, weight=1)
        container.columnconfigure(0, weight=1)

        toolbar = tk.Frame(container, bg=C["bg_surface"], height=40)
        toolbar.grid(row=0, column=0, sticky="ew")
        toolbar.grid_propagate(False)

        tk.Label(toolbar, text="Записи", bg=C["bg_surface"],
                 fg=C["text_muted"], font=("Arial", 8)).pack(
            side="left", padx=12, pady=8)
        self.lbl_count = tk.Label(toolbar, text="0", bg=C["bg_surface"],
                                  fg=C["accent"], font=("Arial", 8, "bold"))
        self.lbl_count.pack(side="left", pady=8)

        tk.Label(toolbar, text="  |  Ризик:", bg=C["bg_surface"],
                 fg=C["text_muted"], font=("Arial", 8)).pack(
            side="left", pady=8)
        self.filter_risk = make_dark_combo(
            toolbar, values=["Всi"] + RISK_TYPES, width=16)
        self.filter_risk.set("Всi")
        self.filter_risk.pack(side="left", padx=6, pady=8)
        self.filter_risk.bind("<<ComboboxSelected>>",
                              lambda _: self._apply_filter())

        tk.Label(toolbar, text="Статус:", bg=C["bg_surface"],
                 fg=C["text_muted"], font=("Arial", 8)).pack(
            side="left", pady=8)
        self.filter_status = make_dark_combo(
            toolbar,
            values=["Всi", "Вiдкрито", "В обробцi", "Закрито", "Вирiшено"],
            width=12)
        self.filter_status.set("Всi")
        self.filter_status.pack(side="left", padx=6, pady=8)
        self.filter_status.bind("<<ComboboxSelected>>",
                                lambda _: self._apply_filter())

        for txt, cmd, bg in [
            ("Переглянути", self._open_selected_detail, C["accent"]),
            ("Дублювати",   self._duplicate_record,     C["bg_surface_alt"]),
            ("Видалити",    self._delete_selected,       C["accent_danger"]),
        ]:
            make_button(toolbar, txt, bg=bg,
                        fg="white" if bg != C["bg_surface_alt"] else C["text_primary"],
                        activebackground=bg, activeforeground="white",
                        font=("Arial", 8), padx=10, pady=3,
                        command=cmd).pack(side="right", padx=4, pady=6)

        tree_f = ttk.Frame(container)
        tree_f.grid(row=1, column=0, sticky="nsew")
        tree_f.rowconfigure(0, weight=1)
        tree_f.columnconfigure(0, weight=1)

        cols = ("id", "entity", "event", "risk", "priority",
                "status", "date", "involved", "desc", "measures")
        self.tree = ttk.Treeview(tree_f, columns=cols,
                                  show="headings", selectmode="browse")
        self.tree.grid(row=0, column=0, sticky="nsew")

        headers = {
            "id":       ("№",           46),
            "entity":   ("Пiдприємство", 155),
            "event":    ("Назва подiї",  185),
            "risk":     ("Тип ризику",   105),
            "priority": ("Прiоритет",     90),
            "status":   ("Статус",         90),
            "date":     ("Дата подiї",    90),
            "involved": ("Задiянi",       130),
            "desc":     ("Опис",          200),
            "measures": ("Заходи",        200),
        }
        for col, (txt, w) in headers.items():
            self.tree.heading(col, text=txt,
                               command=lambda c=col: self._sort_tree(c))
            self.tree.column(col, width=w, anchor="w")

        sy = ttk.Scrollbar(tree_f, orient="vertical",   command=self.tree.yview)
        sx = ttk.Scrollbar(tree_f, orient="horizontal", command=self.tree.xview)
        sy.grid(row=0, column=1, sticky="ns")
        sx.grid(row=1, column=0, sticky="ew")
        self.tree.configure(yscrollcommand=sy.set, xscrollcommand=sx.set)

        self.tree.tag_configure("even", background=C["row_even"])
        self.tree.tag_configure("odd",  background=C["row_odd"])
        for risk, color in RISK_COLORS.items():
            self.tree.tag_configure(f"risk_{risk}", foreground=color)

        self.tree.bind("<<TreeviewSelect>>", self._on_select)
        self.tree.bind("<Double-1>",         self._on_double_click)

        hint_f = tk.Frame(container, bg=C["bg_surface"])
        hint_f.grid(row=2, column=0, sticky="ew")
        tk.Label(hint_f,
                 text="  Подвiйний клiк по рядку — переглянути / редагувати запис",
                 bg=C["bg_surface"], fg=C["text_muted"],
                 font=("Arial", 7, "italic")).pack(
            side="left", padx=8, pady=4)

        detail_f = tk.Frame(container, bg=C["bg_surface"])
        detail_f.grid(row=3, column=0, sticky="ew")
        detail_f.columnconfigure((0, 1), weight=1)

        for ci, (lbl_t, attr) in enumerate(
            [("Опис подiї", "det_desc"),
             ("Вжитi заходи", "det_measures")]
        ):
            sub = tk.Frame(detail_f, bg=C["bg_surface"])
            sub.grid(row=0, column=ci, sticky="nsew",
                     padx=(12 if ci == 0 else 4, 4), pady=8)
            sub.columnconfigure(0, weight=1)
            tk.Label(sub, text=lbl_t, bg=C["bg_surface"],
                     fg=C["text_muted"], font=("Arial", 7, "bold")).grid(
                row=0, column=0, sticky="w")
            t = make_dark_text(sub, height=4, wrap="word", state="disabled")
            t.grid(row=1, column=0, sticky="ew", pady=(2, 0))
            setattr(self, attr, t)

        exp_bar = tk.Frame(container, bg=C["bg_main"])
        exp_bar.grid(row=4, column=0, sticky="ew", padx=8, pady=6)

        make_button(exp_bar, "Експорт CSV",
                    bg=C["bg_surface"], fg=C["text_primary"],
                    activebackground=C["bg_surface_alt"],
                    activeforeground=C["text_primary"],
                    font=("Arial", 8), padx=12, pady=4,
                    command=self._export_csv).pack(side="left", padx=(0, 6))
        if pd:
            make_button(exp_bar, "Експорт Excel",
                        bg=C["accent_success"],
                        activebackground="#16a34a",
                        font=("Arial", 8), padx=12, pady=4,
                        command=self._export_excel).pack(
                side="left", padx=(0, 6))
        make_button(exp_bar, "Iмпорт JSON",
                    bg=C["bg_surface"], fg=C["text_primary"],
                    activebackground=C["bg_surface_alt"],
                    activeforeground=C["text_primary"],
                    font=("Arial", 8), padx=12, pady=4,
                    command=self._import_json).pack(side="left")

    # ------------------------------------------------------------------
    def _on_double_click(self, event: tk.Event) -> None:  # type: ignore[override]
        if not self.tree.selection():
            return
        if self.tree.identify("region", event.x, event.y) != "cell":
            return
        self._open_selected_detail()

    def _open_selected_detail(self) -> None:
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Перегляд", "Оберiть запис для перегляду")
            return
        iid     = sel[0]
        idx_str = self.tree.set(iid, "id")
        rec     = self._find_record(idx_str)
        if not rec:
            return
        EventDetailWindow(
            parent_root=self.frame.winfo_toplevel(),
            record=rec,
            all_records=self.all_records,
            save_callback=lambda old, new: self._on_detail_save(iid, old, new),
            delete_callback=lambda s: self._on_detail_delete(iid, s),
            toast_callback=self._show_toast,
        )

    def _find_record(self, idx_str: str) -> tuple | None:
        for r in self.all_records:
            if (str(r[0]) == idx_str or
                    str(r[0]).lstrip("0") == str(idx_str).lstrip("0")):
                return r
        return None

    def _on_detail_save(self, iid: str, old_id: str,
                        new_record: tuple) -> None:
        for i, r in enumerate(self.all_records):
            if (str(r[0]) == str(old_id) or
                    str(r[0]).lstrip("0") == str(old_id).lstrip("0")):
                self.all_records[i] = new_record
                break
        try:
            self.tree.item(iid, values=(
                new_record[0], new_record[1], new_record[2],
                new_record[4], new_record[9], new_record[10],
                new_record[5], new_record[3], new_record[6],
                new_record[7],
            ))
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
        self.all_records = [r for r in self.all_records
                            if str(r[0]) != str(idx_str)]
        self._recolor_rows()
        self._save_data()
        self._update_count()
        if self.on_data_change:
            self.on_data_change(self.all_records)

    # ------------------------------------------------------------------
    def _sort_tree(self, col: str) -> None:
        data = [(self.tree.set(iid, col), iid)
                for iid in self.tree.get_children("")]
        try:
            data.sort(key=lambda x: (
                float(x[0].replace(",", "")) if x[0] not in ("—", "") else 0))
        except ValueError:
            data.sort(key=lambda x: x[0].lower())
        for i, (_, iid) in enumerate(data):
            self.tree.move(iid, "", i)
        self._recolor_rows()

    def _recolor_rows(self) -> None:
        for i, iid in enumerate(self.tree.get_children()):
            risk     = self.tree.set(iid, "risk")
            base_tag = "even" if i % 2 == 0 else "odd"
            tags     = [base_tag]
            if risk in RISK_COLORS:
                tags.append(f"risk_{risk}")
            self.tree.item(iid, tags=tags)

    # ------------------------------------------------------------------
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
                    row = list(row)
                    if len(row) == 8:
                        row += ["—", "Середнiй", "Вiдкрито"]
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
        data = tuple(list(data) + ["—"] * max(0, 11 - len(data)))
        iid = self.tree.insert("", tk.END, values=(
            data[0], data[1], data[2], data[4], data[9],
            data[10], data[5], data[3], data[6], data[7],
        ))
        self._recolor_rows()
        return iid

    # ------------------------------------------------------------------
    def _get_form_data(self) -> tuple | None:
        detect  = self.ent_detect.get().strip()
        event_d = self.ent_event_date.get().strip()

        for val, label in [(detect, "дати виявлення"),
                           (event_d, "дати подiї")]:
            if val in ("дд.мм.рррр", ""):
                continue
            if not is_valid_date(val):
                messagebox.showwarning(
                    "Помилка",
                    f"Неправильний формат {label} (очiкується дд.мм.рррр)")
                return None

        detect  = "" if detect  == "дд.мм.рррр" else detect
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
            self.cb_status.get().strip()   or "Вiдкрито",
        )

    def _clear_form(self, silent: bool = False) -> None:  # noqa: ARG002
        for w in [self.ent_entity, self.ent_position, self.ent_reporter,
                  self.ent_loss, self.ent_reserve,
                  self.ent_planned, self.ent_refund]:
            w.delete(0, tk.END)
            w.configure(fg=COLORS["text_primary"])
        for w in [self.cb_event, self.cb_risk, self.cb_priority, self.cb_status]:
            w.set("")
        for w in [self.txt_involved, self.txt_impact,
                  self.txt_qualitative, self.txt_description, self.txt_measures]:
            w.delete("1.0", tk.END)
        self.lbl_net.configure(text="0.00", fg=COLORS["accent_success"])
        for e, ph in [(self.ent_detect,     "дд.мм.рррр"),
                      (self.ent_event_date, "дд.мм.рррр")]:
            e.delete(0, tk.END)
            add_placeholder(e, ph)

    def _add_record(self) -> None:
        data = self._get_form_data()
        if not data:
            return
        if not data[1] or not data[2]:
            messagebox.showwarning(
                "Обов'язковi поля",
                "Заповнiть назву пiдприємства та назву подiї")
            return
        self.all_records.append(data)
        self._insert_tree_row(data)
        self._clear_form()
        self._save_data()
        if self.on_data_change:
            self.on_data_change(self.all_records)
        self._update_count()
        self._show_toast("Запис додано")

    # ------------------------------------------------------------------
    def _delete_selected(self) -> None:
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Видалення", "Оберiть запис для видалення")
            return
        iid     = sel[0]
        idx_str = self.tree.set(iid, "id")
        if not messagebox.askyesno(
            "Пiдтвердження",
            f"Видалити запис #{idx_str}?\nЦю дiю не можна скасувати."):
            return
        self.tree.delete(iid)
        self.all_records = [r for r in self.all_records
                            if str(r[0]) != idx_str]
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
        rec = self._find_record(self.tree.set(sel[0], "id"))
        if not rec:
            return
        new_rec = (f"{len(self.all_records)+1:03d}",) + tuple(rec[1:])
        self.all_records.append(new_rec)
        self._insert_tree_row(new_rec)
        self._save_data()
        self._update_count()
        if self.on_data_change:
            self.on_data_change(self.all_records)
        self._show_toast("Запис продубльовано")

    # ------------------------------------------------------------------
    def _apply_filter(self) -> None:
        q      = self.search_var.get().strip().lower()
        risk   = self.filter_risk.get()
        status = self.filter_status.get()
        self.tree.delete(*self.tree.get_children())
        for row in self.all_records:
            if q and q not in " ".join(str(v).lower() for v in row):
                continue
            if risk   != "Всi" and row[4]  != risk:
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

    # ------------------------------------------------------------------
    def _on_select(self, _: object | None = None) -> None:
        sel = self.tree.selection()
        if not sel:
            return
        desc     = self.tree.set(sel[0], "desc")
        measures = self.tree.set(sel[0], "measures")
        for widget, text in [(self.det_desc, desc),
                             (self.det_measures, measures)]:
            widget.configure(state="normal")
            widget.delete("1.0", tk.END)
            widget.insert("1.0", text)
            widget.configure(state="disabled")

    def _update_count(self) -> None:
        self.lbl_count.configure(
            text=f" {len(self.tree.get_children())}")

    def _show_toast(self, msg: str) -> None:
        _show_toast(self.frame, msg)

    # ------------------------------------------------------------------
    def _export_csv(self) -> None:
        if not self.tree.get_children():
            messagebox.showinfo("Експорт", "Таблиця порожня")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV файли", "*.csv")],
            title="Зберегти як CSV")
        if not path:
            return
        try:
            headers = ["ID","Пiдприємство","Назва подiї","Задiянi особи",
                       "Тип ризику","Дата подiї","Опис","Заходи",
                       "Дата виявлення","Прiоритет","Статус"]
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
            title="Зберегти як Excel")
        if not path:
            return
        try:
            headers = ["ID","Пiдприємство","Назва подiї","Задiянi особи",
                       "Тип ризику","Дата подiї","Опис","Заходи",
                       "Дата виявлення","Прiоритет","Статус"]
            n    = len(headers)
            data = [list(r) + [""] * max(0, n - len(r))
                    for r in self.all_records]
            df = pd.DataFrame(data, columns=headers)
            with pd.ExcelWriter(path, engine="openpyxl") as writer:
                df.to_excel(writer, index=False, sheet_name="Реєстр")
                ws = writer.sheets["Реєстр"]
                for col_cells in ws.columns:
                    mx = max(len(str(c.value or "")) for c in col_cells)
                    ws.column_dimensions[
                        col_cells[0].column_letter].width = min(mx + 4, 60)
            self._show_toast("Excel збережено")
        except Exception as e:
            messagebox.showerror("Помилка", str(e))

    def _import_json(self) -> None:
        path = filedialog.askopenfilename(
            filetypes=[("JSON файли", "*.json")], title="Iмпорт JSON")
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
                    row = list(row)           # завжди список, щоб підтримувати присвоєння
                    if len(row) == 8:
                        row += ["—", "Середнiй", "Вiдкрито"]
                    row[0] = f"{len(self.all_records)+1:03d}"
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
#  ВКЛАДКА: АНАЛІТИКА ПОДІЙ
# =============================================================================

class AnalyticsTab:
    """Аналiтика по записах реєстру (дiаграми + таблиця)."""

    def __init__(self, parent: tk.Misc) -> None:
        self.frame   = ttk.Frame(parent)
        self.records: list[tuple] = []
        self._build_ui()

    def _build_ui(self) -> None:
        C = COLORS
        self.frame.rowconfigure(1, weight=1)
        self.frame.columnconfigure(0, weight=1)

        header = tk.Frame(self.frame, bg=C["bg_header"], height=56)
        header.grid(row=0, column=0, sticky="ew")
        header.grid_propagate(False)
        tk.Label(header, text="АНАЛIТИКА ТА ЗВIТИ",
                 bg=C["bg_header"], fg=C["accent_muted"],
                 font=("Arial", 13, "bold")).pack(
            side="left", padx=20, pady=14)
        make_button(header, "Оновити",
                    bg=C["accent"], activebackground=C["accent_soft"],
                    font=("Arial", 9, "bold"), padx=12, pady=4,
                    command=self.refresh).pack(
            side="right", padx=20, pady=12)

        canvas = tk.Canvas(self.frame, bg=C["bg_main"], highlightthickness=0)
        sb = ttk.Scrollbar(self.frame, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=sb.set)
        canvas.grid(row=1, column=0, sticky="nsew")
        sb.grid(row=1, column=1, sticky="ns")

        self.content = tk.Frame(canvas, bg=C["bg_main"])
        cw = canvas.create_window((0, 0), window=self.content, anchor="nw")

        def _conf(_: object) -> None:
            canvas.configure(scrollregion=canvas.bbox("all"))
            canvas.itemconfig(cw, width=canvas.winfo_width())

        self.content.bind("<Configure>", _conf)
        canvas.bind("<Configure>", _conf)
        canvas.bind_all("<MouseWheel>",
                        lambda e: canvas.yview_scroll(
                            int(-1 * e.delta / 120), "units"))
        self.content.columnconfigure(0, weight=1)

        self._build_stat_cards()
        self._build_charts_and_table()

    def _build_stat_cards(self) -> None:
        C = COLORS
        cf = tk.Frame(self.content, bg=C["bg_main"])
        cf.grid(row=0, column=0, sticky="ew", padx=16, pady=(16, 0))
        for i in range(4):
            cf.columnconfigure(i, weight=1)

        self.stat_cards: dict[str, tk.Label] = {}
        defs = [
            ("total",    "Всього записiв",       "0", C["accent"]),
            ("open",     "Вiдкрито / в обробцi", "0", C["accent_danger"]),
            ("critical", "Критичних",             "0", C["accent_warning"]),
            ("closed",   "Закрито / вирiшено",    "0", C["accent_success"]),
        ]
        for ci, (key, title, val, color) in enumerate(defs):
            card = tk.Frame(cf, bg=C["bg_surface"], padx=18, pady=12)
            card.grid(row=0, column=ci, padx=6, sticky="nsew")
            tk.Frame(card, bg=color, height=3).pack(fill="x")
            tk.Label(card, text=title, bg=C["bg_surface"],
                     fg=C["text_muted"], font=("Arial", 8)).pack(
                anchor="w", pady=(8, 2))
            lbl = tk.Label(card, text=val, bg=C["bg_surface"],
                           fg=color, font=("Arial", 22, "bold"))
            lbl.pack(anchor="w")
            self.stat_cards[key] = lbl

    def _build_charts_and_table(self) -> None:
        C = COLORS
        if HAS_MPL:
            charts_row = tk.Frame(self.content, bg=C["bg_main"])
            charts_row.grid(row=1, column=0, sticky="ew", padx=16, pady=16)
            charts_row.columnconfigure((0, 1), weight=1)

            self.fig_left  = Figure(figsize=(5, 3.5), dpi=90,
                                     facecolor=C["bg_surface"])
            self.ax_left   = self.fig_left.add_subplot(111)
            self._style_ax(self.ax_left)
            self.ax_left.set_title("Розподiл за типом ризику",
                                    color=C["text_muted"], fontsize=9)
            fl = tk.Frame(charts_row, bg=C["bg_surface"], padx=8, pady=8)
            fl.grid(row=0, column=0, padx=(0, 8), sticky="nsew")
            self.canvas_left = FigureCanvasTkAgg(self.fig_left, master=fl)
            self.canvas_left.get_tk_widget().pack(fill="both", expand=True)

            self.fig_right = Figure(figsize=(5, 3.5), dpi=90,
                                     facecolor=C["bg_surface"])
            self.ax_right  = self.fig_right.add_subplot(111)
            self._style_ax(self.ax_right)
            self.ax_right.set_title("Записи за статусом",
                                     color=C["text_muted"], fontsize=9)
            fr = tk.Frame(charts_row, bg=C["bg_surface"], padx=8, pady=8)
            fr.grid(row=0, column=1, padx=(8, 0), sticky="nsew")
            self.canvas_right = FigureCanvasTkAgg(self.fig_right, master=fr)
            self.canvas_right.get_tk_widget().pack(fill="both", expand=True)

            self.fig_bottom = Figure(figsize=(10, 3), dpi=90,
                                      facecolor=C["bg_surface"])
            self.ax_bottom  = self.fig_bottom.add_subplot(111)
            self._style_ax(self.ax_bottom)
            self.ax_bottom.set_title(
                "Топ-5 пiдприємств за кiлькiстю подiй",
                color=C["text_muted"], fontsize=9)
            fb = tk.Frame(self.content, bg=C["bg_surface"], padx=8, pady=8)
            fb.grid(row=2, column=0, padx=16, pady=(0, 16), sticky="ew")
            self.canvas_bottom = FigureCanvasTkAgg(self.fig_bottom, master=fb)
            self.canvas_bottom.get_tk_widget().pack(fill="both", expand=True)
        else:
            tk.Label(self.content,
                     text="Встановiть matplotlib:\n  pip install matplotlib",
                     bg=C["bg_main"], fg=C["text_muted"],
                     font=("Arial", 10)).grid(row=1, column=0, pady=40)

        frame = tk.Frame(self.content, bg=C["bg_surface"], padx=16, pady=12)
        frame.grid(row=3, column=0, padx=16, pady=(0, 16), sticky="ew")
        frame.columnconfigure(0, weight=1)
        tk.Label(frame, text="Деталiзована статистика за типом ризику",
                 bg=C["bg_surface"], fg=C["text_muted"],
                 font=("Arial", 9, "bold")).grid(
            row=0, column=0, sticky="w", pady=(0, 8))

        cols = ("risk", "count", "open", "closed")
        self.stats_tree = ttk.Treeview(frame, columns=cols,
                                        show="headings", height=7)
        for col, hdr, w in [("risk", "Тип ризику", 200),
                              ("count", "Всього", 80),
                              ("open",  "Вiдкрито", 80),
                              ("closed","Закрито",  80)]:
            self.stats_tree.heading(col, text=hdr)
            self.stats_tree.column(col, width=w, anchor="w")
        self.stats_tree.grid(row=1, column=0, sticky="ew")

    def _style_ax(self, ax: object) -> None:
        C = COLORS
        ax.set_facecolor(C["bg_surface"])
        ax.tick_params(colors=C["text_muted"], labelsize=8)
        for spine in ax.spines.values():
            spine.set_edgecolor(C["border_soft"])

    def update_data(self, records: list[tuple]) -> None:
        self.records = records
        self.refresh()

    def refresh(self) -> None:
        keys = ["total", "open", "critical", "closed"]
        if not self.records:
            for k in keys:
                self.stat_cards[k].configure(text="0")
            if HAS_MPL:
                self._clear_charts()
            self.stats_tree.delete(*self.stats_tree.get_children())
            return

        C       = COLORS
        records = self.records
        total   = len(records)
        open_c  = sum(1 for r in records
                      if len(r) > 10 and r[10] in ("Вiдкрито", "В обробцi"))
        crit    = sum(1 for r in records
                      if len(r) > 9  and r[9] == "Критичний")
        closed  = sum(1 for r in records
                      if len(r) > 10 and r[10] in ("Закрито", "Вирiшено"))

        self.stat_cards["total"].configure(text=str(total))
        self.stat_cards["open"].configure(text=str(open_c))
        self.stat_cards["critical"].configure(text=str(crit))
        self.stat_cards["closed"].configure(text=str(closed))

        if not HAS_MPL:
            return

        risk_ctr   = Counter(r[4]  for r in records if len(r) > 4)
        status_ctr = Counter(r[10] for r in records if len(r) > 10)
        entity_ctr = Counter(r[1]  for r in records if len(r) > 1 and r[1])

        # Pie
        self.ax_left.clear()
        self._style_ax(self.ax_left)
        self.ax_left.set_title("Розподiл за типом ризику",
                                color=C["text_muted"], fontsize=9)
        if risk_ctr:
            lbls = list(risk_ctr.keys())
            vals = list(risk_ctr.values())
            clrs = [RISK_COLORS.get(l, C["text_muted"]) for l in lbls]
            _, _, autotexts = self.ax_left.pie(
                vals, labels=lbls, autopct="%1.0f%%",
                colors=clrs, startangle=90,
                textprops={"color": C["text_muted"], "fontsize": 7})
            for at in autotexts:
                at.set_fontsize(7); at.set_color("white")
        else:
            self.ax_left.text(0.5, 0.5, "Немає даних",
                              transform=self.ax_left.transAxes,
                              ha="center", va="center", color=C["text_muted"])
        self.canvas_left.draw()

        # Bar
        self.ax_right.clear()
        self._style_ax(self.ax_right)
        self.ax_right.set_title("Записи за статусом",
                                 color=C["text_muted"], fontsize=9)
        if status_ctr:
            s_lbls = list(status_ctr.keys())
            s_vals = list(status_ctr.values())
            s_clrs = [C["accent_danger"], C["accent_warning"],
                      C["accent_success"], C["accent_muted"]][:len(s_lbls)]
            bars = self.ax_right.bar(s_lbls, s_vals,
                                      color=s_clrs, edgecolor="none")
            for bar, val in zip(bars, s_vals, strict=False):
                self.ax_right.text(
                    bar.get_x() + bar.get_width() / 2,
                    bar.get_height() + 0.1,
                    str(val), ha="center", va="bottom",
                    color=C["text_muted"], fontsize=8)
            self.ax_right.tick_params(axis="x", labelrotation=15, labelsize=7)
            self.ax_right.set_ylim(0, max(s_vals) * 1.2 + 1)
        else:
            self.ax_right.text(0.5, 0.5, "Немає даних",
                               transform=self.ax_right.transAxes,
                               ha="center", va="center", color=C["text_muted"])
        self.canvas_right.draw()

        # Barh top-5
        self.ax_bottom.clear()
        self._style_ax(self.ax_bottom)
        self.ax_bottom.set_title(
            "Топ-5 пiдприємств за кiлькiстю подiй",
            color=C["text_muted"], fontsize=9)
        top5 = entity_ctr.most_common(5)
        if top5:
            e_lbls = [e[0][:20] for e in top5]
            e_vals = [e[1] for e in top5]
            bars = self.ax_bottom.barh(e_lbls, e_vals,
                                        color=C["accent"], edgecolor="none")
            for bar, val in zip(bars, e_vals, strict=False):
                self.ax_bottom.text(
                    bar.get_width() + 0.05,
                    bar.get_y() + bar.get_height() / 2,
                    str(val), ha="left", va="center",
                    color=C["text_muted"], fontsize=8)
            self.ax_bottom.tick_params(axis="y", labelsize=8,
                                        colors=C["text_primary"])
        else:
            self.ax_bottom.text(0.5, 0.5, "Немає даних",
                                transform=self.ax_bottom.transAxes,
                                ha="center", va="center",
                                color=C["text_muted"])
        self.canvas_bottom.draw()

        # Таблиця статистики
        self.stats_tree.delete(*self.stats_tree.get_children())
        all_risks = set(RISK_TYPES) | set(r[4] for r in records if len(r) > 4)
        for risk in sorted(all_risks):
            recs  = [r for r in records if len(r) > 4 and r[4] == risk]
            cnt   = len(recs)
            open_ = sum(1 for r in recs
                        if len(r) > 10 and r[10] in ("Вiдкрито", "В обробцi"))
            cl    = sum(1 for r in recs
                        if len(r) > 10 and r[10] in ("Закрито", "Вирiшено"))
            if cnt:
                self.stats_tree.insert("", tk.END,
                                        values=(risk, cnt, open_, cl))

    def _clear_charts(self) -> None:
        if not HAS_MPL:
            return
        for ax in (self.ax_left, self.ax_right, self.ax_bottom):
            ax.clear()
        for cv in (self.canvas_left, self.canvas_right, self.canvas_bottom):
            cv.draw()

    def get_frame(self) -> ttk.Frame:
        return self.frame


# =============================================================================
#  ВКЛАДКА: НАЛАШТУВАННЯ ПОДІЙ
# =============================================================================

class SettingsTab:
    """Вкладка налаштувань модуля 'Реєстр суттєвих подiй'."""

    def __init__(self, parent: tk.Misc) -> None:
        self.frame = ttk.Frame(parent)
        self._build_ui()

    def _build_ui(self) -> None:
        C = COLORS
        self.frame.columnconfigure(0, weight=1)
        self.frame.rowconfigure(1, weight=1)   # виправлено: вага контентного рядка

        header = tk.Frame(self.frame, bg=C["bg_header"], height=56)
        header.grid(row=0, column=0, sticky="ew")
        header.grid_propagate(False)
        tk.Label(header, text="НАЛАШТУВАННЯ РЕЄСТРУ",
                 bg=C["bg_header"], fg=C["text_muted"],
                 font=("Arial", 13, "bold")).pack(
            side="left", padx=20, pady=14)

        content = tk.Frame(self.frame, bg=C["bg_main"])
        content.grid(row=1, column=0, sticky="nsew", padx=40, pady=30)
        content.columnconfigure(0, weight=1)

        self._row(content, 0, "Файл даних:", DATA_FILE, C)
        self._row(content, 1, "Версiя:", "2.1 — Детальне вiкно записiв", C)
        self._row(content, 2, "matplotlib:",
                  "встановлено" if HAS_MPL else "не встановлено", C)
        self._row(content, 3, "pandas:",
                  "встановлено" if pd else "не встановлено", C)

        tk.Label(content, text="Встановлення залежностей:",
                 bg=C["bg_main"], fg=C["text_muted"],
                 font=("Arial", 8, "bold")).grid(
            row=4, column=0, sticky="w", pady=(24, 6))
        tk.Label(content, text="  pip install matplotlib pandas openpyxl",
                 bg=C["bg_surface"], fg=C["accent_muted"],
                 font=("Courier", 9), padx=12, pady=8).grid(
            row=5, column=0, sticky="w")

        tk.Label(content, text="Пiдказки:",
                 bg=C["bg_main"], fg=C["text_muted"],
                 font=("Arial", 8, "bold")).grid(
            row=6, column=0, sticky="w", pady=(24, 6))

        hints = [
            "Подвiйний клiк по рядку таблицi — вiдкрити детальне вiкно запису",
            "У детальному вiкнi доступне редагування та видалення запису",
            "Лiва панель призначена виключно для створення нових записiв",
            "Кнопка 'Переглянути' у тулбарi вiдкриває те саме детальне вiкно",
        ]
        for i, hint in enumerate(hints):
            f = tk.Frame(content, bg=C["bg_main"])
            f.grid(row=7 + i, column=0, sticky="w", pady=2)
            tk.Frame(f, bg=C["accent_success"], width=4, height=4).pack(
                side="left", padx=(0, 8))
            tk.Label(f, text=hint, bg=C["bg_main"], fg=C["text_subtle"],
                     font=("Arial", 8)).pack(side="left")

    def _row(self, parent: tk.Misc, row: int,
             label: str, value: str, C: dict) -> None:
        f = tk.Frame(parent, bg=C["bg_main"])
        f.grid(row=row, column=0, sticky="ew", pady=4)
        tk.Label(f, text=label, bg=C["bg_main"], fg=C["text_muted"],
                 font=("Arial", 9), width=22, anchor="w").pack(side="left")
        tk.Label(f, text=value, bg=C["bg_main"], fg=C["text_primary"],
                 font=("Arial", 9)).pack(side="left")

    def get_frame(self) -> ttk.Frame:
        return self.frame


# =============================================================================
#  СТОРІНКА "РЕЄСТР СУТТЄВИХ ПОДІЙ"
# =============================================================================

class MaterialEventsPage(tk.Frame):
    """Сторiнка 'Реєстр суттєвих подiй' (вкладки + статусбар)."""

    def __init__(self, master: tk.Misc, **kwargs) -> None:
        super().__init__(master, bg=COLORS["bg_main"], **kwargs)

        self.grid_rowconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=0)
        self.grid_columnconfigure(0, weight=1)

        self.notebook = ttk.Notebook(self)
        self.notebook.grid(row=0, column=0, sticky="nsew")

        self.analytics_tab = AnalyticsTab(self.notebook)
        self.registry_tab  = RegistryTab(
            self.notebook,
            on_data_change=self.analytics_tab.update_data)
        self.settings_tab  = SettingsTab(self.notebook)

        self.notebook.add(self.registry_tab.get_frame(),
                          text="  Реєстр подiй  ")
        self.notebook.add(self.analytics_tab.get_frame(),
                          text="  Аналiтика  ")
        self.notebook.add(self.settings_tab.get_frame(),
                          text="  Налаштування  ")

        statusbar = tk.Frame(self, bg=COLORS["bg_header"], height=22)
        statusbar.grid(row=1, column=0, sticky="ew")
        statusbar.grid_propagate(False)

        self._status_lbl = tk.Label(
            statusbar, text="Готово",
            bg=COLORS["bg_header"], fg=COLORS["text_muted"],
            font=("Arial", 7), padx=10)
        self._status_lbl.pack(side="left", pady=3)

        self._time_lbl = tk.Label(
            statusbar, text="",
            bg=COLORS["bg_header"], fg=COLORS["text_muted"],
            font=("Arial", 7), padx=10)
        self._time_lbl.pack(side="right", pady=3)

        self._start_clock()
        self._schedule_autosave()
        self.after(600, lambda: self.analytics_tab.update_data(
            self.registry_tab.all_records))

    def _start_clock(self) -> None:
        self._time_lbl.configure(
            text=datetime.now().strftime("%d.%m.%Y  %H:%M:%S"))
        self.after(1000, self._start_clock)

    def _schedule_autosave(self) -> None:
        try:
            self.registry_tab._save_data()  # noqa: SLF001
            self._status_lbl.configure(
                text=f"Автозбережено о {datetime.now().strftime('%H:%M:%S')}")
        except Exception:
            self._status_lbl.configure(text="Помилка автозбереження")
        self.after(30000, self._schedule_autosave)

    def save_before_exit(self) -> None:
        try:
            self.registry_tab._save_data()  # noqa: SLF001
        except Exception:
            pass


# =============================================================================
#  ДЕТАЛЬНЕ ВІКНО РИЗИКУ
# =============================================================================

class RiskDetailWindow:
    """Спливаюче вiкно для перегляду та редагування запису ризику."""

    RECORD_LEN = 16

    def __init__(
        self,
        parent_root:     tk.Misc,
        record:          tuple,
        all_records:     list[tuple],
        save_callback:   Callable[[str, tuple], None],
        delete_callback: Callable[[str], None],
        toast_callback:  Callable[[str], None],
    ) -> None:
        self.parent_root     = parent_root
        self.record          = list(record)
        self.all_records     = all_records
        self.save_callback   = save_callback
        self.delete_callback = delete_callback
        self.toast_callback  = toast_callback
        self.is_edit_mode    = False
        self._build_window()

    # ------------------------------------------------------------------
    @staticmethod
    def _status_color(status: str) -> str:
        return {
            "Активний":    COLORS["accent_danger"],
            "Монiторинг":  COLORS["accent_warning"],
            "Мiтигований": COLORS["accent"],
            "Закрито":     COLORS["text_muted"],
        }.get(status, COLORS["text_muted"])

    @staticmethod
    def _priority_color(priority: str) -> str:
        return {
            "Критичний": COLORS["accent_danger"],
            "Високий":   COLORS["accent_warning"],
            "Середнiй":  COLORS["accent"],
            "Низький":   COLORS["accent_success"],
        }.get(priority, COLORS["text_primary"])

    # ------------------------------------------------------------------
    def _build_window(self) -> None:
        C = COLORS
        self.win = tk.Toplevel(self.parent_root)
        self.win.title(f"Ризик #{self.record[0]}  —  {self.record[1]}")
        self.win.geometry("780x700")
        self.win.minsize(640, 500)
        self.win.configure(bg=C["bg_main"])
        self.win.grab_set()

        self.win.update_idletasks()
        rx, ry = self.parent_root.winfo_x(), self.parent_root.winfo_y()
        rw, rh = self.parent_root.winfo_width(), self.parent_root.winfo_height()
        ww, wh = 780, 700
        self.win.geometry(f"{ww}x{wh}+{rx+(rw-ww)//2}+{ry+(rh-wh)//2}")

        self.win.rowconfigure(1, weight=1)
        self.win.columnconfigure(0, weight=1)

        # ── Header ──────────────────────────────────────────────────────
        header = tk.Frame(self.win, bg=C["bg_header"], height=58)
        header.grid(row=0, column=0, sticky="ew")
        header.columnconfigure(1, weight=1)
        header.grid_propagate(False)

        score_raw = self.record[7] if len(self.record) > 7 else "0"
        try:
            score_val = int(score_raw)
        except (ValueError, TypeError):
            score_val = 0
        strip_color = _score_color(score_val)
        tk.Frame(header, bg=strip_color, width=4).grid(
            row=0, column=0, sticky="ns")

        title_frame = tk.Frame(header, bg=C["bg_header"])
        title_frame.grid(row=0, column=1, sticky="ew", padx=16, pady=10)

        self.lbl_title = tk.Label(
            title_frame, text=f"Ризик #{self.record[0]}",
            bg=C["bg_header"], fg=C["accent_muted"],
            font=("Arial", 11, "bold"))
        self.lbl_title.pack(anchor="w")

        self.lbl_subtitle = tk.Label(
            title_frame,
            text=self.record[1] if len(self.record) > 1 else "",
            bg=C["bg_header"], fg=C["text_muted"], font=("Arial", 9))
        self.lbl_subtitle.pack(anchor="w")

        status_val = self.record[14] if len(self.record) > 14 else "—"
        sc = self._status_color(status_val)
        self.lbl_status_badge = tk.Label(
            header, text=f"  {status_val}  ",
            bg=sc, fg="white", font=("Arial", 8, "bold"), pady=3)
        self.lbl_status_badge.grid(row=0, column=2, padx=16, pady=18)

        self.lbl_score_badge = tk.Label(
            header,
            text=f"  {_score_label(score_val)} ({score_val})  ",
            bg=strip_color, fg="white",
            font=("Arial", 8, "bold"), pady=3)
        self.lbl_score_badge.grid(row=0, column=3, padx=(0, 16), pady=18)

        # ── Scroll ──────────────────────────────────────────────────────
        canvas = tk.Canvas(self.win, bg=C["bg_main"], highlightthickness=0)
        sb = ttk.Scrollbar(self.win, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=sb.set)
        canvas.grid(row=1, column=0, sticky="nsew")
        sb.grid(row=1, column=1, sticky="ns")

        self.content = tk.Frame(canvas, bg=C["bg_main"])
        self._cw = canvas.create_window((0, 0), window=self.content, anchor="nw")

        def _conf(_: object) -> None:
            canvas.configure(scrollregion=canvas.bbox("all"))
            canvas.itemconfig(self._cw, width=canvas.winfo_width())

        self.content.bind("<Configure>", _conf)
        canvas.bind("<Configure>", _conf)
        canvas.bind_all("<MouseWheel>",
                        lambda e: canvas.yview_scroll(
                            int(-1 * e.delta / 120), "units"))
        self.content.columnconfigure(0, weight=1)
        self.content.columnconfigure(1, weight=1)
        self._build_view_content()

        # ── Bottom buttons ───────────────────────────────────────────────
        btn_bar = tk.Frame(self.win, bg=C["bg_header"], height=46)
        btn_bar.grid(row=2, column=0, columnspan=2, sticky="ew")
        btn_bar.grid_propagate(False)

        left_btns = tk.Frame(btn_bar, bg=C["bg_header"])
        left_btns.pack(side="left", padx=12, pady=8)

        self.btn_edit = make_button(
            left_btns, "Редагувати",
            bg=C["accent_warning"], fg=C["bg_main"],
            activebackground="#d97706", activeforeground="white",
            font=("Arial", 9, "bold"), padx=14, pady=4,
            command=self._toggle_edit_mode)
        self.btn_edit.pack(side="left", padx=(0, 8))

        self.btn_save = make_button(
            left_btns, "Зберегти змiни",
            bg=C["accent_success"],
            activebackground="#16a34a",
            font=("Arial", 9, "bold"), padx=14, pady=4,
            command=self._save_changes)
        self.btn_save.pack_forget()

        self.btn_cancel_edit = make_button(
            left_btns, "Скасувати",
            bg=C["bg_surface"], fg=C["text_muted"],
            activebackground=C["bg_surface_alt"],
            activeforeground=C["text_primary"],
            font=("Arial", 9), padx=12, pady=4,
            command=self._cancel_edit)
        self.btn_cancel_edit.pack_forget()

        right_btns = tk.Frame(btn_bar, bg=C["bg_header"])
        right_btns.pack(side="right", padx=12, pady=8)

        make_button(right_btns, "Видалити",
                    bg=C["accent_danger"],
                    activebackground="#dc2626",
                    font=("Arial", 9, "bold"), padx=14, pady=4,
                    command=self._delete_record).pack(
            side="right", padx=(8, 0))

        make_button(right_btns, "Закрити",
                    bg=C["bg_surface"], fg=C["text_primary"],
                    activebackground=C["bg_surface_alt"],
                    activeforeground=C["text_primary"],
                    font=("Arial", 9), padx=12, pady=4,
                    command=self.win.destroy).pack(side="right")

    # ------------------------------------------------------------------
    def _build_view_content(self) -> None:
        for w in self.content.winfo_children():
            w.destroy()
        r   = self.record
        row = 0

        _build_section_label(self.content,
                              "Iнформацiя про пiдприємство", row)
        row += 1

        _build_info_cell(self.content, "Пiдприємство",
                         r[1] if len(r) > 1 else "—", row, 0)
        p_val = r[13] if len(r) > 13 else "—"
        _build_info_cell(self.content, "Прiоритет", p_val, row, 1,
                         value_color=self._priority_color(p_val))
        row += 1

        _build_section_label(self.content, "Опис ризику", row)
        row += 1

        _build_info_cell(self.content, "Назва ризику",
                         r[2] if len(r) > 2 else "—", row, 0)
        rt_val = r[4] if len(r) > 4 else "—"
        _build_info_cell(self.content, "Тип ризику", rt_val, row, 1,
                         value_color=RISK_COLORS.get(rt_val))
        row += 1

        _build_info_cell(self.content, "Категорiя ризику",
                         r[3] if len(r) > 3 else "—", row, 0)
        _build_info_cell(self.content, "Власник ризику",
                         r[8] if len(r) > 8 else "—", row, 1)
        row += 1

        _build_section_label(self.content, "Оцiнка ризику", row)
        row += 1

        _build_info_cell(self.content, "Iмовiрнiсть",
                         r[5] if len(r) > 5 else "—", row, 0)
        _build_info_cell(self.content, "Вплив",
                         r[6] if len(r) > 6 else "—", row, 1)
        row += 1

        # Score (full-width)
        score_raw = r[7] if len(r) > 7 else "0"
        try:
            score_int = int(score_raw)
        except (ValueError, TypeError):
            score_int = 0
        C       = COLORS
        sc_cell = tk.Frame(self.content, bg=C["bg_surface"], padx=10, pady=8)
        sc_cell.grid(row=row, column=0, columnspan=2, sticky="nsew",
                     padx=8, pady=3)
        sc_cell.columnconfigure(0, weight=1)
        tk.Label(sc_cell,
                 text="Рiвень ризику (Score = Iмовiрнiсть × Вплив)",
                 bg=C["bg_surface"], fg=C["text_subtle"],
                 font=("Arial", 7, "bold")).grid(row=0, column=0, sticky="w")
        tk.Label(sc_cell,
                 text=f"{score_int}  —  {_score_label(score_int)}",
                 bg=C["bg_surface"], fg=_score_color(score_int),
                 font=("Arial", 14, "bold")).grid(
            row=1, column=0, sticky="w", pady=(2, 0))
        row += 1

        res_raw = r[10] if len(r) > 10 else "—"
        res_color = _score_color(int(res_raw)) if str(res_raw).isdigit() else None
        _build_info_cell(self.content, "Залишковий ризик",
                         res_raw, row, 0, value_color=res_color)

        status_val = r[14] if len(r) > 14 else "—"
        _build_info_cell(self.content, "Статус", status_val, row, 1,
                         value_color=self._status_color(status_val))
        row += 1

        _build_section_label(self.content, "Дати", row)
        row += 1
        _build_info_cell(self.content, "Дата виявлення",
                         r[11] if len(r) > 11 else "—", row, 0)
        _build_info_cell(self.content, "Дата перегляду",
                         r[12] if len(r) > 12 else "—", row, 1)
        row += 1

        _build_section_label(self.content, "Деталi", row)
        row += 1
        _build_text_block(self.content, "Заходи контролю",
                          r[9]  if len(r) > 9  else "—", row)
        row += 1
        _build_text_block(self.content, "Детальний опис ризику",
                          r[15] if len(r) > 15 else "—", row)
        row += 1
        tk.Frame(self.content, bg=C["bg_main"], height=12).grid(
            row=row, column=0, columnspan=2)

    # ------------------------------------------------------------------
    def _build_edit_content(self) -> None:
        C   = COLORS
        rec = self.record
        for w in self.content.winfo_children():
            w.destroy()
        self.content.columnconfigure(0, weight=1)

        def section(txt: str, r: int) -> int:
            f = tk.Frame(self.content, bg=C["bg_main"])
            f.grid(row=r, column=0, sticky="ew", padx=8, pady=(14, 4))
            tk.Frame(f, bg=C["accent_warning"], width=3, height=16).pack(
                side="left")
            tk.Label(f, text=txt, bg=C["bg_main"], fg=C["accent_warning"],
                     font=("Arial", 9, "bold")).pack(side="left", padx=8)
            return r + 1

        def lbl(text: str, r: int) -> None:
            tk.Label(self.content, text=text, bg=C["bg_main"],
                     fg=C["text_subtle"], font=("Arial", 8)).grid(
                row=r, column=0, sticky="w", padx=10, pady=(6, 0))

        def mk_e(**kw: object) -> tk.Entry:
            return make_dark_entry(self.content,
                                   accent=C["accent_warning"], **kw)

        def mk_c(values: list[str] | None = None,
                 **kw: object) -> ttk.Combobox:
            return make_dark_combo(self.content, values=values, **kw)

        row = 0
        row = section("Пiдприємство та ризик", row)

        lbl("Скорочена назва пiдприємства:", row); row += 1
        self.e_entity = mk_e()
        self.e_entity.grid(row=row, column=0, sticky="ew",
                           padx=10, pady=(2, 0))
        self.e_entity.insert(0, rec[1] if len(rec) > 1 else "")
        row += 1

        lbl("Назва ризику:", row); row += 1
        self.e_risk_name = mk_e()
        self.e_risk_name.grid(row=row, column=0, sticky="ew",
                              padx=10, pady=(2, 0))
        self.e_risk_name.insert(0, rec[2] if len(rec) > 2 else "")
        row += 1

        lbl("Категорiя ризику:", row); row += 1
        self.e_category = mk_c(values=RISK_CATEGORIES)
        self.e_category.grid(row=row, column=0, sticky="ew",
                             padx=10, pady=(2, 0))
        self.e_category.set(rec[3] if len(rec) > 3 else "")
        row += 1

        lbl("Тип ризику:", row); row += 1
        self.e_risk_type = mk_c(values=RISK_TYPES)
        self.e_risk_type.grid(row=row, column=0, sticky="ew",
                              padx=10, pady=(2, 0))
        self.e_risk_type.set(rec[4] if len(rec) > 4 and rec[4] != "—" else "")
        row += 1

        lbl("Власник ризику:", row); row += 1
        self.e_owner = mk_e()
        self.e_owner.grid(row=row, column=0, sticky="ew",
                          padx=10, pady=(2, 0))
        self.e_owner.insert(0, rec[8] if len(rec) > 8 else "")
        row += 1

        row = section("Оцiнка ризику", row)
        score_f = tk.Frame(self.content, bg=C["bg_main"])
        score_f.grid(row=row, column=0, sticky="ew", padx=10, pady=4)
        score_f.columnconfigure((0, 1, 2), weight=1)
        row += 1

        for ci, (lbl_t, attr, vals, vi) in enumerate([
            ("Iмовiрнiсть:", "e_prob",   PROBABILITY_LEVELS, 5),
            ("Вплив:",        "e_impact", IMPACT_LEVELS,      6),
        ]):
            tk.Label(score_f, text=lbl_t, bg=C["bg_main"],
                     fg=C["text_subtle"], font=("Arial", 8)).grid(
                row=0, column=ci, sticky="w",
                padx=(0 if ci == 0 else 16, 0))
            combo = make_dark_combo(score_f, values=vals, width=22)
            combo.grid(row=1, column=ci, sticky="ew",
                       padx=(0 if ci == 0 else 16, 0), pady=2)
            cur     = rec[vi] if len(rec) > vi else ""
            matched = next((v for v in vals if v.startswith(str(cur)[:1])), "")
            combo.set(cur if cur in vals else matched)
            setattr(self, attr, combo)

        self.lbl_live_score = tk.Label(
            score_f, text="Score: —",
            bg=C["bg_main"], fg=C["text_muted"],
            font=("Arial", 11, "bold"))
        self.lbl_live_score.grid(row=1, column=2, padx=20)

        def _upd_score(_: object = None) -> None:
            try:
                s   = _extract_num(self.e_prob.get()) * \
                      _extract_num(self.e_impact.get())
                col = _score_color(s)
                self.lbl_live_score.configure(
                    text=f"Score: {s}  ({_score_label(s)})", fg=col)
            except Exception:
                self.lbl_live_score.configure(
                    text="Score: —", fg=C["text_muted"])

        self.e_prob.bind("<<ComboboxSelected>>",   _upd_score)
        self.e_impact.bind("<<ComboboxSelected>>", _upd_score)
        _upd_score()

        lbl("Залишковий ризик (1–25):", row); row += 1
        self.e_residual = mk_e()
        self.e_residual.grid(row=row, column=0, sticky="w",
                             padx=10, pady=(2, 0))
        self.e_residual.insert(0, rec[10] if len(rec) > 10 else "")
        row += 1

        row = section("Дати", row)
        date_f = tk.Frame(self.content, bg=C["bg_main"])
        date_f.grid(row=row, column=0, sticky="w", padx=10, pady=4)
        row += 1

        for ci, (lbl_t, attr, vi) in enumerate([
            ("Дата виявлення:", "e_date_id",  11),
            ("Дата перегляду:", "e_date_rev", 12),
        ]):
            tk.Label(date_f, text=lbl_t, bg=C["bg_main"],
                     fg=C["text_subtle"], font=("Arial", 8)).grid(
                row=0, column=ci,
                padx=(0 if ci == 0 else 20, 0), sticky="w")
            e = make_dark_entry(date_f, accent=C["accent_warning"], width=14)
            e.grid(row=1, column=ci,
                   padx=(0 if ci == 0 else 20, 0), pady=2)
            val = rec[vi] if len(rec) > vi else ""
            e.insert(0, val) if (val and val != "—") \
                else add_placeholder(e, "дд.мм.рррр")
            setattr(self, attr, e)

        row = section("Прiоритет та статус", row)
        ps_f = tk.Frame(self.content, bg=C["bg_main"])
        ps_f.grid(row=row, column=0, sticky="w", padx=10, pady=4)
        row += 1

        tk.Label(ps_f, text="Прiоритет:", bg=C["bg_main"],
                 fg=C["text_subtle"], font=("Arial", 8)).grid(
            row=0, column=0, sticky="w")
        self.e_priority = make_dark_combo(
            ps_f,
            values=["Критичний", "Високий", "Середнiй", "Низький"],
            width=14)
        self.e_priority.grid(row=1, column=0, pady=(2, 0), padx=(0, 20))
        self.e_priority.set(rec[13] if len(rec) > 13 else "Середнiй")

        tk.Label(ps_f, text="Статус:", bg=C["bg_main"],
                 fg=C["text_subtle"], font=("Arial", 8)).grid(
            row=0, column=1, sticky="w")
        self.e_status = make_dark_combo(
            ps_f,
            values=["Активний", "Монiторинг", "Мiтигований", "Закрито"],
            width=14)
        self.e_status.grid(row=1, column=1, pady=(2, 0))
        self.e_status.set(rec[14] if len(rec) > 14 else "Активний")

        row = section("Деталi", row)

        lbl("Заходи контролю:", row); row += 1
        self.e_controls = make_dark_text(self.content, height=3, wrap="word")
        self.e_controls.grid(row=row, column=0, sticky="ew",
                             padx=10, pady=(2, 0))
        if len(rec) > 9 and rec[9] and rec[9] != "—":
            self.e_controls.insert("1.0", rec[9])
        row += 1

        lbl("Детальний опис ризику:", row); row += 1
        self.e_description = make_dark_text(self.content, height=4, wrap="word")
        self.e_description.grid(row=row, column=0, sticky="ew",
                                padx=10, pady=(2, 0))
        if len(rec) > 15 and rec[15] and rec[15] != "—":
            self.e_description.insert("1.0", rec[15])
        row += 1

        tk.Frame(self.content, bg=C["bg_main"], height=16).grid(
            row=row, column=0)

    # ------------------------------------------------------------------
    def _toggle_edit_mode(self) -> None:
        self.is_edit_mode = True
        self._build_edit_content()
        self.btn_edit.pack_forget()
        self.btn_save.pack(side="left", padx=(0, 8))
        self.btn_cancel_edit.pack(side="left")
        self.lbl_title.configure(
            text=f"Редагування ризику #{self.record[0]}",
            fg=COLORS["accent_warning"])

    def _cancel_edit(self) -> None:
        self.is_edit_mode = False
        self._build_view_content()
        self.btn_save.pack_forget()
        self.btn_cancel_edit.pack_forget()
        self.btn_edit.pack(side="left", padx=(0, 8))
        self.lbl_title.configure(
            text=f"Ризик #{self.record[0]}",
            fg=COLORS["accent_muted"])

    def _save_changes(self) -> None:
        date_id  = self.e_date_id.get().strip()
        date_rev = self.e_date_rev.get().strip()
        for val, label in [(date_id,  "дати виявлення"),
                           (date_rev, "дати перегляду")]:
            if val in ("дд.мм.рррр", ""):
                continue
            if not is_valid_date(val):
                messagebox.showwarning(
                    "Помилка",
                    f"Неправильний формат {label} (очiкується дд.мм.рррр)",
                    parent=self.win)
                return
        date_id  = "" if date_id  == "дд.мм.рррр" else date_id
        date_rev = "" if date_rev == "дд.мм.рррр" else date_rev

        entity    = self.e_entity.get().strip()
        risk_name = self.e_risk_name.get().strip()
        if not entity or not risk_name:
            messagebox.showwarning(
                "Обов'язковi поля",
                "Заповнiть назву пiдприємства та назву ризику",
                parent=self.win)
            return

        prob_str   = self.e_prob.get().strip()
        impact_str = self.e_impact.get().strip()
        score      = _extract_num(prob_str) * _extract_num(impact_str)

        try:
            residual = int(self.e_residual.get().strip() or "0")
        except ValueError:
            residual = 0

        old_id = self.record[0]
        new_record = (
            self.record[0],
            entity,
            risk_name,
            self.e_category.get().strip()  or "—",
            self.e_risk_type.get().strip() or "—",
            prob_str   or "—",
            impact_str or "—",
            str(score),
            self.e_owner.get().strip(),
            self.e_controls.get("1.0", tk.END).strip(),
            str(residual),
            date_id  or "—",
            date_rev or "—",
            self.e_priority.get().strip() or "Середнiй",
            self.e_status.get().strip()   or "Активний",
            self.e_description.get("1.0", tk.END).strip(),
        )
        self.record = list(new_record)
        self.save_callback(str(old_id), new_record)

        self.lbl_subtitle.configure(text=entity)
        sc  = self._status_color(new_record[14])
        sc2 = _score_color(score)
        self.lbl_status_badge.configure(
            text=f"  {new_record[14]}  ", bg=sc)
        self.lbl_score_badge.configure(
            text=f"  {_score_label(score)} ({score})  ", bg=sc2)
        self._cancel_edit()
        self.toast_callback("Запис оновлено")

    def _delete_record(self) -> None:
        idx_str = self.record[0]
        if not messagebox.askyesno(
            "Пiдтвердження",
            f"Видалити ризик #{idx_str}?\nЦю дiю не можна скасувати.",
            parent=self.win):
            return
        self.delete_callback(str(idx_str))
        self.toast_callback("Запис видалено")
        self.win.destroy()


# =============================================================================
#  ВКЛАДКА: РЕЄСТР РИЗИКІВ
# =============================================================================

class RiskRegistryTab:
    """Вкладка 'Реєстр ризикiв'."""

    EMPTY_RECORD_LEN = 16

    def __init__(
        self,
        parent: tk.Misc,
        on_data_change: Callable[[list[tuple]], None] | None = None,
    ) -> None:
        self.parent         = parent
        self.on_data_change = on_data_change
        self.frame          = ttk.Frame(parent)
        self.all_records:   list[tuple] = []

        self._build_ui()
        self._load_data()

    # ------------------------------------------------------------------
    def _build_ui(self) -> None:
        C = COLORS
        self.frame.rowconfigure(1, weight=1)
        self.frame.columnconfigure(0, weight=1)

        header = tk.Frame(self.frame, bg=C["bg_header"], height=56)
        header.grid(row=0, column=0, sticky="ew")
        header.columnconfigure(1, weight=1)
        header.grid_propagate(False)

        tk.Label(header, text="РЕЄСТР РИЗИКIВ",
                 bg=C["bg_header"], fg=C["accent_muted"],
                 font=("Arial", 13, "bold")).grid(
            row=0, column=0, padx=20, pady=14, sticky="w")

        sf = tk.Frame(header, bg=C["bg_header"])
        sf.grid(row=0, column=1, sticky="e", padx=20)
        tk.Label(sf, text="Пошук:", bg=C["bg_header"],
                 fg=C["text_muted"], font=("Arial", 8)).pack(
            side="left", padx=(0, 6))

        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *_: self._apply_filter())
        tk.Entry(sf, textvariable=self.search_var,
                 bg=C["bg_input"], fg=C["text_primary"],
                 insertbackground=C["text_primary"],
                 relief="flat", bd=2, font=("Arial", 9),
                 width=34).pack(side="left", padx=(0, 8), ipady=2)
        make_button(sf, "Скинути",
                    bg=C["bg_surface"], fg=C["text_muted"],
                    activebackground=C["bg_surface_alt"],
                    activeforeground=C["text_primary"],
                    font=("Arial", 8), padx=8, pady=2,
                    command=self._reset_filter).pack(side="left")

        paned = ttk.PanedWindow(self.frame, orient="horizontal")
        paned.grid(row=1, column=0, sticky="nsew")
        lw = ttk.Frame(paned)
        rw = ttk.Frame(paned)
        paned.add(lw, weight=4)
        paned.add(rw, weight=7)

        self._build_form(lw)
        self._build_table(rw)

    # ------------------------------------------------------------------
    def _build_form(self, container: tk.Misc) -> None:
        C = COLORS
        container.columnconfigure(0, weight=1)
        container.rowconfigure(0, weight=1)

        canvas = tk.Canvas(container, bg=C["bg_main"], highlightthickness=0)
        sb = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=sb.set)
        canvas.grid(row=0, column=0, sticky="nsew")
        sb.grid(row=0, column=1, sticky="ns")

        form = tk.Frame(canvas, bg=C["bg_main"])
        fw   = canvas.create_window((0, 0), window=form, anchor="nw")

        def _conf(_: object) -> None:
            canvas.configure(scrollregion=canvas.bbox("all"))
            canvas.itemconfig(fw, width=canvas.winfo_width())

        form.bind("<Configure>", _conf)
        canvas.bind("<Configure>", _conf)
        canvas.bind_all("<MouseWheel>",
                        lambda e: canvas.yview_scroll(
                            int(-1 * e.delta / 120), "units"))
        form.columnconfigure(0, weight=1)

        def section(txt: str, r: int) -> int:
            f = tk.Frame(form, bg=C["bg_main"])
            f.grid(row=r, column=0, sticky="ew", padx=16, pady=(16, 4))
            tk.Frame(f, bg=C["accent"], width=3, height=16).pack(side="left")
            tk.Label(f, text=txt, bg=C["bg_main"], fg=C["accent"],
                     font=("Arial", 9, "bold")).pack(side="left", padx=8)
            return r + 1

        def lbl(text: str, r: int) -> None:
            tk.Label(form, text=text, bg=C["bg_main"],
                     fg=C["text_subtle"], font=("Arial", 8)).grid(
                row=r, column=0, sticky="w", padx=16, pady=(4, 0))

        def mk_e(**kw: object) -> tk.Entry:
            return make_dark_entry(form, **kw)

        def mk_c(values: list[str] | None = None,
                 **kw: object) -> ttk.Combobox:
            return make_dark_combo(form, values=values, **kw)

        def field(lbl_txt: str, r: int,
                  factory: Callable) -> tuple[tk.Widget, int]:
            lbl(lbl_txt, r)
            w = factory()
            w.grid(row=r + 1, column=0, sticky="ew", padx=16, pady=(2, 0))
            return w, r + 2

        row = 0

        badge_f = tk.Frame(form, bg=C["bg_main"])
        badge_f.grid(row=row, column=0, sticky="ew", padx=16, pady=(12, 0))
        tk.Label(badge_f, text="  + НОВИЙ РИЗИК  ",
                 bg=C["accent"], fg="white",
                 font=("Arial", 8, "bold"), pady=3).pack(anchor="w")
        row += 1

        row = section("Iнформацiя про пiдприємство", row)
        self.ent_entity, row = field(
            "Скорочена назва пiдприємства:", row, mk_e)
        self.ent_owner, row = field("Власник ризику:", row, mk_e)

        row = section("Опис ризику", row)
        self.ent_risk_name, row = field("Назва ризику:", row, mk_e)
        self.cb_category, row = field(
            "Категорiя ризику:", row,
            lambda: mk_c(values=RISK_CATEGORIES))
        self.cb_risk_type, row = field(
            "Тип ризику:", row,
            lambda: mk_c(values=RISK_TYPES))

        row = section("Оцiнка ризику", row)
        score_f = tk.Frame(form, bg=C["bg_main"])
        score_f.grid(row=row, column=0, sticky="ew", padx=16, pady=4)
        score_f.columnconfigure((0, 1), weight=1)
        row += 1

        for ci, (lbl_t, attr, vals) in enumerate([
            ("Iмовiрнiсть (1–5):", "cb_prob",   PROBABILITY_LEVELS),
            ("Вплив (1–5):",        "cb_impact", IMPACT_LEVELS),
        ]):
            tk.Label(score_f, text=lbl_t, bg=C["bg_main"],
                     fg=C["text_subtle"], font=("Arial", 8)).grid(
                row=0, column=ci, sticky="w",
                padx=(0 if ci == 0 else 10, 0))
            combo = make_dark_combo(score_f, values=vals)
            combo.grid(row=1, column=ci, sticky="ew",
                       padx=(0 if ci == 0 else 10, 0), pady=2)
            setattr(self, attr, combo)

        net_f = tk.Frame(form, bg=C["bg_main"])
        net_f.grid(row=row, column=0, sticky="ew", padx=16, pady=(2, 4))
        row += 1
        tk.Label(net_f, text="Рiвень ризику (Score):",
                 bg=C["bg_main"], fg=C["text_muted"],
                 font=("Arial", 8)).pack(side="left")
        self.lbl_score = tk.Label(net_f, text="—",
                                   bg=C["bg_main"], fg=C["accent_success"],
                                   font=("Arial", 13, "bold"))
        self.lbl_score.pack(side="left", padx=10)

        def _upd_score(_: object = None) -> None:
            try:
                s = _extract_num(self.cb_prob.get()) * \
                    _extract_num(self.cb_impact.get())
                self.lbl_score.configure(
                    text=f"{s}  ({_score_label(s)})",
                    fg=_score_color(s))
            except Exception:
                self.lbl_score.configure(text="—", fg=C["text_muted"])

        self.cb_prob.bind("<<ComboboxSelected>>",   _upd_score)
        self.cb_impact.bind("<<ComboboxSelected>>", _upd_score)

        lbl("Залишковий ризик (1–25):", row); row += 1
        self.ent_residual = mk_e()
        self.ent_residual.grid(row=row, column=0, sticky="ew",
                               padx=16, pady=(2, 0))
        row += 1

        row = section("Дати", row)
        date_f = tk.Frame(form, bg=C["bg_main"])
        date_f.grid(row=row, column=0, sticky="w", padx=16, pady=4)
        row += 1
        for ci, (lbl_t, attr) in enumerate([
            ("Дата виявлення:", "ent_date_id"),
            ("Дата перегляду:", "ent_date_rev"),
        ]):
            tk.Label(date_f, text=lbl_t, bg=C["bg_main"],
                     fg=C["text_subtle"], font=("Arial", 8)).grid(
                row=0, column=ci,
                padx=(0 if ci == 0 else 20, 0), sticky="w")
            e = make_dark_entry(date_f, width=14)
            e.grid(row=1, column=ci,
                   padx=(0 if ci == 0 else 20, 0), pady=2)
            add_placeholder(e, "дд.мм.рррр")
            setattr(self, attr, e)

        lbl("Прiоритет:", row); row += 1
        self.cb_priority = mk_c(
            values=["Критичний", "Високий", "Середнiй", "Низький"])
        self.cb_priority.grid(row=row, column=0, sticky="w",
                              padx=16, pady=(2, 0))
        row += 1

        lbl("Статус:", row); row += 1
        self.cb_status = mk_c(
            values=["Активний", "Монiторинг", "Мiтигований", "Закрито"])
        self.cb_status.grid(row=row, column=0, sticky="w",
                            padx=16, pady=(2, 0))
        row += 1

        row = section("Деталi", row)
        for lbl_t, attr, h in [
            ("Заходи контролю:",       "txt_controls",    3),
            ("Детальний опис ризику:", "txt_description", 4),
        ]:
            lbl(lbl_t, row); row += 1
            t = make_dark_text(form, height=h, wrap="word")
            t.grid(row=row, column=0, sticky="ew", padx=16, pady=(2, 0))
            setattr(self, attr, t)
            row += 1

        btn_f = tk.Frame(form, bg=C["bg_main"])
        btn_f.grid(row=row, column=0, sticky="ew", padx=16, pady=(16, 20))
        btn_f.columnconfigure((0, 1), weight=1)

        make_button(btn_f, "Очистити",
                    bg=C["bg_surface"], fg=C["text_muted"],
                    activebackground=C["bg_surface_alt"],
                    activeforeground=C["text_primary"],
                    padx=14, pady=6,
                    command=self._clear_form).grid(
            row=0, column=0, padx=4, sticky="ew")
        make_button(btn_f, "Додати ризик",
                    bg=C["accent"],
                    activebackground=C["accent_soft"],
                    font=("Arial", 9, "bold"), padx=14, pady=6,
                    command=self._add_record).grid(
            row=0, column=1, padx=4, sticky="ew")

    # ------------------------------------------------------------------
    def _build_table(self, container: tk.Misc) -> None:
        C = COLORS
        container.rowconfigure(1, weight=1)
        container.columnconfigure(0, weight=1)

        toolbar = tk.Frame(container, bg=C["bg_surface"], height=40)
        toolbar.grid(row=0, column=0, sticky="ew")
        toolbar.grid_propagate(False)

        tk.Label(toolbar, text="Записи", bg=C["bg_surface"],
                 fg=C["text_muted"], font=("Arial", 8)).pack(
            side="left", padx=12, pady=8)
        self.lbl_count = tk.Label(toolbar, text="0", bg=C["bg_surface"],
                                   fg=C["accent"], font=("Arial", 8, "bold"))
        self.lbl_count.pack(side="left", pady=8)

        tk.Label(toolbar, text="  |  Тип:", bg=C["bg_surface"],
                 fg=C["text_muted"], font=("Arial", 8)).pack(
            side="left", pady=8)
        self.filter_type = make_dark_combo(
            toolbar, values=["Всi"] + RISK_TYPES, width=16)
        self.filter_type.set("Всi")
        self.filter_type.pack(side="left", padx=6, pady=8)
        self.filter_type.bind("<<ComboboxSelected>>",
                              lambda _: self._apply_filter())

        tk.Label(toolbar, text="Статус:", bg=C["bg_surface"],
                 fg=C["text_muted"], font=("Arial", 8)).pack(
            side="left", pady=8)
        self.filter_status = make_dark_combo(
            toolbar,
            values=["Всi", "Активний", "Монiторинг", "Мiтигований", "Закрито"],
            width=12)
        self.filter_status.set("Всi")
        self.filter_status.pack(side="left", padx=6, pady=8)
        self.filter_status.bind("<<ComboboxSelected>>",
                                lambda _: self._apply_filter())

        for txt, cmd, bg in [
            ("Переглянути", self._open_selected_detail, C["accent"]),
            ("Дублювати",   self._duplicate_record,     C["bg_surface_alt"]),
            ("Видалити",    self._delete_selected,       C["accent_danger"]),
        ]:
            make_button(toolbar, txt, bg=bg,
                        fg="white" if bg != C["bg_surface_alt"] else C["text_primary"],
                        activebackground=bg, activeforeground="white",
                        font=("Arial", 8), padx=10, pady=3,
                        command=cmd).pack(side="right", padx=4, pady=6)

        tree_f = ttk.Frame(container)
        tree_f.grid(row=1, column=0, sticky="nsew")
        tree_f.rowconfigure(0, weight=1)
        tree_f.columnconfigure(0, weight=1)

        cols = ("id", "entity", "risk_name", "category", "risk_type",
                "score", "priority", "status", "owner", "date_id")
        self.tree = ttk.Treeview(tree_f, columns=cols,
                                  show="headings", selectmode="browse")
        self.tree.grid(row=0, column=0, sticky="nsew")

        headers = {
            "id":        ("№",            46),
            "entity":    ("Пiдприємство", 150),
            "risk_name": ("Назва ризику", 200),
            "category":  ("Категорiя",    110),
            "risk_type": ("Тип ризику",   110),
            "score":     ("Score",          62),
            "priority":  ("Прiоритет",      88),
            "status":    ("Статус",          90),
            "owner":     ("Власник",        130),
            "date_id":   ("Дата виявл.",   100),
        }
        for col, (txt, w) in headers.items():
            self.tree.heading(col, text=txt,
                               command=lambda c=col: self._sort_tree(c))
            self.tree.column(col, width=w, anchor="w")

        sy = ttk.Scrollbar(tree_f, orient="vertical",   command=self.tree.yview)
        sx = ttk.Scrollbar(tree_f, orient="horizontal", command=self.tree.xview)
        sy.grid(row=0, column=1, sticky="ns")
        sx.grid(row=1, column=0, sticky="ew")
        self.tree.configure(yscrollcommand=sy.set, xscrollcommand=sx.set)

        self.tree.tag_configure("even", background=C["row_even"])
        self.tree.tag_configure("odd",  background=C["row_odd"])
        for risk, color in RISK_COLORS.items():
            self.tree.tag_configure(f"risk_{risk}", foreground=color)
        for tag, color in [
            ("score_low",  COLORS["accent_success"]),
            ("score_mod",  COLORS["accent_warning"]),
            ("score_high", "#f97316"),
            ("score_crit", COLORS["accent_danger"]),
        ]:
            self.tree.tag_configure(tag, foreground=color)

        self.tree.bind("<<TreeviewSelect>>", self._on_select)
        self.tree.bind("<Double-1>",         self._on_double_click)

        hint_f = tk.Frame(container, bg=C["bg_surface"])
        hint_f.grid(row=2, column=0, sticky="ew")
        tk.Label(hint_f,
                 text="  Подвiйний клiк по рядку — переглянути / редагувати ризик",
                 bg=C["bg_surface"], fg=C["text_muted"],
                 font=("Arial", 7, "italic")).pack(
            side="left", padx=8, pady=4)

        detail_f = tk.Frame(container, bg=C["bg_surface"])
        detail_f.grid(row=3, column=0, sticky="ew")
        detail_f.columnconfigure((0, 1), weight=1)

        for ci, (lbl_t, attr) in enumerate([
            ("Заходи контролю", "det_controls"),
            ("Опис ризику",     "det_desc"),
        ]):
            sub = tk.Frame(detail_f, bg=C["bg_surface"])
            sub.grid(row=0, column=ci, sticky="nsew",
                     padx=(12 if ci == 0 else 4, 4), pady=8)
            sub.columnconfigure(0, weight=1)
            tk.Label(sub, text=lbl_t, bg=C["bg_surface"],
                     fg=C["text_muted"], font=("Arial", 7, "bold")).grid(
                row=0, column=0, sticky="w")
            t = make_dark_text(sub, height=4, wrap="word", state="disabled")
            t.grid(row=1, column=0, sticky="ew", pady=(2, 0))
            setattr(self, attr, t)

        exp_bar = tk.Frame(container, bg=C["bg_main"])
        exp_bar.grid(row=4, column=0, sticky="ew", padx=8, pady=6)

        make_button(exp_bar, "Експорт CSV",
                    bg=C["bg_surface"], fg=C["text_primary"],
                    activebackground=C["bg_surface_alt"],
                    activeforeground=C["text_primary"],
                    font=("Arial", 8), padx=12, pady=4,
                    command=self._export_csv).pack(side="left", padx=(0, 6))
        if pd:
            make_button(exp_bar, "Експорт Excel",
                        bg=C["accent_success"],
                        activebackground="#16a34a",
                        font=("Arial", 8), padx=12, pady=4,
                        command=self._export_excel).pack(
                side="left", padx=(0, 6))
        make_button(exp_bar, "Iмпорт JSON",
                    bg=C["bg_surface"], fg=C["text_primary"],
                    activebackground=C["bg_surface_alt"],
                    activeforeground=C["text_primary"],
                    font=("Arial", 8), padx=12, pady=4,
                    command=self._import_json).pack(side="left")

    # ------------------------------------------------------------------
    def _on_double_click(self, event: tk.Event) -> None:  # type: ignore[override]
        if not self.tree.selection():
            return
        if self.tree.identify("region", event.x, event.y) != "cell":
            return
        self._open_selected_detail()

    def _open_selected_detail(self) -> None:
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Перегляд", "Оберiть запис для перегляду")
            return
        iid     = sel[0]
        idx_str = self.tree.set(iid, "id")
        rec     = self._find_record(idx_str)
        if not rec:
            return
        RiskDetailWindow(
            parent_root=self.frame.winfo_toplevel(),
            record=rec,
            all_records=self.all_records,
            save_callback=lambda old, new: self._on_detail_save(iid, old, new),
            delete_callback=lambda s: self._on_detail_delete(iid, s),
            toast_callback=self._show_toast,
        )

    def _find_record(self, idx_str: str) -> tuple | None:
        for r in self.all_records:
            if (str(r[0]) == idx_str or
                    str(r[0]).lstrip("0") == str(idx_str).lstrip("0")):
                return r
        return None

    def _on_detail_save(self, iid: str, old_id: str,
                        new_record: tuple) -> None:
        for i, r in enumerate(self.all_records):
            if (str(r[0]) == str(old_id) or
                    str(r[0]).lstrip("0") == str(old_id).lstrip("0")):
                self.all_records[i] = new_record
                break
        try:
            self.tree.item(iid, values=self._tree_values(new_record))
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
        self.all_records = [r for r in self.all_records
                            if str(r[0]) != str(idx_str)]
        self._recolor_rows()
        self._save_data()
        self._update_count()
        if self.on_data_change:
            self.on_data_change(self.all_records)

    # ------------------------------------------------------------------
    def _sort_tree(self, col: str) -> None:
        data = [(self.tree.set(iid, col), iid)
                for iid in self.tree.get_children("")]
        try:
            data.sort(
                key=lambda x: float(x[0]) if x[0] not in ("—", "") else 0)
        except ValueError:
            data.sort(key=lambda x: x[0].lower())
        for i, (_, iid) in enumerate(data):
            self.tree.move(iid, "", i)
        self._recolor_rows()

    def _recolor_rows(self) -> None:
        for i, iid in enumerate(self.tree.get_children()):
            risk      = self.tree.set(iid, "risk_type")
            score_str = self.tree.set(iid, "score")
            base_tag  = "even" if i % 2 == 0 else "odd"
            tags      = [base_tag]
            if risk in RISK_COLORS:
                tags.append(f"risk_{risk}")
            try:
                s = int(score_str)
                if   s <= 4:  tags.append("score_low")
                elif s <= 9:  tags.append("score_mod")
                elif s <= 16: tags.append("score_high")
                else:         tags.append("score_crit")
            except (ValueError, TypeError):
                pass
            self.tree.item(iid, tags=tags)

    # ------------------------------------------------------------------
    def _load_data(self) -> None:
        self.all_records.clear()
        self.tree.delete(*self.tree.get_children())
        if not os.path.exists(RISK_DATA_FILE):
            return
        try:
            with open(RISK_DATA_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            for row in data:
                if isinstance(row, (list, tuple)):
                    row = self._normalize_record(list(row))
                    self.all_records.append(tuple(row))
                    self._insert_tree_row(tuple(row))
        except Exception as e:
            messagebox.showerror("Помилка завантаження", str(e))
        self._update_count()
        if self.on_data_change:
            self.on_data_change(self.all_records)

    def _save_data(self) -> None:
        try:
            with open(RISK_DATA_FILE, "w", encoding="utf-8") as f:
                json.dump(self.all_records, f, ensure_ascii=False, indent=2)
        except Exception as e:
            messagebox.showerror("Помилка збереження", str(e))

    @staticmethod
    def _normalize_record(row: list | tuple) -> list:
        row = list(row)
        while len(row) < RiskDetailWindow.RECORD_LEN:
            row.append("—")
        return row

    def _tree_values(self, data: tuple) -> tuple:
        d = self._normalize_record(list(data))
        return (d[0], d[1], d[2], d[3], d[4],
                d[7], d[13], d[14], d[8], d[11])

    def _insert_tree_row(self, data: tuple) -> str:
        iid = self.tree.insert("", tk.END, values=self._tree_values(data))
        self._recolor_rows()
        return iid

    # ------------------------------------------------------------------
    def _get_form_data(self) -> tuple | None:
        date_id  = self.ent_date_id.get().strip()
        date_rev = self.ent_date_rev.get().strip()
        for val, label in [(date_id,  "дати виявлення"),
                           (date_rev, "дати перегляду")]:
            if val in ("дд.мм.рррр", ""):
                continue
            if not is_valid_date(val):
                messagebox.showwarning(
                    "Помилка",
                    f"Неправильний формат {label} (очiкується дд.мм.рррр)")
                return None
        date_id  = "" if date_id  == "дд.мм.рррр" else date_id
        date_rev = "" if date_rev == "дд.мм.рррр" else date_rev

        prob_str   = self.cb_prob.get().strip()
        impact_str = self.cb_impact.get().strip()
        score      = _extract_num(prob_str) * _extract_num(impact_str)

        try:
            residual = int(self.ent_residual.get().strip() or "0")
        except ValueError:
            residual = 0

        idx = len(self.all_records) + 1
        return (
            f"{idx:03d}",
            self.ent_entity.get().strip(),
            self.ent_risk_name.get().strip(),
            self.cb_category.get().strip()  or "—",
            self.cb_risk_type.get().strip() or "—",
            prob_str   or "—",
            impact_str or "—",
            str(score),
            self.ent_owner.get().strip(),
            self.txt_controls.get("1.0", tk.END).strip(),
            str(residual),
            date_id  or "—",
            date_rev or "—",
            self.cb_priority.get().strip() or "Середнiй",
            self.cb_status.get().strip()   or "Активний",
            self.txt_description.get("1.0", tk.END).strip(),
        )

    def _clear_form(self) -> None:
        for w in [self.ent_entity, self.ent_owner,
                  self.ent_risk_name, self.ent_residual]:
            w.delete(0, tk.END)
            w.configure(fg=COLORS["text_primary"])
        for w in [self.cb_category, self.cb_risk_type, self.cb_priority,
                  self.cb_status, self.cb_prob, self.cb_impact]:
            w.set("")
        for w in [self.txt_controls, self.txt_description]:
            w.delete("1.0", tk.END)
        self.lbl_score.configure(text="—", fg=COLORS["accent_success"])
        for e, ph in [(self.ent_date_id,  "дд.мм.рррр"),
                      (self.ent_date_rev, "дд.мм.рррр")]:
            e.delete(0, tk.END)
            add_placeholder(e, ph)

    def _add_record(self) -> None:
        data = self._get_form_data()
        if not data:
            return
        if not data[1] or not data[2]:
            messagebox.showwarning(
                "Обов'язковi поля",
                "Заповнiть назву пiдприємства та назву ризику")
            return
        self.all_records.append(data)
        self._insert_tree_row(data)
        self._clear_form()
        self._save_data()
        if self.on_data_change:
            self.on_data_change(self.all_records)
        self._update_count()
        self._show_toast("Ризик додано")

    # ------------------------------------------------------------------
    def _delete_selected(self) -> None:
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Видалення", "Оберiть запис для видалення")
            return
        iid     = sel[0]
        idx_str = self.tree.set(iid, "id")
        if not messagebox.askyesno(
            "Пiдтвердження",
            f"Видалити ризик #{idx_str}?\nЦю дiю не можна скасувати."):
            return
        self.tree.delete(iid)
        self.all_records = [r for r in self.all_records
                            if str(r[0]) != idx_str]
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
        rec = self._find_record(self.tree.set(sel[0], "id"))
        if not rec:
            return
        new_rec = (f"{len(self.all_records)+1:03d}",) + tuple(rec[1:])
        self.all_records.append(new_rec)
        self._insert_tree_row(new_rec)
        self._save_data()
        self._update_count()
        if self.on_data_change:
            self.on_data_change(self.all_records)
        self._show_toast("Запис продубльовано")

    # ------------------------------------------------------------------
    def _apply_filter(self) -> None:
        q      = self.search_var.get().strip().lower()
        r_type = self.filter_type.get()
        status = self.filter_status.get()
        self.tree.delete(*self.tree.get_children())
        for row in self.all_records:
            if q and q not in " ".join(str(v).lower() for v in row):
                continue
            if r_type != "Всi" and (len(row) <= 4 or row[4] != r_type):
                continue
            if status != "Всi" and (len(row) <= 14 or row[14] != status):
                continue
            self._insert_tree_row(row)
        self._update_count()

    def _reset_filter(self) -> None:
        self.search_var.set("")
        self.filter_type.set("Всi")
        self.filter_status.set("Всi")
        self.tree.delete(*self.tree.get_children())
        for row in self.all_records:
            self._insert_tree_row(row)
        self._update_count()

    # ------------------------------------------------------------------
    def _on_select(self, _: object | None = None) -> None:
        sel = self.tree.selection()
        if not sel:
            return
        rec = self._find_record(self.tree.set(sel[0], "id"))
        if not rec:
            return
        controls = rec[9]  if len(rec) > 9  else ""
        desc     = rec[15] if len(rec) > 15 else ""
        for widget, text in [(self.det_controls, controls),
                             (self.det_desc,     desc)]:
            widget.configure(state="normal")
            widget.delete("1.0", tk.END)
            widget.insert("1.0", text or "")
            widget.configure(state="disabled")

    def _update_count(self) -> None:
        self.lbl_count.configure(
            text=f" {len(self.tree.get_children())}")

    def _show_toast(self, msg: str) -> None:
        _show_toast(self.frame, msg)

    # ------------------------------------------------------------------
    _HEADERS = [
        "ID", "Пiдприємство", "Назва ризику", "Категорiя", "Тип ризику",
        "Iмовiрнiсть", "Вплив", "Score", "Власник", "Заходи контролю",
        "Залишковий ризик", "Дата виявлення", "Дата перегляду",
        "Прiоритет", "Статус", "Опис",
    ]

    def _export_csv(self) -> None:
        if not self.tree.get_children():
            messagebox.showinfo("Експорт", "Таблиця порожня")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV файли", "*.csv")],
            title="Зберегти реєстр ризикiв як CSV")
        if not path:
            return
        try:
            with open(path, "w", newline="", encoding="utf-8-sig") as f:
                w = csv.writer(f)
                w.writerow(self._HEADERS)
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
            title="Зберегти реєстр ризикiв як Excel")
        if not path:
            return
        try:
            n    = len(self._HEADERS)
            data = [list(r) + [""] * max(0, n - len(r))
                    for r in self.all_records]
            df = pd.DataFrame(data, columns=self._HEADERS)
            with pd.ExcelWriter(path, engine="openpyxl") as writer:
                df.to_excel(writer, index=False,
                            sheet_name="Реєстр ризикiв")
                ws = writer.sheets["Реєстр ризикiв"]
                for col_cells in ws.columns:
                    mx = max(len(str(c.value or "")) for c in col_cells)
                    ws.column_dimensions[
                        col_cells[0].column_letter].width = min(mx + 4, 60)
            self._show_toast("Excel збережено")
        except Exception as e:
            messagebox.showerror("Помилка", str(e))

    def _import_json(self) -> None:
        path = filedialog.askopenfilename(
            filetypes=[("JSON файли", "*.json")],
            title="Iмпорт JSON")
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
                    row = self._normalize_record(list(row))
                    row[0] = f"{len(self.all_records)+1:03d}"
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
#  ВКЛАДКА: АНАЛІТИКА РИЗИКІВ
# =============================================================================

class RiskAnalyticsTab:
    """Аналiтика реєстру ризикiв."""

    def __init__(self, parent: tk.Misc) -> None:
        self.frame   = ttk.Frame(parent)
        self.records: list[tuple] = []
        self._build_ui()

    def _build_ui(self) -> None:
        C = COLORS
        self.frame.rowconfigure(1, weight=1)
        self.frame.columnconfigure(0, weight=1)

        header = tk.Frame(self.frame, bg=C["bg_header"], height=56)
        header.grid(row=0, column=0, sticky="ew")
        header.grid_propagate(False)
        tk.Label(header, text="АНАЛIТИКА РИЗИКIВ",
                 bg=C["bg_header"], fg=C["accent_muted"],
                 font=("Arial", 13, "bold")).pack(
            side="left", padx=20, pady=14)
        make_button(header, "Оновити",
                    bg=C["accent"], activebackground=C["accent_soft"],
                    font=("Arial", 9, "bold"), padx=12, pady=4,
                    command=self.refresh).pack(
            side="right", padx=20, pady=12)

        canvas = tk.Canvas(self.frame, bg=C["bg_main"], highlightthickness=0)
        sb = ttk.Scrollbar(self.frame, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=sb.set)
        canvas.grid(row=1, column=0, sticky="nsew")
        sb.grid(row=1, column=1, sticky="ns")

        self.content = tk.Frame(canvas, bg=C["bg_main"])
        cw = canvas.create_window((0, 0), window=self.content, anchor="nw")

        def _conf(_: object) -> None:
            canvas.configure(scrollregion=canvas.bbox("all"))
            canvas.itemconfig(cw, width=canvas.winfo_width())

        self.content.bind("<Configure>", _conf)
        canvas.bind("<Configure>", _conf)
        canvas.bind_all("<MouseWheel>",
                        lambda e: canvas.yview_scroll(
                            int(-1 * e.delta / 120), "units"))
        self.content.columnconfigure(0, weight=1)

        self._build_stat_cards()
        self._build_charts_and_table()

    def _build_stat_cards(self) -> None:
        C = COLORS
        cf = tk.Frame(self.content, bg=C["bg_main"])
        cf.grid(row=0, column=0, sticky="ew", padx=16, pady=(16, 0))
        for i in range(5):
            cf.columnconfigure(i, weight=1)

        self.stat_cards: dict[str, tk.Label] = {}
        defs = [
            ("total",     "Всього ризикiв",       "0", C["accent"]),
            ("critical",  "Критичних (Score >16)", "0", C["accent_danger"]),
            ("high",      "Високих (10–16)",        "0", "#f97316"),
            ("active",    "Активних",               "0", C["accent_warning"]),
            ("mitigated", "Мiтигованих",            "0", C["accent_success"]),
        ]
        for ci, (key, title, val, color) in enumerate(defs):
            card = tk.Frame(cf, bg=C["bg_surface"], padx=18, pady=12)
            card.grid(row=0, column=ci, padx=6, sticky="nsew")
            tk.Frame(card, bg=color, height=3).pack(fill="x")
            tk.Label(card, text=title, bg=C["bg_surface"],
                     fg=C["text_muted"], font=("Arial", 8)).pack(
                anchor="w", pady=(8, 2))
            lbl = tk.Label(card, text=val, bg=C["bg_surface"],
                           fg=color, font=("Arial", 22, "bold"))
            lbl.pack(anchor="w")
            self.stat_cards[key] = lbl

    def _build_charts_and_table(self) -> None:
        C = COLORS
        if HAS_MPL:
            cr = tk.Frame(self.content, bg=C["bg_main"])
            cr.grid(row=1, column=0, sticky="ew", padx=16, pady=16)
            cr.columnconfigure((0, 1), weight=1)

            self.fig_left  = Figure(figsize=(5, 3.5), dpi=90,
                                     facecolor=C["bg_surface"])
            self.ax_left   = self.fig_left.add_subplot(111)
            self._style_ax(self.ax_left)
            self.ax_left.set_title("Розподiл за типом ризику",
                                    color=C["text_muted"], fontsize=9)
            fl = tk.Frame(cr, bg=C["bg_surface"], padx=8, pady=8)
            fl.grid(row=0, column=0, padx=(0, 8), sticky="nsew")
            self.canvas_left = FigureCanvasTkAgg(self.fig_left, master=fl)
            self.canvas_left.get_tk_widget().pack(fill="both", expand=True)

            self.fig_right = Figure(figsize=(5, 3.5), dpi=90,
                                     facecolor=C["bg_surface"])
            self.ax_right  = self.fig_right.add_subplot(111)
            self._style_ax(self.ax_right)
            self.ax_right.set_title("Розподiл за рiвнем ризику",
                                     color=C["text_muted"], fontsize=9)
            fr = tk.Frame(cr, bg=C["bg_surface"], padx=8, pady=8)
            fr.grid(row=0, column=1, padx=(8, 0), sticky="nsew")
            self.canvas_right = FigureCanvasTkAgg(self.fig_right, master=fr)
            self.canvas_right.get_tk_widget().pack(fill="both", expand=True)

            self.fig_heat = Figure(figsize=(10, 4), dpi=90,
                                    facecolor=C["bg_surface"])
            self.ax_heat  = self.fig_heat.add_subplot(111)
            self._style_ax(self.ax_heat)
            self.ax_heat.set_title(
                "Матриця ризикiв (Iмовiрнiсть × Вплив)",
                color=C["text_muted"], fontsize=9)
            fh = tk.Frame(self.content, bg=C["bg_surface"], padx=8, pady=8)
            fh.grid(row=2, column=0, padx=16, pady=(0, 16), sticky="ew")
            self.canvas_heat = FigureCanvasTkAgg(self.fig_heat, master=fh)
            self.canvas_heat.get_tk_widget().pack(fill="both", expand=True)
        else:
            tk.Label(self.content,
                     text="Встановiть matplotlib:\n  pip install matplotlib",
                     bg=C["bg_main"], fg=C["text_muted"],
                     font=("Arial", 10)).grid(row=1, column=0, pady=40)

        frame = tk.Frame(self.content, bg=C["bg_surface"], padx=16, pady=12)
        frame.grid(row=3, column=0, padx=16, pady=(0, 16), sticky="ew")
        frame.columnconfigure(0, weight=1)
        tk.Label(frame, text="Деталiзована статистика за типом ризику",
                 bg=C["bg_surface"], fg=C["text_muted"],
                 font=("Arial", 9, "bold")).grid(
            row=0, column=0, sticky="w", pady=(0, 8))

        cols = ("risk_type", "count", "avg_score", "max_score", "active")
        self.stats_tree = ttk.Treeview(frame, columns=cols,
                                        show="headings", height=7)
        for col, hdr, w in [
            ("risk_type",  "Тип ризику",  180),
            ("count",      "Всього",        70),
            ("avg_score",  "Сер. Score",    90),
            ("max_score",  "Макс. Score",   90),
            ("active",     "Активних",      80),
        ]:
            self.stats_tree.heading(col, text=hdr)
            self.stats_tree.column(col, width=w, anchor="w")
        self.stats_tree.grid(row=1, column=0, sticky="ew")

    def _style_ax(self, ax: object) -> None:
        C = COLORS
        ax.set_facecolor(C["bg_surface"])
        ax.tick_params(colors=C["text_muted"], labelsize=8)
        for spine in ax.spines.values():
            spine.set_edgecolor(C["border_soft"])

    def update_data(self, records: list[tuple]) -> None:
        self.records = records
        self.refresh()

    def refresh(self) -> None:
        keys = ["total", "critical", "high", "active", "mitigated"]
        if not self.records:
            for k in keys:
                self.stat_cards[k].configure(text="0")
            if HAS_MPL:
                self._clear_charts()
            self.stats_tree.delete(*self.stats_tree.get_children())
            return

        C       = COLORS
        records = self.records

        total     = len(records)
        critical_ = sum(1 for r in records
                        if len(r) > 7 and str(r[7]).isdigit()
                        and int(r[7]) > 16)
        high_     = sum(1 for r in records
                        if len(r) > 7 and str(r[7]).isdigit()
                        and 10 <= int(r[7]) <= 16)
        active_   = sum(1 for r in records
                        if len(r) > 14 and r[14] == "Активний")
        mitig_    = sum(1 for r in records
                        if len(r) > 14 and r[14] == "Мiтигований")

        self.stat_cards["total"].configure(text=str(total))
        self.stat_cards["critical"].configure(text=str(critical_))
        self.stat_cards["high"].configure(text=str(high_))
        self.stat_cards["active"].configure(text=str(active_))
        self.stat_cards["mitigated"].configure(text=str(mitig_))

        if not HAS_MPL:
            return

        type_ctr = Counter(r[4] for r in records if len(r) > 4)

        # Pie — тип ризику
        self.ax_left.clear()
        self._style_ax(self.ax_left)
        self.ax_left.set_title("Розподiл за типом ризику",
                                color=C["text_muted"], fontsize=9)
        if type_ctr:
            lbls = list(type_ctr.keys())
            vals = list(type_ctr.values())
            clrs = [RISK_COLORS.get(l, C["text_muted"]) for l in lbls]
            _, _, autotexts = self.ax_left.pie(
                vals, labels=lbls, autopct="%1.0f%%",
                colors=clrs, startangle=90,
                textprops={"color": C["text_muted"], "fontsize": 7})
            for at in autotexts:
                at.set_fontsize(7); at.set_color("white")
        else:
            self.ax_left.text(0.5, 0.5, "Немає даних",
                              transform=self.ax_left.transAxes,
                              ha="center", va="center",
                              color=C["text_muted"])
        self.canvas_left.draw()

        # Bar — рівень ризику
        self.ax_right.clear()
        self._style_ax(self.ax_right)
        self.ax_right.set_title("Розподiл за рiвнем ризику",
                                 color=C["text_muted"], fontsize=9)
        level_ctr = {"Низький": 0, "Помiрний": 0,
                     "Високий": 0, "Критичний": 0}
        for r in records:
            if len(r) > 7 and str(r[7]).isdigit():
                level_ctr[_score_label(int(r[7]))] += 1
        if any(level_ctr.values()):
            lbls = list(level_ctr.keys())
            vals = list(level_ctr.values())
            clrs = [C["accent_success"], C["accent_warning"],
                    "#f97316", C["accent_danger"]]
            bars = self.ax_right.bar(lbls, vals, color=clrs, edgecolor="none")
            for bar, val in zip(bars, vals, strict=False):
                if val > 0:
                    self.ax_right.text(
                        bar.get_x() + bar.get_width() / 2,
                        bar.get_height() + 0.1,
                        str(val), ha="center", va="bottom",
                        color=C["text_muted"], fontsize=8)
            self.ax_right.tick_params(axis="x", labelrotation=15, labelsize=7)
            mx = max(vals)
            self.ax_right.set_ylim(0, mx * 1.2 + 1 if mx > 0 else 1)
        else:
            self.ax_right.text(0.5, 0.5, "Немає даних",
                               transform=self.ax_right.transAxes,
                               ha="center", va="center",
                               color=C["text_muted"])
        self.canvas_right.draw()

        # Heatmap
        self.ax_heat.clear()
        self._style_ax(self.ax_heat)
        self.ax_heat.set_title(
            "Матриця ризикiв (Iмовiрнiсть × Вплив)",
            color=C["text_muted"], fontsize=9)
        matrix = [[0] * 5 for _ in range(5)]
        for r in records:
            if len(r) > 6:
                try:
                    prob = _extract_num(r[5])
                    imp  = _extract_num(r[6])
                    if 1 <= prob <= 5 and 1 <= imp <= 5:
                        matrix[5 - prob][imp - 1] += 1
                except (ValueError, IndexError):
                    pass
        if any(any(row) for row in matrix) and np is not None:
            self.ax_heat.imshow(matrix, cmap=plt.cm.RdYlGn_r, aspect="auto")
            self.ax_heat.set_xticks(range(5))
            self.ax_heat.set_yticks(range(5))
            self.ax_heat.set_xticklabels([str(i) for i in range(1, 6)])
            self.ax_heat.set_yticklabels([str(i) for i in range(5, 0, -1)])
            self.ax_heat.set_xlabel("Вплив →",
                                     color=C["text_muted"], fontsize=8)
            self.ax_heat.set_ylabel("Iмовiрнiсть →",
                                     color=C["text_muted"], fontsize=8)
            for i in range(5):
                for j in range(5):
                    if matrix[i][j] > 0:
                        self.ax_heat.text(
                            j, i, str(matrix[i][j]),
                            ha="center", va="center",
                            color="white", fontsize=10, weight="bold")
        else:
            self.ax_heat.text(0.5, 0.5, "Немає даних",
                              transform=self.ax_heat.transAxes,
                              ha="center", va="center",
                              color=C["text_muted"])
        self.canvas_heat.draw()

        # Таблиця статистики
        self.stats_tree.delete(*self.stats_tree.get_children())
        all_types = set(RISK_TYPES) | set(r[4] for r in records if len(r) > 4)
        for rt in sorted(all_types):
            recs = [r for r in records if len(r) > 4 and r[4] == rt]
            cnt  = len(recs)
            if cnt:
                scores    = [int(r[7]) for r in recs
                             if len(r) > 7 and str(r[7]).isdigit()]
                avg_score = sum(scores) / len(scores) if scores else 0
                max_score = max(scores) if scores else 0
                act       = sum(1 for r in recs
                                if len(r) > 14 and r[14] == "Активний")
                self.stats_tree.insert(
                    "", tk.END,
                    values=(rt, cnt, f"{avg_score:.1f}", max_score, act))

    def _clear_charts(self) -> None:
        if not HAS_MPL:
            return
        for ax in (self.ax_left, self.ax_right, self.ax_heat):
            ax.clear()
        for cv in (self.canvas_left, self.canvas_right, self.canvas_heat):
            cv.draw()

    def get_frame(self) -> ttk.Frame:
        return self.frame


# =============================================================================
#  ВКЛАДКА: НАЛАШТУВАННЯ РИЗИКІВ
# =============================================================================

class RiskSettingsTab:
    """Вкладка налаштувань модуля 'Реєстр ризикiв'."""

    def __init__(self, parent: tk.Misc) -> None:
        self.frame = ttk.Frame(parent)
        self._build_ui()

    def _build_ui(self) -> None:
        C = COLORS
        self.frame.columnconfigure(0, weight=1)
        self.frame.rowconfigure(1, weight=1)

        header = tk.Frame(self.frame, bg=C["bg_header"], height=56)
        header.grid(row=0, column=0, sticky="ew")
        header.grid_propagate(False)
        tk.Label(header, text="НАЛАШТУВАННЯ РЕЄСТРУ РИЗИКIВ",
                 bg=C["bg_header"], fg=C["text_muted"],
                 font=("Arial", 13, "bold")).pack(
            side="left", padx=20, pady=14)

        content = tk.Frame(self.frame, bg=C["bg_main"])
        content.grid(row=1, column=0, sticky="nsew", padx=40, pady=30)
        content.columnconfigure(0, weight=1)

        self._row(content, 0, "Файл даних:", RISK_DATA_FILE, C)
        self._row(content, 1, "Версiя:", "1.0 — Реєстр ризикiв", C)
        self._row(content, 2, "matplotlib:",
                                    "встановлено" if HAS_MPL else "не встановлено", C)
        self._row(content, 3, "pandas:",
                  "встановлено" if pd else "не встановлено", C)

        tk.Label(content, text="Встановлення залежностей:",
                 bg=C["bg_main"], fg=C["text_muted"],
                 font=("Arial", 8, "bold")).grid(
            row=4, column=0, sticky="w", pady=(24, 6))
        tk.Label(content, text="  pip install matplotlib pandas openpyxl",
                 bg=C["bg_surface"], fg=C["accent_muted"],
                 font=("Courier", 9), padx=12, pady=8).grid(
            row=5, column=0, sticky="w")

        tk.Label(content, text="Структура запису (16 полiв):",
                 bg=C["bg_main"], fg=C["text_muted"],
                 font=("Arial", 8, "bold")).grid(
            row=6, column=0, sticky="w", pady=(24, 6))

        fields_text = (
            "ID, Пiдприємство, Назва ризику, Категорiя, Тип ризику,\n"
            "Iмовiрнiсть, Вплив, Score, Власник, Заходи контролю,\n"
            "Залишковий ризик, Дата виявлення, Дата перегляду, "
            "Прiоритет, Статус, Опис"
        )
        tk.Label(content, text=fields_text,
                 bg=C["bg_surface"], fg=C["text_subtle"],
                 font=("Arial", 8), justify="left",
                 padx=12, pady=8).grid(row=7, column=0, sticky="w")

        tk.Label(content, text="Пiдказки:",
                 bg=C["bg_main"], fg=C["text_muted"],
                 font=("Arial", 8, "bold")).grid(
            row=8, column=0, sticky="w", pady=(24, 6))

        hints = [
            "Score розраховується автоматично: Iмовiрнiсть × Вплив (1–25)",
            "Рiвнi ризику: Низький (1–4), Помiрний (5–9), Високий (10–16), Критичний (17–25)",
            "Подвiйний клiк по рядку — вiдкрити детальне вiкно ризику",
            "Матриця ризикiв показує розподiл за ймовiрнiстю та впливом",
        ]
        for i, hint in enumerate(hints):
            f = tk.Frame(content, bg=C["bg_main"])
            f.grid(row=9 + i, column=0, sticky="w", pady=2)
            tk.Frame(f, bg=C["accent_success"], width=4, height=4).pack(
                side="left", padx=(0, 8))
            tk.Label(f, text=hint, bg=C["bg_main"], fg=C["text_subtle"],
                     font=("Arial", 8)).pack(side="left")

    def _row(self, parent: tk.Misc, row: int,
             label: str, value: str, C: dict) -> None:
        f = tk.Frame(parent, bg=C["bg_main"])
        f.grid(row=row, column=0, sticky="ew", pady=4)
        tk.Label(f, text=label, bg=C["bg_main"], fg=C["text_muted"],
                 font=("Arial", 9), width=22, anchor="w").pack(side="left")
        tk.Label(f, text=value, bg=C["bg_main"], fg=C["text_primary"],
                 font=("Arial", 9)).pack(side="left")

    def get_frame(self) -> ttk.Frame:
        return self.frame


# =============================================================================
#  СТОРІНКА "РЕЄСТР РИЗИКІВ"
# =============================================================================

class RiskRegisterPage(tk.Frame):
    """Сторiнка 'Реєстр ризикiв' (вкладки + статусбар)."""

    def __init__(self, master: tk.Misc, **kwargs) -> None:
        super().__init__(master, bg=COLORS["bg_main"], **kwargs)

        self.grid_rowconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=0)
        self.grid_columnconfigure(0, weight=1)

        self.notebook = ttk.Notebook(self)
        self.notebook.grid(row=0, column=0, sticky="nsew")

        self.analytics_tab = RiskAnalyticsTab(self.notebook)
        self.registry_tab  = RiskRegistryTab(
            self.notebook,
            on_data_change=self.analytics_tab.update_data)
        self.settings_tab  = RiskSettingsTab(self.notebook)

        self.notebook.add(self.registry_tab.get_frame(),
                          text="  Реєстр ризикiв  ")
        self.notebook.add(self.analytics_tab.get_frame(),
                          text="  Аналiтика  ")
        self.notebook.add(self.settings_tab.get_frame(),
                          text="  Налаштування  ")

        statusbar = tk.Frame(self, bg=COLORS["bg_header"], height=22)
        statusbar.grid(row=1, column=0, sticky="ew")
        statusbar.grid_propagate(False)

        self._status_lbl = tk.Label(
            statusbar, text="Готово",
            bg=COLORS["bg_header"], fg=COLORS["text_muted"],
            font=("Arial", 7), padx=10)
        self._status_lbl.pack(side="left", pady=3)

        self._time_lbl = tk.Label(
            statusbar, text="",
            bg=COLORS["bg_header"], fg=COLORS["text_muted"],
            font=("Arial", 7), padx=10)
        self._time_lbl.pack(side="right", pady=3)

        self._start_clock()
        self._schedule_autosave()
        self.after(600, lambda: self.analytics_tab.update_data(
            self.registry_tab.all_records))

    def _start_clock(self) -> None:
        self._time_lbl.configure(
            text=datetime.now().strftime("%d.%m.%Y  %H:%M:%S"))
        self.after(1000, self._start_clock)

    def _schedule_autosave(self) -> None:
        try:
            self.registry_tab._save_data()  # noqa: SLF001
            self._status_lbl.configure(
                text=f"Автозбережено о {datetime.now().strftime('%H:%M:%S')}")
        except Exception:
            self._status_lbl.configure(text="Помилка автозбереження")
        self.after(30000, self._schedule_autosave)

    def save_before_exit(self) -> None:
        try:
            self.registry_tab._save_data()  # noqa: SLF001
        except Exception:
            pass


# =============================================================================
#  ATLAS APP
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

APP_TITLE      = "ATLAS"
COPYRIGHT_TEXT = "© 2026 Chugaister8"


class AtlasApp(tk.Tk):

    def __init__(self, user_full_name: str) -> None:
        super().__init__()

        self.title("ATLAS | Risk Management System")
        self.geometry("1350x820")
        self.minsize(1150, 700)
        self.configure(bg=COLORS["bg_main"])

        apply_dark_style(self)
        self.option_add("*Font", "Arial 9")

        self._user_full_name = user_full_name
        self._pages: dict[PageKey, tk.Frame] = {}
        self._current_page: PageKey | None = None

        self._build_layout()
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    # ------------------------------------------------------------------
    def _build_layout(self) -> None:
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(1, weight=1)

        self._build_topbar()
        self._build_sidebar()
        self._build_content()

    # ------------------------------------------------------------------
    def _build_topbar(self) -> None:
        C = COLORS
        topbar = tk.Frame(self, bg=C["bg_header"], height=56)
        topbar.grid(row=0, column=0, columnspan=2, sticky="ew")
        topbar.grid_propagate(False)
        topbar.columnconfigure(1, weight=1)

        tk.Label(topbar, text=APP_TITLE,
                 bg=C["bg_header"], fg=C["text_primary"],
                 font=("Arial", 14, "bold")).grid(
            row=0, column=0, padx=20, pady=8, sticky="w")

        tk.Label(topbar, text=self._user_full_name,
                 bg=C["bg_header"], fg=C["text_muted"],
                 font=("Arial", 10)).grid(
            row=0, column=1, padx=8, pady=8, sticky="e")

        make_button(
            topbar, "🔔",
            bg=C["bg_surface"], fg=C["text_primary"],
            activebackground=C["bg_surface_alt"],
            font=("Arial", 14), padx=10, pady=4,
            command=lambda: messagebox.showinfo(
                "Нагадування", "Нагадувань поки немає."),
        ).grid(row=0, column=2, padx=20, pady=8, sticky="e")

    # ------------------------------------------------------------------
    def _build_sidebar(self) -> None:
        C = COLORS
        sidebar = tk.Frame(self, bg=C["bg_sidebar"], width=220)
        sidebar.grid(row=1, column=0, sticky="nsw")
        sidebar.grid_propagate(False)
        sidebar.columnconfigure(0, weight=1)

        tk.Label(sidebar, text="Навiгацiя",
                 bg=C["bg_sidebar"], fg=C["text_muted"],
                 font=("Arial", 10, "bold")).grid(
            row=0, column=0, padx=16, pady=(12, 4), sticky="w")

        nav_frame = tk.Frame(sidebar, bg=C["bg_sidebar"])
        nav_frame.grid(row=1, column=0, sticky="nsew", padx=8)
        nav_frame.columnconfigure(0, weight=1)

        menu_items: list[tuple[PageKey, str]] = [
            ("risk_register",     "Реєстр ризикiв"),
            ("material_events",   "Реєстр суттєвих подiй"),
            ("risk_appetite",     "Ризик апетит"),
            ("analytics",         "Аналiтика"),
            ("reports",           "Звiти"),
            ("risk_coordinators", "Ризик координатори"),
            ("settings",          "Налаштування"),
        ]

        self._nav_buttons: dict[PageKey, tk.Button] = {}
        for i, (key, label) in enumerate(menu_items):
            btn = make_button(
                nav_frame, label,
                bg=C["bg_sidebar"], fg=C["text_muted"],
                activebackground=C["bg_surface"],
                activeforeground=C["text_primary"],
                font=("Arial", 9),
                padx=24, pady=7,
                anchor="w",
                command=lambda k=key: self._on_nav_click(k),
            )
            btn.grid(row=i, column=0, sticky="ew", pady=1)
            self._nav_buttons[key] = btn

        sidebar.grid_rowconfigure(2, weight=1)

        tk.Label(sidebar, text=COPYRIGHT_TEXT,
                 bg=C["bg_sidebar"], fg=C["text_subtle"],
                 font=("Arial", 8)).grid(
            row=3, column=0, pady=6, sticky="s")

    # ------------------------------------------------------------------
    def _build_content(self) -> None:
        self.content_frame = tk.Frame(self, bg=COLORS["bg_main"])
        self.content_frame.grid(row=1, column=1, sticky="nsew")
        self.content_frame.grid_rowconfigure(0, weight=1)
        self.content_frame.grid_columnconfigure(0, weight=1)

        self._on_nav_click("risk_register")   # стартова сторiнка

    # ------------------------------------------------------------------
    def _on_nav_click(self, page_key: PageKey) -> None:
        if self._current_page == page_key:
            return
        self._current_page = page_key

        C = COLORS
        for k, btn in self._nav_buttons.items():
            if k == page_key:
                btn.configure(bg=C["bg_surface"], fg=C["text_primary"])
            else:
                btn.configure(bg=C["bg_sidebar"], fg=C["text_muted"])

        for child in self.content_frame.winfo_children():
            child.grid_forget()

        if page_key not in self._pages:
            if page_key == "material_events":
                page: tk.Frame = MaterialEventsPage(self.content_frame)
            elif page_key == "risk_register":
                page = RiskRegisterPage(self.content_frame)
            else:
                page = self._create_placeholder_page(page_key)
            self._pages[page_key] = page

        self._pages[page_key].grid(row=0, column=0, sticky="nsew")

    # ------------------------------------------------------------------
    def _create_placeholder_page(self, page_key: PageKey) -> tk.Frame:
        C    = COLORS
        page = tk.Frame(self.content_frame, bg=C["bg_main"])
        page.columnconfigure(0, weight=1)
        page.rowconfigure(0, weight=1)

        inner = tk.Frame(page, bg=C["bg_main"])
        inner.place(relx=0.5, rely=0.5, anchor="center")

        # Декоративна лiнiя зверху
        tk.Frame(inner, bg=C["accent"], height=3, width=60).pack(pady=(0, 16))

        tk.Label(
            inner,
            text=page_key.replace("_", " ").title(),
            bg=C["bg_main"], fg=C["text_primary"],
            font=("Arial", 20, "bold"),
        ).pack()

        tk.Label(
            inner,
            text="Модуль в розробцi...",
            bg=C["bg_main"], fg=C["text_muted"],
            font=("Arial", 11),
        ).pack(pady=(8, 0))

        tk.Frame(inner, bg=C["border_soft"], height=1, width=60).pack(
            pady=(16, 0))

        return page

    # ------------------------------------------------------------------
    def _on_close(self) -> None:
        for page in self._pages.values():
            if hasattr(page, "save_before_exit"):
                page.save_before_exit()
        self.destroy()


# =============================================================================
#  MAIN
# =============================================================================

def main() -> int:
    current_user = "Онiщенко Андрiй Сергiйович"
    app = AtlasApp(user_full_name=current_user)
    app.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
