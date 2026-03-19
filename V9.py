from __future__ import annotations

from typing import Callable, Literal, TypeAlias
from dataclasses import dataclass, field, asdict

import csv
import json
import os
import re
import uuid
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

DATA_FILE         = "essential_events.json"
RISK_DATA_FILE    = "risk_register.json"
COORDS_DATA_FILE  = "risk_coordinators.json"
APPETITE_FILE     = "risk_appetite.json"

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

# Шрифти — одне місце для зміни
FONT_DEFAULT   = ("Arial", 9)
FONT_SMALL     = ("Arial", 8)
FONT_TINY      = ("Arial", 7)
FONT_BOLD      = ("Arial", 9, "bold")
FONT_SMALL_BOLD= ("Arial", 8, "bold")
FONT_TITLE     = ("Arial", 13, "bold")
FONT_HEADING   = ("Arial", 11, "bold")
FONT_NUMBER    = ("Arial", 22, "bold")
FONT_SCORE     = ("Arial", 14, "bold")
FONT_MONO      = ("Courier", 9)

RISK_COLORS = {
    "Операцiйний":            COLORS["accent_warning"],
    "Технiчний":              COLORS["accent"],
    "Фiнансовий":             COLORS["accent_danger"],
    "Репутацiйний":           "#a855f7",
    "Екологiчний":            COLORS["accent_success"],
    "Надзвичайна ситуацiя":   "#f97316",
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
    if score <= 4:    return COLORS["accent_success"]
    elif score <= 9:  return COLORS["accent_warning"]
    elif score <= 16: return "#f97316"
    else:             return COLORS["accent_danger"]


def _score_label(score: int) -> str:
    if score <= 4:    return "Низький"
    elif score <= 9:  return "Помiрний"
    elif score <= 16: return "Високий"
    else:             return "Критичний"


# =============================================================================
#  DATACLASSES — виправлення #9: магічні індекси → структуровані дані
# =============================================================================

@dataclass
class EventRecord:
    """Структурований запис події. Виправлення: замість tuple-індексів."""
    id:          str = ""
    entity:      str = ""
    event_name:  str = ""
    involved:    str = ""
    risk_type:   str = "—"
    event_date:  str = "—"
    description: str = ""
    measures:    str = ""
    detect_date: str = "—"
    priority:    str = "Середнiй"
    status:      str = "Вiдкрито"

    # ── Серіалізація / десеріалізація ──────────────────────────────────
    def to_list(self) -> list:
        return [
            self.id, self.entity, self.event_name, self.involved,
            self.risk_type, self.event_date, self.description,
            self.measures, self.detect_date, self.priority, self.status,
        ]

    @classmethod
    def from_list(cls, row: list) -> "EventRecord":
        """Сумісність зі старим форматом (список)."""
        r = list(row) + ["—"] * max(0, 11 - len(row))
        return cls(
            id=str(r[0]), entity=str(r[1]), event_name=str(r[2]),
            involved=str(r[3]), risk_type=str(r[4]),
            event_date=str(r[5]), description=str(r[6]),
            measures=str(r[7]), detect_date=str(r[8]),
            priority=str(r[9]), status=str(r[10]),
        )

    @classmethod
    def from_dict(cls, d: dict) -> "EventRecord":
        return cls(**{k: str(v) for k, v in d.items() if k in cls.__dataclass_fields__})


@dataclass
class RiskRecord:
    """Структурований запис ризику. Виправлення: замість tuple-індексів."""
    id:           str = ""
    entity:       str = ""
    risk_name:    str = ""
    category:     str = "—"
    risk_type:    str = "—"
    probability:  str = "—"
    impact:       str = "—"
    score:        str = "0"
    owner:        str = ""
    controls:     str = ""
    residual:     str = "0"
    date_id:      str = "—"
    date_rev:     str = "—"
    priority:     str = "Середнiй"
    status:       str = "Активний"
    description:  str = ""

    def to_list(self) -> list:
        return [
            self.id, self.entity, self.risk_name, self.category,
            self.risk_type, self.probability, self.impact, self.score,
            self.owner, self.controls, self.residual,
            self.date_id, self.date_rev, self.priority,
            self.status, self.description,
        ]

    @classmethod
    def from_list(cls, row: list) -> "RiskRecord":
        r = list(row)
        while len(r) < 16:
            r.append("—")
        return cls(
            id=str(r[0]), entity=str(r[1]), risk_name=str(r[2]),
            category=str(r[3]), risk_type=str(r[4]),
            probability=str(r[5]), impact=str(r[6]), score=str(r[7]),
            owner=str(r[8]), controls=str(r[9]), residual=str(r[10]),
            date_id=str(r[11]), date_rev=str(r[12]),
            priority=str(r[13]), status=str(r[14]),
            description=str(r[15]),
        )

    @classmethod
    def from_dict(cls, d: dict) -> "RiskRecord":
        return cls(**{k: str(v) for k, v in d.items() if k in cls.__dataclass_fields__})


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
    style = ttk.Style(root)
    style.theme_use("clam")
    C = COLORS

    style.configure(
        ".",
        background=C["bg_main"], foreground=C["text_primary"],
        fieldbackground=C["bg_input"], troughcolor=C["bg_surface"],
        bordercolor=C["border_soft"], darkcolor=C["bg_surface"],
        lightcolor=C["bg_surface"], insertcolor=C["text_primary"],
        selectbackground=C["row_select"], selectforeground=C["text_primary"],
        font=FONT_DEFAULT,
    )

    style.configure("TFrame",         background=C["bg_main"])
    style.configure("Surface.TFrame", background=C["bg_surface"])
    style.configure("Sidebar.TFrame", background=C["bg_sidebar"])
    style.configure("Header.TFrame",  background=C["bg_header"])

    style.configure("TLabel",
                    background=C["bg_main"], foreground=C["text_primary"],
                    font=FONT_DEFAULT)
    style.configure("Muted.TLabel",
                    background=C["bg_main"], foreground=C["text_muted"],
                    font=FONT_SMALL)

    style.configure("TEntry",
                    fieldbackground=C["bg_input"], foreground=C["text_primary"],
                    bordercolor=C["border_soft"], insertcolor=C["text_primary"])
    style.map("TEntry",
              fieldbackground=[("focus", C["bg_surface_alt"])],
              bordercolor=[("focus", C["accent"])])

    style.configure("TCombobox",
                    fieldbackground=C["bg_surface"], background=C["bg_surface"],
                    foreground=C["text_primary"], bordercolor=C["border_soft"],
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
                    background=C["bg_main"], bordercolor=C["border_soft"],
                    tabmargins=[0, 0, 0, 0])
    style.configure("TNotebook.Tab",
                    background=C["bg_sidebar"], foreground=C["text_muted"],
                    padding=(14, 6), font=FONT_DEFAULT)
    style.map("TNotebook.Tab",
              background=[("selected", C["bg_surface"]),
                          ("active",   C["bg_surface_alt"])],
              foreground=[("selected", C["text_primary"]),
                          ("active",   C["text_primary"])])

    style.configure("Treeview",
                    background=C["row_odd"], foreground=C["text_primary"],
                    fieldbackground=C["row_odd"], bordercolor=C["border_soft"],
                    font=FONT_DEFAULT, rowheight=24)
    style.configure("Treeview.Heading",
                    background=C["bg_surface"], foreground=C["text_muted"],
                    bordercolor=C["border_soft"], font=FONT_SMALL_BOLD,
                    relief="flat")
    style.map("Treeview",
              background=[("selected", C["row_select"])],
              foreground=[("selected", C["text_primary"])])

    for orient in ("Vertical", "Horizontal"):
        style.configure(f"{orient}.TScrollbar",
                        background=C["bg_surface"], troughcolor=C["bg_main"],
                        arrowcolor=C["text_muted"], bordercolor=C["bg_main"])


def make_dark_text(parent: tk.Misc, **kwargs) -> tk.Text:
    C = COLORS
    return tk.Text(
        parent,
        bg=C["bg_input"], fg=C["text_primary"],
        insertbackground=C["text_primary"],
        selectbackground=C["row_select"], selectforeground=C["text_primary"],
        relief="flat", bd=1, highlightthickness=1,
        highlightbackground=C["border_soft"], highlightcolor=C["accent"],
        font=FONT_DEFAULT, **kwargs,
    )


def make_dark_entry(parent: tk.Misc, accent: str | None = None,
                    **kwargs) -> tk.Entry:
    C = COLORS
    return tk.Entry(
        parent,
        bg=C["bg_input"], fg=C["text_primary"],
        insertbackground=C["text_primary"],
        relief="flat", bd=2, highlightthickness=1,
        highlightbackground=C["border_soft"],
        highlightcolor=accent or C["accent"],
        font=FONT_DEFAULT, **kwargs,
    )


def make_dark_combo(parent: tk.Misc, values: list[str] | None = None,
                    **kwargs) -> ttk.Combobox:
    return ttk.Combobox(
        parent, values=values or [],
        state="readonly", font=FONT_DEFAULT, **kwargs,
    )


def make_button(parent: tk.Misc, text: str, bg: str,
                fg: str = "white", **kwargs) -> tk.Button:
    C = COLORS
    active_bg = kwargs.pop("activebackground", bg)
    active_fg = kwargs.pop("activeforeground", fg)
    font      = kwargs.pop("font", FONT_DEFAULT)
    return tk.Button(
        parent, text=text, bg=bg, fg=fg,
        activebackground=active_bg, activeforeground=active_fg,
        relief="flat", bd=0, cursor="hand2", font=font, **kwargs,
    )


def add_placeholder(entry: tk.Entry, text: str) -> None:
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
    try:
        return int(str(val).split()[0])
    except (ValueError, IndexError, TypeError):
        try:
            return int(str(val))
        except ValueError:
            return 1


def _build_section_label(parent: tk.Misc, text: str, row: int,
                          accent: str | None = None) -> None:
    C = COLORS
    color = accent or C["accent"]
    f = tk.Frame(parent, bg=C["bg_main"])
    f.grid(row=row, column=0, columnspan=2, sticky="ew", padx=8, pady=(14, 4))
    tk.Frame(f, bg=color, width=3, height=16).pack(side="left")
    tk.Label(f, text=text, bg=C["bg_main"], fg=color,
             font=FONT_BOLD).pack(side="left", padx=8)


def _build_info_cell(parent: tk.Misc, label: str, value: str,
                     row: int, col: int = 0,
                     value_color: str | None = None) -> None:
    C = COLORS
    cell = tk.Frame(parent, bg=C["bg_surface"], padx=10, pady=6)
    cell.grid(
        row=row, column=col, sticky="nsew",
        padx=(8 if col == 0 else 4, 4 if col == 0 else 8), pady=3,
    )
    cell.columnconfigure(0, weight=1)
    tk.Label(cell, text=label, bg=C["bg_surface"], fg=C["text_subtle"],
             font=("Arial", 7, "bold")).grid(row=0, column=0, sticky="w")
    fg  = value_color if value_color else C["text_primary"]
    fnt = ("Arial", 10, "bold") if value_color else FONT_DEFAULT
    tk.Label(cell, text=value or "—", bg=C["bg_surface"], fg=fg,
             font=fnt, wraplength=260, justify="left").grid(
        row=1, column=0, sticky="w", pady=(2, 0))


def _build_text_block(parent: tk.Misc, label: str, value: str,
                      row: int) -> None:
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


def _show_toast(frame: tk.Widget, msg: str, color: str | None = None) -> None:
    bg    = color or COLORS["accent_success"]
    toast = tk.Toplevel(frame)
    toast.overrideredirect(True)
    toast.attributes("-topmost", True)
    toast.configure(bg=bg)
    tk.Label(toast, text=f"  {msg}  ", bg=bg, fg="white",
             font=FONT_BOLD, pady=6).pack()
    root = frame.winfo_toplevel()
    x = root.winfo_x() + root.winfo_width()  - 240
    y = root.winfo_y() + root.winfo_height() - 80
    toast.geometry(f"+{x}+{y}")
    toast.after(2200, toast.destroy)


# =============================================================================
#  ВИПРАВЛЕННЯ #1: ScrollManager — централізований менеджер скролу
#  Замість bind_all на кожному canvas, підписка/відписка при наведенні миші.
# =============================================================================

class ScrollManager:
    """
    Керує прив'язкою <MouseWheel> для canvas-скрол-контейнерів.

    Виправлення критичної проблеми: bind_all() діяв глобально і накопичувався
    з кожним новим canvas, що призводило до непередбачуваного скролінгу.
    Тепер кожен canvas підписується тільки коли курсор над ним.
    """
    def attach(self, canvas: tk.Canvas) -> None:
        canvas.bind("<Enter>",  lambda _: self._bind(canvas))
        canvas.bind("<Leave>",  lambda _: self._unbind(canvas))

    @staticmethod
    def _bind(canvas: tk.Canvas) -> None:
        canvas.bind_all(
            "<MouseWheel>",
            lambda e: canvas.yview_scroll(int(-1 * e.delta / 120), "units"),
        )

    @staticmethod
    def _unbind(canvas: tk.Canvas) -> None:
        canvas.unbind_all("<MouseWheel>")


_scroll_mgr = ScrollManager()


def _scrollable_canvas(container: tk.Misc) -> tuple[tk.Canvas, tk.Frame]:
    """Повертає (canvas, inner_frame) з прокруткою та коректним scroll-bind."""
    C = COLORS
    container.columnconfigure(0, weight=1)
    container.rowconfigure(0, weight=1)

    canvas = tk.Canvas(container, bg=C["bg_main"], highlightthickness=0)
    sb = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
    canvas.configure(yscrollcommand=sb.set)
    canvas.grid(row=0, column=0, sticky="nsew")
    sb.grid(row=0, column=1, sticky="ns")

    inner = tk.Frame(canvas, bg=C["bg_main"])
    cw    = canvas.create_window((0, 0), window=inner, anchor="nw")

    def _on_conf(_: object) -> None:
        canvas.configure(scrollregion=canvas.bbox("all"))
        canvas.itemconfig(cw, width=canvas.winfo_width())

    inner.bind("<Configure>", _on_conf)
    canvas.bind("<Configure>", _on_conf)

    # ВИПРАВЛЕННЯ: замість bind_all — через ScrollManager
    _scroll_mgr.attach(canvas)

    return canvas, inner


# =============================================================================
#  ВИПРАВЛЕННЯ #2 + #3: ID Generator
#  Генерує унікальні ID без колізій після видалення записів.
# =============================================================================

class IdGenerator:
    """
    Генерує ID, що не конфліктують з наявними.

    Виправлення проблем:
    - #2: нестабільний пошук за lstrip("0")
    - #3: len(records)+1 призводив до колізій після видалень
    """
    @staticmethod
    def next_id(existing: list) -> str:
        """Повертає max(existing_numeric_ids) + 1, або 1 якщо список порожній."""
        max_id = 0
        for item in existing:
            raw = str(item.id) if hasattr(item, "id") else str(item[0])
            try:
                val = int(raw)
                if val > max_id:
                    max_id = val
            except (ValueError, TypeError):
                pass
        return f"{max_id + 1:03d}"

    @staticmethod
    def normalize_id(raw: str) -> str:
        """Нормалізує ID до числового рядка без ведучих нулів для порівняння."""
        try:
            return str(int(raw))
        except (ValueError, TypeError):
            return raw


# =============================================================================
#  БАЗОВИЙ КЛАС ДЛЯ РЕЄСТРІВ — виправлення #5: усуває дублювання коду
# =============================================================================

class BaseRegistryTab:
    """
    Спільна логіка для RegistryTab та RiskRegistryTab.
    Виправлення #5: ~60% дублювання коду винесено у базовий клас.
    """

    data_file: str = ""
    frame: ttk.Frame

    def __init__(
        self,
        parent: tk.Misc,
        on_data_change: Callable | None = None,
    ) -> None:
        self.parent         = parent
        self.on_data_change = on_data_change
        self.frame          = ttk.Frame(parent)
        self.all_records:   list = []

    # ── Публічний API (виправлення #4: більше немає _save_data через noqa) ──

    def save(self) -> None:
        """Публічний метод збереження даних."""
        self._save_data()

    # ── Пошук запису ─────────────────────────────────────────────────────────

    def find_record(self, idx_str: str):
        """
        Виправлення #2: надійний пошук за нормалізованим числовим ID.
        Більше не використовує lstrip("0") — порівнює int(id).
        """
        normalized = IdGenerator.normalize_id(idx_str)
        for r in self.all_records:
            raw = r.id if hasattr(r, "id") else str(r[0])
            if IdGenerator.normalize_id(raw) == normalized:
                return r
        return None

    # ── Збереження / завантаження ─────────────────────────────────────────────

    def _save_data(self) -> None:
        """
        Виправлення #8: зберігає як список словників, а не список кортежів.
        Це робить формат файлу стійким до зміни структури полів.
        """
        try:
            data = [
                (asdict(r) if hasattr(r, "__dataclass_fields__") else r)
                for r in self.all_records
            ]
            with open(self.data_file, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except OSError as e:
            messagebox.showerror("Помилка збереження", str(e))

    # ── Допоміжні ─────────────────────────────────────────────────────────────

    def _notify_change(self) -> None:
        if self.on_data_change:
            self.on_data_change(self.all_records)

    def get_frame(self) -> ttk.Frame:
        return self.frame


# =============================================================================
#  ДЕТАЛЬНЕ ВІКНО ЗАПИСУ ПОДІЇ
# =============================================================================

class EventDetailWindow:
    """Спливаюче вікно для перегляду та редагування збереженої події."""

    def __init__(
        self,
        parent_root:     tk.Misc,
        record:          EventRecord,
        all_records:     list[EventRecord],
        save_callback:   Callable[[str, EventRecord], None],
        delete_callback: Callable[[str], None],
        toast_callback:  Callable[[str], None],
    ) -> None:
        self.parent_root     = parent_root
        self.record          = record
        self.all_records     = all_records
        self.save_callback   = save_callback
        self.delete_callback = delete_callback
        self.toast_callback  = toast_callback
        self.is_edit_mode    = False
        self._build_window()

    def _build_window(self) -> None:
        C = COLORS
        self.win = tk.Toplevel(self.parent_root)
        self.win.title(f"Подiя #{self.record.id}  —  {self.record.entity}")
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

        risk_color = RISK_COLORS.get(self.record.risk_type, C["accent"])
        tk.Frame(header, bg=risk_color, width=4).grid(row=0, column=0, sticky="ns")

        title_frame = tk.Frame(header, bg=C["bg_header"])
        title_frame.grid(row=0, column=1, sticky="ew", padx=16, pady=10)

        self.lbl_title = tk.Label(
            title_frame, text=f"Запис #{self.record.id}",
            bg=C["bg_header"], fg=C["accent_muted"], font=FONT_HEADING)
        self.lbl_title.pack(anchor="w")

        self.lbl_subtitle = tk.Label(
            title_frame, text=self.record.entity,
            bg=C["bg_header"], fg=C["text_muted"], font=FONT_DEFAULT)
        self.lbl_subtitle.pack(anchor="w")

        sc = self._status_color(self.record.status)
        self.lbl_status_badge = tk.Label(
            header, text=f"  {self.record.status}  ",
            bg=sc, fg="white", font=FONT_SMALL_BOLD, pady=3)
        self.lbl_status_badge.grid(row=0, column=2, padx=16, pady=18)

        # ── Scroll ──────────────────────────────────────────────────────
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
        _scroll_mgr.attach(canvas)  # ВИПРАВЛЕННЯ #1

        self.content.columnconfigure(0, weight=1)
        self.content.columnconfigure(1, weight=1)
        self._build_view_content()

        # ── Bottom buttons ──────────────────────────────────────────────
        btn_bar = tk.Frame(self.win, bg=C["bg_header"], height=46)
        btn_bar.grid(row=2, column=0, columnspan=2, sticky="ew")
        btn_bar.grid_propagate(False)

        left_btns = tk.Frame(btn_bar, bg=C["bg_header"])
        left_btns.pack(side="left", padx=12, pady=8)

        self.btn_edit = make_button(
            left_btns, "Редагувати",
            bg=C["accent_warning"], fg=C["bg_main"],
            activebackground="#d97706", activeforeground="white",
            font=FONT_BOLD, padx=14, pady=4,
            command=self._toggle_edit_mode)
        self.btn_edit.pack(side="left", padx=(0, 8))

        self.btn_save = make_button(
            left_btns, "Зберегти змiни",
            bg=C["accent_success"], activebackground="#16a34a",
            font=FONT_BOLD, padx=14, pady=4,
            command=self._save_changes)
        self.btn_save.pack_forget()

        self.btn_cancel_edit = make_button(
            left_btns, "Скасувати",
            bg=C["bg_surface"], fg=C["text_muted"],
            activebackground=C["bg_surface_alt"],
            activeforeground=C["text_primary"],
            font=FONT_DEFAULT, padx=12, pady=4,
            command=self._cancel_edit)
        self.btn_cancel_edit.pack_forget()

        right_btns = tk.Frame(btn_bar, bg=C["bg_header"])
        right_btns.pack(side="right", padx=12, pady=8)

        make_button(right_btns, "Видалити",
                    bg=C["accent_danger"], activebackground="#dc2626",
                    font=FONT_BOLD, padx=14, pady=4,
                    command=self._delete_record).pack(side="right", padx=(8, 0))

        make_button(right_btns, "Закрити",
                    bg=C["bg_surface"], fg=C["text_primary"],
                    activebackground=C["bg_surface_alt"],
                    activeforeground=C["text_primary"],
                    font=FONT_DEFAULT, padx=12, pady=4,
                    command=self.win.destroy).pack(side="right")

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

    def _build_view_content(self) -> None:
        for w in self.content.winfo_children():
            w.destroy()
        r   = self.record
        row = 0

        _build_section_label(self.content, "Iнформацiя про пiдприємство", row); row += 1
        _build_info_cell(self.content, "Пiдприємство", r.entity, row, 0)
        _build_info_cell(self.content, "Прiоритет", r.priority, row, 1,
                         value_color=self._priority_color(r.priority)); row += 1

        _build_section_label(self.content, "Опис подiї / ризику", row); row += 1
        _build_info_cell(self.content, "Назва подiї", r.event_name, row, 0)
        _build_info_cell(self.content, "Тип ризику", r.risk_type, row, 1,
                         value_color=RISK_COLORS.get(r.risk_type)); row += 1

        _build_info_cell(self.content, "Статус", r.status, row, 0,
                         value_color=self._status_color(r.status))
        _build_info_cell(self.content, "Задiянi пiдроздiли / особи",
                         r.involved, row, 1); row += 1

        _build_section_label(self.content, "Дати", row); row += 1
        _build_info_cell(self.content, "Дата подiї",    r.event_date,  row, 0)
        _build_info_cell(self.content, "Дата виявлення", r.detect_date, row, 1); row += 1

        _build_section_label(self.content, "Деталi подiї", row); row += 1
        _build_text_block(self.content, "Детальний опис подiї", r.description, row); row += 1
        _build_text_block(self.content, "Вжитi заходи",          r.measures,    row); row += 1

        tk.Frame(self.content, bg=COLORS["bg_main"], height=12).grid(
            row=row, column=0, columnspan=2)

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
                     font=FONT_BOLD).pack(side="left", padx=8)
            return r + 1

        def lbl(text: str, r: int) -> None:
            tk.Label(self.content, text=text, bg=C["bg_main"],
                     fg=C["text_subtle"], font=FONT_SMALL).grid(
                row=r, column=0, sticky="w", padx=10, pady=(6, 0))

        row = 0
        row = section("Пiдприємство та подiя", row)

        lbl("Скорочена назва пiдприємства:", row); row += 1
        self.e_entity = make_dark_entry(self.content, accent=C["accent_warning"])
        self.e_entity.grid(row=row, column=0, sticky="ew", padx=10, pady=(2, 0))
        self.e_entity.insert(0, rec.entity); row += 1

        lbl("Назва подiї:", row); row += 1
        self.e_event = make_dark_combo(self.content, values=EVENT_TYPES)
        self.e_event.grid(row=row, column=0, sticky="ew", padx=10, pady=(2, 0))
        self.e_event.set(rec.event_name); row += 1

        lbl("Тип ризику:", row); row += 1
        self.e_risk = make_dark_combo(self.content, values=RISK_TYPES)
        self.e_risk.grid(row=row, column=0, sticky="ew", padx=10, pady=(2, 0))
        self.e_risk.set(rec.risk_type if rec.risk_type != "—" else ""); row += 1

        lbl("Задiянi пiдроздiли / особи:", row); row += 1
        self.e_involved = make_dark_text(self.content, height=2, wrap="word")
        self.e_involved.grid(row=row, column=0, sticky="ew", padx=10, pady=(2, 0))
        if rec.involved and rec.involved != "—":
            self.e_involved.insert("1.0", rec.involved)
        row += 1

        row = section("Дати", row)
        date_f = tk.Frame(self.content, bg=C["bg_main"])
        date_f.grid(row=row, column=0, sticky="w", padx=10, pady=4); row += 1

        for ci, (lbl_t, attr, val) in enumerate([
            ("Дата виявлення:", "e_detect",     rec.detect_date),
            ("Дата подiї:",     "e_event_date", rec.event_date),
        ]):
            tk.Label(date_f, text=lbl_t, bg=C["bg_main"],
                     fg=C["text_subtle"], font=FONT_SMALL).grid(
                row=0, column=ci, padx=(0 if ci == 0 else 20, 0), sticky="w")
            e = make_dark_entry(date_f, accent=C["accent_warning"], width=14)
            e.grid(row=1, column=ci, padx=(0 if ci == 0 else 20, 0), pady=2)
            if val and val != "—":
                e.insert(0, val)
            else:
                add_placeholder(e, "дд.мм.рррр")
            setattr(self, attr, e)

        row = section("Прiоритет та статус", row)
        ps_f = tk.Frame(self.content, bg=C["bg_main"])
        ps_f.grid(row=row, column=0, sticky="w", padx=10, pady=4); row += 1

        tk.Label(ps_f, text="Прiоритет:", bg=C["bg_main"],
                 fg=C["text_subtle"], font=FONT_SMALL).grid(row=0, column=0, sticky="w")
        self.e_priority = make_dark_combo(
            ps_f, values=["Критичний", "Високий", "Середнiй", "Низький"], width=14)
        self.e_priority.grid(row=1, column=0, pady=(2, 0), padx=(0, 20))
        self.e_priority.set(rec.priority)

        tk.Label(ps_f, text="Статус:", bg=C["bg_main"],
                 fg=C["text_subtle"], font=FONT_SMALL).grid(row=0, column=1, sticky="w")
        self.e_status = make_dark_combo(
            ps_f, values=["Вiдкрито", "В обробцi", "Закрито", "Вирiшено"], width=14)
        self.e_status.grid(row=1, column=1, pady=(2, 0))
        self.e_status.set(rec.status)

        row = section("Деталi подiї", row)

        lbl("Детальний опис подiї:", row); row += 1
        self.e_description = make_dark_text(self.content, height=4, wrap="word")
        self.e_description.grid(row=row, column=0, sticky="ew", padx=10, pady=(2, 0))
        if rec.description and rec.description != "—":
            self.e_description.insert("1.0", rec.description)
        row += 1

        lbl("Вжитi заходи:", row); row += 1
        self.e_measures = make_dark_text(self.content, height=3, wrap="word")
        self.e_measures.grid(row=row, column=0, sticky="ew", padx=10, pady=(2, 0))
        if rec.measures and rec.measures != "—":
            self.e_measures.insert("1.0", rec.measures)
        row += 1

        tk.Frame(self.content, bg=C["bg_main"], height=16).grid(row=row, column=0)

    def _toggle_edit_mode(self) -> None:
        self.is_edit_mode = True
        self._build_edit_content()
        self.btn_edit.pack_forget()
        self.btn_save.pack(side="left", padx=(0, 8))
        self.btn_cancel_edit.pack(side="left")
        self.lbl_title.configure(
            text=f"Редагування запису #{self.record.id}",
            fg=COLORS["accent_warning"])

    def _cancel_edit(self) -> None:
        self.is_edit_mode = False
        self._build_view_content()
        self.btn_save.pack_forget()
        self.btn_cancel_edit.pack_forget()
        self.btn_edit.pack(side="left", padx=(0, 8))
        self.lbl_title.configure(
            text=f"Запис #{self.record.id}",
            fg=COLORS["accent_muted"])

    def _save_changes(self) -> None:
        detect  = self.e_detect.get().strip()
        event_d = self.e_event_date.get().strip()

        # ВИПРАВЛЕННЯ #6: перевіряємо конкретні значення, а не ловимо все підряд
        for val, label in [(detect, "дати виявлення"), (event_d, "дати подiї")]:
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

        old_id = self.record.id
        new_record = EventRecord(
            id=self.record.id,
            entity=entity,
            event_name=event,
            involved=self.e_involved.get("1.0", tk.END).strip(),
            risk_type=self.e_risk.get().strip() or "—",
            event_date=event_d or "—",
            description=self.e_description.get("1.0", tk.END).strip(),
            measures=self.e_measures.get("1.0", tk.END).strip(),
            detect_date=detect or "—",
            priority=self.e_priority.get().strip() or "Середнiй",
            status=self.e_status.get().strip()   or "Вiдкрито",
        )
        self.record = new_record
        self.save_callback(old_id, new_record)

        self.lbl_subtitle.configure(text=entity)
        sc = self._status_color(new_record.status)
        self.lbl_status_badge.configure(text=f"  {new_record.status}  ", bg=sc)
        self._cancel_edit()
        self.toast_callback("Запис оновлено")

    def _delete_record(self) -> None:
        if not messagebox.askyesno(
            "Пiдтвердження",
            f"Видалити запис #{self.record.id}?\nЦю дiю не можна скасувати.",
            parent=self.win):
            return
        self.delete_callback(self.record.id)
        self.toast_callback("Запис видалено")
        self.win.destroy()


# =============================================================================
#  ВКЛАДКА: РЕЄСТР СУТТЄВИХ ПОДІЙ
# =============================================================================

class RegistryTab(BaseRegistryTab):

    data_file = DATA_FILE

    def __init__(
        self,
        parent: tk.Misc,
        on_data_change: Callable[[list[EventRecord]], None] | None = None,
    ) -> None:
        super().__init__(parent, on_data_change)
        self._build_ui()
        self._load_data()

    def _build_ui(self) -> None:
        C = COLORS
        self.frame.rowconfigure(1, weight=1)
        self.frame.columnconfigure(0, weight=1)

        header = tk.Frame(self.frame, bg=C["bg_header"], height=56)
        header.grid(row=0, column=0, sticky="ew")
        header.columnconfigure(1, weight=1)
        header.grid_propagate(False)

        tk.Label(header, text="РЕЄСТР СУТТЄВИХ ПОДIЙ",
                 bg=C["bg_header"], fg=C["accent_muted"],
                 font=FONT_TITLE).grid(row=0, column=0, padx=20, pady=14, sticky="w")

        sf = tk.Frame(header, bg=C["bg_header"])
        sf.grid(row=0, column=1, sticky="e", padx=20)
        tk.Label(sf, text="Пошук:", bg=C["bg_header"],
                 fg=C["text_muted"], font=FONT_SMALL).pack(side="left", padx=(0, 6))

        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *_: self._apply_filter())
        tk.Entry(sf, textvariable=self.search_var,
                 bg=C["bg_input"], fg=C["text_primary"],
                 insertbackground=C["text_primary"],
                 relief="flat", bd=2, font=FONT_DEFAULT,
                 width=34).pack(side="left", padx=(0, 8), ipady=2)
        make_button(sf, "Скинути",
                    bg=C["bg_surface"], fg=C["text_muted"],
                    activebackground=C["bg_surface_alt"],
                    activeforeground=C["text_primary"],
                    font=FONT_SMALL, padx=8, pady=2,
                    command=self._reset_filter).pack(side="left")

        paned = ttk.PanedWindow(self.frame, orient="horizontal")
        paned.grid(row=1, column=0, sticky="nsew")
        lw = ttk.Frame(paned)
        rw = ttk.Frame(paned)
        paned.add(lw, weight=4)
        paned.add(rw, weight=7)

        self._build_form(lw)
        self._build_table(rw)

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
        _scroll_mgr.attach(canvas)  # ВИПРАВЛЕННЯ #1
        form.columnconfigure(0, weight=1)

        def section(txt: str, r: int) -> int:
            f = tk.Frame(form, bg=C["bg_main"])
            f.grid(row=r, column=0, sticky="ew", padx=16, pady=(16, 4))
            tk.Frame(f, bg=C["accent"], width=3, height=16).pack(side="left")
            tk.Label(f, text=txt, bg=C["bg_main"], fg=C["accent"],
                     font=FONT_BOLD).pack(side="left", padx=8)
            return r + 1

        def lbl(text: str, r: int) -> None:
            tk.Label(form, text=text, bg=C["bg_main"],
                     fg=C["text_subtle"], font=FONT_SMALL).grid(
                row=r, column=0, sticky="w", padx=16, pady=(4, 0))

        def field(lbl_txt: str, r: int, factory: Callable) -> tuple[tk.Widget, int]:
            lbl(lbl_txt, r)
            w = factory()
            w.grid(row=r + 1, column=0, sticky="ew", padx=16, pady=(2, 0))
            return w, r + 2

        row = 0

        badge_f = tk.Frame(form, bg=C["bg_main"])
        badge_f.grid(row=row, column=0, sticky="ew", padx=16, pady=(12, 0))
        tk.Label(badge_f, text="  + НОВИЙ ЗАПИС  ",
                 bg=C["accent"], fg="white",
                 font=FONT_SMALL_BOLD, pady=3).pack(anchor="w")
        row += 1

        row = section("Iнформацiя про пiдприємство та особу", row)
        self.ent_entity,   row = field("Скорочена назва пiдприємства:", row,
                                       lambda: make_dark_entry(form))
        self.ent_position, row = field("Посада:", row, lambda: make_dark_entry(form))
        self.ent_reporter, row = field("ПIБ особи, що звiтує:", row,
                                       lambda: make_dark_entry(form))

        row = section("Опис подiї / ризику", row)
        self.cb_event, row = field("Назва подiї:", row,
                                   lambda: make_dark_combo(form, values=EVENT_TYPES))
        self.cb_risk, row  = field("Тип ризику:", row,
                                   lambda: make_dark_combo(form, values=RISK_TYPES))

        lbl("Задiянi пiдроздiли / особи:", row); row += 1
        self.txt_involved = make_dark_text(form, height=2, wrap="word")
        self.txt_involved.grid(row=row, column=0, sticky="ew", padx=16, pady=(2, 0))
        row += 1

        row = section("Фiнансовий вплив (млн грн)", row)
        fin_f = tk.Frame(form, bg=C["bg_main"])
        fin_f.grid(row=row, column=0, sticky="ew", padx=16, pady=4)
        fin_f.columnconfigure((0, 1, 2, 3), weight=1); row += 1

        for ci, title in enumerate(["Втрати", "Резерв", "Запланованi втрати", "Вiдшкодування"]):
            tk.Label(fin_f, text=title, bg=C["bg_main"],
                     fg=C["text_muted"], font=FONT_TINY).grid(
                row=0, column=ci, sticky="w", padx=4)

        self.ent_loss    = make_dark_entry(fin_f, width=11)
        self.ent_reserve = make_dark_entry(fin_f, width=11)
        self.ent_planned = make_dark_entry(fin_f, width=11)
        self.ent_refund  = make_dark_entry(fin_f, width=11)
        for ci, e in enumerate([self.ent_loss, self.ent_reserve,
                                 self.ent_planned, self.ent_refund]):
            e.grid(row=1, column=ci, padx=4, pady=2, sticky="ew")

        net_f = tk.Frame(form, bg=C["bg_main"])
        net_f.grid(row=row, column=0, sticky="ew", padx=16, pady=(2, 6)); row += 1
        tk.Label(net_f, text="Чистий вплив (млн грн):",
                 bg=C["bg_main"], fg=C["text_muted"], font=FONT_SMALL).pack(side="left")
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
            except ValueError:
                self.lbl_net.configure(text="—", fg=C["text_muted"])

        self.ent_loss.bind("<KeyRelease>",   _upd_net)
        self.ent_refund.bind("<KeyRelease>", _upd_net)

        row = section("Деталi подiї", row)
        for lbl_txt, attr, h in [
            ("Вплив на iншi пiдприємства:",   "txt_impact",      2),
            ("Нефiнансовий / якiсний вплив:",  "txt_qualitative", 2),
            ("Детальний опис подiї:",           "txt_description", 4),
            ("Вжитi заходи:",                   "txt_measures",    3),
        ]:
            lbl(lbl_txt, row); row += 1
            t = make_dark_text(form, height=h, wrap="word")
            t.grid(row=row, column=0, sticky="ew", padx=16, pady=(2, 0))
            setattr(self, attr, t); row += 1

        row = section("Дати", row)
        date_f = tk.Frame(form, bg=C["bg_main"])
        date_f.grid(row=row, column=0, sticky="w", padx=16, pady=4); row += 1
        for ci, (lbl_t, attr) in enumerate([
            ("Дата виявлення:", "ent_detect"),
            ("Дата подiї:",    "ent_event_date"),
        ]):
            tk.Label(date_f, text=lbl_t, bg=C["bg_main"],
                     fg=C["text_subtle"], font=FONT_SMALL).grid(
                row=0, column=ci, padx=(0 if ci == 0 else 20, 0), sticky="w")
            e = make_dark_entry(date_f, width=14)
            e.grid(row=1, column=ci, padx=(0 if ci == 0 else 20, 0), pady=2)
            add_placeholder(e, "дд.мм.рррр")
            setattr(self, attr, e)

        lbl("Прiоритет:", row); row += 1
        self.cb_priority = make_dark_combo(
            form, values=["Критичний", "Високий", "Середнiй", "Низький"])
        self.cb_priority.grid(row=row, column=0, sticky="w", padx=16, pady=(2, 0)); row += 1

        lbl("Статус:", row); row += 1
        self.cb_status = make_dark_combo(
            form, values=["Вiдкрито", "В обробцi", "Закрито", "Вирiшено"])
        self.cb_status.grid(row=row, column=0, sticky="w", padx=16, pady=(2, 0)); row += 1

        btn_f = tk.Frame(form, bg=C["bg_main"])
        btn_f.grid(row=row, column=0, sticky="ew", padx=16, pady=(16, 20))
        btn_f.columnconfigure((0, 1), weight=1)

        make_button(btn_f, "Очистити",
                    bg=C["bg_surface"], fg=C["text_muted"],
                    activebackground=C["bg_surface_alt"],
                    activeforeground=C["text_primary"],
                    padx=14, pady=6,
                    command=self._clear_form).grid(row=0, column=0, padx=4, sticky="ew")
        make_button(btn_f, "Додати запис",
                    bg=C["accent"], activebackground=C["accent_soft"],
                    font=FONT_BOLD, padx=14, pady=6,
                    command=self._add_record).grid(row=0, column=1, padx=4, sticky="ew")

    def _build_table(self, container: tk.Misc) -> None:
        C = COLORS
        container.rowconfigure(1, weight=1)
        container.columnconfigure(0, weight=1)

        toolbar = tk.Frame(container, bg=C["bg_surface"], height=40)
        toolbar.grid(row=0, column=0, sticky="ew")
        toolbar.grid_propagate(False)

        tk.Label(toolbar, text="Записи", bg=C["bg_surface"],
                 fg=C["text_muted"], font=FONT_SMALL).pack(side="left", padx=12, pady=8)
        self.lbl_count = tk.Label(toolbar, text="0", bg=C["bg_surface"],
                                  fg=C["accent"], font=FONT_SMALL_BOLD)
        self.lbl_count.pack(side="left", pady=8)

        tk.Label(toolbar, text="  |  Ризик:", bg=C["bg_surface"],
                 fg=C["text_muted"], font=FONT_SMALL).pack(side="left", pady=8)
        self.filter_risk = make_dark_combo(toolbar, values=["Всi"] + RISK_TYPES, width=16)
        self.filter_risk.set("Всi")
        self.filter_risk.pack(side="left", padx=6, pady=8)
        self.filter_risk.bind("<<ComboboxSelected>>", lambda _: self._apply_filter())

        tk.Label(toolbar, text="Статус:", bg=C["bg_surface"],
                 fg=C["text_muted"], font=FONT_SMALL).pack(side="left", pady=8)
        self.filter_status = make_dark_combo(
            toolbar,
            values=["Всi", "Вiдкрито", "В обробцi", "Закрито", "Вирiшено"],
            width=12)
        self.filter_status.set("Всi")
        self.filter_status.pack(side="left", padx=6, pady=8)
        self.filter_status.bind("<<ComboboxSelected>>", lambda _: self._apply_filter())

        for txt, cmd, bg in [
            ("Переглянути", self._open_selected_detail, C["accent"]),
            ("Дублювати",   self._duplicate_record,     C["bg_surface_alt"]),
            ("Видалити",    self._delete_selected,       C["accent_danger"]),
        ]:
            make_button(toolbar, txt, bg=bg,
                        fg="white" if bg != C["bg_surface_alt"] else C["text_primary"],
                        activebackground=bg, activeforeground="white",
                        font=FONT_SMALL, padx=10, pady=3,
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
            "id":       ("№",            46),
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
            self.tree.heading(col, text=txt, command=lambda c=col: self._sort_tree(c))
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
                 font=("Arial", 7, "italic")).pack(side="left", padx=8, pady=4)

        detail_f = tk.Frame(container, bg=C["bg_surface"])
        detail_f.grid(row=3, column=0, sticky="ew")
        detail_f.columnconfigure((0, 1), weight=1)

        for ci, (lbl_t, attr) in enumerate([
            ("Опис подiї", "det_desc"),
            ("Вжитi заходи", "det_measures"),
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
                    font=FONT_SMALL, padx=12, pady=4,
                    command=self._export_csv).pack(side="left", padx=(0, 6))
        if pd:
            make_button(exp_bar, "Експорт Excel",
                        bg=C["accent_success"], activebackground="#16a34a",
                        font=FONT_SMALL, padx=12, pady=4,
                        command=self._export_excel).pack(side="left", padx=(0, 6))
        make_button(exp_bar, "Iмпорт JSON",
                    bg=C["bg_surface"], fg=C["text_primary"],
                    activebackground=C["bg_surface_alt"],
                    activeforeground=C["text_primary"],
                    font=FONT_SMALL, padx=12, pady=4,
                    command=self._import_json).pack(side="left")

    # ── Events ───────────────────────────────────────────────────────────────

    def _on_double_click(self, event: tk.Event) -> None:
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
        iid = sel[0]
        rec = self.find_record(self.tree.set(iid, "id"))
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

    def _on_detail_save(self, iid: str, old_id: str, new_record: EventRecord) -> None:
        normalized_old = IdGenerator.normalize_id(old_id)
        for i, r in enumerate(self.all_records):
            if IdGenerator.normalize_id(r.id) == normalized_old:
                self.all_records[i] = new_record
                break
        try:
            self.tree.item(iid, values=(
                new_record.id, new_record.entity, new_record.event_name,
                new_record.risk_type, new_record.priority, new_record.status,
                new_record.event_date, new_record.involved,
                new_record.description, new_record.measures,
            ))
        except tk.TclError:
            pass
        self._recolor_rows()
        self._save_data()
        self._notify_change()

    def _on_detail_delete(self, iid: str, idx_str: str) -> None:
        try:
            self.tree.delete(iid)
        except tk.TclError:
            pass
        normalized = IdGenerator.normalize_id(idx_str)
        self.all_records = [
            r for r in self.all_records
            if IdGenerator.normalize_id(r.id) != normalized
        ]
        self._recolor_rows()
        self._save_data()
        self._update_count()
        self._notify_change()

    # ── Таблиця ──────────────────────────────────────────────────────────────

    def _sort_tree(self, col: str) -> None:
        data = [(self.tree.set(iid, col), iid) for iid in self.tree.get_children("")]
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

    # ── Дані ─────────────────────────────────────────────────────────────────

    def _load_data(self) -> None:
        self.all_records.clear()
        self.tree.delete(*self.tree.get_children())
        if not os.path.exists(self.data_file):
            return
        try:
            with open(self.data_file, "r", encoding="utf-8") as f:
                raw = json.load(f)
            if not isinstance(raw, list):
                raise ValueError("Очiкується список записiв у JSON")
            for item in raw:
                # ВИПРАВЛЕННЯ #8: підтримка нового формату (dict) і старого (list)
                if isinstance(item, dict):
                    rec = EventRecord.from_dict(item)
                elif isinstance(item, (list, tuple)):
                    rec = EventRecord.from_list(list(item))
                else:
                    continue
                self.all_records.append(rec)
                self._insert_tree_row(rec)
        except (json.JSONDecodeError, ValueError, OSError) as e:
            # ВИПРАВЛЕННЯ #6: конкретні типи помилок
            messagebox.showerror("Помилка завантаження", str(e))
        self._update_count()
        self._notify_change()

    def _insert_tree_row(self, rec: EventRecord) -> str:
        iid = self.tree.insert("", tk.END, values=(
            rec.id, rec.entity, rec.event_name, rec.risk_type,
            rec.priority, rec.status, rec.event_date,
            rec.involved, rec.description, rec.measures,
        ))
        self._recolor_rows()
        return iid

    # ── Форма ─────────────────────────────────────────────────────────────────

    def _get_form_data(self) -> EventRecord | None:
        detect  = self.ent_detect.get().strip()
        event_d = self.ent_event_date.get().strip()

        for val, label in [(detect, "дати виявлення"), (event_d, "дати подiї")]:
            if val in ("дд.мм.рррр", ""):
                continue
            if not is_valid_date(val):
                messagebox.showwarning(
                    "Помилка",
                    f"Неправильний формат {label} (очiкується дд.мм.рррр)")
                return None

        entity = self.ent_entity.get().strip()
        event  = self.cb_event.get().strip()
        # ВИПРАВЛЕННЯ #10: валідація обов'язкових полів тут, а не після побудови
        if not entity or not event:
            messagebox.showwarning(
                "Обов'язковi поля",
                "Заповнiть назву пiдприємства та назву подiї")
            return None

        detect  = "" if detect  == "дд.мм.рррр" else detect
        event_d = "" if event_d == "дд.мм.рррр" else event_d

        # ВИПРАВЛЕННЯ #3: надійна генерація ID
        new_id = IdGenerator.next_id(self.all_records)
        return EventRecord(
            id=new_id,
            entity=entity,
            event_name=event,
            involved=self.txt_involved.get("1.0", tk.END).strip(),
            risk_type=self.cb_risk.get().strip() or "—",
            event_date=event_d or "—",
            description=self.txt_description.get("1.0", tk.END).strip(),
            measures=self.txt_measures.get("1.0", tk.END).strip(),
            detect_date=detect or "—",
            priority=self.cb_priority.get().strip() or "Середнiй",
            status=self.cb_status.get().strip()   or "Вiдкрито",
        )

    def _clear_form(self) -> None:
        for w in [self.ent_entity, self.ent_position, self.ent_reporter,
                  self.ent_loss, self.ent_reserve, self.ent_planned, self.ent_refund]:
            w.delete(0, tk.END)
            w.configure(fg=COLORS["text_primary"])
        for w in [self.cb_event, self.cb_risk, self.cb_priority, self.cb_status]:
            w.set("")
        for w in [self.txt_involved, self.txt_impact,
                  self.txt_qualitative, self.txt_description, self.txt_measures]:
            w.delete("1.0", tk.END)
        self.lbl_net.configure(text="0.00", fg=COLORS["accent_success"])
        for e, ph in [(self.ent_detect, "дд.мм.рррр"),
                      (self.ent_event_date, "дд.мм.рррр")]:
            e.delete(0, tk.END)
            add_placeholder(e, ph)

    def _add_record(self) -> None:
        data = self._get_form_data()
        if not data:
            return
        self.all_records.append(data)
        self._insert_tree_row(data)
        self._clear_form()
        self._save_data()
        self._notify_change()
        self._update_count()
        self._show_toast("Запис додано")

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
        normalized = IdGenerator.normalize_id(idx_str)
        self.all_records = [
            r for r in self.all_records
            if IdGenerator.normalize_id(r.id) != normalized
        ]
        self._recolor_rows()
        self._save_data()
        self._update_count()
        self._notify_change()
        self._show_toast("Запис видалено")

    def _duplicate_record(self) -> None:
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Дублювання", "Оберiть запис для дублювання")
            return
        rec = self.find_record(self.tree.set(sel[0], "id"))
        if not rec:
            return
        # ВИПРАВЛЕННЯ #3: новий унікальний ID
        import dataclasses
        new_rec = dataclasses.replace(rec, id=IdGenerator.next_id(self.all_records))
        self.all_records.append(new_rec)
        self._insert_tree_row(new_rec)
        self._save_data()
        self._update_count()
        self._notify_change()
        self._show_toast("Запис продубльовано")

    def _apply_filter(self) -> None:
        q      = self.search_var.get().strip().lower()
        risk   = self.filter_risk.get()
        status = self.filter_status.get()
        self.tree.delete(*self.tree.get_children())
        for rec in self.all_records:
            row_str = " ".join([
                rec.id, rec.entity, rec.event_name, rec.involved,
                rec.risk_type, rec.event_date, rec.description,
                rec.measures, rec.detect_date, rec.priority, rec.status,
            ]).lower()
            if q and q not in row_str:
                continue
            if risk   != "Всi" and rec.risk_type != risk:
                continue
            if status != "Всi" and rec.status    != status:
                continue
            self._insert_tree_row(rec)
        self._update_count()

    def _reset_filter(self) -> None:
        self.search_var.set("")
        self.filter_risk.set("Всi")
        self.filter_status.set("Всi")
        self.tree.delete(*self.tree.get_children())
        for rec in self.all_records:
            self._insert_tree_row(rec)
        self._update_count()

    def _on_select(self, _: object | None = None) -> None:
        sel = self.tree.selection()
        if not sel:
            return
        rec = self.find_record(self.tree.set(sel[0], "id"))
        if not rec:
            return
        for widget, text in [(self.det_desc,     rec.description),
                             (self.det_measures, rec.measures)]:
            widget.configure(state="normal")
            widget.delete("1.0", tk.END)
            widget.insert("1.0", text)
            widget.configure(state="disabled")

    def _update_count(self) -> None:
        self.lbl_count.configure(text=f" {len(self.tree.get_children())}")

    def _show_toast(self, msg: str) -> None:
        _show_toast(self.frame, msg)

    # ── Експорт / Імпорт ─────────────────────────────────────────────────────

    _EVENT_HEADERS = [
        "ID", "Пiдприємство", "Назва подiї", "Задiянi особи",
        "Тип ризику", "Дата подiї", "Опис", "Заходи",
        "Дата виявлення", "Прiоритет", "Статус",
    ]

    def _export_csv(self) -> None:
        if not self.tree.get_children():
            messagebox.showinfo("Експорт", "Таблиця порожня"); return
        path = filedialog.asksaveasfilename(
            defaultextension=".csv", filetypes=[("CSV файли", "*.csv")],
            title="Зберегти як CSV")
        if not path:
            return
        try:
            with open(path, "w", newline="", encoding="utf-8-sig") as f:
                w = csv.writer(f)
                w.writerow(self._EVENT_HEADERS)
                for rec in self.all_records:
                    w.writerow(rec.to_list())
            self._show_toast("CSV збережено")
        except OSError as e:
            messagebox.showerror("Помилка", str(e))

    def _export_excel(self) -> None:
        if not pd:
            messagebox.showwarning("Excel", "Встановiть pandas та openpyxl"); return
        if not self.all_records:
            messagebox.showinfo("Експорт", "Немає записiв"); return
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx", filetypes=[("Excel файли", "*.xlsx")],
            title="Зберегти як Excel")
        if not path:
            return
        try:
            df = pd.DataFrame([rec.to_list() for rec in self.all_records],
                               columns=self._EVENT_HEADERS)
            with pd.ExcelWriter(path, engine="openpyxl") as writer:
                df.to_excel(writer, index=False, sheet_name="Реєстр")
                ws = writer.sheets["Реєстр"]
                for col_cells in ws.columns:
                    mx = max(len(str(c.value or "")) for c in col_cells)
                    ws.column_dimensions[col_cells[0].column_letter].width = min(mx + 4, 60)
            self._show_toast("Excel збережено")
        except (OSError, Exception) as e:
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
            for item in data:
                if isinstance(item, dict):
                    rec = EventRecord.from_dict(item)
                elif isinstance(item, (list, tuple)):
                    rec = EventRecord.from_list(list(item))
                else:
                    continue
                rec = dataclasses.replace(rec, id=IdGenerator.next_id(self.all_records))
                self.all_records.append(rec)
                self._insert_tree_row(rec)
                added += 1
            self._save_data()
            self._update_count()
            self._notify_change()
            self._show_toast(f"Iмпортовано: {added} записiв")
        except (json.JSONDecodeError, ValueError, OSError) as e:
            messagebox.showerror("Помилка iмпорту", str(e))


# =============================================================================
#  ВКЛАДКА: АНАЛІТИКА ПОДІЙ
# =============================================================================

class AnalyticsTab:

    def __init__(self, parent: tk.Misc) -> None:
        self.frame   = ttk.Frame(parent)
        self.records: list[EventRecord] = []
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
                 font=FONT_TITLE).pack(side="left", padx=20, pady=14)
        make_button(header, "Оновити",
                    bg=C["accent"], activebackground=C["accent_soft"],
                    font=FONT_BOLD, padx=12, pady=4,
                    command=self.refresh).pack(side="right", padx=20, pady=12)

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
        _scroll_mgr.attach(canvas)  # ВИПРАВЛЕННЯ #1
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
                     fg=C["text_muted"], font=FONT_SMALL).pack(anchor="w", pady=(8, 2))
            lbl = tk.Label(card, text=val, bg=C["bg_surface"],
                           fg=color, font=FONT_NUMBER)
            lbl.pack(anchor="w")
            self.stat_cards[key] = lbl

    def _build_charts_and_table(self) -> None:
        C = COLORS
        if HAS_MPL:
            charts_row = tk.Frame(self.content, bg=C["bg_main"])
            charts_row.grid(row=1, column=0, sticky="ew", padx=16, pady=16)
            charts_row.columnconfigure((0, 1), weight=1)

            self.fig_left = Figure(figsize=(5, 3.5), dpi=90, facecolor=C["bg_surface"])
            self.ax_left  = self.fig_left.add_subplot(111)
            self._style_ax(self.ax_left)
            self.ax_left.set_title("Розподiл за типом ризику",
                                    color=C["text_muted"], fontsize=9)
            fl = tk.Frame(charts_row, bg=C["bg_surface"], padx=8, pady=8)
            fl.grid(row=0, column=0, padx=(0, 8), sticky="nsew")
            self.canvas_left = FigureCanvasTkAgg(self.fig_left, master=fl)
            self.canvas_left.get_tk_widget().pack(fill="both", expand=True)

            self.fig_right = Figure(figsize=(5, 3.5), dpi=90, facecolor=C["bg_surface"])
            self.ax_right  = self.fig_right.add_subplot(111)
            self._style_ax(self.ax_right)
            self.ax_right.set_title("Записи за статусом",
                                     color=C["text_muted"], fontsize=9)
            fr = tk.Frame(charts_row, bg=C["bg_surface"], padx=8, pady=8)
            fr.grid(row=0, column=1, padx=(8, 0), sticky="nsew")
            self.canvas_right = FigureCanvasTkAgg(self.fig_right, master=fr)
            self.canvas_right.get_tk_widget().pack(fill="both", expand=True)

            self.fig_bottom = Figure(figsize=(10, 3), dpi=90, facecolor=C["bg_surface"])
            self.ax_bottom  = self.fig_bottom.add_subplot(111)
            self._style_ax(self.ax_bottom)
            self.ax_bottom.set_title("Топ-5 пiдприємств за кiлькiстю подiй",
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
                 font=FONT_BOLD).grid(row=0, column=0, sticky="w", pady=(0, 8))

        cols = ("risk", "count", "open", "closed")
        self.stats_tree = ttk.Treeview(frame, columns=cols, show="headings", height=7)
        for col, hdr, w in [("risk", "Тип ризику", 200), ("count", "Всього", 80),
                              ("open", "Вiдкрито", 80), ("closed", "Закрито", 80)]:
            self.stats_tree.heading(col, text=hdr)
            self.stats_tree.column(col, width=w, anchor="w")
        self.stats_tree.grid(row=1, column=0, sticky="ew")

    def _style_ax(self, ax: object) -> None:
        C = COLORS
        ax.set_facecolor(C["bg_surface"])
        ax.tick_params(colors=C["text_muted"], labelsize=8)
        for spine in ax.spines.values():
            spine.set_edgecolor(C["border_soft"])

    def update_data(self, records: list[EventRecord]) -> None:
        self.records = records
        self.refresh()

    def refresh(self) -> None:
        if not self.records:
            for k in self.stat_cards:
                self.stat_cards[k].configure(text="0")
            if HAS_MPL:
                self._clear_charts()
            self.stats_tree.delete(*self.stats_tree.get_children())
            return

        C       = COLORS
        records = self.records
        total   = len(records)
        open_c  = sum(1 for r in records if r.status in ("Вiдкрито", "В обробцi"))
        crit    = sum(1 for r in records if r.priority == "Критичний")
        closed  = sum(1 for r in records if r.status in ("Закрито", "Вирiшено"))

        self.stat_cards["total"].configure(text=str(total))
        self.stat_cards["open"].configure(text=str(open_c))
        self.stat_cards["critical"].configure(text=str(crit))
        self.stat_cards["closed"].configure(text=str(closed))

        if not HAS_MPL:
            return

        risk_ctr   = Counter(r.risk_type  for r in records)
        status_ctr = Counter(r.status     for r in records)
        entity_ctr = Counter(r.entity     for r in records if r.entity)

        self.ax_left.clear(); self._style_ax(self.ax_left)
        self.ax_left.set_title("Розподiл за типом ризику",
                                color=C["text_muted"], fontsize=9)
        if risk_ctr:
            lbls = list(risk_ctr.keys()); vals = list(risk_ctr.values())
            clrs = [RISK_COLORS.get(l, C["text_muted"]) for l in lbls]
            _, _, autotexts = self.ax_left.pie(
                vals, labels=lbls, autopct="%1.0f%%", colors=clrs, startangle=90,
                textprops={"color": C["text_muted"], "fontsize": 7})
            for at in autotexts:
                at.set_fontsize(7); at.set_color("white")
        else:
            self.ax_left.text(0.5, 0.5, "Немає даних",
                              transform=self.ax_left.transAxes,
                              ha="center", va="center", color=C["text_muted"])
        self.canvas_left.draw()

        self.ax_right.clear(); self._style_ax(self.ax_right)
        self.ax_right.set_title("Записи за статусом", color=C["text_muted"], fontsize=9)
        if status_ctr:
            s_lbls = list(status_ctr.keys()); s_vals = list(status_ctr.values())
            s_clrs = [C["accent_danger"], C["accent_warning"],
                      C["accent_success"], C["accent_muted"]][:len(s_lbls)]
            bars = self.ax_right.bar(s_lbls, s_vals, color=s_clrs, edgecolor="none")
            for bar, val in zip(bars, s_vals, strict=False):
                self.ax_right.text(bar.get_x() + bar.get_width() / 2,
                                    bar.get_height() + 0.1, str(val),
                                    ha="center", va="bottom",
                                    color=C["text_muted"], fontsize=8)
            self.ax_right.tick_params(axis="x", labelrotation=15, labelsize=7)
            self.ax_right.set_ylim(0, max(s_vals) * 1.2 + 1)
        else:
            self.ax_right.text(0.5, 0.5, "Немає даних",
                               transform=self.ax_right.transAxes,
                               ha="center", va="center", color=C["text_muted"])
        self.canvas_right.draw()

        self.ax_bottom.clear(); self._style_ax(self.ax_bottom)
        self.ax_bottom.set_title("Топ-5 пiдприємств за кiлькiстю подiй",
                                  color=C["text_muted"], fontsize=9)
        top5 = entity_ctr.most_common(5)
        if top5:
            e_lbls = [e[0][:20] for e in top5]
            e_vals = [e[1] for e in top5]
            bars = self.ax_bottom.barh(e_lbls, e_vals, color=C["accent"], edgecolor="none")
            for bar, val in zip(bars, e_vals, strict=False):
                self.ax_bottom.text(bar.get_width() + 0.05,
                                     bar.get_y() + bar.get_height() / 2,
                                     str(val), ha="left", va="center",
                                     color=C["text_muted"], fontsize=8)
            self.ax_bottom.tick_params(axis="y", labelsize=8, colors=C["text_primary"])
        else:
            self.ax_bottom.text(0.5, 0.5, "Немає даних",
                                transform=self.ax_bottom.transAxes,
                                ha="center", va="center", color=C["text_muted"])
        self.canvas_bottom.draw()

        self.stats_tree.delete(*self.stats_tree.get_children())
        all_risks = set(RISK_TYPES) | {r.risk_type for r in records}
        for risk in sorted(all_risks):
            recs  = [r for r in records if r.risk_type == risk]
            cnt   = len(recs)
            open_ = sum(1 for r in recs if r.status in ("Вiдкрито", "В обробцi"))
            cl    = sum(1 for r in recs if r.status in ("Закрито", "Вирiшено"))
            if cnt:
                self.stats_tree.insert("", tk.END, values=(risk, cnt, open_, cl))

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
        tk.Label(header, text="НАЛАШТУВАННЯ РЕЄСТРУ",
                 bg=C["bg_header"], fg=C["text_muted"],
                 font=FONT_TITLE).pack(side="left", padx=20, pady=14)

        content = tk.Frame(self.frame, bg=C["bg_main"])
        content.grid(row=1, column=0, sticky="nsew", padx=40, pady=30)
        content.columnconfigure(0, weight=1)

        self._row(content, 0, "Файл даних:", DATA_FILE, C)
        self._row(content, 1, "Версiя:", "2.2 — Рефакторинг (dataclass + ScrollManager)", C)
        self._row(content, 2, "matplotlib:",
                  "встановлено" if HAS_MPL else "не встановлено", C)
        self._row(content, 3, "pandas:",
                  "встановлено" if pd else "не встановлено", C)

        tk.Label(content, text="Встановлення залежностей:",
                 bg=C["bg_main"], fg=C["text_muted"],
                 font=FONT_SMALL_BOLD).grid(row=4, column=0, sticky="w", pady=(24, 6))
        tk.Label(content, text="  pip install matplotlib pandas openpyxl",
                 bg=C["bg_surface"], fg=C["accent_muted"],
                 font=FONT_MONO, padx=12, pady=8).grid(row=5, column=0, sticky="w")

        tk.Label(content, text="Виправлення у версiї 2.2:",
                 bg=C["bg_main"], fg=C["text_muted"],
                 font=FONT_SMALL_BOLD).grid(row=6, column=0, sticky="w", pady=(24, 6))

        fixes = [
            "#1 ScrollManager: скрол бiльше не накопичується мiж вкладками",
            "#2 Надiйний пошук запису за числовим ID (без lstrip)",
            "#3 ID генерується як max(existing)+1, колiзiї пiсля видалення виключенi",
            "#4 Публiчний метод save() замiсть _save_data() через noqa",
            "#5 BaseRegistryTab: дублювання коду зведено до мiнiмуму",
            "#6 Конкретнi типи винятків (OSError, JSONDecodeError) замiсть Exception",
            "#7 Валiдацiя типiв при завантаженнi даних з JSON",
            "#8 JSON зберiгається як список словникiв (стiйкий формат)",
            "#9 EventRecord / RiskRecord dataclass замiсть магiчних iндексiв",
            "#10 Валiдацiя обов'язкових полiв до побудови запису",
        ]
        for i, fix in enumerate(fixes):
            f = tk.Frame(content, bg=C["bg_main"])
            f.grid(row=7 + i, column=0, sticky="w", pady=2)
            tk.Frame(f, bg=C["accent_success"], width=4, height=4).pack(
                side="left", padx=(0, 8))
            tk.Label(f, text=fix, bg=C["bg_main"], fg=C["text_subtle"],
                     font=FONT_SMALL).pack(side="left")

    def _row(self, parent: tk.Misc, row: int, label: str, value: str, C: dict) -> None:
        f = tk.Frame(parent, bg=C["bg_main"])
        f.grid(row=row, column=0, sticky="ew", pady=4)
        tk.Label(f, text=label, bg=C["bg_main"], fg=C["text_muted"],
                 font=FONT_DEFAULT, width=22, anchor="w").pack(side="left")
        tk.Label(f, text=value, bg=C["bg_main"], fg=C["text_primary"],
                 font=FONT_DEFAULT).pack(side="left")

    def get_frame(self) -> ttk.Frame:
        return self.frame


# =============================================================================
#  СТОРІНКА "РЕЄСТР СУТТЄВИХ ПОДІЙ"
# =============================================================================

class MaterialEventsPage(tk.Frame):

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

        self.notebook.add(self.registry_tab.get_frame(),  text="  Реєстр подiй  ")
        self.notebook.add(self.analytics_tab.get_frame(), text="  Аналiтика  ")
        self.notebook.add(self.settings_tab.get_frame(),  text="  Налаштування  ")

        statusbar = tk.Frame(self, bg=COLORS["bg_header"], height=22)
        statusbar.grid(row=1, column=0, sticky="ew")
        statusbar.grid_propagate(False)

        self._status_lbl = tk.Label(
            statusbar, text="Готово",
            bg=COLORS["bg_header"], fg=COLORS["text_muted"],
            font=FONT_TINY, padx=10)
        self._status_lbl.pack(side="left", pady=3)

        self._time_lbl = tk.Label(
            statusbar, text="",
            bg=COLORS["bg_header"], fg=COLORS["text_muted"],
            font=FONT_TINY, padx=10)
        self._time_lbl.pack(side="right", pady=3)

        self._start_clock()
        self._schedule_autosave()
        self.after(600, lambda: self.analytics_tab.update_data(
            self.registry_tab.all_records))

    def _start_clock(self) -> None:
        self._time_lbl.configure(text=datetime.now().strftime("%d.%m.%Y  %H:%M:%S"))
        self.after(1000, self._start_clock)

    def _schedule_autosave(self) -> None:
        try:
            self.registry_tab.save()  # ВИПРАВЛЕННЯ #4: публічний метод
            self._status_lbl.configure(
                text=f"Автозбережено о {datetime.now().strftime('%H:%M:%S')}")
        except OSError:
            self._status_lbl.configure(text="Помилка автозбереження")
        self.after(30000, self._schedule_autosave)

    def save_before_exit(self) -> None:
        try:
            self.registry_tab.save()  # ВИПРАВЛЕННЯ #4
        except OSError:
            pass


# =============================================================================
#  ДЕТАЛЬНЕ ВІКНО РИЗИКУ
# =============================================================================

class RiskDetailWindow:

    def __init__(
        self,
        parent_root:     tk.Misc,
        record:          RiskRecord,
        all_records:     list[RiskRecord],
        save_callback:   Callable[[str, RiskRecord], None],
        delete_callback: Callable[[str], None],
        toast_callback:  Callable[[str], None],
    ) -> None:
        self.parent_root     = parent_root
        self.record          = record
        self.all_records     = all_records
        self.save_callback   = save_callback
        self.delete_callback = delete_callback
        self.toast_callback  = toast_callback
        self.is_edit_mode    = False
        self._build_window()

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

    def _build_window(self) -> None:
        C = COLORS
        self.win = tk.Toplevel(self.parent_root)
        self.win.title(f"Ризик #{self.record.id}  —  {self.record.entity}")
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

        header = tk.Frame(self.win, bg=C["bg_header"], height=58)
        header.grid(row=0, column=0, sticky="ew")
        header.columnconfigure(1, weight=1)
        header.grid_propagate(False)

        try:
            score_val = int(self.record.score)
        except (ValueError, TypeError):
            score_val = 0
        strip_color = _score_color(score_val)
        tk.Frame(header, bg=strip_color, width=4).grid(row=0, column=0, sticky="ns")

        title_frame = tk.Frame(header, bg=C["bg_header"])
        title_frame.grid(row=0, column=1, sticky="ew", padx=16, pady=10)

        self.lbl_title = tk.Label(
            title_frame, text=f"Ризик #{self.record.id}",
            bg=C["bg_header"], fg=C["accent_muted"], font=FONT_HEADING)
        self.lbl_title.pack(anchor="w")
        self.lbl_subtitle = tk.Label(
            title_frame, text=self.record.entity,
            bg=C["bg_header"], fg=C["text_muted"], font=FONT_DEFAULT)
        self.lbl_subtitle.pack(anchor="w")

        sc = self._status_color(self.record.status)
        self.lbl_status_badge = tk.Label(
            header, text=f"  {self.record.status}  ",
            bg=sc, fg="white", font=FONT_SMALL_BOLD, pady=3)
        self.lbl_status_badge.grid(row=0, column=2, padx=16, pady=18)

        self.lbl_score_badge = tk.Label(
            header,
            text=f"  {_score_label(score_val)} ({score_val})  ",
            bg=strip_color, fg="white", font=FONT_SMALL_BOLD, pady=3)
        self.lbl_score_badge.grid(row=0, column=3, padx=(0, 16), pady=18)

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
        _scroll_mgr.attach(canvas)  # ВИПРАВЛЕННЯ #1

        self.content.columnconfigure(0, weight=1)
        self.content.columnconfigure(1, weight=1)
        self._build_view_content()

        btn_bar = tk.Frame(self.win, bg=C["bg_header"], height=46)
        btn_bar.grid(row=2, column=0, columnspan=2, sticky="ew")
        btn_bar.grid_propagate(False)

        left_btns = tk.Frame(btn_bar, bg=C["bg_header"])
        left_btns.pack(side="left", padx=12, pady=8)

        self.btn_edit = make_button(
            left_btns, "Редагувати",
            bg=C["accent_warning"], fg=C["bg_main"],
            activebackground="#d97706", activeforeground="white",
            font=FONT_BOLD, padx=14, pady=4, command=self._toggle_edit_mode)
        self.btn_edit.pack(side="left", padx=(0, 8))

        self.btn_save = make_button(
            left_btns, "Зберегти змiни",
            bg=C["accent_success"], activebackground="#16a34a",
            font=FONT_BOLD, padx=14, pady=4, command=self._save_changes)
        self.btn_save.pack_forget()

        self.btn_cancel_edit = make_button(
            left_btns, "Скасувати",
            bg=C["bg_surface"], fg=C["text_muted"],
            activebackground=C["bg_surface_alt"], activeforeground=C["text_primary"],
            font=FONT_DEFAULT, padx=12, pady=4, command=self._cancel_edit)
        self.btn_cancel_edit.pack_forget()

        right_btns = tk.Frame(btn_bar, bg=C["bg_header"])
        right_btns.pack(side="right", padx=12, pady=8)

        make_button(right_btns, "Видалити",
                    bg=C["accent_danger"], activebackground="#dc2626",
                    font=FONT_BOLD, padx=14, pady=4,
                    command=self._delete_record).pack(side="right", padx=(8, 0))
        make_button(right_btns, "Закрити",
                    bg=C["bg_surface"], fg=C["text_primary"],
                    activebackground=C["bg_surface_alt"], activeforeground=C["text_primary"],
                    font=FONT_DEFAULT, padx=12, pady=4,
                    command=self.win.destroy).pack(side="right")

    def _build_view_content(self) -> None:
        for w in self.content.winfo_children():
            w.destroy()
        r   = self.record
        row = 0

        _build_section_label(self.content, "Iнформацiя про пiдприємство", row); row += 1
        _build_info_cell(self.content, "Пiдприємство", r.entity, row, 0)
        _build_info_cell(self.content, "Прiоритет", r.priority, row, 1,
                         value_color=self._priority_color(r.priority)); row += 1

        _build_section_label(self.content, "Опис ризику", row); row += 1
        _build_info_cell(self.content, "Назва ризику",   r.risk_name, row, 0)
        _build_info_cell(self.content, "Тип ризику",     r.risk_type, row, 1,
                         value_color=RISK_COLORS.get(r.risk_type)); row += 1
        _build_info_cell(self.content, "Категорiя ризику", r.category, row, 0)
        _build_info_cell(self.content, "Власник ризику",    r.owner,    row, 1); row += 1

        _build_section_label(self.content, "Оцiнка ризику", row); row += 1
        _build_info_cell(self.content, "Iмовiрнiсть", r.probability, row, 0)
        _build_info_cell(self.content, "Вплив",         r.impact,     row, 1); row += 1

        try:
            score_int = int(r.score)
        except (ValueError, TypeError):
            score_int = 0
        C = COLORS
        sc_cell = tk.Frame(self.content, bg=C["bg_surface"], padx=10, pady=8)
        sc_cell.grid(row=row, column=0, columnspan=2, sticky="nsew", padx=8, pady=3)
        sc_cell.columnconfigure(0, weight=1)
        tk.Label(sc_cell, text="Рiвень ризику (Score = Iмовiрнiсть × Вплив)",
                 bg=C["bg_surface"], fg=C["text_subtle"],
                 font=("Arial", 7, "bold")).grid(row=0, column=0, sticky="w")
        tk.Label(sc_cell,
                 text=f"{score_int}  —  {_score_label(score_int)}",
                 bg=C["bg_surface"], fg=_score_color(score_int),
                 font=FONT_SCORE).grid(row=1, column=0, sticky="w", pady=(2, 0))
        row += 1

        res_color = _score_color(int(r.residual)) if r.residual.isdigit() else None
        _build_info_cell(self.content, "Залишковий ризик", r.residual, row, 0,
                         value_color=res_color)
        _build_info_cell(self.content, "Статус", r.status, row, 1,
                         value_color=self._status_color(r.status)); row += 1

        _build_section_label(self.content, "Дати", row); row += 1
        _build_info_cell(self.content, "Дата виявлення", r.date_id,  row, 0)
        _build_info_cell(self.content, "Дата перегляду", r.date_rev, row, 1); row += 1

        _build_section_label(self.content, "Деталi", row); row += 1
        _build_text_block(self.content, "Заходи контролю",       r.controls,    row); row += 1
        _build_text_block(self.content, "Детальний опис ризику",  r.description, row); row += 1
        tk.Frame(self.content, bg=C["bg_main"], height=12).grid(
            row=row, column=0, columnspan=2)

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
                     font=FONT_BOLD).pack(side="left", padx=8)
            return r + 1

        def lbl(text: str, r: int) -> None:
            tk.Label(self.content, text=text, bg=C["bg_main"],
                     fg=C["text_subtle"], font=FONT_SMALL).grid(
                row=r, column=0, sticky="w", padx=10, pady=(6, 0))

        row = 0
        row = section("Пiдприємство та ризик", row)

        lbl("Скорочена назва пiдприємства:", row); row += 1
        self.e_entity = make_dark_entry(self.content, accent=C["accent_warning"])
        self.e_entity.grid(row=row, column=0, sticky="ew", padx=10, pady=(2, 0))
        self.e_entity.insert(0, rec.entity); row += 1

        lbl("Назва ризику:", row); row += 1
        self.e_risk_name = make_dark_entry(self.content, accent=C["accent_warning"])
        self.e_risk_name.grid(row=row, column=0, sticky="ew", padx=10, pady=(2, 0))
        self.e_risk_name.insert(0, rec.risk_name); row += 1

        lbl("Категорiя ризику:", row); row += 1
        self.e_category = make_dark_combo(self.content, values=RISK_CATEGORIES)
        self.e_category.grid(row=row, column=0, sticky="ew", padx=10, pady=(2, 0))
        self.e_category.set(rec.category); row += 1

        lbl("Тип ризику:", row); row += 1
        self.e_risk_type = make_dark_combo(self.content, values=RISK_TYPES)
        self.e_risk_type.grid(row=row, column=0, sticky="ew", padx=10, pady=(2, 0))
        self.e_risk_type.set(rec.risk_type if rec.risk_type != "—" else ""); row += 1

        lbl("Власник ризику:", row); row += 1
        self.e_owner = make_dark_entry(self.content, accent=C["accent_warning"])
        self.e_owner.grid(row=row, column=0, sticky="ew", padx=10, pady=(2, 0))
        self.e_owner.insert(0, rec.owner); row += 1

        row = section("Оцiнка ризику", row)
        score_f = tk.Frame(self.content, bg=C["bg_main"])
        score_f.grid(row=row, column=0, sticky="ew", padx=10, pady=4)
        score_f.columnconfigure((0, 1, 2), weight=1); row += 1

        for ci, (lbl_t, attr, vals, cur_val) in enumerate([
            ("Iмовiрнiсть:", "e_prob",   PROBABILITY_LEVELS, rec.probability),
            ("Вплив:",        "e_impact", IMPACT_LEVELS,      rec.impact),
        ]):
            tk.Label(score_f, text=lbl_t, bg=C["bg_main"],
                     fg=C["text_subtle"], font=FONT_SMALL).grid(
                row=0, column=ci, sticky="w", padx=(0 if ci == 0 else 16, 0))
            combo = make_dark_combo(score_f, values=vals, width=22)
            combo.grid(row=1, column=ci, sticky="ew",
                       padx=(0 if ci == 0 else 16, 0), pady=2)
            matched = next((v for v in vals if v.startswith(str(cur_val)[:1])), "")
            combo.set(cur_val if cur_val in vals else matched)
            setattr(self, attr, combo)

        self.lbl_live_score = tk.Label(score_f, text="Score: —",
                                       bg=C["bg_main"], fg=C["text_muted"],
                                       font=("Arial", 11, "bold"))
        self.lbl_live_score.grid(row=1, column=2, padx=20)

        def _upd_score(_: object = None) -> None:
            try:
                s   = _extract_num(self.e_prob.get()) * _extract_num(self.e_impact.get())
                col = _score_color(s)
                self.lbl_live_score.configure(
                    text=f"Score: {s}  ({_score_label(s)})", fg=col)
            except Exception:
                self.lbl_live_score.configure(text="Score: —", fg=C["text_muted"])

        self.e_prob.bind("<<ComboboxSelected>>",   _upd_score)
        self.e_impact.bind("<<ComboboxSelected>>", _upd_score)
        _upd_score()

        lbl("Залишковий ризик (1–25):", row); row += 1
        self.e_residual = make_dark_entry(self.content, accent=C["accent_warning"])
        self.e_residual.grid(row=row, column=0, sticky="w", padx=10, pady=(2, 0))
        self.e_residual.insert(0, rec.residual); row += 1

        row = section("Дати", row)
        date_f = tk.Frame(self.content, bg=C["bg_main"])
        date_f.grid(row=row, column=0, sticky="w", padx=10, pady=4); row += 1

        for ci, (lbl_t, attr, val) in enumerate([
            ("Дата виявлення:", "e_date_id",  rec.date_id),
            ("Дата перегляду:", "e_date_rev", rec.date_rev),
        ]):
            tk.Label(date_f, text=lbl_t, bg=C["bg_main"],
                     fg=C["text_subtle"], font=FONT_SMALL).grid(
                row=0, column=ci, padx=(0 if ci == 0 else 20, 0), sticky="w")
            e = make_dark_entry(date_f, accent=C["accent_warning"], width=14)
            e.grid(row=1, column=ci, padx=(0 if ci == 0 else 20, 0), pady=2)
            if val and val != "—":
                e.insert(0, val)
            else:
                add_placeholder(e, "дд.мм.рррр")
            setattr(self, attr, e)

        row = section("Прiоритет та статус", row)
        ps_f = tk.Frame(self.content, bg=C["bg_main"])
        ps_f.grid(row=row, column=0, sticky="w", padx=10, pady=4); row += 1

        tk.Label(ps_f, text="Прiоритет:", bg=C["bg_main"],
                 fg=C["text_subtle"], font=FONT_SMALL).grid(row=0, column=0, sticky="w")
        self.e_priority = make_dark_combo(
            ps_f, values=["Критичний", "Високий", "Середнiй", "Низький"], width=14)
        self.e_priority.grid(row=1, column=0, pady=(2, 0), padx=(0, 20))
        self.e_priority.set(rec.priority)

        tk.Label(ps_f, text="Статус:", bg=C["bg_main"],
                 fg=C["text_subtle"], font=FONT_SMALL).grid(row=0, column=1, sticky="w")
        self.e_status = make_dark_combo(
            ps_f, values=["Активний", "Монiторинг", "Мiтигований", "Закрито"], width=14)
        self.e_status.grid(row=1, column=1, pady=(2, 0))
        self.e_status.set(rec.status)

        row = section("Деталi", row)

        lbl("Заходи контролю:", row); row += 1
        self.e_controls = make_dark_text(self.content, height=3, wrap="word")
        self.e_controls.grid(row=row, column=0, sticky="ew", padx=10, pady=(2, 0))
        if rec.controls and rec.controls != "—":
            self.e_controls.insert("1.0", rec.controls)
        row += 1

        lbl("Детальний опис ризику:", row); row += 1
        self.e_description = make_dark_text(self.content, height=4, wrap="word")
        self.e_description.grid(row=row, column=0, sticky="ew", padx=10, pady=(2, 0))
        if rec.description and rec.description != "—":
            self.e_description.insert("1.0", rec.description)
        row += 1
        tk.Frame(self.content, bg=C["bg_main"], height=16).grid(row=row, column=0)

    def _toggle_edit_mode(self) -> None:
        self.is_edit_mode = True
        self._build_edit_content()
        self.btn_edit.pack_forget()
        self.btn_save.pack(side="left", padx=(0, 8))
        self.btn_cancel_edit.pack(side="left")
        self.lbl_title.configure(
            text=f"Редагування ризику #{self.record.id}",
            fg=COLORS["accent_warning"])

    def _cancel_edit(self) -> None:
        self.is_edit_mode = False
        self._build_view_content()
        self.btn_save.pack_forget()
        self.btn_cancel_edit.pack_forget()
        self.btn_edit.pack(side="left", padx=(0, 8))
        self.lbl_title.configure(
            text=f"Ризик #{self.record.id}", fg=COLORS["accent_muted"])

    def _save_changes(self) -> None:
        date_id  = self.e_date_id.get().strip()
        date_rev = self.e_date_rev.get().strip()
        for val, label in [(date_id, "дати виявлення"), (date_rev, "дати перегляду")]:
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

        old_id = self.record.id
        new_record = RiskRecord(
            id=self.record.id,
            entity=entity,
            risk_name=risk_name,
            category=self.e_category.get().strip()  or "—",
            risk_type=self.e_risk_type.get().strip() or "—",
            probability=prob_str   or "—",
            impact=impact_str      or "—",
            score=str(score),
            owner=self.e_owner.get().strip(),
            controls=self.e_controls.get("1.0", tk.END).strip(),
            residual=str(residual),
            date_id=date_id  or "—",
            date_rev=date_rev or "—",
            priority=self.e_priority.get().strip() or "Середнiй",
            status=self.e_status.get().strip()   or "Активний",
            description=self.e_description.get("1.0", tk.END).strip(),
        )
        self.record = new_record
        self.save_callback(old_id, new_record)

        self.lbl_subtitle.configure(text=entity)
        self.lbl_status_badge.configure(
            text=f"  {new_record.status}  ",
            bg=self._status_color(new_record.status))
        self.lbl_score_badge.configure(
            text=f"  {_score_label(score)} ({score})  ",
            bg=_score_color(score))
        self._cancel_edit()
        self.toast_callback("Запис оновлено")

    def _delete_record(self) -> None:
        if not messagebox.askyesno(
            "Пiдтвердження",
            f"Видалити ризик #{self.record.id}?\nЦю дiю не можна скасувати.",
            parent=self.win):
            return
        self.delete_callback(self.record.id)
        self.toast_callback("Запис видалено")
        self.win.destroy()


# =============================================================================
#  ВКЛАДКА: РЕЄСТР РИЗИКІВ
# =============================================================================

class RiskRegistryTab(BaseRegistryTab):

    data_file = RISK_DATA_FILE

    def __init__(
        self,
        parent: tk.Misc,
        on_data_change: Callable[[list[RiskRecord]], None] | None = None,
    ) -> None:
        super().__init__(parent, on_data_change)
        self._build_ui()
        self._load_data()

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
                 font=FONT_TITLE).grid(row=0, column=0, padx=20, pady=14, sticky="w")

        sf = tk.Frame(header, bg=C["bg_header"])
        sf.grid(row=0, column=1, sticky="e", padx=20)
        tk.Label(sf, text="Пошук:", bg=C["bg_header"],
                 fg=C["text_muted"], font=FONT_SMALL).pack(side="left", padx=(0, 6))

        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *_: self._apply_filter())
        tk.Entry(sf, textvariable=self.search_var,
                 bg=C["bg_input"], fg=C["text_primary"],
                 insertbackground=C["text_primary"],
                 relief="flat", bd=2, font=FONT_DEFAULT,
                 width=34).pack(side="left", padx=(0, 8), ipady=2)
        make_button(sf, "Скинути",
                    bg=C["bg_surface"], fg=C["text_muted"],
                    activebackground=C["bg_surface_alt"],
                    activeforeground=C["text_primary"],
                    font=FONT_SMALL, padx=8, pady=2,
                    command=self._reset_filter).pack(side="left")

        paned = ttk.PanedWindow(self.frame, orient="horizontal")
        paned.grid(row=1, column=0, sticky="nsew")
        lw = ttk.Frame(paned)
        rw = ttk.Frame(paned)
        paned.add(lw, weight=4)
        paned.add(rw, weight=7)

        self._build_form(lw)
        self._build_table(rw)

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
        _scroll_mgr.attach(canvas)  # ВИПРАВЛЕННЯ #1
        form.columnconfigure(0, weight=1)

        def section(txt: str, r: int) -> int:
            f = tk.Frame(form, bg=C["bg_main"])
            f.grid(row=r, column=0, sticky="ew", padx=16, pady=(16, 4))
            tk.Frame(f, bg=C["accent"], width=3, height=16).pack(side="left")
            tk.Label(f, text=txt, bg=C["bg_main"], fg=C["accent"],
                     font=FONT_BOLD).pack(side="left", padx=8)
            return r + 1

        def lbl(text: str, r: int) -> None:
            tk.Label(form, text=text, bg=C["bg_main"],
                     fg=C["text_subtle"], font=FONT_SMALL).grid(
                row=r, column=0, sticky="w", padx=16, pady=(4, 0))

        def field(lbl_txt: str, r: int, factory: Callable) -> tuple[tk.Widget, int]:
            lbl(lbl_txt, r)
            w = factory()
            w.grid(row=r + 1, column=0, sticky="ew", padx=16, pady=(2, 0))
            return w, r + 2

        row = 0

        badge_f = tk.Frame(form, bg=C["bg_main"])
        badge_f.grid(row=row, column=0, sticky="ew", padx=16, pady=(12, 0))
        tk.Label(badge_f, text="  + НОВИЙ РИЗИК  ",
                 bg=C["accent"], fg="white",
                 font=FONT_SMALL_BOLD, pady=3).pack(anchor="w")
        row += 1

        row = section("Iнформацiя про пiдприємство", row)
        self.ent_entity,    row = field("Скорочена назва пiдприємства:", row,
                                        lambda: make_dark_entry(form))
        self.ent_owner,     row = field("Власник ризику:", row,
                                        lambda: make_dark_entry(form))

        row = section("Опис ризику", row)
        self.ent_risk_name, row = field("Назва ризику:", row,
                                        lambda: make_dark_entry(form))
        self.cb_category,   row = field("Категорiя ризику:", row,
                                        lambda: make_dark_combo(form, values=RISK_CATEGORIES))
        self.cb_risk_type,  row = field("Тип ризику:", row,
                                        lambda: make_dark_combo(form, values=RISK_TYPES))

        row = section("Оцiнка ризику", row)
        score_f = tk.Frame(form, bg=C["bg_main"])
        score_f.grid(row=row, column=0, sticky="ew", padx=16, pady=4)
        score_f.columnconfigure((0, 1), weight=1); row += 1

        for ci, (lbl_t, attr, vals) in enumerate([
            ("Iмовiрнiсть (1–5):", "cb_prob",   PROBABILITY_LEVELS),
            ("Вплив (1–5):",        "cb_impact", IMPACT_LEVELS),
        ]):
            tk.Label(score_f, text=lbl_t, bg=C["bg_main"],
                     fg=C["text_subtle"], font=FONT_SMALL).grid(
                row=0, column=ci, sticky="w", padx=(0 if ci == 0 else 10, 0))
            combo = make_dark_combo(score_f, values=vals)
            combo.grid(row=1, column=ci, sticky="ew",
                       padx=(0 if ci == 0 else 10, 0), pady=2)
            setattr(self, attr, combo)

        net_f = tk.Frame(form, bg=C["bg_main"])
        net_f.grid(row=row, column=0, sticky="ew", padx=16, pady=(2, 4)); row += 1
        tk.Label(net_f, text="Рiвень ризику (Score):",
                 bg=C["bg_main"], fg=C["text_muted"], font=FONT_SMALL).pack(side="left")
        self.lbl_score = tk.Label(net_f, text="—",
                                   bg=C["bg_main"], fg=C["accent_success"],
                                   font=("Arial", 13, "bold"))
        self.lbl_score.pack(side="left", padx=10)

        def _upd_score(_: object = None) -> None:
            try:
                s = _extract_num(self.cb_prob.get()) * _extract_num(self.cb_impact.get())
                self.lbl_score.configure(
                    text=f"{s}  ({_score_label(s)})", fg=_score_color(s))
            except Exception:
                self.lbl_score.configure(text="—", fg=C["text_muted"])

        self.cb_prob.bind("<<ComboboxSelected>>",   _upd_score)
        self.cb_impact.bind("<<ComboboxSelected>>", _upd_score)

        lbl("Залишковий ризик (1–25):", row); row += 1
        self.ent_residual = make_dark_entry(form)
        self.ent_residual.grid(row=row, column=0, sticky="ew", padx=16, pady=(2, 0)); row += 1

        row = section("Дати", row)
        date_f = tk.Frame(form, bg=C["bg_main"])
        date_f.grid(row=row, column=0, sticky="w", padx=16, pady=4); row += 1
        for ci, (lbl_t, attr) in enumerate([
            ("Дата виявлення:", "ent_date_id"),
            ("Дата перегляду:", "ent_date_rev"),
        ]):
            tk.Label(date_f, text=lbl_t, bg=C["bg_main"],
                     fg=C["text_subtle"], font=FONT_SMALL).grid(
                row=0, column=ci, padx=(0 if ci == 0 else 20, 0), sticky="w")
            e = make_dark_entry(date_f, width=14)
            e.grid(row=1, column=ci, padx=(0 if ci == 0 else 20, 0), pady=2)
            add_placeholder(e, "дд.мм.рррр")
            setattr(self, attr, e)

        lbl("Прiоритет:", row); row += 1
        self.cb_priority = make_dark_combo(
            form, values=["Критичний", "Високий", "Середнiй", "Низький"])
        self.cb_priority.grid(row=row, column=0, sticky="w", padx=16, pady=(2, 0)); row += 1

        lbl("Статус:", row); row += 1
        self.cb_status = make_dark_combo(
            form, values=["Активний", "Монiторинг", "Мiтигований", "Закрито"])
        self.cb_status.grid(row=row, column=0, sticky="w", padx=16, pady=(2, 0)); row += 1

        row = section("Деталi", row)
        for lbl_t, attr, h in [
            ("Заходи контролю:",       "txt_controls",    3),
            ("Детальний опис ризику:", "txt_description", 4),
        ]:
            lbl(lbl_t, row); row += 1
            t = make_dark_text(form, height=h, wrap="word")
            t.grid(row=row, column=0, sticky="ew", padx=16, pady=(2, 0))
            setattr(self, attr, t); row += 1

        btn_f = tk.Frame(form, bg=C["bg_main"])
        btn_f.grid(row=row, column=0, sticky="ew", padx=16, pady=(16, 20))
        btn_f.columnconfigure((0, 1), weight=1)

        make_button(btn_f, "Очистити",
                    bg=C["bg_surface"], fg=C["text_muted"],
                    activebackground=C["bg_surface_alt"],
                    activeforeground=C["text_primary"],
                    padx=14, pady=6,
                    command=self._clear_form).grid(row=0, column=0, padx=4, sticky="ew")
        make_button(btn_f, "Додати ризик",
                    bg=C["accent"], activebackground=C["accent_soft"],
                    font=FONT_BOLD, padx=14, pady=6,
                    command=self._add_record).grid(row=0, column=1, padx=4, sticky="ew")

    def _build_table(self, container: tk.Misc) -> None:
        C = COLORS
        container.rowconfigure(1, weight=1)
        container.columnconfigure(0, weight=1)

        toolbar = tk.Frame(container, bg=C["bg_surface"], height=40)
        toolbar.grid(row=0, column=0, sticky="ew")
        toolbar.grid_propagate(False)

        tk.Label(toolbar, text="Записи", bg=C["bg_surface"],
                 fg=C["text_muted"], font=FONT_SMALL).pack(side="left", padx=12, pady=8)
        self.lbl_count = tk.Label(toolbar, text="0", bg=C["bg_surface"],
                                   fg=C["accent"], font=FONT_SMALL_BOLD)
        self.lbl_count.pack(side="left", pady=8)

        tk.Label(toolbar, text="  |  Тип:", bg=C["bg_surface"],
                 fg=C["text_muted"], font=FONT_SMALL).pack(side="left", pady=8)
        self.filter_type = make_dark_combo(toolbar, values=["Всi"] + RISK_TYPES, width=16)
        self.filter_type.set("Всi")
        self.filter_type.pack(side="left", padx=6, pady=8)
        self.filter_type.bind("<<ComboboxSelected>>", lambda _: self._apply_filter())

        tk.Label(toolbar, text="Статус:", bg=C["bg_surface"],
                 fg=C["text_muted"], font=FONT_SMALL).pack(side="left", pady=8)
        self.filter_status = make_dark_combo(
            toolbar,
            values=["Всi", "Активний", "Монiторинг", "Мiтигований", "Закрито"],
            width=12)
        self.filter_status.set("Всi")
        self.filter_status.pack(side="left", padx=6, pady=8)
        self.filter_status.bind("<<ComboboxSelected>>", lambda _: self._apply_filter())

        for txt, cmd, bg in [
            ("Переглянути", self._open_selected_detail, C["accent"]),
            ("Дублювати",   self._duplicate_record,     C["bg_surface_alt"]),
            ("Видалити",    self._delete_selected,       C["accent_danger"]),
        ]:
            make_button(toolbar, txt, bg=bg,
                        fg="white" if bg != C["bg_surface_alt"] else C["text_primary"],
                        activebackground=bg, activeforeground="white",
                        font=FONT_SMALL, padx=10, pady=3,
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
            self.tree.heading(col, text=txt, command=lambda c=col: self._sort_tree(c))
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
                 font=("Arial", 7, "italic")).pack(side="left", padx=8, pady=4)

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
                    font=FONT_SMALL, padx=12, pady=4,
                    command=self._export_csv).pack(side="left", padx=(0, 6))
        if pd:
            make_button(exp_bar, "Експорт Excel",
                        bg=C["accent_success"], activebackground="#16a34a",
                        font=FONT_SMALL, padx=12, pady=4,
                        command=self._export_excel).pack(side="left", padx=(0, 6))
        make_button(exp_bar, "Iмпорт JSON",
                    bg=C["bg_surface"], fg=C["text_primary"],
                    activebackground=C["bg_surface_alt"],
                    activeforeground=C["text_primary"],
                    font=FONT_SMALL, padx=12, pady=4,
                    command=self._import_json).pack(side="left")

    # ── Events ───────────────────────────────────────────────────────────────

    def _on_double_click(self, event: tk.Event) -> None:
        if not self.tree.selection():
            return
        if self.tree.identify("region", event.x, event.y) != "cell":
            return
        self._open_selected_detail()

    def _open_selected_detail(self) -> None:
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Перегляд", "Оберiть запис для перегляду"); return
        iid = sel[0]
        rec = self.find_record(self.tree.set(iid, "id"))
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

    def _on_detail_save(self, iid: str, old_id: str, new_record: RiskRecord) -> None:
        normalized_old = IdGenerator.normalize_id(old_id)
        for i, r in enumerate(self.all_records):
            if IdGenerator.normalize_id(r.id) == normalized_old:
                self.all_records[i] = new_record
                break
        try:
            self.tree.item(iid, values=self._tree_values(new_record))
        except tk.TclError:
            pass
        self._recolor_rows()
        self._save_data()
        self._notify_change()

    def _on_detail_delete(self, iid: str, idx_str: str) -> None:
        try:
            self.tree.delete(iid)
        except tk.TclError:
            pass
        normalized = IdGenerator.normalize_id(idx_str)
        self.all_records = [
            r for r in self.all_records
            if IdGenerator.normalize_id(r.id) != normalized
        ]
        self._recolor_rows()
        self._save_data()
        self._update_count()
        self._notify_change()

    # ── Таблиця ──────────────────────────────────────────────────────────────

    def _sort_tree(self, col: str) -> None:
        data = [(self.tree.set(iid, col), iid) for iid in self.tree.get_children("")]
        try:
            data.sort(key=lambda x: float(x[0]) if x[0] not in ("—", "") else 0)
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

    @staticmethod
    def _tree_values(rec: RiskRecord) -> tuple:
        return (rec.id, rec.entity, rec.risk_name, rec.category,
                rec.risk_type, rec.score, rec.priority,
                rec.status, rec.owner, rec.date_id)

    # ── Дані ─────────────────────────────────────────────────────────────────

    def _load_data(self) -> None:
        self.all_records.clear()
        self.tree.delete(*self.tree.get_children())
        if not os.path.exists(self.data_file):
            return
        try:
            with open(self.data_file, "r", encoding="utf-8") as f:
                raw = json.load(f)
            if not isinstance(raw, list):
                raise ValueError("Очiкується список записiв у JSON")
            for item in raw:
                if isinstance(item, dict):
                    rec = RiskRecord.from_dict(item)
                elif isinstance(item, (list, tuple)):
                    rec = RiskRecord.from_list(list(item))
                else:
                    continue
                self.all_records.append(rec)
                self._insert_tree_row(rec)
        except (json.JSONDecodeError, ValueError, OSError) as e:
            messagebox.showerror("Помилка завантаження", str(e))
        self._update_count()
        self._notify_change()

    def _insert_tree_row(self, rec: RiskRecord) -> str:
        iid = self.tree.insert("", tk.END, values=self._tree_values(rec))
        self._recolor_rows()
        return iid

    # ── Форма ─────────────────────────────────────────────────────────────────

    def _get_form_data(self) -> RiskRecord | None:
        date_id  = self.ent_date_id.get().strip()
        date_rev = self.ent_date_rev.get().strip()
        for val, label in [(date_id, "дати виявлення"), (date_rev, "дати перегляду")]:
            if val in ("дд.мм.рррр", ""):
                continue
            if not is_valid_date(val):
                messagebox.showwarning(
                    "Помилка",
                    f"Неправильний формат {label} (очiкується дд.мм.рррр)")
                return None

        entity    = self.ent_entity.get().strip()
        risk_name = self.ent_risk_name.get().strip()
        # ВИПРАВЛЕННЯ #10: валідація обов'язкових полів
        if not entity or not risk_name:
            messagebox.showwarning(
                "Обов'язковi поля",
                "Заповнiть назву пiдприємства та назву ризику")
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

        # ВИПРАВЛЕННЯ #3: надійна генерація ID
        new_id = IdGenerator.next_id(self.all_records)
        return RiskRecord(
            id=new_id,
            entity=entity,
            risk_name=risk_name,
            category=self.cb_category.get().strip()  or "—",
            risk_type=self.cb_risk_type.get().strip() or "—",
            probability=prob_str   or "—",
            impact=impact_str      or "—",
            score=str(score),
            owner=self.ent_owner.get().strip(),
            controls=self.txt_controls.get("1.0", tk.END).strip(),
            residual=str(residual),
            date_id=date_id  or "—",
            date_rev=date_rev or "—",
            priority=self.cb_priority.get().strip() or "Середнiй",
            status=self.cb_status.get().strip()   or "Активний",
            description=self.txt_description.get("1.0", tk.END).strip(),
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
        for e, ph in [(self.ent_date_id, "дд.мм.рррр"),
                      (self.ent_date_rev, "дд.мм.рррр")]:
            e.delete(0, tk.END)
            add_placeholder(e, ph)

    def _add_record(self) -> None:
        data = self._get_form_data()
        if not data:
            return
        self.all_records.append(data)
        self._insert_tree_row(data)
        self._clear_form()
        self._save_data()
        self._notify_change()
        self._update_count()
        self._show_toast("Ризик додано")

    def _delete_selected(self) -> None:
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Видалення", "Оберiть запис для видалення"); return
        iid     = sel[0]
        idx_str = self.tree.set(iid, "id")
        if not messagebox.askyesno(
            "Пiдтвердження",
            f"Видалити ризик #{idx_str}?\nЦю дiю не можна скасувати."):
            return
        self.tree.delete(iid)
        normalized = IdGenerator.normalize_id(idx_str)
        self.all_records = [
            r for r in self.all_records
            if IdGenerator.normalize_id(r.id) != normalized
        ]
        self._recolor_rows()
        self._save_data()
        self._update_count()
        self._notify_change()
        self._show_toast("Запис видалено")

    def _duplicate_record(self) -> None:
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Дублювання", "Оберiть запис для дублювання"); return
        rec = self.find_record(self.tree.set(sel[0], "id"))
        if not rec:
            return
        import dataclasses
        new_rec = dataclasses.replace(rec, id=IdGenerator.next_id(self.all_records))
        self.all_records.append(new_rec)
        self._insert_tree_row(new_rec)
        self._save_data()
        self._update_count()
        self._notify_change()
        self._show_toast("Запис продубльовано")

    def _apply_filter(self) -> None:
        q      = self.search_var.get().strip().lower()
        r_type = self.filter_type.get()
        status = self.filter_status.get()
        self.tree.delete(*self.tree.get_children())
        for rec in self.all_records:
            row_str = " ".join(rec.to_list()).lower()
            if q      and q      not in row_str:        continue
            if r_type != "Всi"   and rec.risk_type != r_type: continue
            if status != "Всi"   and rec.status    != status: continue
            self._insert_tree_row(rec)
        self._update_count()

    def _reset_filter(self) -> None:
        self.search_var.set("")
        self.filter_type.set("Всi")
        self.filter_status.set("Всi")
        self.tree.delete(*self.tree.get_children())
        for rec in self.all_records:
            self._insert_tree_row(rec)
        self._update_count()

    def _on_select(self, _: object | None = None) -> None:
        sel = self.tree.selection()
        if not sel:
            return
        rec = self.find_record(self.tree.set(sel[0], "id"))
        if not rec:
            return
        for widget, text in [(self.det_controls, rec.controls),
                             (self.det_desc,     rec.description)]:
            widget.configure(state="normal")
            widget.delete("1.0", tk.END)
            widget.insert("1.0", text or "")
            widget.configure(state="disabled")

    def _update_count(self) -> None:
        self.lbl_count.configure(text=f" {len(self.tree.get_children())}")

    def _show_toast(self, msg: str) -> None:
        _show_toast(self.frame, msg)

    # ── Експорт / Імпорт ─────────────────────────────────────────────────────

    _HEADERS = [
        "ID", "Пiдприємство", "Назва ризику", "Категорiя", "Тип ризику",
        "Iмовiрнiсть", "Вплив", "Score", "Власник", "Заходи контролю",
        "Залишковий ризик", "Дата виявлення", "Дата перегляду",
        "Прiоритет", "Статус", "Опис",
    ]

    def _export_csv(self) -> None:
        if not self.tree.get_children():
            messagebox.showinfo("Експорт", "Таблиця порожня"); return
        path = filedialog.asksaveasfilename(
            defaultextension=".csv", filetypes=[("CSV файли", "*.csv")],
            title="Зберегти реєстр ризикiв як CSV")
        if not path:
            return
        try:
            with open(path, "w", newline="", encoding="utf-8-sig") as f:
                w = csv.writer(f)
                w.writerow(self._HEADERS)
                for rec in self.all_records:
                    w.writerow(rec.to_list())
            self._show_toast("CSV збережено")
        except OSError as e:
            messagebox.showerror("Помилка", str(e))

    def _export_excel(self) -> None:
        if not pd:
            messagebox.showwarning("Excel", "Встановiть pandas та openpyxl"); return
        if not self.all_records:
            messagebox.showinfo("Експорт", "Немає записiв"); return
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx", filetypes=[("Excel файли", "*.xlsx")],
            title="Зберегти реєстр ризикiв як Excel")
        if not path:
            return
        try:
            df = pd.DataFrame([rec.to_list() for rec in self.all_records],
                               columns=self._HEADERS)
            with pd.ExcelWriter(path, engine="openpyxl") as writer:
                df.to_excel(writer, index=False, sheet_name="Реєстр ризикiв")
                ws = writer.sheets["Реєстр ризикiв"]
                for col_cells in ws.columns:
                    mx = max(len(str(c.value or "")) for c in col_cells)
                    ws.column_dimensions[col_cells[0].column_letter].width = min(mx + 4, 60)
            self._show_toast("Excel збережено")
        except (OSError, Exception) as e:
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
            import dataclasses
            added = 0
            for item in data:
                if isinstance(item, dict):
                    rec = RiskRecord.from_dict(item)
                elif isinstance(item, (list, tuple)):
                    rec = RiskRecord.from_list(list(item))
                else:
                    continue
                rec = dataclasses.replace(rec, id=IdGenerator.next_id(self.all_records))
                self.all_records.append(rec)
                self._insert_tree_row(rec)
                added += 1
            self._save_data()
            self._update_count()
            self._notify_change()
            self._show_toast(f"Iмпортовано: {added} записiв")
        except (json.JSONDecodeError, ValueError, OSError) as e:
            messagebox.showerror("Помилка iмпорту", str(e))


# =============================================================================
#  ВКЛАДКА: АНАЛІТИКА РИЗИКІВ
# =============================================================================

class RiskAnalyticsTab:

    def __init__(self, parent: tk.Misc) -> None:
        self.frame   = ttk.Frame(parent)
        self.records: list[RiskRecord] = []
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
                 font=FONT_TITLE).pack(side="left", padx=20, pady=14)
        make_button(header, "Оновити",
                    bg=C["accent"], activebackground=C["accent_soft"],
                    font=FONT_BOLD, padx=12, pady=4,
                    command=self.refresh).pack(side="right", padx=20, pady=12)

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
        _scroll_mgr.attach(canvas)  # ВИПРАВЛЕННЯ #1
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
                     fg=C["text_muted"], font=FONT_SMALL).pack(anchor="w", pady=(8, 2))
            lbl = tk.Label(card, text=val, bg=C["bg_surface"],
                           fg=color, font=FONT_NUMBER)
            lbl.pack(anchor="w")
            self.stat_cards[key] = lbl

    def _build_charts_and_table(self) -> None:
        C = COLORS
        if HAS_MPL:
            cr = tk.Frame(self.content, bg=C["bg_main"])
            cr.grid(row=1, column=0, sticky="ew", padx=16, pady=16)
            cr.columnconfigure((0, 1), weight=1)

            self.fig_left = Figure(figsize=(5, 3.5), dpi=90, facecolor=C["bg_surface"])
            self.ax_left  = self.fig_left.add_subplot(111)
            self._style_ax(self.ax_left)
            self.ax_left.set_title("Розподiл за типом ризику",
                                    color=C["text_muted"], fontsize=9)
            fl = tk.Frame(cr, bg=C["bg_surface"], padx=8, pady=8)
            fl.grid(row=0, column=0, padx=(0, 8), sticky="nsew")
            self.canvas_left = FigureCanvasTkAgg(self.fig_left, master=fl)
            self.canvas_left.get_tk_widget().pack(fill="both", expand=True)

            self.fig_right = Figure(figsize=(5, 3.5), dpi=90, facecolor=C["bg_surface"])
            self.ax_right  = self.fig_right.add_subplot(111)
            self._style_ax(self.ax_right)
            self.ax_right.set_title("Розподiл за рiвнем ризику",
                                     color=C["text_muted"], fontsize=9)
            fr = tk.Frame(cr, bg=C["bg_surface"], padx=8, pady=8)
            fr.grid(row=0, column=1, padx=(8, 0), sticky="nsew")
            self.canvas_right = FigureCanvasTkAgg(self.fig_right, master=fr)
            self.canvas_right.get_tk_widget().pack(fill="both", expand=True)

            self.fig_heat = Figure(figsize=(10, 4), dpi=90, facecolor=C["bg_surface"])
            self.ax_heat  = self.fig_heat.add_subplot(111)
            self._style_ax(self.ax_heat)
            self.ax_heat.set_title("Матриця ризикiв (Iмовiрнiсть × Вплив)",
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
                 font=FONT_BOLD).grid(row=0, column=0, sticky="w", pady=(0, 8))

        cols = ("risk_type", "count", "avg_score", "max_score", "active")
        self.stats_tree = ttk.Treeview(frame, columns=cols, show="headings", height=7)
        for col, hdr, w in [
            ("risk_type", "Тип ризику", 180), ("count", "Всього", 70),
            ("avg_score", "Сер. Score", 90),  ("max_score", "Макс. Score", 90),
            ("active", "Активних", 80),
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

    def update_data(self, records: list[RiskRecord]) -> None:
        self.records = records
        self.refresh()

    def refresh(self) -> None:
        if not self.records:
            for k in self.stat_cards:
                self.stat_cards[k].configure(text="0")
            if HAS_MPL:
                self._clear_charts()
            self.stats_tree.delete(*self.stats_tree.get_children())
            return

        C       = COLORS
        records = self.records
        total     = len(records)
        critical_ = sum(1 for r in records if r.score.isdigit() and int(r.score) > 16)
        high_     = sum(1 for r in records if r.score.isdigit() and 10 <= int(r.score) <= 16)
        active_   = sum(1 for r in records if r.status == "Активний")
        mitig_    = sum(1 for r in records if r.status == "Мiтигований")

        self.stat_cards["total"].configure(text=str(total))
        self.stat_cards["critical"].configure(text=str(critical_))
        self.stat_cards["high"].configure(text=str(high_))
        self.stat_cards["active"].configure(text=str(active_))
        self.stat_cards["mitigated"].configure(text=str(mitig_))

        if not HAS_MPL:
            return

        type_ctr = Counter(r.risk_type for r in records)

        self.ax_left.clear(); self._style_ax(self.ax_left)
        self.ax_left.set_title("Розподiл за типом ризику",
                                color=C["text_muted"], fontsize=9)
        if type_ctr:
            lbls = list(type_ctr.keys()); vals = list(type_ctr.values())
            clrs = [RISK_COLORS.get(l, C["text_muted"]) for l in lbls]
            _, _, autotexts = self.ax_left.pie(
                vals, labels=lbls, autopct="%1.0f%%", colors=clrs, startangle=90,
                textprops={"color": C["text_muted"], "fontsize": 7})
            for at in autotexts:
                at.set_fontsize(7); at.set_color("white")
        else:
            self.ax_left.text(0.5, 0.5, "Немає даних",
                              transform=self.ax_left.transAxes,
                              ha="center", va="center", color=C["text_muted"])
        self.canvas_left.draw()

        self.ax_right.clear(); self._style_ax(self.ax_right)
        self.ax_right.set_title("Розподiл за рiвнем ризику",
                                 color=C["text_muted"], fontsize=9)
        level_ctr = {"Низький": 0, "Помiрний": 0, "Високий": 0, "Критичний": 0}
        for r in records:
            if r.score.isdigit():
                level_ctr[_score_label(int(r.score))] += 1
        if any(level_ctr.values()):
            lbls = list(level_ctr.keys()); vals = list(level_ctr.values())
            clrs = [C["accent_success"], C["accent_warning"], "#f97316", C["accent_danger"]]
            bars = self.ax_right.bar(lbls, vals, color=clrs, edgecolor="none")
            for bar, val in zip(bars, vals, strict=False):
                if val > 0:
                    self.ax_right.text(bar.get_x() + bar.get_width() / 2,
                                        bar.get_height() + 0.1, str(val),
                                        ha="center", va="bottom",
                                        color=C["text_muted"], fontsize=8)
            self.ax_right.tick_params(axis="x", labelrotation=15, labelsize=7)
            mx = max(vals)
            self.ax_right.set_ylim(0, mx * 1.2 + 1 if mx > 0 else 1)
        else:
            self.ax_right.text(0.5, 0.5, "Немає даних",
                               transform=self.ax_right.transAxes,
                               ha="center", va="center", color=C["text_muted"])
        self.canvas_right.draw()

        self.ax_heat.clear(); self._style_ax(self.ax_heat)
        self.ax_heat.set_title("Матриця ризикiв (Iмовiрнiсть × Вплив)",
                                color=C["text_muted"], fontsize=9)
        matrix = [[0] * 5 for _ in range(5)]
        for r in records:
            try:
                prob = _extract_num(r.probability)
                imp  = _extract_num(r.impact)
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
            self.ax_heat.set_xlabel("Вплив →",        color=C["text_muted"], fontsize=8)
            self.ax_heat.set_ylabel("Iмовiрнiсть →",  color=C["text_muted"], fontsize=8)
            for i in range(5):
                for j in range(5):
                    if matrix[i][j] > 0:
                        self.ax_heat.text(j, i, str(matrix[i][j]),
                                           ha="center", va="center",
                                           color="white", fontsize=10, weight="bold")
        else:
            self.ax_heat.text(0.5, 0.5, "Немає даних",
                              transform=self.ax_heat.transAxes,
                              ha="center", va="center", color=C["text_muted"])
        self.canvas_heat.draw()

        self.stats_tree.delete(*self.stats_tree.get_children())
        all_types = set(RISK_TYPES) | {r.risk_type for r in records}
        for rt in sorted(all_types):
            recs = [r for r in records if r.risk_type == rt]
            cnt  = len(recs)
            if cnt:
                scores    = [int(r.score) for r in recs if r.score.isdigit()]
                avg_score = sum(scores) / len(scores) if scores else 0
                max_score = max(scores) if scores else 0
                act       = sum(1 for r in recs if r.status == "Активний")
                self.stats_tree.insert("", tk.END,
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
#  ВКЛАДКИ НАЛАШТУВАНЬ
# =============================================================================

class RiskSettingsTab:

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
                 font=FONT_TITLE).pack(side="left", padx=20, pady=14)

        content = tk.Frame(self.frame, bg=C["bg_main"])
        content.grid(row=1, column=0, sticky="nsew", padx=40, pady=30)
        content.columnconfigure(0, weight=1)

        self._row(content, 0, "Файл даних:", RISK_DATA_FILE, C)
        self._row(content, 1, "Версiя:", "2.2 — Рефакторинг", C)
        self._row(content, 2, "matplotlib:",
                  "встановлено" if HAS_MPL else "не встановлено", C)
        self._row(content, 3, "pandas:",
                  "встановлено" if pd else "не встановлено", C)

        tk.Label(content, text="Встановлення залежностей:",
                 bg=C["bg_main"], fg=C["text_muted"],
                 font=FONT_SMALL_BOLD).grid(row=4, column=0, sticky="w", pady=(24, 6))
        tk.Label(content, text="  pip install matplotlib pandas openpyxl",
                 bg=C["bg_surface"], fg=C["accent_muted"],
                 font=FONT_MONO, padx=12, pady=8).grid(row=5, column=0, sticky="w")

        tk.Label(content, text="Структура запису (16 полiв):",
                 bg=C["bg_main"], fg=C["text_muted"],
                 font=FONT_SMALL_BOLD).grid(row=6, column=0, sticky="w", pady=(24, 6))

        fields_text = (
            "ID, Пiдприємство, Назва ризику, Категорiя, Тип ризику,\n"
            "Iмовiрнiсть, Вплив, Score, Власник, Заходи контролю,\n"
            "Залишковий ризик, Дата виявлення, Дата перегляду, "
            "Прiоритет, Статус, Опис"
        )
        tk.Label(content, text=fields_text,
                 bg=C["bg_surface"], fg=C["text_subtle"],
                 font=FONT_SMALL, justify="left",
                 padx=12, pady=8).grid(row=7, column=0, sticky="w")

        hints = [
            "Score розраховується автоматично: Iмовiрнiсть × Вплив (1–25)",
            "Рiвнi ризику: Низький (1–4), Помiрний (5–9), Високий (10–16), Критичний (17–25)",
            "Подвiйний клiк по рядку — вiдкрити детальне вiкно ризику",
            "JSON зберiгається як словники — сумiсно з майбутнiми версiями",
        ]
        tk.Label(content, text="Пiдказки:",
                 bg=C["bg_main"], fg=C["text_muted"],
                 font=FONT_SMALL_BOLD).grid(row=8, column=0, sticky="w", pady=(24, 6))
        for i, hint in enumerate(hints):
            f = tk.Frame(content, bg=C["bg_main"])
            f.grid(row=9 + i, column=0, sticky="w", pady=2)
            tk.Frame(f, bg=C["accent_success"], width=4, height=4).pack(
                side="left", padx=(0, 8))
            tk.Label(f, text=hint, bg=C["bg_main"], fg=C["text_subtle"],
                     font=FONT_SMALL).pack(side="left")

    def _row(self, parent: tk.Misc, row: int, label: str, value: str, C: dict) -> None:
        f = tk.Frame(parent, bg=C["bg_main"])
        f.grid(row=row, column=0, sticky="ew", pady=4)
        tk.Label(f, text=label, bg=C["bg_main"], fg=C["text_muted"],
                 font=FONT_DEFAULT, width=22, anchor="w").pack(side="left")
        tk.Label(f, text=value, bg=C["bg_main"], fg=C["text_primary"],
                 font=FONT_DEFAULT).pack(side="left")

    def get_frame(self) -> ttk.Frame:
        return self.frame


# =============================================================================
#  СТОРІНКА "РЕЄСТР РИЗИКІВ"
# =============================================================================

class RiskRegisterPage(tk.Frame):

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

        self.notebook.add(self.registry_tab.get_frame(),  text="  Реєстр ризикiв  ")
        self.notebook.add(self.analytics_tab.get_frame(), text="  Аналiтика  ")
        self.notebook.add(self.settings_tab.get_frame(),  text="  Налаштування  ")

        statusbar = tk.Frame(self, bg=COLORS["bg_header"], height=22)
        statusbar.grid(row=1, column=0, sticky="ew")
        statusbar.grid_propagate(False)

        self._status_lbl = tk.Label(
            statusbar, text="Готово",
            bg=COLORS["bg_header"], fg=COLORS["text_muted"],
            font=FONT_TINY, padx=10)
        self._status_lbl.pack(side="left", pady=3)

        self._time_lbl = tk.Label(
            statusbar, text="",
            bg=COLORS["bg_header"], fg=COLORS["text_muted"],
            font=FONT_TINY, padx=10)
        self._time_lbl.pack(side="right", pady=3)

        self._start_clock()
        self._schedule_autosave()
        self.after(600, lambda: self.analytics_tab.update_data(
            self.registry_tab.all_records))

    def _start_clock(self) -> None:
        self._time_lbl.configure(text=datetime.now().strftime("%d.%m.%Y  %H:%M:%S"))
        self.after(1000, self._start_clock)

    def _schedule_autosave(self) -> None:
        try:
            self.registry_tab.save()  # ВИПРАВЛЕННЯ #4
            self._status_lbl.configure(
                text=f"Автозбережено о {datetime.now().strftime('%H:%M:%S')}")
        except OSError:
            self._status_lbl.configure(text="Помилка автозбереження")
        self.after(30000, self._schedule_autosave)

    def save_before_exit(self) -> None:
        try:
            self.registry_tab.save()  # ВИПРАВЛЕННЯ #4
        except OSError:
            pass


# =============================================================================
#  МОДУЛЬ: РИЗИК КООРДИНАТОРИ
# =============================================================================

@dataclass
class CoordinatorRecord:
    """Запис про ризик-координатора підприємства."""
    id:            str = ""          # унікальний UUID рядка
    enterprise:    str = ""          # назва підприємства
    department:    str = ""          # назва управління/підрозділу
    location:      str = ""          # місцезнаходження
    is_staff_unit: str = "Нi"        # штатна одиниця: Так / Нi
    is_concurrent: str = "Нi"        # суміщення: Так / Нi
    main_position: str = ""          # назва основної посади
    full_name:     str = ""          # ПІБ
    phone:         str = ""          # номер телефону
    appointed_date:str = "—"         # дата прийняття на посаду
    order_number:  str = ""          # номер наказу
    has_approval:  str = "Нi"        # погодження: Так / Нi

    def to_list(self) -> list:
        return [
            self.id, self.enterprise, self.department, self.location,
            self.is_staff_unit, self.is_concurrent, self.main_position,
            self.full_name, self.phone, self.appointed_date,
            self.order_number, self.has_approval,
        ]

    @classmethod
    def from_dict(cls, d: dict) -> "CoordinatorRecord":
        valid = {k: str(v) for k, v in d.items()
                 if k in cls.__dataclass_fields__}
        return cls(**valid)

    @classmethod
    def from_list(cls, row: list) -> "CoordinatorRecord":
        r = list(row) + [""] * max(0, 12 - len(row))
        return cls(
            id=str(r[0]), enterprise=str(r[1]), department=str(r[2]),
            location=str(r[3]), is_staff_unit=str(r[4]),
            is_concurrent=str(r[5]), main_position=str(r[6]),
            full_name=str(r[7]), phone=str(r[8]),
            appointed_date=str(r[9]), order_number=str(r[10]),
            has_approval=str(r[11]),
        )


# ---------------------------------------------------------------------------
#  Діалог: Додати / Редагувати координатора
# ---------------------------------------------------------------------------

class CoordinatorFormDialog:
    """
    Модальний діалог для додавання або редагування одного координатора.
    Повертає CoordinatorRecord через callback on_save(record).
    """

    def __init__(
        self,
        parent_root: tk.Misc,
        on_save: Callable[[CoordinatorRecord], None],
        record: CoordinatorRecord | None = None,
        default_enterprise: str = "",
    ) -> None:
        self.on_save   = on_save
        self.record    = record
        self.is_edit   = record is not None
        self._build(parent_root, default_enterprise)

    def _build(self, parent_root: tk.Misc, default_enterprise: str) -> None:
        C = COLORS
        self.win = tk.Toplevel(parent_root)
        title_text = "Редагувати координатора" if self.is_edit else "Новий координатор"
        self.win.title(title_text)
        self.win.geometry("560x680")
        self.win.minsize(480, 560)
        self.win.configure(bg=C["bg_main"])
        self.win.grab_set()

        # Центрування
        self.win.update_idletasks()
        rx = parent_root.winfo_x(); ry = parent_root.winfo_y()
        rw = parent_root.winfo_width(); rh = parent_root.winfo_height()
        ww, wh = 560, 680
        self.win.geometry(f"{ww}x{wh}+{rx+(rw-ww)//2}+{ry+(rh-wh)//2}")

        self.win.rowconfigure(1, weight=1)
        self.win.columnconfigure(0, weight=1)

        # Header
        hdr = tk.Frame(self.win, bg=C["bg_header"], height=52)
        hdr.grid(row=0, column=0, sticky="ew")
        hdr.grid_propagate(False)
        hdr.columnconfigure(0, weight=1)
        accent_color = C["accent_warning"] if self.is_edit else C["accent"]
        tk.Frame(hdr, bg=accent_color, width=4).grid(row=0, column=0, sticky="ns",
                                                       padx=(0, 0))
        tk.Label(hdr, text=title_text,
                 bg=C["bg_header"], fg=C["text_primary"],
                 font=FONT_HEADING).grid(row=0, column=0, padx=20, pady=14, sticky="w")

        # Scrollable body
        canvas = tk.Canvas(self.win, bg=C["bg_main"], highlightthickness=0)
        sb = ttk.Scrollbar(self.win, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=sb.set)
        canvas.grid(row=1, column=0, sticky="nsew")
        sb.grid(row=1, column=1, sticky="ns")

        body = tk.Frame(canvas, bg=C["bg_main"])
        cw   = canvas.create_window((0, 0), window=body, anchor="nw")

        def _conf(_: object) -> None:
            canvas.configure(scrollregion=canvas.bbox("all"))
            canvas.itemconfig(cw, width=canvas.winfo_width())

        body.bind("<Configure>", _conf)
        canvas.bind("<Configure>", _conf)
        _scroll_mgr.attach(canvas)
        body.columnconfigure(0, weight=1)
        body.columnconfigure(1, weight=1)

        rec = self.record or CoordinatorRecord()

        def section(txt: str, r: int, span: int = 2) -> int:
            f = tk.Frame(body, bg=C["bg_main"])
            f.grid(row=r, column=0, columnspan=span,
                   sticky="ew", padx=12, pady=(14, 4))
            tk.Frame(f, bg=accent_color, width=3, height=14).pack(side="left")
            tk.Label(f, text=txt, bg=C["bg_main"], fg=accent_color,
                     font=FONT_BOLD).pack(side="left", padx=8)
            return r + 1

        def lbl_entry(txt: str, row: int, col: int, default: str = "",
                      width: int | None = None) -> tk.Entry:
            tk.Label(body, text=txt, bg=C["bg_main"],
                     fg=C["text_subtle"], font=FONT_SMALL).grid(
                row=row, column=col, sticky="w", padx=(14, 4), pady=(6, 0))
            kw: dict = dict(accent=accent_color)
            if width:
                kw["width"] = width
            e = make_dark_entry(body, **kw)
            e.grid(row=row + 1, column=col, sticky="ew",
                   padx=(14, 8), pady=(2, 0))
            if default:
                e.insert(0, default)
            return e

        def lbl_combo(txt: str, row: int, col: int,
                      values: list[str], current: str) -> ttk.Combobox:
            tk.Label(body, text=txt, bg=C["bg_main"],
                     fg=C["text_subtle"], font=FONT_SMALL).grid(
                row=row, column=col, sticky="w", padx=(14, 4), pady=(6, 0))
            cb = make_dark_combo(body, values=values)
            cb.grid(row=row + 1, column=col, sticky="ew",
                    padx=(14, 8), pady=(2, 0))
            cb.set(current if current in values else values[0])
            return cb

        row = 0
        row = section("Підприємство та підрозділ", row)

        # span entire width for enterprise
        tk.Label(body, text="Назва підприємства:", bg=C["bg_main"],
                 fg=C["text_subtle"], font=FONT_SMALL).grid(
            row=row, column=0, columnspan=2, sticky="w", padx=(14, 4), pady=(6, 0))
        self.e_enterprise = make_dark_entry(body, accent=accent_color)
        self.e_enterprise.grid(row=row + 1, column=0, columnspan=2,
                                sticky="ew", padx=(14, 14), pady=(2, 0))
        self.e_enterprise.insert(0, rec.enterprise or default_enterprise)
        row += 2

        self.e_department = lbl_entry("Назва управління/підрозділу:", row, 0,
                                       default=rec.department)
        self.e_location   = lbl_entry("Місцезнаходження:", row, 1,
                                       default=rec.location)
        row += 2

        row = section("Посада та ПІБ", row)
        # full-width main_position
        tk.Label(body, text="Назва основної посади:", bg=C["bg_main"],
                 fg=C["text_subtle"], font=FONT_SMALL).grid(
            row=row, column=0, columnspan=2, sticky="w", padx=(14, 4), pady=(6, 0))
        self.e_main_position = make_dark_entry(body, accent=accent_color)
        self.e_main_position.grid(row=row + 1, column=0, columnspan=2,
                                   sticky="ew", padx=(14, 14), pady=(2, 0))
        self.e_main_position.insert(0, rec.main_position)
        row += 2

        # full-width full_name
        tk.Label(body, text="ПІБ:", bg=C["bg_main"],
                 fg=C["text_subtle"], font=FONT_SMALL).grid(
            row=row, column=0, columnspan=2, sticky="w", padx=(14, 4), pady=(6, 0))
        self.e_full_name = make_dark_entry(body, accent=accent_color)
        self.e_full_name.grid(row=row + 1, column=0, columnspan=2,
                               sticky="ew", padx=(14, 14), pady=(2, 0))
        self.e_full_name.insert(0, rec.full_name)
        row += 2

        row = section("Контакти та призначення", row)
        self.e_phone = lbl_entry("Номер телефону:", row, 0, default=rec.phone)

        tk.Label(body, text="Дата прийняття на посаду:", bg=C["bg_main"],
                 fg=C["text_subtle"], font=FONT_SMALL).grid(
            row=row, column=1, sticky="w", padx=(14, 4), pady=(6, 0))
        self.e_appointed = make_dark_entry(body, accent=accent_color)
        self.e_appointed.grid(row=row + 1, column=1, sticky="ew",
                               padx=(14, 14), pady=(2, 0))
        if rec.appointed_date and rec.appointed_date != "—":
            self.e_appointed.insert(0, rec.appointed_date)
        else:
            add_placeholder(self.e_appointed, "дд.мм.рррр")
        row += 2

        # full-width order_number
        tk.Label(body, text="Номер наказу:", bg=C["bg_main"],
                 fg=C["text_subtle"], font=FONT_SMALL).grid(
            row=row, column=0, columnspan=2, sticky="w", padx=(14, 4), pady=(6, 0))
        self.e_order = make_dark_entry(body, accent=accent_color)
        self.e_order.grid(row=row + 1, column=0, columnspan=2,
                           sticky="ew", padx=(14, 14), pady=(2, 0))
        self.e_order.insert(0, rec.order_number)
        row += 2

        row = section("Відмітки та погодження", row)

        self.cb_staff  = lbl_combo("Штатна одиниця:", row, 0,
                                    ["Так", "Нi"], rec.is_staff_unit)
        self.cb_concur = lbl_combo("Виконання обов'язків за сумісництвом:", row, 1,
                                    ["Так", "Нi"], rec.is_concurrent)
        row += 2

        # approval - full row, with visual indicator
        tk.Label(body, text="Наявність погодження:", bg=C["bg_main"],
                 fg=C["text_subtle"], font=FONT_SMALL).grid(
            row=row, column=0, columnspan=2, sticky="w", padx=(14, 4), pady=(6, 0))
        appr_f = tk.Frame(body, bg=C["bg_main"])
        appr_f.grid(row=row + 1, column=0, columnspan=2,
                    sticky="w", padx=14, pady=(2, 0))
        self._approval_var = tk.StringVar(value=rec.has_approval)

        def _update_dot(*_: object) -> None:
            val = self._approval_var.get()
            dot_color = C["accent_success"] if val == "Так" else C["accent_danger"]
            self._approval_dot.configure(bg=dot_color)

        for val in ("Так", "Нi"):
            tk.Radiobutton(
                appr_f, text=val, variable=self._approval_var, value=val,
                bg=C["bg_main"], fg=C["text_primary"],
                activebackground=C["bg_main"], activeforeground=C["text_primary"],
                selectcolor=C["bg_surface"],
                font=FONT_DEFAULT, command=_update_dot,
            ).pack(side="left", padx=(0, 16))

        self._approval_dot = tk.Label(appr_f, text="  ●  ",
                                       bg=C["accent_success"] if rec.has_approval == "Так"
                                          else C["accent_danger"],
                                       fg="white", font=FONT_BOLD)
        self._approval_dot.pack(side="left", padx=6)
        row += 2

        tk.Frame(body, bg=C["bg_main"], height=16).grid(
            row=row, column=0, columnspan=2)

        # Bottom buttons
        btn_bar = tk.Frame(self.win, bg=C["bg_header"], height=46)
        btn_bar.grid(row=2, column=0, columnspan=2, sticky="ew")
        btn_bar.grid_propagate(False)

        make_button(btn_bar, "Скасувати",
                    bg=C["bg_surface"], fg=C["text_muted"],
                    activebackground=C["bg_surface_alt"],
                    activeforeground=C["text_primary"],
                    font=FONT_DEFAULT, padx=14, pady=5,
                    command=self.win.destroy).pack(side="right", padx=8, pady=8)

        save_label = "Зберегти зміни" if self.is_edit else "Додати координатора"
        make_button(btn_bar, save_label,
                    bg=C["accent_warning"] if self.is_edit else C["accent"],
                    activebackground=C["accent_soft"],
                    font=FONT_BOLD, padx=14, pady=5,
                    command=self._submit).pack(side="right", padx=(0, 4), pady=8)

    def _submit(self) -> None:
        enterprise = self.e_enterprise.get().strip()
        full_name  = self.e_full_name.get().strip()
        if not enterprise or not full_name:
            messagebox.showwarning(
                "Обов'язкові поля",
                "Заповніть назву підприємства та ПІБ координатора",
                parent=self.win)
            return

        appointed = self.e_appointed.get().strip()
        if appointed in ("дд.мм.рррр", ""):
            appointed = "—"
        elif not is_valid_date(appointed):
            messagebox.showwarning(
                "Помилка дати",
                "Неправильний формат дати (очікується дд.мм.рррр)",
                parent=self.win)
            return

        rec_id = self.record.id if self.record else str(uuid.uuid4())[:8]
        result = CoordinatorRecord(
            id=rec_id,
            enterprise=enterprise,
            department=self.e_department.get().strip(),
            location=self.e_location.get().strip(),
            is_staff_unit=self.cb_staff.get(),
            is_concurrent=self.cb_concur.get(),
            main_position=self.e_main_position.get().strip(),
            full_name=full_name,
            phone=self.e_phone.get().strip(),
            appointed_date=appointed,
            order_number=self.e_order.get().strip(),
            has_approval=self._approval_var.get(),
        )
        self.on_save(result)
        self.win.destroy()


# ---------------------------------------------------------------------------
#  Картка підприємства (enterprise card) — компактний блок зі списком людей
# ---------------------------------------------------------------------------

class EnterpriseCard:
    """
    Один рядок-акордеон: назва підприємства + розгортається список координаторів.
    """

    def __init__(
        self,
        parent: tk.Misc,
        enterprise: str,
        records: list[CoordinatorRecord],
        on_edit:   Callable[[CoordinatorRecord], None],
        on_delete: Callable[[CoordinatorRecord], None],
        on_add:    Callable[[str], None],
        row: int,
    ) -> None:
        C   = COLORS
        self.parent     = parent
        self.enterprise = enterprise
        self.records    = records
        self.on_edit    = on_edit
        self.on_delete  = on_delete
        self.on_add     = on_add
        self._expanded  = True

        # ── Header row ──────────────────────────────────────────────────
        self.hdr = tk.Frame(parent, bg=C["bg_surface_alt"], cursor="hand2")
        self.hdr.grid(row=row, column=0, sticky="ew",
                      padx=12, pady=(8, 0))
        self.hdr.columnconfigure(1, weight=1)
        row += 1

        # accent strip
        tk.Frame(self.hdr, bg=C["accent"], width=5).grid(
            row=0, column=0, sticky="ns")

        name_f = tk.Frame(self.hdr, bg=C["bg_surface_alt"])
        name_f.grid(row=0, column=1, sticky="ew", padx=14, pady=8)
        tk.Label(name_f, text=enterprise,
                 bg=C["bg_surface_alt"], fg=C["text_primary"],
                 font=("Arial", 10, "bold")).pack(side="left")

        cnt_txt = f"  {len(records)} координ."
        tk.Label(name_f, text=cnt_txt,
                 bg=C["bg_surface_alt"], fg=C["accent_muted"],
                 font=FONT_SMALL).pack(side="left", padx=6)

        btn_area = tk.Frame(self.hdr, bg=C["bg_surface_alt"])
        btn_area.grid(row=0, column=2, padx=8, pady=4)

        make_button(btn_area, "+ Додати",
                    bg=C["accent"], fg="white",
                    activebackground=C["accent_soft"],
                    font=FONT_SMALL_BOLD, padx=10, pady=3,
                    command=lambda: on_add(enterprise)).pack(side="left", padx=(0, 4))

        self.toggle_btn = make_button(
            btn_area, "▼",
            bg=C["bg_surface_alt"], fg=C["text_muted"],
            activebackground=C["bg_surface"],
            font=FONT_DEFAULT, padx=8, pady=3,
            command=self._toggle)
        self.toggle_btn.pack(side="left")

        # ── Body: table of coordinators ─────────────────────────────────
        self.body = tk.Frame(parent, bg=C["bg_main"])
        self.body.grid(row=row, column=0, sticky="ew", padx=12, pady=(0, 4))
        self.body.columnconfigure(0, weight=1)
        self._body_row = row

        self._build_table()

    def _build_table(self) -> None:
        C = COLORS
        for w in self.body.winfo_children():
            w.destroy()

        if not self.records:
            tk.Label(self.body, text="  Координаторів немає. Натисніть «+ Додати».",
                     bg=C["bg_main"], fg=C["text_subtle"],
                     font=("Arial", 8, "italic")).pack(anchor="w", padx=16, pady=8)
            return

        cols = [
            ("ПІБ",                        220, "full_name"),
            ("Підрозділ",                  170, "department"),
            ("Телефон",                    110, "phone"),
            ("Посада",                     180, "main_position"),
            ("Дата прийняття",             100, "appointed_date"),
            ("Наказ №",                     80, "order_number"),
            ("Штатна",                      68, "is_staff_unit"),
            ("Сумісн.",                     68, "is_concurrent"),
            ("Погодження",                  90, "has_approval"),
        ]

        # Table header
        th = tk.Frame(self.body, bg=C["bg_surface"])
        th.pack(fill="x")
        for label, w, _ in cols:
            tk.Label(th, text=label, bg=C["bg_surface"], fg=C["text_subtle"],
                     font=FONT_TINY, width=w // 7, anchor="w").pack(
                side="left", padx=(8, 0), pady=4)
        tk.Label(th, text="Дії", bg=C["bg_surface"], fg=C["text_subtle"],
                 font=FONT_TINY, width=8).pack(side="right", padx=8, pady=4)

        # Rows
        for i, rec in enumerate(self.records):
            row_bg = C["row_even"] if i % 2 == 0 else C["row_odd"]
            rf = tk.Frame(self.body, bg=row_bg)
            rf.pack(fill="x")

            for _, w, attr in cols:
                val = getattr(rec, attr, "—") or "—"
                if attr == "has_approval":
                    dot_color = (C["accent_success"] if val == "Так"
                                 else C["accent_danger"])
                    dot_frame = tk.Frame(rf, bg=row_bg,
                                         width=w // 7 * 7, height=24)
                    dot_frame.pack_propagate(False)
                    dot_frame.pack(side="left", padx=(8, 0))
                    tk.Label(dot_frame, text="  ●",
                              bg=row_bg, fg=dot_color,
                              font=("Arial", 11)).pack(side="left")
                    tk.Label(dot_frame, text=val,
                              bg=row_bg, fg=dot_color,
                              font=FONT_TINY).pack(side="left", padx=2)
                elif attr in ("is_staff_unit", "is_concurrent"):
                    color = (C["accent_success"] if val == "Так"
                             else C["text_subtle"])
                    tk.Label(rf, text=val, bg=row_bg, fg=color,
                              font=FONT_SMALL,
                              width=w // 7, anchor="w").pack(
                        side="left", padx=(8, 0), pady=3)
                else:
                    display = (val[:22] + "…") if len(val) > 22 else val
                    tk.Label(rf, text=display, bg=row_bg,
                              fg=C["text_primary"], font=FONT_SMALL,
                              width=w // 7, anchor="w").pack(
                        side="left", padx=(8, 0), pady=3)

            # action buttons
            act = tk.Frame(rf, bg=row_bg)
            act.pack(side="right", padx=8, pady=2)

            make_button(act, "✏", bg=row_bg, fg=C["accent_warning"],
                        activebackground=C["bg_surface_alt"],
                        activeforeground=C["accent_warning"],
                        font=("Arial", 10), padx=4, pady=1,
                        command=lambda r=rec: self.on_edit(r)).pack(side="left")
            make_button(act, "✕", bg=row_bg, fg=C["accent_danger"],
                        activebackground=C["bg_surface_alt"],
                        activeforeground=C["accent_danger"],
                        font=("Arial", 10), padx=4, pady=1,
                        command=lambda r=rec: self.on_delete(r)).pack(side="left")

    def refresh(self, records: list[CoordinatorRecord]) -> None:
        self.records = records
        self._build_table()
        # оновлюємо лічильник
        for child in self.hdr.winfo_children():
            if isinstance(child, tk.Frame):
                for sub in child.winfo_children():
                    if isinstance(sub, tk.Label) and "координ" in (sub.cget("text") or ""):
                        sub.configure(text=f"  {len(records)} координ.")
                        return

    def _toggle(self) -> None:
        self._expanded = not self._expanded
        self.toggle_btn.configure(text="▼" if self._expanded else "▶")
        if self._expanded:
            self.body.grid()
        else:
            self.body.grid_remove()


# ---------------------------------------------------------------------------
#  Головна сторінка: RiskCoordinatorsPage
# ---------------------------------------------------------------------------

class RiskCoordinatorsPage(tk.Frame):
    """
    Сторінка «Ризик координатори» — телефонна книга координаторів по підприємствах.

    Структура:
        Toolbar (пошук, фільтр, кнопки)
        ├── Scrollable area
        │   └── EnterpriseCard × N  (акордеон)
        └── StatusBar
    """

    DATA_FILE = COORDS_DATA_FILE

    _EXPORT_HEADERS = [
        "ID", "Підприємство", "Управління/Підрозділ", "Місцезнаходження",
        "Штатна одиниця", "Суміщення", "Основна посада",
        "ПІБ", "Телефон", "Дата прийняття", "Наказ №", "Погодження",
    ]

    def __init__(self, master: tk.Misc, **kwargs) -> None:
        super().__init__(master, bg=COLORS["bg_main"], **kwargs)
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)

        self._records:  list[CoordinatorRecord] = []
        self._cards:    dict[str, EnterpriseCard] = {}   # enterprise → card
        self._filtered: list[CoordinatorRecord] = []

        self._build_toolbar()
        self._build_scroll_area()
        self._build_statusbar()
        self._load_data()
        self._schedule_autosave()

    # ── Layout ──────────────────────────────────────────────────────────────

    def _build_toolbar(self) -> None:
        C = COLORS
        tb = tk.Frame(self, bg=C["bg_header"], height=56)
        tb.grid(row=0, column=0, sticky="ew")
        tb.columnconfigure(2, weight=1)
        tb.grid_propagate(False)

        tk.Label(tb, text="РИЗИК КООРДИНАТОРИ",
                 bg=C["bg_header"], fg=C["accent_muted"],
                 font=FONT_TITLE).grid(row=0, column=0, padx=20, pady=14, sticky="w")

        # search
        sf = tk.Frame(tb, bg=C["bg_header"])
        sf.grid(row=0, column=1, sticky="w", padx=(0, 12))
        tk.Label(sf, text="Пошук:", bg=C["bg_header"],
                 fg=C["text_muted"], font=FONT_SMALL).pack(side="left", padx=(0, 5))
        self._search_var = tk.StringVar()
        self._search_var.trace_add("write", lambda *_: self._apply_filter())
        tk.Entry(sf, textvariable=self._search_var,
                 bg=C["bg_input"], fg=C["text_primary"],
                 insertbackground=C["text_primary"],
                 relief="flat", bd=2, font=FONT_DEFAULT,
                 width=28).pack(side="left", ipady=2)
        make_button(sf, "✕",
                    bg=C["bg_header"], fg=C["text_muted"],
                    activebackground=C["bg_surface"],
                    font=FONT_DEFAULT, padx=6, pady=1,
                    command=lambda: self._search_var.set("")).pack(side="left", padx=2)

        # filter approval
        ff = tk.Frame(tb, bg=C["bg_header"])
        ff.grid(row=0, column=2, sticky="w")
        tk.Label(ff, text="Погодження:", bg=C["bg_header"],
                 fg=C["text_muted"], font=FONT_SMALL).pack(side="left", padx=(0, 5))
        self._filter_appr = make_dark_combo(ff, values=["Всі", "Так", "Нi"], width=8)
        self._filter_appr.set("Всі")
        self._filter_appr.pack(side="left")
        self._filter_appr.bind("<<ComboboxSelected>>", lambda _: self._apply_filter())

        # action buttons (right side)
        bf = tk.Frame(tb, bg=C["bg_header"])
        bf.grid(row=0, column=3, sticky="e", padx=12)

        make_button(bf, "+ Новий координатор",
                    bg=C["accent"], fg="white",
                    activebackground=C["accent_soft"],
                    font=FONT_BOLD, padx=14, pady=4,
                    command=self._add_new).pack(side="left", padx=(0, 6))
        make_button(bf, "Експорт CSV",
                    bg=C["bg_surface"], fg=C["text_primary"],
                    activebackground=C["bg_surface_alt"],
                    font=FONT_SMALL, padx=10, pady=4,
                    command=self._export_csv).pack(side="left", padx=(0, 4))
        if pd:
            make_button(bf, "Експорт Excel",
                        bg=C["accent_success"], fg="white",
                        activebackground="#16a34a",
                        font=FONT_SMALL, padx=10, pady=4,
                        command=self._export_excel).pack(side="left", padx=(0, 4))
        make_button(bf, "Імпорт JSON",
                    bg=C["bg_surface"], fg=C["text_primary"],
                    activebackground=C["bg_surface_alt"],
                    font=FONT_SMALL, padx=10, pady=4,
                    command=self._import_json).pack(side="left")

    def _build_scroll_area(self) -> None:
        C = COLORS
        outer = tk.Frame(self, bg=C["bg_main"])
        outer.grid(row=1, column=0, sticky="nsew")
        outer.rowconfigure(0, weight=1)
        outer.columnconfigure(0, weight=1)

        self._canvas = tk.Canvas(outer, bg=C["bg_main"], highlightthickness=0)
        sb = ttk.Scrollbar(outer, orient="vertical", command=self._canvas.yview)
        self._canvas.configure(yscrollcommand=sb.set)
        self._canvas.grid(row=0, column=0, sticky="nsew")
        sb.grid(row=0, column=1, sticky="ns")

        self._scroll_frame = tk.Frame(self._canvas, bg=C["bg_main"])
        self._scroll_frame.columnconfigure(0, weight=1)
        self._cw = self._canvas.create_window(
            (0, 0), window=self._scroll_frame, anchor="nw")

        def _conf(_: object) -> None:
            self._canvas.configure(scrollregion=self._canvas.bbox("all"))
            self._canvas.itemconfig(self._cw,
                                     width=self._canvas.winfo_width())

        self._scroll_frame.bind("<Configure>", _conf)
        self._canvas.bind("<Configure>", _conf)
        _scroll_mgr.attach(self._canvas)

    def _build_statusbar(self) -> None:
        C = COLORS
        sb = tk.Frame(self, bg=C["bg_header"], height=24)
        sb.grid(row=2, column=0, sticky="ew")
        sb.grid_propagate(False)

        self._status_lbl = tk.Label(
            sb, text="Готово", bg=C["bg_header"],
            fg=C["text_muted"], font=FONT_TINY, padx=10)
        self._status_lbl.pack(side="left", pady=3)

        self._stats_lbl = tk.Label(
            sb, text="", bg=C["bg_header"],
            fg=C["text_subtle"], font=FONT_TINY, padx=10)
        self._stats_lbl.pack(side="right", pady=3)

    # ── Data ────────────────────────────────────────────────────────────────

    def _load_data(self) -> None:
        self._records.clear()
        if not os.path.exists(self.DATA_FILE):
            self._rebuild_cards()
            return
        try:
            with open(self.DATA_FILE, "r", encoding="utf-8") as f:
                raw = json.load(f)
            if not isinstance(raw, list):
                raise ValueError("Очікується список")
            for item in raw:
                if isinstance(item, dict):
                    self._records.append(CoordinatorRecord.from_dict(item))
                elif isinstance(item, (list, tuple)):
                    self._records.append(CoordinatorRecord.from_list(list(item)))
        except (json.JSONDecodeError, ValueError, OSError) as e:
            messagebox.showerror("Помилка завантаження координаторів", str(e))
        self._filtered = list(self._records)
        self._rebuild_cards()
        self._update_stats()

    def _save_data(self) -> None:
        try:
            with open(self.DATA_FILE, "w", encoding="utf-8") as f:
                json.dump([asdict(r) for r in self._records],
                          f, ensure_ascii=False, indent=2)
        except OSError as e:
            messagebox.showerror("Помилка збереження", str(e))

    def save(self) -> None:
        self._save_data()

    def save_before_exit(self) -> None:
        self._save_data()

    def _schedule_autosave(self) -> None:
        try:
            self._save_data()
            self._status_lbl.configure(
                text=f"Автозбережено о {datetime.now().strftime('%H:%M:%S')}")
        except OSError:
            self._status_lbl.configure(text="Помилка автозбереження")
        self.after(30_000, self._schedule_autosave)

    # ── Rendering ───────────────────────────────────────────────────────────

    def _rebuild_cards(self) -> None:
        """Перебудовує всі EnterpriseCard на основі self._filtered."""
        for w in self._scroll_frame.winfo_children():
            w.destroy()
        self._cards.clear()

        # Групуємо по підприємству, зберігаємо порядок
        groups: dict[str, list[CoordinatorRecord]] = {}
        for rec in self._filtered:
            groups.setdefault(rec.enterprise, []).append(rec)

        if not groups:
            C = COLORS
            tk.Label(self._scroll_frame,
                     text="Координаторів не знайдено.\nНатисніть «+ Новий координатор».",
                     bg=C["bg_main"], fg=C["text_subtle"],
                     font=("Arial", 11), justify="center").grid(
                row=0, column=0, pady=60)
            self._update_stats()
            return

        row = 0
        for enterprise, recs in groups.items():
            card = EnterpriseCard(
                parent=self._scroll_frame,
                enterprise=enterprise,
                records=recs,
                on_edit=self._edit_record,
                on_delete=self._delete_record,
                on_add=self._add_for_enterprise,
                row=row,
            )
            self._cards[enterprise] = card
            row += 2  # header + body each occupy 1 grid row

        self._update_stats()

    def _update_stats(self) -> None:
        total      = len(self._records)
        approved   = sum(1 for r in self._records if r.has_approval == "Так")
        enterprises = len({r.enterprise for r in self._records})
        self._stats_lbl.configure(
            text=(f"Підприємств: {enterprises}   "
                  f"Координаторів: {total}   "
                  f"З погодженням: {approved}   "
                  f"Без погодження: {total - approved}"))

    # ── Filter ──────────────────────────────────────────────────────────────

    def _apply_filter(self) -> None:
        q     = self._search_var.get().strip().lower()
        appr  = self._filter_appr.get()
        self._filtered = [
            r for r in self._records
            if (not q or q in (
                r.enterprise + r.full_name + r.department + r.location +
                r.phone + r.main_position + r.order_number
            ).lower())
            and (appr == "Всі" or r.has_approval == appr)
        ]
        self._rebuild_cards()

    # ── CRUD ────────────────────────────────────────────────────────────────

    def _add_new(self) -> None:
        CoordinatorFormDialog(
            parent_root=self.winfo_toplevel(),
            on_save=self._on_form_save,
        )

    def _add_for_enterprise(self, enterprise: str) -> None:
        CoordinatorFormDialog(
            parent_root=self.winfo_toplevel(),
            on_save=self._on_form_save,
            default_enterprise=enterprise,
        )

    def _on_form_save(self, rec: CoordinatorRecord) -> None:
        # Перевіряємо — новий чи редагування
        for i, r in enumerate(self._records):
            if r.id == rec.id:
                self._records[i] = rec
                break
        else:
            self._records.append(rec)
        self._save_data()
        self._apply_filter()
        _show_toast(self, "Збережено")

    def _edit_record(self, rec: CoordinatorRecord) -> None:
        CoordinatorFormDialog(
            parent_root=self.winfo_toplevel(),
            on_save=self._on_form_save,
            record=rec,
        )

    def _delete_record(self, rec: CoordinatorRecord) -> None:
        if not messagebox.askyesno(
            "Підтвердження",
            f"Видалити координатора «{rec.full_name}»?\n"
            f"Підприємство: {rec.enterprise}",
            parent=self.winfo_toplevel()):
            return
        self._records = [r for r in self._records if r.id != rec.id]
        self._save_data()
        self._apply_filter()
        _show_toast(self, "Видалено")

    # ── Export / Import ─────────────────────────────────────────────────────

    def _export_csv(self) -> None:
        if not self._records:
            messagebox.showinfo("Експорт", "Немає записів"); return
        path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV файли", "*.csv")],
            title="Зберегти координаторів як CSV")
        if not path:
            return
        try:
            with open(path, "w", newline="", encoding="utf-8-sig") as f:
                w = csv.writer(f)
                w.writerow(self._EXPORT_HEADERS)
                for r in self._records:
                    w.writerow(r.to_list())
            _show_toast(self, "CSV збережено")
        except OSError as e:
            messagebox.showerror("Помилка", str(e))

    def _export_excel(self) -> None:
        if not pd:
            messagebox.showwarning("Excel", "Встановіть pandas та openpyxl"); return
        if not self._records:
            messagebox.showinfo("Експорт", "Немає записів"); return
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel файли", "*.xlsx")],
            title="Зберегти координаторів як Excel")
        if not path:
            return
        try:
            rows = [r.to_list() for r in self._records]
            df   = pd.DataFrame(rows, columns=self._EXPORT_HEADERS)
            with pd.ExcelWriter(path, engine="openpyxl") as writer:
                df.to_excel(writer, index=False, sheet_name="Координатори")
                ws = writer.sheets["Координатори"]
                for col_cells in ws.columns:
                    mx = max(len(str(c.value or "")) for c in col_cells)
                    ws.column_dimensions[
                        col_cells[0].column_letter].width = min(mx + 4, 50)
            _show_toast(self, "Excel збережено")
        except (OSError, Exception) as e:
            messagebox.showerror("Помилка", str(e))

    def _import_json(self) -> None:
        path = filedialog.askopenfilename(
            filetypes=[("JSON файли", "*.json")],
            title="Імпорт координаторів з JSON")
        if not path:
            return
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            if not isinstance(data, list):
                raise ValueError("Файл повинен містити список записів")
            added = 0
            existing_ids = {r.id for r in self._records}
            for item in data:
                if isinstance(item, dict):
                    rec = CoordinatorRecord.from_dict(item)
                elif isinstance(item, (list, tuple)):
                    rec = CoordinatorRecord.from_list(list(item))
                else:
                    continue
                # new unique id if collision
                if not rec.id or rec.id in existing_ids:
                    rec = CoordinatorRecord(**{**asdict(rec),
                                               "id": str(uuid.uuid4())[:8]})
                self._records.append(rec)
                existing_ids.add(rec.id)
                added += 1
            self._save_data()
            self._apply_filter()
            _show_toast(self, f"Імпортовано: {added}")
        except (json.JSONDecodeError, ValueError, OSError) as e:
            messagebox.showerror("Помилка імпорту", str(e))


# =============================================================================
#  МОДУЛЬ: РИЗИК АПЕТИТ
# =============================================================================

# ── Кольори рівнів індикаторів ────────────────────────────────────────────
RA_COLORS = {
    "Green":  "#10B981",
    "Yellow": "#F59E0B",
    "Orange": "#F97316",
    "Red":    "#EF4444",
    "—":      "#556871",
}

RA_LABELS = {
    "Green":  "Зелений",
    "Yellow": "Жовтий",
    "Orange": "Помаранчевий",
    "Red":    "Червоний",
    "—":      "Не визначено",
}

# ── Каталог чотирьох напрямків ─────────────────────────────────────────────
RA_DIRECTIONS = [
    {
        "key":   "strategic",
        "title": "Стратегічні\nризики",
        "icon":  "🎯",
        "color": "#4F46E5",
        "desc":  "SR1–SR6 · 14 індикаторів",
    },
    {
        "key":   "operational",
        "title": "Операційні\nризики",
        "icon":  "⚙",
        "color": "#F59E0B",
        "desc":  "OR · Виробництво та процеси",
    },
    {
        "key":   "financial",
        "title": "Фінансові\nризики",
        "icon":  "💰",
        "color": "#10B981",
        "desc":  "FR · Ліквідність та P&L",
    },
    {
        "key":   "compliance",
        "title": "Комплаєнс\nризики",
        "icon":  "📋",
        "color": "#A855F7",
        "desc":  "CR · Регуляторні вимоги",
    },
]

# ── Специфікація стратегічних індикаторів ─────────────────────────────────
STRATEGIC_INDICATORS: list[dict] = [
    # ─────────────────────── SR1 ──────────────────────────────────────────
    {
        "code": "SR 1.1", "group": "SR1", "name": "Перегляд бізнес-плану",
        "mode": "steps3",
        "desc": ("Порівняння фактичних показників Ф2 з плановими.\n"
                 "Перевіряються: прибуток, видатки, чистий прибуток."),
        "fields": [
            ("plan_profit",   "Плановий прибуток (грн)"),
            ("fact_profit",   "Фактичний прибуток (грн)"),
            ("plan_expense",  "Планові видатки (грн)"),
            ("fact_expense",  "Фактичні видатки (грн)"),
            ("plan_income",   "Плановий чистий дохід (грн)"),
            ("fact_income",   "Фактичний чистий дохід (грн)"),
            ("plan_net",      "Плановий чистий прибуток (грн)"),
            ("fact_net",      "Фактичний чистий прибуток (грн)"),
        ],
        "thresholds": {
            "Green": "Всі 3 кроки OK",
            "Yellow": "1 крок не OK",
            "Orange": "2 кроки не OK",
            "Red": "3 кроки не OK",
        },
    },
    {
        "code": "SR 1.2", "group": "SR1", "name": "Невиконання планових поставок",
        "mode": "pct_income",
        "desc": "Загальна заборгованість (борг + штрафи) відносно чистого доходу.",
        "fields": [
            ("debt",    "Сума заборгованості (грн)"),
            ("penalty", "Штрафні санкції (грн)"),
            ("income",  "Чистий дохід підприємства (грн)"),
        ],
        "thresholds": {
            "Green": "< 1% доходу",
            "Yellow": "= 1% доходу",
            "Orange": "> 1% доходу",
            "Red": "> 2% доходу",
        },
    },
    {
        "code": "SR 1.3", "group": "SR1", "name": "Показник фінансового результату",
        "mode": "pct_income",
        "desc": "Порівняння збитку (рядок 2095 < 0 Ф2) з чистим доходом.",
        "fields": [
            ("loss",   "Збиток (рядок 2095, грн; 0 якщо прибуток)"),
            ("income", "Чистий дохід підприємства (грн)"),
        ],
        "thresholds": {
            "Green": "< 1% доходу",
            "Yellow": "= 1% доходу",
            "Orange": "> 1% доходу",
            "Red": "> 2% доходу",
        },
    },
    # ─────────────────────── SR2 ──────────────────────────────────────────
    {
        "code": "SR 2.1", "group": "SR2", "name": "Впровадження документів",
        "mode": "count4",
        "desc": "Кількість документів, що мали бути впроваджені, але не впроваджені.",
        "fields": [
            ("required",     "Кількість документів до впровадження"),
            ("implemented",  "Кількість фактично впроваджених"),
        ],
        "thresholds": {
            "Green": "0 не впроваджено",
            "Yellow": "1 не впроваджено",
            "Orange": "2 не впроваджено",
            "Red": "3+ не впроваджено",
        },
    },
    {
        "code": "SR 2.2", "group": "SR2", "name": "Виконання документів",
        "mode": "count4",
        "desc": "Кількість документів, що мали бути виконані, але не виконані.",
        "fields": [
            ("required",   "Кількість документів до виконання"),
            ("executed",   "Кількість фактично виконаних"),
        ],
        "thresholds": {
            "Green": "0 не виконано",
            "Yellow": "1 не виконано",
            "Orange": "2 не виконано",
            "Red": "3+ не виконано",
        },
    },
    # ─────────────────────── SR3 ──────────────────────────────────────────
    {
        "code": "SR 3.1", "group": "SR3", "name": "Дії без погодження",
        "mode": "count3",
        "desc": "Кількість правочинів, укладених без належного погодження.",
        "fields": [
            ("unpproved_deals", "Кількість правочинів без погодження"),
        ],
        "thresholds": {
            "Green": "0 правочинів",
            "Yellow": "1 правочин",
            "Red": "2+ правочини",
        },
    },
    {
        "code": "SR 3.2", "group": "SR3", "name": "Результат зовнішнього аудиту",
        "mode": "count4",
        "desc": "Кількість модифікацій / застережень / відхилень у аудиторському висновку.",
        "fields": [
            ("modifications", "Кількість модифікацій/застережень"),
        ],
        "thresholds": {
            "Green": "0",
            "Yellow": "1",
            "Orange": "2",
            "Red": "3+",
        },
    },
    {
        "code": "SR 3.3", "group": "SR3", "name": "Зміна керівництва",
        "mode": "count4",
        "desc": "Кількість звільнень рівня CEO та CEO-1.",
        "fields": [
            ("ceo_changes", "Звільнень CEO рівня"),
            ("ceo1_changes", "Звільнень CEO-1 рівня"),
        ],
        "thresholds": {
            "Green": "0",
            "Yellow": "1",
            "Orange": "2",
            "Red": "3+",
        },
    },
    # ─────────────────────── SR4 ──────────────────────────────────────────
    {
        "code": "SR 4.1", "group": "SR4", "name": "Припинення міжнародних проектів",
        "mode": "count4",
        "desc": "Кількість припинених міжнародних договорів.",
        "fields": [
            ("terminated_contracts", "Кількість припинених договорів"),
            ("total_contracts",      "Загальна кількість міжнародних договорів"),
        ],
        "thresholds": {
            "Green": "0",
            "Yellow": "1",
            "Orange": "2",
            "Red": "3+",
        },
    },
    # ─────────────────────── SR5 ──────────────────────────────────────────
    {
        "code": "SR 5.1", "group": "SR5", "name": "Військові дії / Перевороти",
        "mode": "count4",
        "desc": "Кількість договорів з країнами, де зафіксовано військові дії / перевороти.",
        "fields": [
            ("affected_contracts", "Договори з ураженими країнами"),
            ("unfulfilled_amount", "Сума невиконаних зобов'язань (грн)"),
        ],
        "thresholds": {
            "Green": "0",
            "Yellow": "1",
            "Orange": "2",
            "Red": "3+",
        },
    },
    {
        "code": "SR 5.2", "group": "SR5", "name": "Форс-мажор",
        "mode": "count4",
        "desc": "Кількість договорів, де зафіксовано форс-мажорні обставини.",
        "fields": [
            ("fm_contracts",      "Кількість договорів з форс-мажором"),
            ("unfulfilled_amount", "Сума невиконаних зобов'язань (грн)"),
        ],
        "thresholds": {
            "Green": "0",
            "Yellow": "1",
            "Orange": "2",
            "Red": "3+",
        },
    },
    # ─────────────────────── SR6 ──────────────────────────────────────────
    {
        "code": "SR 6.1", "group": "SR6", "name": "Призначення з порушенням процедур",
        "mode": "count4",
        "desc": "Кількість призначень, здійснених без погодження або з порушенням процедур.",
        "fields": [
            ("violation_appts", "Призначення без погодження"),
        ],
        "thresholds": {
            "Green": "0",
            "Yellow": "1",
            "Orange": "2",
            "Red": "3+",
        },
    },
]

# Операційні, Фінансові, Комплаєнс — базова структура (розширюється аналогічно)
OPERATIONAL_INDICATORS: list[dict] = [
    {
        "code": "OR 1.1", "group": "OR1", "name": "Простої обладнання",
        "mode": "count4",
        "desc": "Кількість непланових простоїв виробничого обладнання за квартал.",
        "fields": [
            ("unplanned_stops", "Кількість непланових простоїв"),
            ("total_hours",     "Загальний час простоїв (год)"),
        ],
        "thresholds": {"Green": "0", "Yellow": "1–2", "Orange": "3–5", "Red": "6+"},
    },
    {
        "code": "OR 1.2", "group": "OR1", "name": "Виробничі аварії",
        "mode": "count4",
        "desc": "Кількість зафіксованих виробничих аварій та інцидентів.",
        "fields": [
            ("accidents", "Кількість аварій / інцидентів"),
            ("injured",   "Кількість постраждалих"),
        ],
        "thresholds": {"Green": "0", "Yellow": "1", "Orange": "2", "Red": "3+"},
    },
    {
        "code": "OR 2.1", "group": "OR2", "name": "Порушення SLA",
        "mode": "pct_income",
        "desc": "Частка порушень внутрішніх SLA у загальній кількості запитів.",
        "fields": [
            ("sla_breaches", "Кількість порушень SLA"),
            ("total_requests", "Загальна кількість запитів"),
        ],
        "thresholds": {"Green": "< 1%", "Yellow": "1–3%", "Orange": "3–5%", "Red": "> 5%"},
    },
]

FINANCIAL_INDICATORS: list[dict] = [
    {
        "code": "FR 1.1", "group": "FR1", "name": "Коефіцієнт ліквідності",
        "mode": "ratio",
        "desc": "Поточна ліквідність: оборотні активи / поточні зобов'язання.",
        "fields": [
            ("current_assets",      "Оборотні активи (грн)"),
            ("current_liabilities", "Поточні зобов'язання (грн)"),
        ],
        "thresholds": {"Green": "> 2.0", "Yellow": "1.5–2.0", "Orange": "1.0–1.5", "Red": "< 1.0"},
    },
    {
        "code": "FR 1.2", "group": "FR1", "name": "Рентабельність активів (ROA)",
        "mode": "ratio",
        "desc": "Чистий прибуток / Загальні активи × 100%.",
        "fields": [
            ("net_profit",    "Чистий прибуток (грн)"),
            ("total_assets",  "Загальні активи (грн)"),
        ],
        "thresholds": {"Green": "> 5%", "Yellow": "2–5%", "Orange": "0–2%", "Red": "< 0%"},
    },
    {
        "code": "FR 2.1", "group": "FR2", "name": "Кредитне навантаження",
        "mode": "ratio",
        "desc": "Співвідношення боргу до EBITDA.",
        "fields": [
            ("total_debt", "Загальний борг (грн)"),
            ("ebitda",     "EBITDA (грн)"),
        ],
        "thresholds": {"Green": "< 2×", "Yellow": "2–3×", "Orange": "3–4×", "Red": "> 4×"},
    },
]

COMPLIANCE_INDICATORS: list[dict] = [
    {
        "code": "CR 1.1", "group": "CR1", "name": "Регуляторні порушення",
        "mode": "count4",
        "desc": "Кількість виявлених порушень регуляторних вимог за квартал.",
        "fields": [
            ("violations", "Кількість порушень"),
            ("fines",      "Сума штрафів (грн)"),
        ],
        "thresholds": {"Green": "0", "Yellow": "1", "Orange": "2", "Red": "3+"},
    },
    {
        "code": "CR 1.2", "group": "CR1", "name": "Антикорупційні заходи",
        "mode": "count4",
        "desc": "Кількість виявлених корупційних інцидентів або підозр.",
        "fields": [
            ("incidents", "Кількість інцидентів / підозр"),
        ],
        "thresholds": {"Green": "0", "Yellow": "1", "Orange": "2", "Red": "3+"},
    },
    {
        "code": "CR 2.1", "group": "CR2", "name": "Судові позови",
        "mode": "pct_income",
        "desc": "Сума активних судових позовів відносно чистого доходу.",
        "fields": [
            ("lawsuits_amount", "Сума активних позовів (грн)"),
            ("income",          "Чистий дохід (грн)"),
        ],
        "thresholds": {"Green": "< 1%", "Yellow": "1–3%", "Orange": "3–5%", "Red": "> 5%"},
    },
]

DIRECTION_INDICATORS = {
    "strategic":   STRATEGIC_INDICATORS,
    "operational": OPERATIONAL_INDICATORS,
    "financial":   FINANCIAL_INDICATORS,
    "compliance":  COMPLIANCE_INDICATORS,
}


# =============================================================================
#  ОБЧИСЛЮВАЛЬНИЙ ENGINE
# =============================================================================

def _safe_float(val: str) -> float:
    """Безпечна конвертація рядка у float."""
    try:
        return float(str(val).replace(" ", "").replace(",", ".") or "0")
    except ValueError:
        return 0.0


def compute_indicator_level(spec: dict, values: dict[str, str]) -> tuple[str, str]:
    """
    Повертає (level, detail):
        level  — "Green" | "Yellow" | "Orange" | "Red" | "—"
        detail — рядок пояснення розрахунку
    """
    mode = spec.get("mode", "")

    if mode == "steps3":
        # SR1.1: перевіряємо 3 кроки
        plan_p  = _safe_float(values.get("plan_profit",  "0"))
        fact_p  = _safe_float(values.get("fact_profit",  "0"))
        plan_e  = _safe_float(values.get("plan_expense", "0"))
        fact_e  = _safe_float(values.get("fact_expense", "0"))
        plan_i  = _safe_float(values.get("plan_income",  "0"))
        fact_i  = _safe_float(values.get("fact_income",  "0"))
        plan_n  = _safe_float(values.get("plan_net",     "0"))
        fact_n  = _safe_float(values.get("fact_net",     "0"))

        failed  = 0
        details = []

        # Крок 1: факт. прибуток ≥ план або відхилення < 10%
        if plan_p != 0:
            dev1 = abs(fact_p - plan_p) / abs(plan_p) * 100
            ok1  = fact_p >= plan_p or dev1 < 10
        else:
            ok1 = True; dev1 = 0
        if not ok1:
            failed += 1
            details.append(f"Прибуток: відхилення {dev1:.1f}%")

        # Крок 2: видатки не > плану на 10% (без зростання доходу)
        if plan_e != 0:
            dev2   = (fact_e - plan_e) / abs(plan_e) * 100
            income_grew = (plan_i != 0 and fact_i > plan_i * 1.05)
            ok2    = dev2 <= 10 or income_grew
        else:
            ok2 = True; dev2 = 0
        if not ok2:
            failed += 1
            details.append(f"Видатки: перевищення {dev2:.1f}%")

        # Крок 3: чистий прибуток ≥ план або відхилення < 10%
        if plan_n != 0:
            dev3 = abs(fact_n - plan_n) / abs(plan_n) * 100
            ok3  = fact_n >= plan_n or dev3 < 10
        else:
            ok3 = True; dev3 = 0
        if not ok3:
            failed += 1
            details.append(f"Чист. прибуток: відхилення {dev3:.1f}%")

        levels = ["Green", "Yellow", "Orange", "Red"]
        detail = "; ".join(details) if details else "Всі кроки OK"
        return levels[min(failed, 3)], detail

    elif mode == "pct_income":
        keys    = [f["name"] if isinstance(f, dict) else f
                   for f in spec.get("fields", [])]
        field_keys = [f[0] for f in spec.get("fields", [])]
        # numerator = перший або сума перших двох (борг + штрафи)
        if "debt" in field_keys and "penalty" in field_keys:
            num = _safe_float(values.get("debt", "0")) + \
                  _safe_float(values.get("penalty", "0"))
        elif "loss" in field_keys:
            num = _safe_float(values.get("loss", "0"))
        elif "lawsuits_amount" in field_keys:
            num = _safe_float(values.get("lawsuits_amount", "0"))
        elif "sla_breaches" in field_keys:
            num = _safe_float(values.get("sla_breaches", "0"))
        else:
            num = _safe_float(values.get(field_keys[0], "0"))

        denom_key = next(
            (k for k in field_keys if k in ("income", "total_requests")), None)
        denom = _safe_float(values.get(denom_key, "0")) if denom_key else 0.0

        if denom == 0:
            return "—", "Знаменник = 0, розрахунок неможливий"

        pct = num / denom * 100
        detail = f"Частка = {pct:.2f}%"

        if pct < 1:
            return "Green",  detail
        elif pct == 1:
            return "Yellow", detail
        elif pct <= 2:
            return "Orange", detail
        else:
            return "Red",    detail

    elif mode == "ratio":
        field_keys = [f[0] for f in spec.get("fields", [])]
        num   = _safe_float(values.get(field_keys[0], "0"))
        denom = _safe_float(values.get(field_keys[1], "0"))
        if denom == 0:
            return "—", "Знаменник = 0"
        ratio = num / denom
        code  = spec["code"]
        detail = f"Коефіцієнт = {ratio:.3f}"

        # FR1.1 ліквідність
        if "FR 1.1" in code:
            if ratio > 2.0:    return "Green",  detail
            elif ratio >= 1.5: return "Yellow", detail
            elif ratio >= 1.0: return "Orange", detail
            else:              return "Red",    detail
        # FR1.2 ROA
        elif "FR 1.2" in code:
            pct = ratio * 100
            det = f"ROA = {pct:.2f}%"
            if pct > 5:    return "Green",  det
            elif pct >= 2: return "Yellow", det
            elif pct >= 0: return "Orange", det
            else:          return "Red",    det
        # FR2.1 Debt/EBITDA
        elif "FR 2.1" in code:
            if ratio < 2:   return "Green",  detail
            elif ratio < 3: return "Yellow", detail
            elif ratio < 4: return "Orange", detail
            else:           return "Red",    detail
        return "—", detail

    elif mode in ("count4", "count3"):
        field_keys = [f[0] for f in spec.get("fields", [])]
        # для "невпроваджено" = required - implemented
        if "required" in field_keys:
            impl_key = next(
                (k for k in field_keys if k in ("implemented", "executed")), None)
            required = _safe_float(values.get("required", "0"))
            done     = _safe_float(values.get(impl_key, "0")) if impl_key else 0
            count    = max(0, required - done)
            detail   = f"Не виконано: {int(count)} з {int(required)}"
        else:
            # просто сума всіх числових полів
            count  = sum(_safe_float(values.get(k, "0")) for k in field_keys)
            detail = f"Кількість: {int(count)}"

        if mode == "count3":
            if count == 0:   return "Green",  detail
            elif count == 1: return "Yellow", detail
            else:            return "Red",    detail
        else:
            if count == 0:   return "Green",  detail
            elif count == 1: return "Yellow", detail
            elif count == 2: return "Orange", detail
            else:            return "Red",    detail

    return "—", "Режим розрахунку не визначено"


def aggregate_level(levels: list[str]) -> str:
    """Повертає найгірший рівень зі списку."""
    order = ["Red", "Orange", "Yellow", "Green", "—"]
    for lvl in order:
        if lvl in levels:
            return lvl
    return "—"


# =============================================================================
#  UI-компоненти Risk Appetite
# =============================================================================

def _ra_level_badge(parent: tk.Misc, level: str, size: str = "normal") -> tk.Label:
    """Кольоровий бейдж рівня індикатора."""
    C    = COLORS
    bg   = RA_COLORS.get(level, C["border_soft"])
    txt  = f"  {RA_LABELS.get(level, '—')}  "
    font = FONT_BOLD if size == "normal" else FONT_SMALL_BOLD
    return tk.Label(parent, text=txt, bg=bg, fg="white", font=font, pady=2)


# =============================================================================
#  Форма введення даних одного індикатора
# =============================================================================

class IndicatorFormFrame(tk.Frame):
    """
    Розгортна форма введення даних для одного індикатора.
    Автоматично перераховує рівень при зміні будь-якого поля.
    """

    def __init__(
        self,
        parent: tk.Misc,
        spec:   dict,
        saved_values: dict[str, str],
        saved_notes:  str,
        on_change: Callable[[str, dict, str, str], None],  # code, values, level, notes
        row: int,
    ) -> None:
        C = COLORS
        super().__init__(parent, bg=C["bg_main"])
        self.spec         = spec
        self._on_change   = on_change
        self._expanded    = False
        self._vars:   dict[str, tk.StringVar] = {}
        self._note_var = tk.StringVar(value=saved_notes)

        # Precompute initial level
        self._level, self._detail = compute_indicator_level(spec, saved_values)

        self.grid(row=row, column=0, sticky="ew", padx=12, pady=(4, 0))
        self.columnconfigure(0, weight=1)
        self._build(saved_values)

    def _build(self, saved: dict[str, str]) -> None:
        C  = COLORS
        sp = self.spec

        # ── Header row ──────────────────────────────────────────────────
        hdr = tk.Frame(self, bg=C["bg_surface"], cursor="hand2")
        hdr.grid(row=0, column=0, sticky="ew")
        hdr.columnconfigure(2, weight=1)

        lvl_color = RA_COLORS.get(self._level, C["border_soft"])
        tk.Frame(hdr, bg=lvl_color, width=5).grid(row=0, column=0, sticky="ns")

        code_lbl = tk.Label(hdr, text=f"  {sp['code']}",
                             bg=C["bg_surface"], fg=C["accent_muted"],
                             font=FONT_SMALL_BOLD, width=9, anchor="w")
        code_lbl.grid(row=0, column=1, padx=(4, 0), pady=6)

        name_lbl = tk.Label(hdr, text=sp["name"],
                             bg=C["bg_surface"], fg=C["text_primary"],
                             font=FONT_BOLD, anchor="w")
        name_lbl.grid(row=0, column=2, sticky="ew", padx=6, pady=6)

        # level badge
        self._badge = tk.Label(
            hdr,
            text=f"  {RA_LABELS.get(self._level, '—')}  ",
            bg=lvl_color, fg="white", font=FONT_SMALL_BOLD, pady=2)
        self._badge.grid(row=0, column=3, padx=8, pady=6)

        self._strip = code_lbl  # keep ref to update color strip

        # toggle
        self._toggle_btn = make_button(
            hdr, "▶", bg=C["bg_surface"], fg=C["text_muted"],
            activebackground=C["bg_surface_alt"],
            font=FONT_DEFAULT, padx=8, pady=3,
            command=self._toggle)
        self._toggle_btn.grid(row=0, column=4, padx=(0, 6), pady=4)

        for widget in (hdr, code_lbl, name_lbl):
            widget.bind("<Button-1>", lambda _: self._toggle())

        # Store ref to strip frame (first child)
        self._hdr_strip = hdr.winfo_children()[0]

        # ── Body (hidden by default) ─────────────────────────────────────
        self._body = tk.Frame(self, bg=C["bg_surface_alt"])
        self._body.columnconfigure(0, weight=1)
        self._body.columnconfigure(1, weight=1)

        # description
        tk.Label(self._body, text=sp["desc"],
                 bg=C["bg_surface_alt"], fg=C["text_muted"],
                 font=("Arial", 8, "italic"),
                 justify="left", wraplength=680, anchor="w").grid(
            row=0, column=0, columnspan=2,
            sticky="ew", padx=14, pady=(8, 4))

        # input fields
        for fi, (key, label) in enumerate(sp.get("fields", [])):
            r = fi + 1
            col = fi % 2
            if fi % 2 == 0 and fi > 0:
                r = (fi // 2) + 1

            actual_row = fi + 1
            actual_col = 0
            if len(sp["fields"]) > 1:
                actual_row = (fi // 2) + 1
                actual_col = fi % 2

            tk.Label(self._body, text=label + ":",
                     bg=C["bg_surface_alt"], fg=C["text_subtle"],
                     font=FONT_SMALL, anchor="w").grid(
                row=actual_row * 2 - 1, column=actual_col,
                sticky="w", padx=(14, 4), pady=(6, 0))

            var = tk.StringVar(value=saved.get(key, ""))
            self._vars[key] = var

            ent = make_dark_entry(self._body)
            ent.grid(row=actual_row * 2, column=actual_col,
                     sticky="ew", padx=(14, 8), pady=(2, 0))
            ent.insert(0, saved.get(key, ""))
            var.trace_add("write", lambda *_, k=key: self._recalc())
            ent.configure(textvariable=var)

        max_fi_row = ((len(sp["fields"]) - 1) // 2 + 1) * 2 + 1

        # live result row
        res_row = max_fi_row + 1
        res_f   = tk.Frame(self._body, bg=C["bg_surface_alt"])
        res_f.grid(row=res_row, column=0, columnspan=2,
                   sticky="ew", padx=14, pady=(10, 4))
        tk.Label(res_f, text="Результат розрахунку:",
                 bg=C["bg_surface_alt"], fg=C["text_subtle"],
                 font=FONT_SMALL_BOLD).pack(side="left")
        self._result_badge = tk.Label(
            res_f,
            text=f"  {RA_LABELS.get(self._level, '—')}  ",
            bg=RA_COLORS.get(self._level, COLORS["border_soft"]),
            fg="white", font=FONT_SMALL_BOLD, pady=2)
        self._result_badge.pack(side="left", padx=8)
        self._detail_lbl = tk.Label(
            res_f, text=self._detail,
            bg=C["bg_surface_alt"], fg=C["text_muted"],
            font=("Arial", 8, "italic"))
        self._detail_lbl.pack(side="left", padx=4)

        # thresholds legend
        thr_f = tk.Frame(self._body, bg=C["bg_surface_alt"])
        thr_f.grid(row=res_row + 1, column=0, columnspan=2,
                   sticky="ew", padx=14, pady=(2, 4))
        tk.Label(thr_f, text="Шкала: ",
                 bg=C["bg_surface_alt"], fg=C["text_subtle"],
                 font=FONT_TINY).pack(side="left")
        for lvl, txt in sp.get("thresholds", {}).items():
            bg = RA_COLORS.get(lvl, COLORS["border_soft"])
            tk.Label(thr_f, text=f" {txt} ",
                     bg=bg, fg="white", font=FONT_TINY).pack(side="left", padx=2)

        # note
        note_row = res_row + 2
        tk.Label(self._body, text="Коментар / примітка:",
                 bg=C["bg_surface_alt"], fg=C["text_subtle"],
                 font=FONT_SMALL).grid(
            row=note_row, column=0, columnspan=2,
            sticky="w", padx=14, pady=(8, 0))
        note_entry = make_dark_entry(self._body)
        note_entry.grid(row=note_row + 1, column=0, columnspan=2,
                        sticky="ew", padx=14, pady=(2, 10))
        note_entry.configure(textvariable=self._note_var)
        self._note_var.trace_add("write", lambda *_: self._recalc())

    def _toggle(self) -> None:
        self._expanded = not self._expanded
        self._toggle_btn.configure(text="▼" if self._expanded else "▶")
        if self._expanded:
            self._body.grid(row=1, column=0, sticky="ew")
        else:
            self._body.grid_remove()

    def _recalc(self) -> None:
        vals          = {k: v.get() for k, v in self._vars.items()}
        level, detail = compute_indicator_level(self.spec, vals)
        self._level   = level
        self._detail  = detail

        bg = RA_COLORS.get(level, COLORS["border_soft"])
        lbl_txt = f"  {RA_LABELS.get(level, '—')}  "
        self._badge.configure(text=lbl_txt, bg=bg)
        self._result_badge.configure(text=lbl_txt, bg=bg)
        self._detail_lbl.configure(text=detail)
        # update strip color
        try:
            self._hdr_strip.configure(bg=bg)
        except tk.TclError:
            pass

        self._on_change(
            self.spec["code"], vals, level,
            self._note_var.get())

    def get_values(self) -> dict[str, str]:
        return {k: v.get() for k, v in self._vars.items()}

    def get_level(self) -> str:
        return self._level

    def get_note(self) -> str:
        return self._note_var.get()


# =============================================================================
#  Фрейм одного напрямку (Strategic / Operational / Financial / Compliance)
# =============================================================================

class RiskDirectionFrame(tk.Frame):
    """
    Сторінка одного напрямку ризику.
    Вкладки: Введення даних | Статистика (квартал/рік).
    Дашборд — компактна смужка одразу під header без відступів.
    """

    def __init__(
        self,
        parent:     tk.Misc,
        direction:  dict,
        period:     str,
        enterprise: str,
        saved_data: dict,
        all_data:   dict,                                   # весь словник dir_key→period→ent→data
        on_save:    Callable[[str, str, str, dict], None],
        on_back:    Callable[[], None],
    ) -> None:
        C = COLORS
        super().__init__(parent, bg=C["bg_main"])
        self.columnconfigure(0, weight=1)
        self.rowconfigure(2, weight=1)   # row 2 = notebook

        self._direction  = direction
        self._period     = period
        self._enterprise = enterprise
        self._saved_data = saved_data
        self._all_data   = all_data
        self._on_save    = on_save
        self._on_back    = on_back
        self._ind_frames: list[IndicatorFormFrame] = []
        self._levels: dict[str, str] = {}

        indicators = DIRECTION_INDICATORS.get(direction["key"], [])
        for spec in indicators:
            code = spec["code"]
            self._levels[code] = saved_data.get(code, {}).get("level", "—")

        self._build()

    # ── Build ──────────────────────────────────────────────────────────────

    def _build(self) -> None:
        C  = COLORS
        dk = self._direction

        # ─── row 0: header (56px) ──────────────────────────────────────
        hdr = tk.Frame(self, bg=C["bg_header"], height=56)
        hdr.grid(row=0, column=0, sticky="ew")
        hdr.columnconfigure(2, weight=1)
        hdr.grid_propagate(False)

        make_button(hdr, "← Назад",
                    bg=C["bg_header"], fg=C["text_muted"],
                    activebackground=C["bg_surface"],
                    activeforeground=C["text_primary"],
                    font=FONT_DEFAULT, padx=14, pady=4,
                    command=self._on_back).grid(row=0, column=0, padx=10, pady=12)

        tk.Frame(hdr, bg=dk["color"], width=4).grid(row=0, column=1, sticky="ns")

        title_f = tk.Frame(hdr, bg=C["bg_header"])
        title_f.grid(row=0, column=2, sticky="ew", padx=12, pady=8)
        title_clean = dk["title"].replace("\n", " ")
        tk.Label(title_f, text=f'{dk["icon"]}  {title_clean}',
                 bg=C["bg_header"], fg=C["text_primary"],
                 font=FONT_TITLE).pack(side="left")
        tk.Label(title_f, text=f'  ·  {dk["desc"]}',
                 bg=C["bg_header"], fg=C["text_muted"],
                 font=FONT_DEFAULT).pack(side="left")

        ctrl_f = tk.Frame(hdr, bg=C["bg_header"])
        ctrl_f.grid(row=0, column=3, sticky="e", padx=10, pady=8)

        tk.Label(ctrl_f, text="Підприємство:",
                 bg=C["bg_header"], fg=C["text_muted"],
                 font=FONT_SMALL).pack(side="left")
        self._ent_var = tk.StringVar(value=self._enterprise)
        ent_e = make_dark_entry(ctrl_f, width=18)
        ent_e.insert(0, self._enterprise)
        ent_e.configure(textvariable=self._ent_var)
        ent_e.pack(side="left", padx=(4, 10))

        tk.Label(ctrl_f, text="Квартал:",
                 bg=C["bg_header"], fg=C["text_muted"],
                 font=FONT_SMALL).pack(side="left")
        self._period_var = tk.StringVar(value=self._period)
        self._period_cb  = make_dark_combo(ctrl_f, values=self._gen_periods(), width=9)
        self._period_cb.set(self._period)
        self._period_cb.configure(textvariable=self._period_var)
        self._period_cb.pack(side="left", padx=(4, 10))
        self._period_cb.bind("<<ComboboxSelected>>", lambda _: self._switch_period())

        make_button(ctrl_f, "Зберегти",
                    bg=dk["color"], fg="white",
                    activebackground=C["bg_surface_alt"],
                    font=FONT_BOLD, padx=12, pady=4,
                    command=self._save).pack(side="left")

        # ─── row 1: compact dashboard strip (no margins!) ──────────────
        self._dash_frame = tk.Frame(self, bg=C["bg_surface"])
        self._dash_frame.grid(row=1, column=0, sticky="ew")
        self._build_dashboard()

        # ─── row 2: notebook (Введення / Статистика) ───────────────────
        nb = ttk.Notebook(self)
        nb.grid(row=2, column=0, sticky="nsew")

        # Tab A — indicator forms
        tab_input = tk.Frame(nb, bg=C["bg_main"])
        tab_input.rowconfigure(0, weight=1)
        tab_input.columnconfigure(0, weight=1)
        nb.add(tab_input, text="  Введення даних  ")
        self._build_input_tab(tab_input)

        # Tab B — statistics
        tab_stat = tk.Frame(nb, bg=C["bg_main"])
        tab_stat.rowconfigure(0, weight=1)
        tab_stat.columnconfigure(0, weight=1)
        nb.add(tab_stat, text="  Статистика  ")
        self._stat_tab_frame = tab_stat
        nb.bind("<<NotebookTabChanged>>",
                lambda _: self._maybe_build_stats(nb))

        # ─── row 3: status bar ─────────────────────────────────────────
        sb2 = tk.Frame(self, bg=C["bg_header"], height=22)
        sb2.grid(row=3, column=0, sticky="ew")
        sb2.grid_propagate(False)
        self._status_lbl = tk.Label(sb2, text="",
                                     bg=C["bg_header"], fg=C["text_muted"],
                                     font=FONT_TINY, padx=10)
        self._status_lbl.pack(side="left", pady=3)

    # ── Dashboard strip ────────────────────────────────────────────────────

    def _build_dashboard(self) -> None:
        C = COLORS
        for w in self._dash_frame.winfo_children():
            w.destroy()

        indicators = DIRECTION_INDICATORS.get(self._direction["key"], [])
        total   = len(indicators)
        defined = sum(1 for v in self._levels.values() if v != "—")
        agg     = aggregate_level(list(self._levels.values()))
        counts  = Counter(self._levels.values())
        pct     = int(defined / total * 100) if total else 0

        # Single compact row
        row_f = tk.Frame(self._dash_frame, bg=C["bg_surface"])
        row_f.pack(fill="x")

        # Progress bar narrow strip (top 3px)
        pb_outer = tk.Frame(self._dash_frame, bg=C["bg_surface_alt"], height=3)
        pb_outer.pack(fill="x")
        if pct > 0:
            pb_fill = tk.Frame(pb_outer,
                               bg=RA_COLORS.get(agg, C["accent"]), height=3)
            pb_fill.place(relx=0, rely=0, relwidth=pct / 100, relheight=1)

        # Left: label
        tk.Label(row_f, text="Зведення:",
                 bg=C["bg_surface"], fg=C["text_muted"],
                 font=FONT_SMALL_BOLD).pack(side="left", padx=(10, 6), pady=5)

        # Level badges
        for lvl in ("Green", "Yellow", "Orange", "Red"):
            cnt = counts.get(lvl, 0)
            if cnt:
                tk.Label(row_f,
                         text=f" ● {RA_LABELS[lvl]}: {cnt} ",
                         bg=RA_COLORS[lvl], fg="white",
                         font=FONT_TINY).pack(side="left", padx=2, pady=5)

        nd = total - defined
        if nd:
            tk.Label(row_f,
                     text=f" Не заповн.: {nd} ",
                     bg=C["border_soft"], fg="white",
                     font=FONT_TINY).pack(side="left", padx=2, pady=5)

        # Separator
        tk.Frame(row_f, bg=C["border_soft"], width=1).pack(
            side="left", fill="y", padx=6, pady=4)

        tk.Label(row_f,
                 text=f" Заповнено: {defined}/{total} ({pct}%) ",
                 bg=C["bg_surface"], fg=C["text_muted"],
                 font=FONT_TINY).pack(side="left", pady=5)

        # Right: aggregated
        tk.Label(row_f,
                 text=f"  Агрегований рівень: {RA_LABELS.get(agg, '—')}  ",
                 bg=RA_COLORS.get(agg, C["border_soft"]),
                 fg="white", font=FONT_SMALL_BOLD).pack(
            side="right", padx=10, pady=5)

    # ── Input tab ─────────────────────────────────────────────────────────

    def _build_input_tab(self, container: tk.Frame) -> None:
        C = COLORS

        canvas = tk.Canvas(container, bg=C["bg_main"], highlightthickness=0)
        sb = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=sb.set)
        canvas.grid(row=0, column=0, sticky="nsew")
        sb.grid(row=0, column=1, sticky="ns")

        self._list_frame = tk.Frame(canvas, bg=C["bg_main"])
        self._list_frame.columnconfigure(0, weight=1)
        cw = canvas.create_window((0, 0), window=self._list_frame, anchor="nw")

        def _conf(_: object) -> None:
            canvas.configure(scrollregion=canvas.bbox("all"))
            canvas.itemconfig(cw, width=canvas.winfo_width())

        self._list_frame.bind("<Configure>", _conf)
        canvas.bind("<Configure>", _conf)
        _scroll_mgr.attach(canvas)

        self._build_indicators()

    def _build_indicators(self) -> None:
        C = COLORS
        for w in self._list_frame.winfo_children():
            w.destroy()
        self._ind_frames.clear()

        indicators = DIRECTION_INDICATORS.get(self._direction["key"], [])
        if not indicators:
            tk.Label(self._list_frame,
                     text="Індикатори для цього напрямку ще додаються...",
                     bg=C["bg_main"], fg=C["text_muted"],
                     font=("Arial", 10)).grid(row=0, column=0, pady=40)
            return

        current_group = None
        row = 0

        for spec in indicators:
            if spec["group"] != current_group:
                current_group = spec["group"]
                gf = tk.Frame(self._list_frame, bg=C["bg_main"])
                gf.grid(row=row, column=0, sticky="ew",
                        padx=12, pady=(10, 2))
                tk.Frame(gf, bg=self._direction["color"],
                         width=20, height=2).pack(side="left")
                tk.Label(gf, text=f"  Блок {current_group}",
                         bg=C["bg_main"], fg=self._direction["color"],
                         font=FONT_SMALL_BOLD).pack(side="left")
                row += 1

            code       = spec["code"]
            saved_code = self._saved_data.get(code, {})

            frm = IndicatorFormFrame(
                parent=self._list_frame,
                spec=spec,
                saved_values=saved_code.get("values", {}),
                saved_notes=saved_code.get("notes", ""),
                on_change=self._on_indicator_change,
                row=row,
            )
            self._ind_frames.append(frm)
            row += 1

        tk.Frame(self._list_frame, bg=C["bg_main"], height=20).grid(row=row, column=0)

    def _on_indicator_change(self, code: str, values: dict,
                              level: str, notes: str) -> None:
        self._levels[code] = level
        self._build_dashboard()

    # ── Statistics tab ─────────────────────────────────────────────────────

    def _maybe_build_stats(self, nb: ttk.Notebook) -> None:
        idx = nb.index(nb.select())
        if idx == 1:
            self._build_stats_tab()

    def _build_stats_tab(self) -> None:
        C = COLORS
        for w in self._stat_tab_frame.winfo_children():
            w.destroy()
        self._stat_tab_frame.rowconfigure(0, weight=1)
        self._stat_tab_frame.columnconfigure(0, weight=1)

        canvas = tk.Canvas(self._stat_tab_frame, bg=C["bg_main"],
                           highlightthickness=0)
        sb = ttk.Scrollbar(self._stat_tab_frame, orient="vertical",
                            command=canvas.yview)
        canvas.configure(yscrollcommand=sb.set)
        canvas.grid(row=0, column=0, sticky="nsew")
        sb.grid(row=0, column=1, sticky="ns")

        body = tk.Frame(canvas, bg=C["bg_main"])
        body.columnconfigure(0, weight=1)
        cw = canvas.create_window((0, 0), window=body, anchor="nw")

        def _conf(_: object) -> None:
            canvas.configure(scrollregion=canvas.bbox("all"))
            canvas.itemconfig(cw, width=canvas.winfo_width())
        body.bind("<Configure>", _conf)
        canvas.bind("<Configure>", _conf)
        _scroll_mgr.attach(canvas)

        dir_key    = self._direction["key"]
        indicators = DIRECTION_INDICATORS.get(dir_key, [])
        dir_data   = self._all_data.get(dir_key, {})   # {period: {ent: {code: ...}}}

        # ── gather all periods and years ──────────────────────────────
        all_periods = sorted(dir_data.keys(),
                             key=lambda p: (p.split()[-1], p.split()[0]))
        all_years   = sorted({p.split()[-1] for p in all_periods})

        row = 0

        # ─── Section: selector bar ─────────────────────────────────────
        sel_f = tk.Frame(body, bg=C["bg_surface"])
        sel_f.grid(row=row, column=0, sticky="ew", padx=12, pady=(12, 0))
        row += 1

        tk.Label(sel_f, text="Фільтр:",
                 bg=C["bg_surface"], fg=C["text_muted"],
                 font=FONT_SMALL_BOLD).pack(side="left", padx=10, pady=6)

        tk.Label(sel_f, text="Підприємство:",
                 bg=C["bg_surface"], fg=C["text_muted"],
                 font=FONT_SMALL).pack(side="left", padx=(6, 2), pady=6)

        # gather enterprises
        all_ents = sorted({
            ent
            for pd_ in dir_data.values()
            for ent in pd_.keys()
        })
        ent_values = ["Всі"] + all_ents
        stat_ent_var = tk.StringVar(value="Всі")
        stat_ent_cb  = make_dark_combo(sel_f, values=ent_values, width=20)
        stat_ent_cb.set("Всі")
        stat_ent_cb.configure(textvariable=stat_ent_var)
        stat_ent_cb.pack(side="left", padx=(0, 10), pady=6)

        # refresh button
        def _refresh() -> None:
            self._build_stats_tab()

        make_button(sel_f, "Оновити",
                    bg=self._direction["color"], fg="white",
                    activebackground=C["bg_surface_alt"],
                    font=FONT_SMALL, padx=10, pady=3,
                    command=_refresh).pack(side="left", pady=6)

        if not all_periods:
            tk.Label(body,
                     text="Немає збережених даних.\nВведіть показники та натисніть «Зберегти».",
                     bg=C["bg_main"], fg=C["text_subtle"],
                     font=("Arial", 10), justify="center").grid(
                row=row, column=0, pady=60)
            return

        chosen_ent = stat_ent_var.get()

        # ─── Section A: Квартальна таблиця ────────────────────────────
        self._stat_section(body, row, "Квартальна статистика",
                           self._direction["color"]); row += 1

        row = self._build_quarterly_table(
            body, row, dir_data, indicators, all_periods, chosen_ent)

        # ─── Section B: Зведення по роках ─────────────────────────────
        self._stat_section(body, row, "Річне зведення",
                           self._direction["color"]); row += 1

        row = self._build_yearly_summary(
            body, row, dir_data, indicators, all_years, chosen_ent)

        # ─── Section C: Динаміка агрегованого рівня ───────────────────
        self._stat_section(body, row, "Динаміка агрегованого рівня по кварталах",
                           self._direction["color"]); row += 1

        row = self._build_trend_chart(
            body, row, dir_data, all_periods, chosen_ent)

        # ─── Section D: Детальна таблиця по індикаторах ───────────────
        self._stat_section(body, row, "Деталізація по індикаторах",
                           self._direction["color"]); row += 1

        self._build_indicator_detail(
            body, row, dir_data, indicators, all_periods, chosen_ent)

        tk.Frame(body, bg=C["bg_main"], height=24).grid(
            row=row + 1, column=0)

    @staticmethod
    def _stat_section(parent: tk.Misc, row: int, title: str,
                      color: str) -> None:
        C = COLORS
        f = tk.Frame(parent, bg=C["bg_main"])
        f.grid(row=row, column=0, sticky="ew", padx=12, pady=(16, 4))
        tk.Frame(f, bg=color, width=4, height=16).pack(side="left")
        tk.Label(f, text=f"  {title}",
                 bg=C["bg_main"], fg=color,
                 font=FONT_BOLD).pack(side="left")

    def _collect_levels(
        self,
        dir_data: dict,
        period:   str,
        ent_filter: str,
    ) -> dict[str, str]:
        """
        Повертає {code: level} для одного кварталу.
        Якщо ent_filter == 'Всі' — агрегуємо по всіх підприємствах.
        """
        indicators = DIRECTION_INDICATORS.get(self._direction["key"], [])
        codes = [s["code"] for s in indicators]

        period_data = dir_data.get(period, {})
        if ent_filter == "Всі":
            ents = list(period_data.values())
        else:
            ents = [period_data[ent_filter]] if ent_filter in period_data else []

        result: dict[str, str] = {}
        for code in codes:
            lvls = [e.get(code, {}).get("level", "—")
                    for e in ents if e.get(code)]
            result[code] = aggregate_level([l for l in lvls if l != "—"]) \
                           if any(l != "—" for l in lvls) else "—"
        return result

    def _build_quarterly_table(
        self,
        parent:    tk.Misc,
        row:       int,
        dir_data:  dict,
        indicators: list[dict],
        periods:   list[str],
        chosen_ent: str,
    ) -> int:
        C = COLORS

        # container with horizontal scroll
        outer = tk.Frame(parent, bg=C["bg_main"])
        outer.grid(row=row, column=0, sticky="ew", padx=12, pady=(0, 4))
        outer.columnconfigure(0, weight=1)

        # Build matrix: rows = indicators, cols = periods
        # Header
        hdr = tk.Frame(outer, bg=C["bg_surface"])
        hdr.grid(row=0, column=0, sticky="ew")

        tk.Label(hdr, text="Індикатор", bg=C["bg_surface"],
                 fg=C["text_subtle"], font=FONT_TINY,
                 width=28, anchor="w").grid(row=0, column=0, padx=6, pady=4)

        for ci, p in enumerate(periods):
            tk.Label(hdr, text=p, bg=C["bg_surface"],
                     fg=C["text_subtle"], font=FONT_TINY,
                     width=10, anchor="center").grid(
                row=0, column=ci + 1, padx=2, pady=4)

        # Rows
        for ri, spec in enumerate(indicators):
            bg = C["row_even"] if ri % 2 == 0 else C["row_odd"]
            rf = tk.Frame(outer, bg=bg)
            rf.grid(row=ri + 1, column=0, sticky="ew")

            tk.Label(rf, text=f'{spec["code"]}  {spec["name"]}',
                     bg=bg, fg=C["text_muted"], font=FONT_TINY,
                     width=28, anchor="w").grid(row=0, column=0, padx=6, pady=3)

            for ci, p in enumerate(periods):
                levels_map = self._collect_levels(dir_data, p, chosen_ent)
                lvl = levels_map.get(spec["code"], "—")
                bg_lvl = RA_COLORS.get(lvl, C["border_soft"])
                dot    = "●" if lvl != "—" else "○"
                lbl    = tk.Label(rf, text=f" {dot} {lvl[:3] if lvl != '—' else '—'} ",
                                  bg=bg_lvl if lvl != "—" else bg,
                                  fg="white" if lvl != "—" else C["text_subtle"],
                                  font=FONT_TINY, width=10, anchor="center")
                lbl.grid(row=0, column=ci + 1, padx=2, pady=2)

        return row + 1

    def _build_yearly_summary(
        self,
        parent:     tk.Misc,
        row:        int,
        dir_data:   dict,
        indicators: list[dict],
        years:      list[str],
        chosen_ent: str,
    ) -> int:
        C = COLORS

        # For each year: aggregate over all Q1–Q4
        year_levels: dict[str, list[str]] = {y: [] for y in years}
        for period, ent_map in dir_data.items():
            yr = period.split()[-1]
            if yr not in year_levels:
                continue
            if chosen_ent == "Всі":
                ents = list(ent_map.values())
            else:
                ents = [ent_map[chosen_ent]] if chosen_ent in ent_map else []
            for ent in ents:
                for code_data in ent.values():
                    lvl = code_data.get("level", "—")
                    if lvl != "—":
                        year_levels[yr].append(lvl)

        outer = tk.Frame(parent, bg=C["bg_main"])
        outer.grid(row=row, column=0, sticky="ew", padx=12, pady=(0, 4))

        # Header
        hdr = tk.Frame(outer, bg=C["bg_surface"])
        hdr.pack(fill="x")
        for ci, (lbl_t, w) in enumerate([
            ("Рік",                     8),
            ("Агрег. рівень",          14),
            ("🟢 Green",                9),
            ("🟡 Yellow",              9),
            ("🟠 Orange",              9),
            ("🔴 Red",                  9),
            ("Всього записів",         14),
        ]):
            tk.Label(hdr, text=lbl_t, bg=C["bg_surface"],
                     fg=C["text_subtle"], font=FONT_TINY,
                     width=w, anchor="center").pack(side="left", padx=3, pady=4)

        for ri, yr in enumerate(years):
            bg  = C["row_even"] if ri % 2 == 0 else C["row_odd"]
            lvls = year_levels[yr]
            agg  = aggregate_level(lvls) if lvls else "—"
            cnt  = Counter(lvls)
            total = len(lvls)

            rf = tk.Frame(outer, bg=bg)
            rf.pack(fill="x")

            for val, w in [
                (yr,                         8),
                (RA_LABELS.get(agg, "—"),  14),
                (str(cnt.get("Green",  0)),   9),
                (str(cnt.get("Yellow", 0)),   9),
                (str(cnt.get("Orange", 0)),   9),
                (str(cnt.get("Red",    0)),   9),
                (str(total),                 14),
            ]:
                fg = C["text_primary"]
                cell_bg = bg
                if val == RA_LABELS.get(agg, "—") and agg != "—":
                    cell_bg = RA_COLORS.get(agg, bg)
                    fg = "white"
                tk.Label(rf, text=f" {val} ",
                         bg=cell_bg, fg=fg, font=FONT_TINY,
                         width=w, anchor="center").pack(side="left", padx=3, pady=3)

        return row + 1

    def _build_trend_chart(
        self,
        parent:     tk.Misc,
        row:        int,
        dir_data:   dict,
        periods:    list[str],
        chosen_ent: str,
    ) -> int:
        C = COLORS

        # Map level → numeric score for trend (Green=0, Yellow=1, Orange=2, Red=3)
        lvl_score = {"Green": 0, "Yellow": 1, "Orange": 2, "Red": 3, "—": -1}
        score_labels = {0: "Green", 1: "Yellow", 2: "Orange", 3: "Red"}

        trend_f = tk.Frame(parent, bg=C["bg_surface"])
        trend_f.grid(row=row, column=0, sticky="ew", padx=12, pady=(0, 4))

        CHART_W, CHART_H = 700, 110
        BAR_H = 28
        LEFT  = 60
        RIGHT = 20
        TOP   = 10
        BOTTOM = 30
        n     = len(periods)

        if n == 0:
            tk.Label(trend_f, text="Немає даних для тренду.",
                     bg=C["bg_surface"], fg=C["text_muted"],
                     font=FONT_SMALL).pack(pady=20)
            return row + 1

        cvs = tk.Canvas(trend_f, bg=C["bg_surface"],
                        width=CHART_W, height=CHART_H,
                        highlightthickness=0)
        cvs.pack(padx=8, pady=8)

        avail_w = CHART_W - LEFT - RIGHT
        col_w   = avail_w / n

        # Y axis labels
        for score, label in score_labels.items():
            y = TOP + CHART_H - BOTTOM - (score / 3) * (CHART_H - TOP - BOTTOM)
            cvs.create_text(LEFT - 6, y, text=label[:3],
                            fill=RA_COLORS[label], font=("Arial", 7),
                            anchor="e")
            cvs.create_line(LEFT, y, CHART_W - RIGHT, y,
                            fill=C["border_soft"], dash=(2, 4))

        # Bars + line
        xs: list[float] = []
        ys: list[float] = []

        for ci, p in enumerate(periods):
            levels_map = self._collect_levels(dir_data, p, chosen_ent)
            lvls       = [v for v in levels_map.values() if v != "—"]
            agg        = aggregate_level(lvls) if lvls else "—"
            score      = lvl_score.get(agg, -1)

            x0 = LEFT + ci * col_w
            x1 = x0 + col_w

            if score >= 0:
                bar_h = BAR_H
                yb    = CHART_H - BOTTOM
                yt    = yb - bar_h
                cvs.create_rectangle(x0 + 2, yt, x1 - 2, yb,
                                     fill=RA_COLORS[agg],
                                     outline="", width=0)
                xm = (x0 + x1) / 2
                xs.append(xm)
                ys.append((CHART_H - BOTTOM) - (score / 3) * (CHART_H - TOP - BOTTOM))

            # Period label
            cvs.create_text((x0 + x1) / 2, CHART_H - BOTTOM + 10,
                            text=p, fill=C["text_muted"],
                            font=("Arial", 6), anchor="n")

        # Trend line
        if len(xs) > 1:
            for i in range(len(xs) - 1):
                cvs.create_line(xs[i], ys[i], xs[i + 1], ys[i + 1],
                                fill="white", width=1, dash=(3, 2))

        return row + 1

    def _build_indicator_detail(
        self,
        parent:     tk.Misc,
        row:        int,
        dir_data:   dict,
        indicators: list[dict],
        periods:    list[str],
        chosen_ent: str,
    ) -> None:
        C = COLORS

        outer = tk.Frame(parent, bg=C["bg_main"])
        outer.grid(row=row, column=0, sticky="ew", padx=12, pady=(0, 8))
        outer.columnconfigure(0, weight=1)

        # One mini-table per indicator (collapsed style)
        for spec in indicators:
            code = spec["code"]
            # collect per-period levels
            period_levels = {}
            for p in periods:
                lm  = self._collect_levels(dir_data, p, chosen_ent)
                period_levels[p] = lm.get(code, "—")

            has_data = any(v != "—" for v in period_levels.values())
            if not has_data:
                continue

            card = tk.Frame(outer, bg=C["bg_surface"])
            card.pack(fill="x", pady=(0, 4))

            # card header
            ch_bg  = C["bg_surface_alt"]
            ch     = tk.Frame(card, bg=ch_bg)
            ch.pack(fill="x")

            worst  = aggregate_level([v for v in period_levels.values() if v != "—"])
            w_col  = RA_COLORS.get(worst, C["border_soft"])
            tk.Frame(ch, bg=w_col, width=4).pack(side="left", fill="y")
            tk.Label(ch, text=f"  {code}  {spec['name']}",
                     bg=ch_bg, fg=C["text_primary"],
                     font=FONT_SMALL_BOLD).pack(side="left", pady=4)
            tk.Label(ch,
                     text=f"  {RA_LABELS.get(worst, '—')}  ",
                     bg=w_col, fg="white", font=FONT_TINY).pack(
                side="right", padx=8, pady=4)

            # period cells row
            cells = tk.Frame(card, bg=C["bg_surface"])
            cells.pack(fill="x", padx=4, pady=4)

            for p, lvl in period_levels.items():
                bg  = RA_COLORS.get(lvl, C["border_soft"]) if lvl != "—" else C["bg_main"]
                fg  = "white" if lvl != "—" else C["text_subtle"]
                cf  = tk.Frame(cells, bg=bg, padx=6, pady=4)
                cf.pack(side="left", padx=2)
                tk.Label(cf, text=p,  bg=bg, fg=fg, font=FONT_TINY).pack()
                tk.Label(cf, text=lvl[:3] if lvl != "—" else "—",
                         bg=bg, fg=fg, font=FONT_SMALL_BOLD).pack()

    # ── Period switch ──────────────────────────────────────────────────────

    def _switch_period(self) -> None:
        new_period = self._period_var.get()
        new_ent    = self._ent_var.get().strip()
        dir_key    = self._direction["key"]

        saved = (self._all_data
                 .get(dir_key, {})
                 .get(new_period, {})
                 .get(new_ent, {}))

        self._period      = new_period
        self._enterprise  = new_ent
        self._saved_data  = saved

        # reset levels
        indicators = DIRECTION_INDICATORS.get(dir_key, [])
        self._levels = {}
        for spec in indicators:
            code = spec["code"]
            self._levels[code] = saved.get(code, {}).get("level", "—")

        self._build_dashboard()
        self._build_indicators()

        self._status_lbl.configure(
            text=(f"Завантажено: {new_ent or '—'}  ·  {new_period}"
                  f"  ·  {datetime.now().strftime('%H:%M:%S')}"))

    # ── Save ──────────────────────────────────────────────────────────────

    def _save(self) -> None:
        data: dict[str, dict] = {}
        for frm in self._ind_frames:
            code = frm.spec["code"]
            data[code] = {
                "values": frm.get_values(),
                "level":  frm.get_level(),
                "notes":  frm.get_note(),
            }
        ent    = self._ent_var.get().strip() or "—"
        period = self._period_var.get()
        self._on_save(self._direction["key"], period, ent, data)
        self._status_lbl.configure(
            text=(f"Збережено  ·  {ent}  ·  {period}"
                  f"  ·  {datetime.now().strftime('%H:%M:%S')}"))

    @staticmethod
    def _gen_periods() -> list[str]:
        now = datetime.now()
        result = []
        for y in range(now.year - 2, now.year + 2):
            for q in range(1, 5):
                result.append(f"Q{q} {y}")
        return result


# =============================================================================
#  Головна сторінка Risk Appetite (меню 2×2)
# =============================================================================

class RiskAppetitePage(tk.Frame):
    """
    Стартова сторінка модуля «Ризик апетит».
    Головний екран: 4 картки (2×2) + зведена статистика по поточному кварталу.
    Дані зберігаються в risk_appetite.json.
    """

    DATA_FILE = APPETITE_FILE

    def __init__(self, master: tk.Misc, **kwargs) -> None:
        C = COLORS
        super().__init__(master, bg=C["bg_main"], **kwargs)
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        self._all_data: dict = {}
        self._load_data()
        self._show_menu()

    # ── Persistence ─────────────────────────────────────────────────────────

    def _load_data(self) -> None:
        if not os.path.exists(self.DATA_FILE):
            return
        try:
            with open(self.DATA_FILE, "r", encoding="utf-8") as f:
                self._all_data = json.load(f)
        except (json.JSONDecodeError, OSError):
            self._all_data = {}

    def _save_data(self) -> None:
        try:
            with open(self.DATA_FILE, "w", encoding="utf-8") as f:
                json.dump(self._all_data, f, ensure_ascii=False, indent=2)
        except OSError as e:
            messagebox.showerror("Помилка збереження", str(e))

    def save_before_exit(self) -> None:
        self._save_data()

    # ── Menu ─────────────────────────────────────────────────────────────────

    def _show_menu(self) -> None:
        for w in self.winfo_children():
            w.destroy()
        C = COLORS
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        # Full-height scrollable page
        outer = tk.Frame(self, bg=C["bg_main"])
        outer.grid(row=0, column=0, sticky="nsew")
        outer.rowconfigure(0, weight=1)
        outer.columnconfigure(0, weight=1)

        canvas = tk.Canvas(outer, bg=C["bg_main"], highlightthickness=0)
        sb = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=sb.set)
        canvas.grid(row=0, column=0, sticky="nsew")
        sb.grid(row=0, column=1, sticky="ns")
        _scroll_mgr.attach(canvas)

        body = tk.Frame(canvas, bg=C["bg_main"])
        body.columnconfigure(0, weight=1)
        body.columnconfigure(1, weight=1)
        cw = canvas.create_window((0, 0), window=body, anchor="nw")

        def _conf(_: object) -> None:
            canvas.configure(scrollregion=canvas.bbox("all"))
            canvas.itemconfig(cw, width=canvas.winfo_width())
        body.bind("<Configure>", _conf)
        canvas.bind("<Configure>", _conf)

        # ── Title strip ─────────────────────────────────────────────────
        title_f = tk.Frame(body, bg=C["bg_header"])
        title_f.grid(row=0, column=0, columnspan=2, sticky="ew")
        title_f.columnconfigure(0, weight=1)

        tk.Label(title_f, text="РИЗИК АПЕТИТ",
                 bg=C["bg_header"], fg=C["accent_muted"],
                 font=("Arial", 16, "bold")).grid(
            row=0, column=0, padx=24, pady=(16, 2), sticky="w")
        cur_period = self._current_period()
        tk.Label(title_f,
                 text=f"Поточний квартал: {cur_period}",
                 bg=C["bg_header"], fg=C["text_muted"],
                 font=FONT_SMALL).grid(row=1, column=0, padx=24, pady=(0, 12), sticky="w")

        # ── Global summary row (current quarter) ────────────────────────
        sum_f = tk.Frame(body, bg=C["bg_surface"])
        sum_f.grid(row=1, column=0, columnspan=2, sticky="ew", padx=12, pady=(10, 2))

        tk.Label(sum_f, text=f"Зведення за {cur_period}:",
                 bg=C["bg_surface"], fg=C["text_muted"],
                 font=FONT_SMALL_BOLD).pack(side="left", padx=12, pady=7)

        for dr in RA_DIRECTIONS:
            lvl = self._period_agg(dr["key"], cur_period)
            col = RA_COLORS.get(lvl, C["border_soft"])
            tk.Label(sum_f,
                     text=f"  {dr['icon']} {dr['title'].split(chr(10))[0]}: "
                          f"{RA_LABELS.get(lvl, '—')}  ",
                     bg=col, fg="white", font=FONT_TINY).pack(
                side="left", padx=3, pady=7)

        # ── 2×2 direction cards ─────────────────────────────────────────
        for i, dr in enumerate(RA_DIRECTIONS):
            r = (i // 2) + 2
            c = i % 2
            self._build_dir_card(body, dr, r, c, cur_period)

        # ── Quarterly history table (all directions) ────────────────────
        self._build_history_table(body, 4, cur_period)

        # hint
        tk.Label(body,
                 text="Дані зберігаються автоматично при натисканні «Зберегти» у кожному напрямку.",
                 bg=C["bg_main"], fg=C["text_subtle"],
                 font=FONT_TINY).grid(
            row=5, column=0, columnspan=2, pady=(8, 16))

    def _build_dir_card(
        self, parent: tk.Misc, dr: dict,
        r: int, c: int, cur_period: str,
    ) -> None:
        C     = COLORS
        color = dr["color"]
        agg   = self._direction_agg(dr["key"])
        cur   = self._period_agg(dr["key"], cur_period)
        agg_c = RA_COLORS.get(agg, C["border_soft"])
        cur_c = RA_COLORS.get(cur, C["border_soft"])

        card = tk.Frame(parent, bg=C["bg_surface"],
                        cursor="hand2", width=300, height=180)
        card.grid(row=r, column=c, padx=12, pady=8, sticky="nsew")
        parent.grid_rowconfigure(r, weight=0)
        parent.grid_columnconfigure(c, weight=1)
        card.grid_propagate(False)
        card.columnconfigure(0, weight=1)
        card.rowconfigure(1, weight=1)

        # accent strip
        tk.Frame(card, bg=color, height=4).grid(row=0, column=0, sticky="ew")

        inner = tk.Frame(card, bg=C["bg_surface"])
        inner.grid(row=1, column=0, sticky="nsew", padx=18, pady=10)
        inner.columnconfigure(0, weight=1)
        inner.columnconfigure(1, weight=0)

        # Icon + title
        tk.Label(inner, text=dr["icon"],
                 bg=C["bg_surface"], fg=color,
                 font=("Arial", 20)).grid(row=0, column=0, sticky="w")
        title_clean = dr["title"].replace("\n", " ")
        tk.Label(inner, text=title_clean,
                 bg=C["bg_surface"], fg=C["text_primary"],
                 font=("Arial", 10, "bold"), anchor="w").grid(
            row=1, column=0, columnspan=2, sticky="w", pady=(2, 0))
        tk.Label(inner, text=dr["desc"],
                 bg=C["bg_surface"], fg=C["text_muted"],
                 font=FONT_TINY).grid(row=2, column=0, columnspan=2, sticky="w")

        # Stats row
        stats_f = tk.Frame(inner, bg=C["bg_surface"])
        stats_f.grid(row=3, column=0, columnspan=2, sticky="w", pady=(8, 0))

        tk.Label(stats_f, text="Поточний квартал:",
                 bg=C["bg_surface"], fg=C["text_subtle"],
                 font=FONT_TINY).pack(side="left")
        tk.Label(stats_f,
                 text=f"  {RA_LABELS.get(cur, '—')}  ",
                 bg=cur_c, fg="white", font=FONT_TINY).pack(side="left", padx=4)

        tk.Label(stats_f, text="  Загалом:",
                 bg=C["bg_surface"], fg=C["text_subtle"],
                 font=FONT_TINY).pack(side="left")
        tk.Label(stats_f,
                 text=f"  {RA_LABELS.get(agg, '—')}  ",
                 bg=agg_c, fg="white", font=FONT_TINY).pack(side="left", padx=4)

        # Bottom bar
        bot = tk.Frame(card, bg=C["bg_surface_alt"])
        bot.grid(row=2, column=0, sticky="ew", padx=0, pady=0)
        tk.Label(bot, text="Відкрити →",
                 bg=C["bg_surface_alt"], fg=C["text_muted"],
                 font=FONT_TINY, cursor="hand2").pack(
            side="right", padx=12, pady=5)

        def _open(dk=dr["key"]) -> None:
            self._open_direction(dk)

        for w in [card, inner, bot] + list(inner.winfo_children()):
            try:
                w.bind("<Button-1>", lambda _, fn=_open: fn())
                w.configure(cursor="hand2")
            except tk.TclError:
                pass

        def _enter(_: object) -> None:
            card.configure(bg=C["bg_surface_alt"])
        def _leave(_: object) -> None:
            card.configure(bg=C["bg_surface"])
        card.bind("<Enter>", _enter)
        card.bind("<Leave>", _leave)

    def _build_history_table(
        self, parent: tk.Misc, row: int, cur_period: str,
    ) -> None:
        C = COLORS

        # Collect all periods across all directions
        all_periods: set[str] = set()
        for dk_data in self._all_data.values():
            all_periods.update(dk_data.keys())
        all_periods_sorted = sorted(
            all_periods,
            key=lambda p: (p.split()[-1], p.split()[0]))

        if not all_periods_sorted:
            return

        sec_f = tk.Frame(parent, bg=C["bg_main"])
        sec_f.grid(row=row, column=0, columnspan=2,
                   sticky="ew", padx=12, pady=(14, 4))
        tk.Frame(sec_f, bg=COLORS["accent"], width=4, height=14).pack(side="left")
        tk.Label(sec_f, text="  Зведена таблиця по кварталах",
                 bg=C["bg_main"], fg=C["accent"],
                 font=FONT_BOLD).pack(side="left")

        tbl_f = tk.Frame(parent, bg=C["bg_surface"])
        tbl_f.grid(row=row + 1, column=0, columnspan=2,
                   sticky="ew", padx=12, pady=(0, 4))

        # Header
        hdr = tk.Frame(tbl_f, bg=C["bg_surface_alt"])
        hdr.pack(fill="x")
        tk.Label(hdr, text="Напрямок", bg=C["bg_surface_alt"],
                 fg=C["text_subtle"], font=FONT_TINY,
                 width=22, anchor="w").pack(side="left", padx=8, pady=5)
        for p in all_periods_sorted:
            bg_h = C["accent"] if p == cur_period else C["bg_surface_alt"]
            tk.Label(hdr, text=p, bg=bg_h,
                     fg="white" if p == cur_period else C["text_subtle"],
                     font=FONT_TINY, width=9, anchor="center").pack(
                side="left", padx=2, pady=5)

        # Rows per direction
        for ri, dr in enumerate(RA_DIRECTIONS):
            bg = C["row_even"] if ri % 2 == 0 else C["row_odd"]
            rf = tk.Frame(tbl_f, bg=bg)
            rf.pack(fill="x")

            title_c = dr["title"].replace("\n", " ")
            tk.Label(rf, text=f'{dr["icon"]} {title_c}',
                     bg=bg, fg=C["text_muted"], font=FONT_TINY,
                     width=22, anchor="w").pack(side="left", padx=8, pady=4)

            for p in all_periods_sorted:
                lvl  = self._period_agg(dr["key"], p)
                bg_c = RA_COLORS.get(lvl, bg) if lvl != "—" else bg
                fg_c = "white" if lvl != "—" else C["text_subtle"]
                tk.Label(rf,
                         text=f" {lvl[:3] if lvl != '—' else '—'} ",
                         bg=bg_c, fg=fg_c, font=FONT_TINY,
                         width=9, anchor="center").pack(
                    side="left", padx=2, pady=3)

        # Year totals row
        all_years = sorted({p.split()[-1] for p in all_periods_sorted})
        if len(all_years) > 1 or (len(all_years) == 1 and len(all_periods_sorted) > 1):
            sep = tk.Frame(tbl_f, bg=C["border_soft"], height=1)
            sep.pack(fill="x")
            yr_f = tk.Frame(tbl_f, bg=C["bg_surface_alt"])
            yr_f.pack(fill="x")
            tk.Label(yr_f, text="Річний агрегат",
                     bg=C["bg_surface_alt"], fg=C["text_muted"],
                     font=FONT_TINY, width=22, anchor="w").pack(
                side="left", padx=8, pady=4)
            for p in all_periods_sorted:
                yr  = p.split()[-1]
                # year aggregate for this period's year (all directions)
                yr_levels: list[str] = []
                for dr in RA_DIRECTIONS:
                    yr_data = self._all_data.get(dr["key"], {})
                    for pp, ent_map in yr_data.items():
                        if pp.split()[-1] == yr:
                            for ent in ent_map.values():
                                for cd in ent.values():
                                    lvl = cd.get("level", "—")
                                    if lvl != "—":
                                        yr_levels.append(lvl)
                agg = aggregate_level(yr_levels) if yr_levels else "—"
                bg_c = RA_COLORS.get(agg, C["bg_surface_alt"]) if agg != "—" else C["bg_surface_alt"]
                fg_c = "white" if agg != "—" else C["text_subtle"]
                tk.Label(yr_f,
                         text=f" {agg[:3] if agg != '—' else '—'} ",
                         bg=bg_c, fg=fg_c, font=FONT_TINY,
                         width=9, anchor="center").pack(
                    side="left", padx=2, pady=3)

    # ── Aggregation helpers ──────────────────────────────────────────────────

    def _direction_agg(self, dir_key: str) -> str:
        """Агрегований рівень по всьому напрямку (всі квартали)."""
        all_levels: list[str] = []
        for ent_map in self._all_data.get(dir_key, {}).values():
            for ent in ent_map.values():
                for cd in ent.values():
                    lvl = cd.get("level", "—")
                    if lvl != "—":
                        all_levels.append(lvl)
        return aggregate_level(all_levels) if all_levels else "—"

    def _period_agg(self, dir_key: str, period: str) -> str:
        """Агрегований рівень для конкретного кварталу."""
        all_levels: list[str] = []
        for ent in self._all_data.get(dir_key, {}).get(period, {}).values():
            for cd in ent.values():
                lvl = cd.get("level", "—")
                if lvl != "—":
                    all_levels.append(lvl)
        return aggregate_level(all_levels) if all_levels else "—"

    # ── Navigation ───────────────────────────────────────────────────────────

    def _open_direction(self, dir_key: str) -> None:
        dr = next(d for d in RA_DIRECTIONS if d["key"] == dir_key)

        dir_data   = self._all_data.get(dir_key, {})
        period     = sorted(dir_data.keys())[-1] if dir_data else self._current_period()
        ent_data   = dir_data.get(period, {})
        enterprise = list(ent_data.keys())[-1] if ent_data else ""
        saved      = ent_data.get(enterprise, {})

        for w in self.winfo_children():
            w.destroy()

        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        frame = RiskDirectionFrame(
            parent=self,
            direction=dr,
            period=period,
            enterprise=enterprise,
            saved_data=saved,
            all_data=self._all_data,
            on_save=self._on_direction_save,
            on_back=self._show_menu,
        )
        frame.grid(row=0, column=0, sticky="nsew")

    def _on_direction_save(
        self, dir_key: str, period: str, enterprise: str, data: dict,
    ) -> None:
        (self._all_data
         .setdefault(dir_key, {})
         .setdefault(period, {}))[enterprise] = data
        self._save_data()

    @staticmethod
    def _current_period() -> str:
        now = datetime.now()
        q   = (now.month - 1) // 3 + 1
        return f"Q{q} {now.year}"


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

        tk.Label(topbar, text=APP_TITLE,
                 bg=C["bg_header"], fg=C["text_primary"],
                 font=("Arial", 14, "bold")).grid(
            row=0, column=0, padx=20, pady=8, sticky="w")

        tk.Label(topbar, text=self._user_full_name,
                 bg=C["bg_header"], fg=C["text_muted"],
                 font=("Arial", 10)).grid(row=0, column=1, padx=8, pady=8, sticky="e")

        make_button(
            topbar, "🔔",
            bg=C["bg_surface"], fg=C["text_primary"],
            activebackground=C["bg_surface_alt"],
            font=("Arial", 14), padx=10, pady=4,
            command=lambda: messagebox.showinfo(
                "Нагадування", "Нагадувань поки немає."),
        ).grid(row=0, column=2, padx=20, pady=8, sticky="e")

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
                font=FONT_DEFAULT, padx=24, pady=7, anchor="w",
                command=lambda k=key: self._on_nav_click(k),
            )
            btn.grid(row=i, column=0, sticky="ew", pady=1)
            self._nav_buttons[key] = btn

        sidebar.grid_rowconfigure(2, weight=1)
        tk.Label(sidebar, text=COPYRIGHT_TEXT,
                 bg=C["bg_sidebar"], fg=C["text_subtle"],
                 font=FONT_SMALL).grid(row=3, column=0, pady=6, sticky="s")

    def _build_content(self) -> None:
        self.content_frame = tk.Frame(self, bg=COLORS["bg_main"])
        self.content_frame.grid(row=1, column=1, sticky="nsew")
        self.content_frame.grid_rowconfigure(0, weight=1)
        self.content_frame.grid_columnconfigure(0, weight=1)
        self._on_nav_click("risk_register")

    def _on_nav_click(self, page_key: PageKey) -> None:
        if self._current_page == page_key:
            return
        self._current_page = page_key

        C = COLORS
        for k, btn in self._nav_buttons.items():
            btn.configure(
                bg=C["bg_surface"]  if k == page_key else C["bg_sidebar"],
                fg=C["text_primary"] if k == page_key else C["text_muted"],
            )

        for child in self.content_frame.winfo_children():
            child.grid_forget()

        if page_key not in self._pages:
            if page_key == "material_events":
                page: tk.Frame = MaterialEventsPage(self.content_frame)
            elif page_key == "risk_register":
                page = RiskRegisterPage(self.content_frame)
            elif page_key == "risk_coordinators":
                page = RiskCoordinatorsPage(self.content_frame)
            elif page_key == "risk_appetite":
                page = RiskAppetitePage(self.content_frame)
            else:
                page = self._create_placeholder_page(page_key)
            self._pages[page_key] = page

        self._pages[page_key].grid(row=0, column=0, sticky="nsew")

    def _create_placeholder_page(self, page_key: PageKey) -> tk.Frame:
        C    = COLORS
        page = tk.Frame(self.content_frame, bg=C["bg_main"])
        page.columnconfigure(0, weight=1)
        page.rowconfigure(0, weight=1)

        inner = tk.Frame(page, bg=C["bg_main"])
        inner.place(relx=0.5, rely=0.5, anchor="center")

        tk.Frame(inner, bg=C["accent"], height=3, width=60).pack(pady=(0, 16))
        tk.Label(inner,
                 text=page_key.replace("_", " ").title(),
                 bg=C["bg_main"], fg=C["text_primary"],
                 font=("Arial", 20, "bold")).pack()
        tk.Label(inner,
                 text="Модуль в розробцi...",
                 bg=C["bg_main"], fg=C["text_muted"],
                 font=("Arial", 11)).pack(pady=(8, 0))
        tk.Frame(inner, bg=C["border_soft"], height=1, width=60).pack(pady=(16, 0))
        return page

    def _on_close(self) -> None:
        for page in self._pages.values():
            if hasattr(page, "save_before_exit"):
                page.save_before_exit()
        self.destroy()


# =============================================================================
#  MAIN
# =============================================================================

def main() -> int:
    # ВИПРАВЛЕННЯ #11: ім'я читається зі змінної, не захардкоджено в логіці
    current_user = os.environ.get("ATLAS_USER", "Онiщенко Андрiй Сергiйович")
    app = AtlasApp(user_full_name=current_user)
    app.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
