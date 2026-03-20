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

FONT_DEFAULT    = ("Arial", 9)
FONT_SMALL      = ("Arial", 8)
FONT_TINY       = ("Arial", 7)
FONT_BOLD       = ("Arial", 9, "bold")
FONT_SMALL_BOLD = ("Arial", 8, "bold")
FONT_TITLE      = ("Arial", 13, "bold")
FONT_HEADING    = ("Arial", 11, "bold")
FONT_NUMBER     = ("Arial", 22, "bold")
FONT_SCORE      = ("Arial", 14, "bold")
FONT_MONO       = ("Courier", 9)

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
#  DATACLASSES
# =============================================================================

@dataclass
class EventRecord:
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

    def to_list(self) -> list:
        return [self.id, self.entity, self.event_name, self.involved,
                self.risk_type, self.event_date, self.description,
                self.measures, self.detect_date, self.priority, self.status]

    @classmethod
    def from_list(cls, row: list) -> "EventRecord":
        r = list(row) + ["—"] * max(0, 11 - len(row))
        return cls(id=str(r[0]), entity=str(r[1]), event_name=str(r[2]),
                   involved=str(r[3]), risk_type=str(r[4]),
                   event_date=str(r[5]), description=str(r[6]),
                   measures=str(r[7]), detect_date=str(r[8]),
                   priority=str(r[9]), status=str(r[10]))

    @classmethod
    def from_dict(cls, d: dict) -> "EventRecord":
        return cls(**{k: str(v) for k, v in d.items()
                      if k in cls.__dataclass_fields__})


@dataclass
class RiskRecord:
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
        return [self.id, self.entity, self.risk_name, self.category,
                self.risk_type, self.probability, self.impact, self.score,
                self.owner, self.controls, self.residual,
                self.date_id, self.date_rev, self.priority,
                self.status, self.description]

    @classmethod
    def from_list(cls, row: list) -> "RiskRecord":
        r = list(row)
        while len(r) < 16:
            r.append("—")
        return cls(id=str(r[0]), entity=str(r[1]), risk_name=str(r[2]),
                   category=str(r[3]), risk_type=str(r[4]),
                   probability=str(r[5]), impact=str(r[6]), score=str(r[7]),
                   owner=str(r[8]), controls=str(r[9]), residual=str(r[10]),
                   date_id=str(r[11]), date_rev=str(r[12]),
                   priority=str(r[13]), status=str(r[14]),
                   description=str(r[15]))

    @classmethod
    def from_dict(cls, d: dict) -> "RiskRecord":
        return cls(**{k: str(v) for k, v in d.items()
                      if k in cls.__dataclass_fields__})


# =============================================================================
#  ХЕЛПЕРИ ТА СТИЛІ
# =============================================================================

def is_valid_date(s: str) -> bool:
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
    style.configure(".", background=C["bg_main"], foreground=C["text_primary"],
                    fieldbackground=C["bg_input"], troughcolor=C["bg_surface"],
                    bordercolor=C["border_soft"], darkcolor=C["bg_surface"],
                    lightcolor=C["bg_surface"], insertcolor=C["text_primary"],
                    selectbackground=C["row_select"],
                    selectforeground=C["text_primary"], font=FONT_DEFAULT)
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
                    fieldbackground=C["row_odd"],
                    bordercolor=C["border_soft"],
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
                        arrowcolor=C["text_muted"],
                        bordercolor=C["bg_main"])


def make_dark_text(parent: tk.Misc, **kwargs) -> tk.Text:
    C = COLORS
    return tk.Text(parent, bg=C["bg_input"], fg=C["text_primary"],
                   insertbackground=C["text_primary"],
                   selectbackground=C["row_select"],
                   selectforeground=C["text_primary"],
                   relief="flat", bd=1, highlightthickness=1,
                   highlightbackground=C["border_soft"],
                   highlightcolor=C["accent"],
                   font=FONT_DEFAULT, **kwargs)


def make_dark_entry(parent: tk.Misc, accent: str | None = None,
                    **kwargs) -> tk.Entry:
    C = COLORS
    return tk.Entry(parent, bg=C["bg_input"], fg=C["text_primary"],
                    insertbackground=C["text_primary"],
                    relief="flat", bd=2, highlightthickness=1,
                    highlightbackground=C["border_soft"],
                    highlightcolor=accent or C["accent"],
                    font=FONT_DEFAULT, **kwargs)


def make_dark_combo(parent: tk.Misc, values: list[str] | None = None,
                    **kwargs) -> ttk.Combobox:
    return ttk.Combobox(parent, values=values or [],
                        state="readonly", font=FONT_DEFAULT, **kwargs)


def make_button(parent: tk.Misc, text: str, bg: str,
                fg: str = "white", **kwargs) -> tk.Button:
    C = COLORS
    active_bg = kwargs.pop("activebackground", bg)
    active_fg = kwargs.pop("activeforeground", fg)
    font      = kwargs.pop("font", FONT_DEFAULT)
    return tk.Button(parent, text=text, bg=bg, fg=fg,
                     activebackground=active_bg, activeforeground=active_fg,
                     relief="flat", bd=0, cursor="hand2",
                     font=font, **kwargs)


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


def _safe_num(val: object) -> float:
    try:
        return float(str(val).replace(" ", "").replace(",", ".") or "0")
    except (ValueError, TypeError):
        return 0.0


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
    cell.grid(row=row, column=col, sticky="nsew",
              padx=(8 if col == 0 else 4, 4 if col == 0 else 8), pady=3)
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
    cell.grid(row=row, column=0, columnspan=2, sticky="nsew", padx=8, pady=3)
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
#  SCROLL MANAGER
# =============================================================================

class ScrollManager:
    def attach(self, canvas: tk.Canvas) -> None:
        canvas.bind("<Enter>",  lambda _: self._bind(canvas))
        canvas.bind("<Leave>",  lambda _: self._unbind(canvas))

    @staticmethod
    def _bind(canvas: tk.Canvas) -> None:
        canvas.bind_all("<MouseWheel>",
            lambda e: canvas.yview_scroll(int(-1 * e.delta / 120), "units"))

    @staticmethod
    def _unbind(canvas: tk.Canvas) -> None:
        canvas.unbind_all("<MouseWheel>")


_scroll_mgr = ScrollManager()


def _scrollable_canvas(container: tk.Misc) -> tuple[tk.Canvas, tk.Frame]:
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
    _scroll_mgr.attach(canvas)
    return canvas, inner


# =============================================================================
#  ID GENERATOR
# =============================================================================

class IdGenerator:
    @staticmethod
    def next_id(existing: list) -> str:
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
        try:
            return str(int(raw))
        except (ValueError, TypeError):
            return raw


# =============================================================================
#  BASE REGISTRY TAB
# =============================================================================

class BaseRegistryTab:
    data_file: str = ""
    frame: ttk.Frame

    def __init__(self, parent: tk.Misc,
                 on_data_change: Callable | None = None) -> None:
        self.parent         = parent
        self.on_data_change = on_data_change
        self.frame          = ttk.Frame(parent)
        self.all_records:   list = []

    def save(self) -> None:
        self._save_data()

    def find_record(self, idx_str: str):
        normalized = IdGenerator.normalize_id(idx_str)
        for r in self.all_records:
            raw = r.id if hasattr(r, "id") else str(r[0])
            if IdGenerator.normalize_id(raw) == normalized:
                return r
        return None

    def _save_data(self) -> None:
        try:
            data = [(asdict(r) if hasattr(r, "__dataclass_fields__") else r)
                    for r in self.all_records]
            with open(self.data_file, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except OSError as e:
            messagebox.showerror("Помилка збереження", str(e))

    def _notify_change(self) -> None:
        if self.on_data_change:
            self.on_data_change(self.all_records)

    def get_frame(self) -> ttk.Frame:
        return self.frame

# =============================================================================
#  EVENT DETAIL WINDOW
# =============================================================================

class EventDetailWindow:
    def __init__(self, parent_root: tk.Misc, record: EventRecord,
                 all_records: list[EventRecord],
                 save_callback: Callable[[str, EventRecord], None],
                 delete_callback: Callable[[str], None],
                 toast_callback: Callable[[str], None]) -> None:
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

        header = tk.Frame(self.win, bg=C["bg_header"], height=58)
        header.grid(row=0, column=0, sticky="ew")
        header.columnconfigure(1, weight=1)
        header.grid_propagate(False)
        risk_color = RISK_COLORS.get(self.record.risk_type, C["accent"])
        tk.Frame(header, bg=risk_color, width=4).grid(row=0, column=0, sticky="ns")
        title_frame = tk.Frame(header, bg=C["bg_header"])
        title_frame.grid(row=0, column=1, sticky="ew", padx=16, pady=10)
        self.lbl_title = tk.Label(title_frame,
            text=f"Запис #{self.record.id}",
            bg=C["bg_header"], fg=C["accent_muted"], font=FONT_HEADING)
        self.lbl_title.pack(anchor="w")
        self.lbl_subtitle = tk.Label(title_frame, text=self.record.entity,
            bg=C["bg_header"], fg=C["text_muted"], font=FONT_DEFAULT)
        self.lbl_subtitle.pack(anchor="w")
        sc = self._status_color(self.record.status)
        self.lbl_status_badge = tk.Label(header,
            text=f"  {self.record.status}  ",
            bg=sc, fg="white", font=FONT_SMALL_BOLD, pady=3)
        self.lbl_status_badge.grid(row=0, column=2, padx=16, pady=18)

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
        _scroll_mgr.attach(canvas)
        self.content.columnconfigure(0, weight=1)
        self.content.columnconfigure(1, weight=1)
        self._build_view_content()

        btn_bar = tk.Frame(self.win, bg=C["bg_header"], height=46)
        btn_bar.grid(row=2, column=0, columnspan=2, sticky="ew")
        btn_bar.grid_propagate(False)
        left_btns = tk.Frame(btn_bar, bg=C["bg_header"])
        left_btns.pack(side="left", padx=12, pady=8)
        self.btn_edit = make_button(left_btns, "Редагувати",
            bg=C["accent_warning"], fg=C["bg_main"],
            activebackground="#d97706", activeforeground="white",
            font=FONT_BOLD, padx=14, pady=4, command=self._toggle_edit_mode)
        self.btn_edit.pack(side="left", padx=(0, 8))
        self.btn_save = make_button(left_btns, "Зберегти змiни",
            bg=C["accent_success"], activebackground="#16a34a",
            font=FONT_BOLD, padx=14, pady=4, command=self._save_changes)
        self.btn_save.pack_forget()
        self.btn_cancel_edit = make_button(left_btns, "Скасувати",
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

    @staticmethod
    def _status_color(status: str) -> str:
        return {"Вiдкрито": COLORS["accent_danger"], "В обробцi": COLORS["accent_warning"],
                "Закрито": COLORS["text_muted"], "Вирiшено": COLORS["accent_success"]
                }.get(status, COLORS["text_muted"])

    @staticmethod
    def _priority_color(priority: str) -> str:
        return {"Критичний": COLORS["accent_danger"], "Високий": COLORS["accent_warning"],
                "Середнiй": COLORS["accent"], "Низький": COLORS["accent_success"]
                }.get(priority, COLORS["text_primary"])

    def _build_view_content(self) -> None:
        for w in self.content.winfo_children():
            w.destroy()
        r = self.record
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
        _build_info_cell(self.content, "Задiянi пiдроздiли / особи", r.involved, row, 1); row += 1
        _build_section_label(self.content, "Дати", row); row += 1
        _build_info_cell(self.content, "Дата подiї",     r.event_date,  row, 0)
        _build_info_cell(self.content, "Дата виявлення", r.detect_date, row, 1); row += 1
        _build_section_label(self.content, "Деталi подiї", row); row += 1
        _build_text_block(self.content, "Детальний опис подiї", r.description, row); row += 1
        _build_text_block(self.content, "Вжитi заходи",          r.measures,    row); row += 1
        tk.Frame(self.content, bg=COLORS["bg_main"], height=12).grid(
            row=row, column=0, columnspan=2)

    def _build_edit_content(self) -> None:
        C = COLORS
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
        self.e_priority = make_dark_combo(ps_f,
            values=["Критичний", "Високий", "Середнiй", "Низький"], width=14)
        self.e_priority.grid(row=1, column=0, pady=(2, 0), padx=(0, 20))
        self.e_priority.set(rec.priority)
        tk.Label(ps_f, text="Статус:", bg=C["bg_main"],
                 fg=C["text_subtle"], font=FONT_SMALL).grid(row=0, column=1, sticky="w")
        self.e_status = make_dark_combo(ps_f,
            values=["Вiдкрито", "В обробцi", "Закрито", "Вирiшено"], width=14)
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
        self.lbl_title.configure(text=f"Редагування запису #{self.record.id}",
                                  fg=COLORS["accent_warning"])

    def _cancel_edit(self) -> None:
        self.is_edit_mode = False
        self._build_view_content()
        self.btn_save.pack_forget()
        self.btn_cancel_edit.pack_forget()
        self.btn_edit.pack(side="left", padx=(0, 8))
        self.lbl_title.configure(text=f"Запис #{self.record.id}",
                                  fg=COLORS["accent_muted"])

    def _save_changes(self) -> None:
        detect  = self.e_detect.get().strip()
        event_d = self.e_event_date.get().strip()
        for val, label in [(detect, "дати виявлення"), (event_d, "дати подiї")]:
            if val in ("дд.мм.рррр", ""):
                continue
            if not is_valid_date(val):
                messagebox.showwarning("Помилка",
                    f"Неправильний формат {label} (очiкується дд.мм.рррр)",
                    parent=self.win)
                return
        detect  = "" if detect  == "дд.мм.рррр" else detect
        event_d = "" if event_d == "дд.мм.рррр" else event_d
        entity = self.e_entity.get().strip()
        event  = self.e_event.get().strip()
        if not entity or not event:
            messagebox.showwarning("Обов'язковi поля",
                "Заповнiть назву пiдприємства та назву подiї", parent=self.win)
            return
        old_id = self.record.id
        new_record = EventRecord(
            id=self.record.id, entity=entity, event_name=event,
            involved=self.e_involved.get("1.0", tk.END).strip(),
            risk_type=self.e_risk.get().strip() or "—",
            event_date=event_d or "—",
            description=self.e_description.get("1.0", tk.END).strip(),
            measures=self.e_measures.get("1.0", tk.END).strip(),
            detect_date=detect or "—",
            priority=self.e_priority.get().strip() or "Середнiй",
            status=self.e_status.get().strip()   or "Вiдкрито")
        self.record = new_record
        self.save_callback(old_id, new_record)
        self.lbl_subtitle.configure(text=entity)
        sc = self._status_color(new_record.status)
        self.lbl_status_badge.configure(text=f"  {new_record.status}  ", bg=sc)
        self._cancel_edit()
        self.toast_callback("Запис оновлено")

    def _delete_record(self) -> None:
        if not messagebox.askyesno("Пiдтвердження",
            f"Видалити запис #{self.record.id}?\nЦю дiю не можна скасувати.",
            parent=self.win):
            return
        self.delete_callback(self.record.id)
        self.toast_callback("Запис видалено")
        self.win.destroy()


# =============================================================================
#  REGISTRY TAB (Events)
# =============================================================================

class RegistryTab(BaseRegistryTab):
    data_file = DATA_FILE

    def __init__(self, parent: tk.Misc,
                 on_data_change: Callable[[list[EventRecord]], None] | None = None) -> None:
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
        make_button(sf, "Скинути", bg=C["bg_surface"], fg=C["text_muted"],
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
        _scroll_mgr.attach(canvas)
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
        self.cb_priority = make_dark_combo(form,
            values=["Критичний", "Високий", "Середнiй", "Низький"])
        self.cb_priority.grid(row=row, column=0, sticky="w", padx=16, pady=(2, 0)); row += 1
        lbl("Статус:", row); row += 1
        self.cb_status = make_dark_combo(form,
            values=["Вiдкрито", "В обробцi", "Закрито", "Вирiшено"])
        self.cb_status.grid(row=row, column=0, sticky="w", padx=16, pady=(2, 0)); row += 1

        btn_f = tk.Frame(form, bg=C["bg_main"])
        btn_f.grid(row=row, column=0, sticky="ew", padx=16, pady=(16, 20))
        btn_f.columnconfigure((0, 1), weight=1)
        make_button(btn_f, "Очистити", bg=C["bg_surface"], fg=C["text_muted"],
                    activebackground=C["bg_surface_alt"],
                    activeforeground=C["text_primary"],
                    padx=14, pady=6, command=self._clear_form
                    ).grid(row=0, column=0, padx=4, sticky="ew")
        make_button(btn_f, "Додати запис", bg=C["accent"],
                    activebackground=C["accent_soft"],
                    font=FONT_BOLD, padx=14, pady=6, command=self._add_record
                    ).grid(row=0, column=1, padx=4, sticky="ew")

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
        self.filter_status = make_dark_combo(toolbar,
            values=["Всi", "Вiдкрито", "В обробцi", "Закрито", "Вирiшено"], width=12)
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
                        font=FONT_SMALL, padx=10, pady=3, command=cmd
                        ).pack(side="right", padx=4, pady=6)

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
            "id": ("№", 46), "entity": ("Пiдприємство", 155),
            "event": ("Назва подiї", 185), "risk": ("Тип ризику", 105),
            "priority": ("Прiоритет", 90), "status": ("Статус", 90),
            "date": ("Дата подiї", 90), "involved": ("Задiянi", 130),
            "desc": ("Опис", 200), "measures": ("Заходи", 200),
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
        make_button(exp_bar, "Експорт CSV", bg=C["bg_surface"], fg=C["text_primary"],
                    activebackground=C["bg_surface_alt"],
                    activeforeground=C["text_primary"],
                    font=FONT_SMALL, padx=12, pady=4,
                    command=self._export_csv).pack(side="left", padx=(0, 6))
        if pd:
            make_button(exp_bar, "Експорт Excel", bg=C["accent_success"],
                        activebackground="#16a34a",
                        font=FONT_SMALL, padx=12, pady=4,
                        command=self._export_excel).pack(side="left", padx=(0, 6))
        make_button(exp_bar, "Iмпорт JSON", bg=C["bg_surface"], fg=C["text_primary"],
                    activebackground=C["bg_surface_alt"],
                    activeforeground=C["text_primary"],
                    font=FONT_SMALL, padx=12, pady=4,
                    command=self._import_json).pack(side="left")

    def _on_double_click(self, event: tk.Event) -> None:
        if not self.tree.selection(): return
        if self.tree.identify("region", event.x, event.y) != "cell": return
        self._open_selected_detail()

    def _open_selected_detail(self) -> None:
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Перегляд", "Оберiть запис для перегляду"); return
        iid = sel[0]
        rec = self.find_record(self.tree.set(iid, "id"))
        if not rec: return
        EventDetailWindow(
            parent_root=self.frame.winfo_toplevel(), record=rec,
            all_records=self.all_records,
            save_callback=lambda old, new: self._on_detail_save(iid, old, new),
            delete_callback=lambda s: self._on_detail_delete(iid, s),
            toast_callback=self._show_toast)

    def _on_detail_save(self, iid: str, old_id: str, new_record: EventRecord) -> None:
        normalized_old = IdGenerator.normalize_id(old_id)
        for i, r in enumerate(self.all_records):
            if IdGenerator.normalize_id(r.id) == normalized_old:
                self.all_records[i] = new_record; break
        try:
            self.tree.item(iid, values=(
                new_record.id, new_record.entity, new_record.event_name,
                new_record.risk_type, new_record.priority, new_record.status,
                new_record.event_date, new_record.involved,
                new_record.description, new_record.measures))
        except tk.TclError: pass
        self._recolor_rows()
        self._save_data()
        self._notify_change()

    def _on_detail_delete(self, iid: str, idx_str: str) -> None:
        try: self.tree.delete(iid)
        except tk.TclError: pass
        normalized = IdGenerator.normalize_id(idx_str)
        self.all_records = [r for r in self.all_records
                            if IdGenerator.normalize_id(r.id) != normalized]
        self._recolor_rows()
        self._save_data()
        self._update_count()
        self._notify_change()

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

    def _load_data(self) -> None:
        self.all_records.clear()
        self.tree.delete(*self.tree.get_children())
        if not os.path.exists(self.data_file): return
        try:
            with open(self.data_file, "r", encoding="utf-8") as f:
                raw = json.load(f)
            if not isinstance(raw, list):
                raise ValueError("Очiкується список записiв у JSON")
            for item in raw:
                if isinstance(item, dict):
                    rec = EventRecord.from_dict(item)
                elif isinstance(item, (list, tuple)):
                    rec = EventRecord.from_list(list(item))
                else: continue
                self.all_records.append(rec)
                self._insert_tree_row(rec)
        except (json.JSONDecodeError, ValueError, OSError) as e:
            messagebox.showerror("Помилка завантаження", str(e))
        self._update_count()
        self._notify_change()

    def _insert_tree_row(self, rec: EventRecord) -> str:
        iid = self.tree.insert("", tk.END, values=(
            rec.id, rec.entity, rec.event_name, rec.risk_type,
            rec.priority, rec.status, rec.event_date,
            rec.involved, rec.description, rec.measures))
        self._recolor_rows()
        return iid

    def _get_form_data(self) -> EventRecord | None:
        detect  = self.ent_detect.get().strip()
        event_d = self.ent_event_date.get().strip()
        for val, label in [(detect, "дати виявлення"), (event_d, "дати подiї")]:
            if val in ("дд.мм.рррр", ""): continue
            if not is_valid_date(val):
                messagebox.showwarning("Помилка",
                    f"Неправильний формат {label} (очiкується дд.мм.рррр)")
                return None
        entity = self.ent_entity.get().strip()
        event  = self.cb_event.get().strip()
        if not entity or not event:
            messagebox.showwarning("Обов'язковi поля",
                "Заповнiть назву пiдприємства та назву подiї")
            return None
        detect  = "" if detect  == "дд.мм.рррр" else detect
        event_d = "" if event_d == "дд.мм.рррр" else event_d
        new_id  = IdGenerator.next_id(self.all_records)
        return EventRecord(
            id=new_id, entity=entity, event_name=event,
            involved=self.txt_involved.get("1.0", tk.END).strip(),
            risk_type=self.cb_risk.get().strip() or "—",
            event_date=event_d or "—",
            description=self.txt_description.get("1.0", tk.END).strip(),
            measures=self.txt_measures.get("1.0", tk.END).strip(),
            detect_date=detect or "—",
            priority=self.cb_priority.get().strip() or "Середнiй",
            status=self.cb_status.get().strip()   or "Вiдкрито")

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
        if not data: return
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
            messagebox.showinfo("Видалення", "Оберiть запис для видалення"); return
        iid     = sel[0]
        idx_str = self.tree.set(iid, "id")
        if not messagebox.askyesno("Пiдтвердження",
            f"Видалити запис #{idx_str}?\nЦю дiю не можна скасувати."): return
        self.tree.delete(iid)
        normalized = IdGenerator.normalize_id(idx_str)
        self.all_records = [r for r in self.all_records
                            if IdGenerator.normalize_id(r.id) != normalized]
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
        if not rec: return
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
            row_str = " ".join([rec.id, rec.entity, rec.event_name, rec.involved,
                                rec.risk_type, rec.event_date, rec.description,
                                rec.measures, rec.detect_date,
                                rec.priority, rec.status]).lower()
            if q      and q      not in row_str:         continue
            if risk   != "Всi"   and rec.risk_type != risk:   continue
            if status != "Всi"   and rec.status    != status: continue
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
        if not sel: return
        rec = self.find_record(self.tree.set(sel[0], "id"))
        if not rec: return
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

    _EVENT_HEADERS = ["ID", "Пiдприємство", "Назва подiї", "Задiянi особи",
                      "Тип ризику", "Дата подiї", "Опис", "Заходи",
                      "Дата виявлення", "Прiоритет", "Статус"]

    def _export_csv(self) -> None:
        if not self.tree.get_children():
            messagebox.showinfo("Експорт", "Таблиця порожня"); return
        path = filedialog.asksaveasfilename(defaultextension=".csv",
            filetypes=[("CSV файли", "*.csv")], title="Зберегти як CSV")
        if not path: return
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
        path = filedialog.asksaveasfilename(defaultextension=".xlsx",
            filetypes=[("Excel файли", "*.xlsx")], title="Зберегти як Excel")
        if not path: return
        try:
            df = pd.DataFrame([rec.to_list() for rec in self.all_records],
                               columns=self._EVENT_HEADERS)
            with pd.ExcelWriter(path, engine="openpyxl") as writer:
                df.to_excel(writer, index=False, sheet_name="Реєстр")
                ws = writer.sheets["Реєстр"]
                for col_cells in ws.columns:
                    mx = max(len(str(c.value or "")) for c in col_cells)
                    ws.column_dimensions[
                        col_cells[0].column_letter].width = min(mx + 4, 60)
            self._show_toast("Excel збережено")
        except (OSError, Exception) as e:
            messagebox.showerror("Помилка", str(e))

    def _import_json(self) -> None:
        path = filedialog.askopenfilename(filetypes=[("JSON файли", "*.json")],
                                          title="Iмпорт JSON")
        if not path: return
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            if not isinstance(data, list):
                raise ValueError("Файл повинен мiстити список записiв")
            import dataclasses
            added = 0
            for item in data:
                if isinstance(item, dict):
                    rec = EventRecord.from_dict(item)
                elif isinstance(item, (list, tuple)):
                    rec = EventRecord.from_list(list(item))
                else: continue
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
#  ANALYTICS TAB (Events)
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
        make_button(header, "Оновити", bg=C["accent"],
                    activebackground=C["accent_soft"],
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
        _scroll_mgr.attach(canvas)
        self.content.columnconfigure(0, weight=1)
        self._build_stat_cards()
        self._build_charts_and_table()

    def _build_stat_cards(self) -> None:
        C = COLORS
        cf = tk.Frame(self.content, bg=C["bg_main"])
        cf.grid(row=0, column=0, sticky="ew", padx=16, pady=(16, 0))
        for i in range(4): cf.columnconfigure(i, weight=1)
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
            if HAS_MPL: self._clear_charts()
            self.stats_tree.delete(*self.stats_tree.get_children())
            return
        C = COLORS
        records = self.records
        total  = len(records)
        open_c = sum(1 for r in records if r.status in ("Вiдкрито", "В обробцi"))
        crit   = sum(1 for r in records if r.priority == "Критичний")
        closed = sum(1 for r in records if r.status in ("Закрито", "Вирiшено"))
        self.stat_cards["total"].configure(text=str(total))
        self.stat_cards["open"].configure(text=str(open_c))
        self.stat_cards["critical"].configure(text=str(crit))
        self.stat_cards["closed"].configure(text=str(closed))
        if not HAS_MPL: return
        risk_ctr   = Counter(r.risk_type for r in records)
        status_ctr = Counter(r.status    for r in records)
        entity_ctr = Counter(r.entity    for r in records if r.entity)
        self.ax_left.clear(); self._style_ax(self.ax_left)
        self.ax_left.set_title("Розподiл за типом ризику", color=C["text_muted"], fontsize=9)
        if risk_ctr:
            lbls = list(risk_ctr.keys()); vals = list(risk_ctr.values())
            clrs = [RISK_COLORS.get(l, C["text_muted"]) for l in lbls]
            _, _, autotexts = self.ax_left.pie(vals, labels=lbls, autopct="%1.0f%%",
                colors=clrs, startangle=90,
                textprops={"color": C["text_muted"], "fontsize": 7})
            for at in autotexts:
                at.set_fontsize(7); at.set_color("white")
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
        self.canvas_right.draw()
        self.ax_bottom.clear(); self._style_ax(self.ax_bottom)
        self.ax_bottom.set_title("Топ-5 пiдприємств за кiлькiстю подiй",
                                  color=C["text_muted"], fontsize=9)
        top5 = entity_ctr.most_common(5)
        if top5:
            e_lbls = [e[0][:20] for e in top5]; e_vals = [e[1] for e in top5]
            bars = self.ax_bottom.barh(e_lbls, e_vals, color=C["accent"], edgecolor="none")
            for bar, val in zip(bars, e_vals, strict=False):
                self.ax_bottom.text(bar.get_width() + 0.05,
                                     bar.get_y() + bar.get_height() / 2,
                                     str(val), ha="left", va="center",
                                     color=C["text_muted"], fontsize=8)
            self.ax_bottom.tick_params(axis="y", labelsize=8,
                                        colors=C["text_primary"])
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
        if not HAS_MPL: return
        for ax in (self.ax_left, self.ax_right, self.ax_bottom):
            ax.clear()
        for cv in (self.canvas_left, self.canvas_right, self.canvas_bottom):
            cv.draw()

    def get_frame(self) -> ttk.Frame:
        return self.frame


# =============================================================================
#  SETTINGS TAB (Events)
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
        self._row(content, 1, "Версiя:", "2.3 — OR інтегровано в ризик-апетит", C)
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
#  MATERIAL EVENTS PAGE
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
        self.registry_tab  = RegistryTab(self.notebook,
            on_data_change=self.analytics_tab.update_data)
        self.settings_tab  = SettingsTab(self.notebook)
        self.notebook.add(self.registry_tab.get_frame(),  text="  Реєстр подiй  ")
        self.notebook.add(self.analytics_tab.get_frame(), text="  Аналiтика  ")
        self.notebook.add(self.settings_tab.get_frame(),  text="  Налаштування  ")
        statusbar = tk.Frame(self, bg=COLORS["bg_header"], height=22)
        statusbar.grid(row=1, column=0, sticky="ew")
        statusbar.grid_propagate(False)
        self._status_lbl = tk.Label(statusbar, text="Готово",
            bg=COLORS["bg_header"], fg=COLORS["text_muted"],
            font=FONT_TINY, padx=10)
        self._status_lbl.pack(side="left", pady=3)
        self._time_lbl = tk.Label(statusbar, text="",
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
            self.registry_tab.save()
            self._status_lbl.configure(
                text=f"Автозбережено о {datetime.now().strftime('%H:%M:%S')}")
        except OSError:
            self._status_lbl.configure(text="Помилка автозбереження")
        self.after(30000, self._schedule_autosave)

    def save_before_exit(self) -> None:
        try:
            self.registry_tab.save()
        except OSError:
            pass



# =============================================================================
#  RISK DETAIL WINDOW
# =============================================================================

class RiskDetailWindow:
    def __init__(self, parent_root: tk.Misc, record: RiskRecord,
                 all_records: list[RiskRecord],
                 save_callback: Callable[[str, RiskRecord], None],
                 delete_callback: Callable[[str], None],
                 toast_callback: Callable[[str], None]) -> None:
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
        return {"Активний": COLORS["accent_danger"], "Монiторинг": COLORS["accent_warning"],
                "Мiтигований": COLORS["accent"], "Закрито": COLORS["text_muted"]
                }.get(status, COLORS["text_muted"])

    @staticmethod
    def _priority_color(priority: str) -> str:
        return {"Критичний": COLORS["accent_danger"], "Високий": COLORS["accent_warning"],
                "Середнiй": COLORS["accent"], "Низький": COLORS["accent_success"]
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
        self.lbl_title = tk.Label(title_frame, text=f"Ризик #{self.record.id}",
            bg=C["bg_header"], fg=C["accent_muted"], font=FONT_HEADING)
        self.lbl_title.pack(anchor="w")
        self.lbl_subtitle = tk.Label(title_frame, text=self.record.entity,
            bg=C["bg_header"], fg=C["text_muted"], font=FONT_DEFAULT)
        self.lbl_subtitle.pack(anchor="w")
        sc = self._status_color(self.record.status)
        self.lbl_status_badge = tk.Label(header,
            text=f"  {self.record.status}  ",
            bg=sc, fg="white", font=FONT_SMALL_BOLD, pady=3)
        self.lbl_status_badge.grid(row=0, column=2, padx=16, pady=18)
        self.lbl_score_badge = tk.Label(header,
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
        _scroll_mgr.attach(canvas)
        self.content.columnconfigure(0, weight=1)
        self.content.columnconfigure(1, weight=1)
        self._build_view_content()

        btn_bar = tk.Frame(self.win, bg=C["bg_header"], height=46)
        btn_bar.grid(row=2, column=0, columnspan=2, sticky="ew")
        btn_bar.grid_propagate(False)
        left_btns = tk.Frame(btn_bar, bg=C["bg_header"])
        left_btns.pack(side="left", padx=12, pady=8)
        self.btn_edit = make_button(left_btns, "Редагувати",
            bg=C["accent_warning"], fg=C["bg_main"],
            activebackground="#d97706", activeforeground="white",
            font=FONT_BOLD, padx=14, pady=4, command=self._toggle_edit_mode)
        self.btn_edit.pack(side="left", padx=(0, 8))
        self.btn_save = make_button(left_btns, "Зберегти змiни",
            bg=C["accent_success"], activebackground="#16a34a",
            font=FONT_BOLD, padx=14, pady=4, command=self._save_changes)
        self.btn_save.pack_forget()
        self.btn_cancel_edit = make_button(left_btns, "Скасувати",
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
        for w in self.content.winfo_children(): w.destroy()
        r = self.record
        row = 0
        _build_section_label(self.content, "Iнформацiя про пiдприємство", row); row += 1
        _build_info_cell(self.content, "Пiдприємство", r.entity, row, 0)
        _build_info_cell(self.content, "Прiоритет", r.priority, row, 1,
                         value_color=self._priority_color(r.priority)); row += 1
        _build_section_label(self.content, "Опис ризику", row); row += 1
        _build_info_cell(self.content, "Назва ризику",     r.risk_name, row, 0)
        _build_info_cell(self.content, "Тип ризику",       r.risk_type, row, 1,
                         value_color=RISK_COLORS.get(r.risk_type)); row += 1
        _build_info_cell(self.content, "Категорiя ризику", r.category, row, 0)
        _build_info_cell(self.content, "Власник ризику",   r.owner,    row, 1); row += 1
        _build_section_label(self.content, "Оцiнка ризику", row); row += 1
        _build_info_cell(self.content, "Iмовiрнiсть", r.probability, row, 0)
        _build_info_cell(self.content, "Вплив",         r.impact,     row, 1); row += 1
        try: score_int = int(r.score)
        except (ValueError, TypeError): score_int = 0
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
        _build_text_block(self.content, "Заходи контролю",      r.controls,    row); row += 1
        _build_text_block(self.content, "Детальний опис ризику", r.description, row); row += 1
        tk.Frame(self.content, bg=COLORS["bg_main"], height=12).grid(
            row=row, column=0, columnspan=2)

    def _build_edit_content(self) -> None:
        C = COLORS
        rec = self.record
        for w in self.content.winfo_children(): w.destroy()
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
            bg=C["bg_main"], fg=C["text_muted"], font=("Arial", 11, "bold"))
        self.lbl_live_score.grid(row=1, column=2, padx=20)

        def _upd_score(_: object = None) -> None:
            try:
                s   = _extract_num(self.e_prob.get()) * _extract_num(self.e_impact.get())
                col = _score_color(s)
                self.lbl_live_score.configure(
                    text=f"Score: {s}  ({_score_label(s)})", fg=col)
            except Exception: pass

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
            if val and val != "—": e.insert(0, val)
            else: add_placeholder(e, "дд.мм.рррр")
            setattr(self, attr, e)

        row = section("Прiоритет та статус", row)
        ps_f = tk.Frame(self.content, bg=C["bg_main"])
        ps_f.grid(row=row, column=0, sticky="w", padx=10, pady=4); row += 1
        tk.Label(ps_f, text="Прiоритет:", bg=C["bg_main"],
                 fg=C["text_subtle"], font=FONT_SMALL).grid(row=0, column=0, sticky="w")
        self.e_priority = make_dark_combo(ps_f,
            values=["Критичний", "Високий", "Середнiй", "Низький"], width=14)
        self.e_priority.grid(row=1, column=0, pady=(2, 0), padx=(0, 20))
        self.e_priority.set(rec.priority)
        tk.Label(ps_f, text="Статус:", bg=C["bg_main"],
                 fg=C["text_subtle"], font=FONT_SMALL).grid(row=0, column=1, sticky="w")
        self.e_status = make_dark_combo(ps_f,
            values=["Активний", "Монiторинг", "Мiтигований", "Закрито"], width=14)
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
        self.lbl_title.configure(text=f"Редагування ризику #{self.record.id}",
                                  fg=COLORS["accent_warning"])

    def _cancel_edit(self) -> None:
        self.is_edit_mode = False
        self._build_view_content()
        self.btn_save.pack_forget()
        self.btn_cancel_edit.pack_forget()
        self.btn_edit.pack(side="left", padx=(0, 8))
        self.lbl_title.configure(text=f"Ризик #{self.record.id}",
                                  fg=COLORS["accent_muted"])

    def _save_changes(self) -> None:
        date_id  = self.e_date_id.get().strip()
        date_rev = self.e_date_rev.get().strip()
        for val, label in [(date_id, "дати виявлення"), (date_rev, "дати перегляду")]:
            if val in ("дд.мм.рррр", ""): continue
            if not is_valid_date(val):
                messagebox.showwarning("Помилка",
                    f"Неправильний формат {label} (очiкується дд.мм.рррр)",
                    parent=self.win)
                return
        date_id  = "" if date_id  == "дд.мм.рррр" else date_id
        date_rev = "" if date_rev == "дд.мм.рррр" else date_rev
        entity    = self.e_entity.get().strip()
        risk_name = self.e_risk_name.get().strip()
        if not entity or not risk_name:
            messagebox.showwarning("Обов'язковi поля",
                "Заповнiть назву пiдприємства та назву ризику", parent=self.win)
            return
        prob_str   = self.e_prob.get().strip()
        impact_str = self.e_impact.get().strip()
        score      = _extract_num(prob_str) * _extract_num(impact_str)
        try: residual = int(self.e_residual.get().strip() or "0")
        except ValueError: residual = 0
        old_id = self.record.id
        new_record = RiskRecord(
            id=self.record.id, entity=entity, risk_name=risk_name,
            category=self.e_category.get().strip()  or "—",
            risk_type=self.e_risk_type.get().strip() or "—",
            probability=prob_str   or "—", impact=impact_str or "—",
            score=str(score), owner=self.e_owner.get().strip(),
            controls=self.e_controls.get("1.0", tk.END).strip(),
            residual=str(residual), date_id=date_id or "—",
            date_rev=date_rev or "—",
            priority=self.e_priority.get().strip() or "Середнiй",
            status=self.e_status.get().strip()   or "Активний",
            description=self.e_description.get("1.0", tk.END).strip())
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
        if not messagebox.askyesno("Пiдтвердження",
            f"Видалити ризик #{self.record.id}?\nЦю дiю не можна скасувати.",
            parent=self.win): return
        self.delete_callback(self.record.id)
        self.toast_callback("Запис видалено")
        self.win.destroy()


# =============================================================================
#  RISK REGISTRY TAB
# =============================================================================

class RiskRegistryTab(BaseRegistryTab):
    data_file = RISK_DATA_FILE

    def __init__(self, parent: tk.Misc,
                 on_data_change: Callable[[list[RiskRecord]], None] | None = None) -> None:
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
        make_button(sf, "Скинути", bg=C["bg_surface"], fg=C["text_muted"],
                    activebackground=C["bg_surface_alt"],
                    activeforeground=C["text_primary"],
                    font=FONT_SMALL, padx=8, pady=2,
                    command=self._reset_filter).pack(side="left")
        paned = ttk.PanedWindow(self.frame, orient="horizontal")
        paned.grid(row=1, column=0, sticky="nsew")
        lw = ttk.Frame(paned); rw = ttk.Frame(paned)
        paned.add(lw, weight=4); paned.add(rw, weight=7)
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
        def _conf(_):
            canvas.configure(scrollregion=canvas.bbox("all"))
            canvas.itemconfig(fw, width=canvas.winfo_width())
        form.bind("<Configure>", _conf)
        canvas.bind("<Configure>", _conf)
        _scroll_mgr.attach(canvas)
        form.columnconfigure(0, weight=1)

        def section(txt, r):
            f = tk.Frame(form, bg=C["bg_main"])
            f.grid(row=r, column=0, sticky="ew", padx=16, pady=(16, 4))
            tk.Frame(f, bg=C["accent"], width=3, height=16).pack(side="left")
            tk.Label(f, text=txt, bg=C["bg_main"], fg=C["accent"],
                     font=FONT_BOLD).pack(side="left", padx=8)
            return r + 1

        def lbl(text, r):
            tk.Label(form, text=text, bg=C["bg_main"],
                     fg=C["text_subtle"], font=FONT_SMALL).grid(
                row=r, column=0, sticky="w", padx=16, pady=(4, 0))

        def field(lbl_txt, r, factory):
            lbl(lbl_txt, r); w = factory()
            w.grid(row=r+1, column=0, sticky="ew", padx=16, pady=(2, 0))
            return w, r+2

        row = 0
        badge_f = tk.Frame(form, bg=C["bg_main"])
        badge_f.grid(row=row, column=0, sticky="ew", padx=16, pady=(12, 0))
        tk.Label(badge_f, text="  + НОВИЙ РИЗИК  ",
                 bg=C["accent"], fg="white",
                 font=FONT_SMALL_BOLD, pady=3).pack(anchor="w")
        row += 1

        row = section("Iнформацiя про пiдприємство", row)
        self.ent_entity,  row = field("Скорочена назва пiдприємства:", row,
                                       lambda: make_dark_entry(form))
        self.ent_owner,   row = field("Власник ризику:", row,
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
                row=0, column=ci, sticky="w", padx=(0 if ci==0 else 10, 0))
            combo = make_dark_combo(score_f, values=vals)
            combo.grid(row=1, column=ci, sticky="ew",
                       padx=(0 if ci==0 else 10, 0), pady=2)
            setattr(self, attr, combo)
        net_f = tk.Frame(form, bg=C["bg_main"])
        net_f.grid(row=row, column=0, sticky="ew", padx=16, pady=(2, 4)); row += 1
        tk.Label(net_f, text="Рiвень ризику (Score):",
                 bg=C["bg_main"], fg=C["text_muted"], font=FONT_SMALL).pack(side="left")
        self.lbl_score = tk.Label(net_f, text="—",
                                   bg=C["bg_main"], fg=C["accent_success"],
                                   font=("Arial", 13, "bold"))
        self.lbl_score.pack(side="left", padx=10)

        def _upd_score(_=None):
            try:
                s = _extract_num(self.cb_prob.get()) * _extract_num(self.cb_impact.get())
                self.lbl_score.configure(
                    text=f"{s}  ({_score_label(s)})", fg=_score_color(s))
            except Exception: pass

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
                row=0, column=ci, padx=(0 if ci==0 else 20, 0), sticky="w")
            e = make_dark_entry(date_f, width=14)
            e.grid(row=1, column=ci, padx=(0 if ci==0 else 20, 0), pady=2)
            add_placeholder(e, "дд.мм.рррр")
            setattr(self, attr, e)

        lbl("Прiоритет:", row); row += 1
        self.cb_priority = make_dark_combo(form,
            values=["Критичний", "Високий", "Середнiй", "Низький"])
        self.cb_priority.grid(row=row, column=0, sticky="w", padx=16, pady=(2, 0)); row += 1
        lbl("Статус:", row); row += 1
        self.cb_status = make_dark_combo(form,
            values=["Активний", "Монiторинг", "Мiтигований", "Закрито"])
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
        make_button(btn_f, "Очистити", bg=C["bg_surface"], fg=C["text_muted"],
                    activebackground=C["bg_surface_alt"],
                    activeforeground=C["text_primary"],
                    padx=14, pady=6, command=self._clear_form
                    ).grid(row=0, column=0, padx=4, sticky="ew")
        make_button(btn_f, "Додати ризик", bg=C["accent"],
                    activebackground=C["accent_soft"],
                    font=FONT_BOLD, padx=14, pady=6, command=self._add_record
                    ).grid(row=0, column=1, padx=4, sticky="ew")

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
        self.filter_status = make_dark_combo(toolbar,
            values=["Всi", "Активний", "Монiторинг", "Мiтигований", "Закрито"], width=12)
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
                        font=FONT_SMALL, padx=10, pady=3, command=cmd
                        ).pack(side="right", padx=4, pady=6)

        tree_f = ttk.Frame(container)
        tree_f.grid(row=1, column=0, sticky="nsew")
        tree_f.rowconfigure(0, weight=1); tree_f.columnconfigure(0, weight=1)
        cols = ("id", "entity", "risk_name", "category", "risk_type",
                "score", "priority", "status", "owner", "date_id")
        self.tree = ttk.Treeview(tree_f, columns=cols,
                                  show="headings", selectmode="browse")
        self.tree.grid(row=0, column=0, sticky="nsew")
        headers = {
            "id": ("№", 46), "entity": ("Пiдприємство", 150),
            "risk_name": ("Назва ризику", 200), "category": ("Категорiя", 110),
            "risk_type": ("Тип ризику", 110), "score": ("Score", 62),
            "priority": ("Прiоритет", 88), "status": ("Статус", 90),
            "owner": ("Власник", 130), "date_id": ("Дата виявл.", 100),
        }
        for col, (txt, w) in headers.items():
            self.tree.heading(col, text=txt, command=lambda c=col: self._sort_tree(c))
            self.tree.column(col, width=w, anchor="w")
        sy = ttk.Scrollbar(tree_f, orient="vertical",   command=self.tree.yview)
        sx = ttk.Scrollbar(tree_f, orient="horizontal", command=self.tree.xview)
        sy.grid(row=0, column=1, sticky="ns"); sx.grid(row=1, column=0, sticky="ew")
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
        tk.Label(hint_f, text="  Подвiйний клiк по рядку — переглянути / редагувати ризик",
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
                     padx=(12 if ci==0 else 4, 4), pady=8)
            sub.columnconfigure(0, weight=1)
            tk.Label(sub, text=lbl_t, bg=C["bg_surface"],
                     fg=C["text_muted"], font=("Arial", 7, "bold")).grid(
                row=0, column=0, sticky="w")
            t = make_dark_text(sub, height=4, wrap="word", state="disabled")
            t.grid(row=1, column=0, sticky="ew", pady=(2, 0))
            setattr(self, attr, t)

        exp_bar = tk.Frame(container, bg=C["bg_main"])
        exp_bar.grid(row=4, column=0, sticky="ew", padx=8, pady=6)
        make_button(exp_bar, "Експорт CSV", bg=C["bg_surface"], fg=C["text_primary"],
                    activebackground=C["bg_surface_alt"], activeforeground=C["text_primary"],
                    font=FONT_SMALL, padx=12, pady=4,
                    command=self._export_csv).pack(side="left", padx=(0, 6))
        if pd:
            make_button(exp_bar, "Експорт Excel", bg=C["accent_success"],
                        activebackground="#16a34a",
                        font=FONT_SMALL, padx=12, pady=4,
                        command=self._export_excel).pack(side="left", padx=(0, 6))
        make_button(exp_bar, "Iмпорт JSON", bg=C["bg_surface"], fg=C["text_primary"],
                    activebackground=C["bg_surface_alt"], activeforeground=C["text_primary"],
                    font=FONT_SMALL, padx=12, pady=4,
                    command=self._import_json).pack(side="left")

    def _on_double_click(self, event: tk.Event) -> None:
        if not self.tree.selection(): return
        if self.tree.identify("region", event.x, event.y) != "cell": return
        self._open_selected_detail()

    def _open_selected_detail(self) -> None:
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Перегляд", "Оберiть запис для перегляду"); return
        iid = sel[0]
        rec = self.find_record(self.tree.set(iid, "id"))
        if not rec: return
        RiskDetailWindow(
            parent_root=self.frame.winfo_toplevel(), record=rec,
            all_records=self.all_records,
            save_callback=lambda old, new: self._on_detail_save(iid, old, new),
            delete_callback=lambda s: self._on_detail_delete(iid, s),
            toast_callback=self._show_toast)

    def _on_detail_save(self, iid, old_id, new_record):
        normalized_old = IdGenerator.normalize_id(old_id)
        for i, r in enumerate(self.all_records):
            if IdGenerator.normalize_id(r.id) == normalized_old:
                self.all_records[i] = new_record; break
        try: self.tree.item(iid, values=self._tree_values(new_record))
        except tk.TclError: pass
        self._recolor_rows(); self._save_data(); self._notify_change()

    def _on_detail_delete(self, iid, idx_str):
        try: self.tree.delete(iid)
        except tk.TclError: pass
        normalized = IdGenerator.normalize_id(idx_str)
        self.all_records = [r for r in self.all_records
                            if IdGenerator.normalize_id(r.id) != normalized]
        self._recolor_rows(); self._save_data(); self._update_count(); self._notify_change()

    def _sort_tree(self, col):
        data = [(self.tree.set(iid, col), iid) for iid in self.tree.get_children("")]
        try: data.sort(key=lambda x: float(x[0]) if x[0] not in ("—", "") else 0)
        except ValueError: data.sort(key=lambda x: x[0].lower())
        for i, (_, iid) in enumerate(data): self.tree.move(iid, "", i)
        self._recolor_rows()

    def _recolor_rows(self):
        for i, iid in enumerate(self.tree.get_children()):
            risk      = self.tree.set(iid, "risk_type")
            score_str = self.tree.set(iid, "score")
            base_tag  = "even" if i % 2 == 0 else "odd"
            tags      = [base_tag]
            if risk in RISK_COLORS: tags.append(f"risk_{risk}")
            try:
                s = int(score_str)
                if   s <= 4:  tags.append("score_low")
                elif s <= 9:  tags.append("score_mod")
                elif s <= 16: tags.append("score_high")
                else:         tags.append("score_crit")
            except (ValueError, TypeError): pass
            self.tree.item(iid, tags=tags)

    @staticmethod
    def _tree_values(rec: RiskRecord) -> tuple:
        return (rec.id, rec.entity, rec.risk_name, rec.category,
                rec.risk_type, rec.score, rec.priority,
                rec.status, rec.owner, rec.date_id)

    def _load_data(self):
        self.all_records.clear()
        self.tree.delete(*self.tree.get_children())
        if not os.path.exists(self.data_file): return
        try:
            with open(self.data_file, "r", encoding="utf-8") as f:
                raw = json.load(f)
            if not isinstance(raw, list):
                raise ValueError("Очiкується список записiв у JSON")
            for item in raw:
                if isinstance(item, dict): rec = RiskRecord.from_dict(item)
                elif isinstance(item, (list, tuple)): rec = RiskRecord.from_list(list(item))
                else: continue
                self.all_records.append(rec)
                self._insert_tree_row(rec)
        except (json.JSONDecodeError, ValueError, OSError) as e:
            messagebox.showerror("Помилка завантаження", str(e))
        self._update_count(); self._notify_change()

    def _insert_tree_row(self, rec):
        iid = self.tree.insert("", tk.END, values=self._tree_values(rec))
        self._recolor_rows()
        return iid

    def _get_form_data(self):
        date_id  = self.ent_date_id.get().strip()
        date_rev = self.ent_date_rev.get().strip()
        for val, label in [(date_id, "дати виявлення"), (date_rev, "дати перегляду")]:
            if val in ("дд.мм.рррр", ""): continue
            if not is_valid_date(val):
                messagebox.showwarning("Помилка",
                    f"Неправильний формат {label} (очiкується дд.мм.рррр)")
                return None
        entity    = self.ent_entity.get().strip()
        risk_name = self.ent_risk_name.get().strip()
        if not entity or not risk_name:
            messagebox.showwarning("Обов\'язковi поля",
                "Заповнiть назву пiдприємства та назву ризику")
            return None
        date_id  = "" if date_id  == "дд.мм.рррр" else date_id
        date_rev = "" if date_rev == "дд.мм.рррр" else date_rev
        prob_str   = self.cb_prob.get().strip()
        impact_str = self.cb_impact.get().strip()
        score      = _extract_num(prob_str) * _extract_num(impact_str)
        try: residual = int(self.ent_residual.get().strip() or "0")
        except ValueError: residual = 0
        new_id = IdGenerator.next_id(self.all_records)
        return RiskRecord(
            id=new_id, entity=entity, risk_name=risk_name,
            category=self.cb_category.get().strip()  or "—",
            risk_type=self.cb_risk_type.get().strip() or "—",
            probability=prob_str   or "—", impact=impact_str or "—",
            score=str(score), owner=self.ent_owner.get().strip(),
            controls=self.txt_controls.get("1.0", tk.END).strip(),
            residual=str(residual), date_id=date_id or "—",
            date_rev=date_rev or "—",
            priority=self.cb_priority.get().strip() or "Середнiй",
            status=self.cb_status.get().strip()   or "Активний",
            description=self.txt_description.get("1.0", tk.END).strip())

    def _clear_form(self):
        for w in [self.ent_entity, self.ent_owner, self.ent_risk_name, self.ent_residual]:
            w.delete(0, tk.END); w.configure(fg=COLORS["text_primary"])
        for w in [self.cb_category, self.cb_risk_type, self.cb_priority,
                  self.cb_status, self.cb_prob, self.cb_impact]:
            w.set("")
        for w in [self.txt_controls, self.txt_description]: w.delete("1.0", tk.END)
        self.lbl_score.configure(text="—", fg=COLORS["accent_success"])
        for e, ph in [(self.ent_date_id, "дд.мм.рррр"), (self.ent_date_rev, "дд.мм.рррр")]:
            e.delete(0, tk.END); add_placeholder(e, ph)

    def _add_record(self):
        data = self._get_form_data()
        if not data: return
        self.all_records.append(data); self._insert_tree_row(data)
        self._clear_form(); self._save_data(); self._notify_change()
        self._update_count(); self._show_toast("Ризик додано")

    def _delete_selected(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Видалення", "Оберiть запис для видалення"); return
        iid = sel[0]; idx_str = self.tree.set(iid, "id")
        if not messagebox.askyesno("Пiдтвердження",
            f"Видалити ризик #{idx_str}?\nЦю дiю не можна скасувати."): return
        self.tree.delete(iid)
        normalized = IdGenerator.normalize_id(idx_str)
        self.all_records = [r for r in self.all_records
                            if IdGenerator.normalize_id(r.id) != normalized]
        self._recolor_rows(); self._save_data(); self._update_count()
        self._notify_change(); self._show_toast("Запис видалено")

    def _duplicate_record(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Дублювання", "Оберiть запис для дублювання"); return
        rec = self.find_record(self.tree.set(sel[0], "id"))
        if not rec: return
        import dataclasses
        new_rec = dataclasses.replace(rec, id=IdGenerator.next_id(self.all_records))
        self.all_records.append(new_rec); self._insert_tree_row(new_rec)
        self._save_data(); self._update_count(); self._notify_change()
        self._show_toast("Запис продубльовано")

    def _apply_filter(self):
        q = self.search_var.get().strip().lower()
        r_type = self.filter_type.get(); status = self.filter_status.get()
        self.tree.delete(*self.tree.get_children())
        for rec in self.all_records:
            row_str = " ".join(rec.to_list()).lower()
            if q      and q      not in row_str:           continue
            if r_type != "Всi"   and rec.risk_type != r_type: continue
            if status != "Всi"   and rec.status    != status: continue
            self._insert_tree_row(rec)
        self._update_count()

    def _reset_filter(self):
        self.search_var.set(""); self.filter_type.set("Всi"); self.filter_status.set("Всi")
        self.tree.delete(*self.tree.get_children())
        for rec in self.all_records: self._insert_tree_row(rec)
        self._update_count()

    def _on_select(self, _=None):
        sel = self.tree.selection()
        if not sel: return
        rec = self.find_record(self.tree.set(sel[0], "id"))
        if not rec: return
        for widget, text in [(self.det_controls, rec.controls),
                             (self.det_desc,     rec.description)]:
            widget.configure(state="normal")
            widget.delete("1.0", tk.END)
            widget.insert("1.0", text or "")
            widget.configure(state="disabled")

    def _update_count(self):
        self.lbl_count.configure(text=f" {len(self.tree.get_children())}")

    def _show_toast(self, msg):
        _show_toast(self.frame, msg)

    _HEADERS = ["ID", "Пiдприємство", "Назва ризику", "Категорiя", "Тип ризику",
                "Iмовiрнiсть", "Вплив", "Score", "Власник", "Заходи контролю",
                "Залишковий ризик", "Дата виявлення", "Дата перегляду",
                "Прiоритет", "Статус", "Опис"]

    def _export_csv(self):
        if not self.tree.get_children():
            messagebox.showinfo("Експорт", "Таблиця порожня"); return
        path = filedialog.asksaveasfilename(defaultextension=".csv",
            filetypes=[("CSV файли", "*.csv")], title="Зберегти реєстр ризикiв як CSV")
        if not path: return
        try:
            with open(path, "w", newline="", encoding="utf-8-sig") as f:
                w = csv.writer(f); w.writerow(self._HEADERS)
                for rec in self.all_records: w.writerow(rec.to_list())
            self._show_toast("CSV збережено")
        except OSError as e: messagebox.showerror("Помилка", str(e))

    def _export_excel(self):
        if not pd:
            messagebox.showwarning("Excel", "Встановiть pandas та openpyxl"); return
        if not self.all_records:
            messagebox.showinfo("Експорт", "Немає записiв"); return
        path = filedialog.asksaveasfilename(defaultextension=".xlsx",
            filetypes=[("Excel файли", "*.xlsx")],
            title="Зберегти реєстр ризикiв як Excel")
        if not path: return
        try:
            df = pd.DataFrame([rec.to_list() for rec in self.all_records],
                               columns=self._HEADERS)
            with pd.ExcelWriter(path, engine="openpyxl") as writer:
                df.to_excel(writer, index=False, sheet_name="Реєстр ризикiв")
                ws = writer.sheets["Реєстр ризикiв"]
                for col_cells in ws.columns:
                    mx = max(len(str(c.value or "")) for c in col_cells)
                    ws.column_dimensions[col_cells[0].column_letter].width = min(mx+4, 60)
            self._show_toast("Excel збережено")
        except (OSError, Exception) as e: messagebox.showerror("Помилка", str(e))

    def _import_json(self):
        path = filedialog.askopenfilename(filetypes=[("JSON файли", "*.json")],
                                          title="Iмпорт JSON")
        if not path: return
        try:
            with open(path, "r", encoding="utf-8") as f: data = json.load(f)
            if not isinstance(data, list):
                raise ValueError("Файл повинен мiстити список записiв")
            import dataclasses; added = 0
            for item in data:
                if isinstance(item, dict): rec = RiskRecord.from_dict(item)
                elif isinstance(item, (list, tuple)): rec = RiskRecord.from_list(list(item))
                else: continue
                rec = dataclasses.replace(rec, id=IdGenerator.next_id(self.all_records))
                self.all_records.append(rec); self._insert_tree_row(rec); added += 1
            self._save_data(); self._update_count(); self._notify_change()
            self._show_toast(f"Iмпортовано: {added} записiв")
        except (json.JSONDecodeError, ValueError, OSError) as e:
            messagebox.showerror("Помилка iмпорту", str(e))


# =============================================================================
#  RISK ANALYTICS TAB
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
        make_button(header, "Оновити", bg=C["accent"],
                    activebackground=C["accent_soft"],
                    font=FONT_BOLD, padx=12, pady=4,
                    command=self.refresh).pack(side="right", padx=20, pady=12)
        canvas = tk.Canvas(self.frame, bg=C["bg_main"], highlightthickness=0)
        sb = ttk.Scrollbar(self.frame, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=sb.set)
        canvas.grid(row=1, column=0, sticky="nsew")
        sb.grid(row=1, column=1, sticky="ns")
        self.content = tk.Frame(canvas, bg=C["bg_main"])
        cw = canvas.create_window((0, 0), window=self.content, anchor="nw")
        def _conf(_):
            canvas.configure(scrollregion=canvas.bbox("all"))
            canvas.itemconfig(cw, width=canvas.winfo_width())
        self.content.bind("<Configure>", _conf)
        canvas.bind("<Configure>", _conf)
        _scroll_mgr.attach(canvas)
        self.content.columnconfigure(0, weight=1)
        self._build_stat_cards()
        self._build_charts_and_table()

    def _build_stat_cards(self) -> None:
        C = COLORS
        cf = tk.Frame(self.content, bg=C["bg_main"])
        cf.grid(row=0, column=0, sticky="ew", padx=16, pady=(16, 0))
        for i in range(5): cf.columnconfigure(i, weight=1)
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
            self.ax_left.set_title("Розподiл за типом ризику", color=C["text_muted"], fontsize=9)
            fl = tk.Frame(cr, bg=C["bg_surface"], padx=8, pady=8)
            fl.grid(row=0, column=0, padx=(0, 8), sticky="nsew")
            self.canvas_left = FigureCanvasTkAgg(self.fig_left, master=fl)
            self.canvas_left.get_tk_widget().pack(fill="both", expand=True)
            self.fig_right = Figure(figsize=(5, 3.5), dpi=90, facecolor=C["bg_surface"])
            self.ax_right  = self.fig_right.add_subplot(111)
            self._style_ax(self.ax_right)
            self.ax_right.set_title("Розподiл за рiвнем ризику", color=C["text_muted"], fontsize=9)
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

    def _style_ax(self, ax) -> None:
        C = COLORS
        ax.set_facecolor(C["bg_surface"])
        ax.tick_params(colors=C["text_muted"], labelsize=8)
        for spine in ax.spines.values():
            spine.set_edgecolor(C["border_soft"])

    def update_data(self, records: list[RiskRecord]) -> None:
        self.records = records; self.refresh()

    def refresh(self) -> None:
        if not self.records:
            for k in self.stat_cards: self.stat_cards[k].configure(text="0")
            if HAS_MPL: self._clear_charts()
            self.stats_tree.delete(*self.stats_tree.get_children()); return
        C = COLORS; records = self.records
        total    = len(records)
        critical = sum(1 for r in records if r.score.isdigit() and int(r.score) > 16)
        high_    = sum(1 for r in records if r.score.isdigit() and 10 <= int(r.score) <= 16)
        active_  = sum(1 for r in records if r.status == "Активний")
        mitig_   = sum(1 for r in records if r.status == "Мiтигований")
        self.stat_cards["total"].configure(text=str(total))
        self.stat_cards["critical"].configure(text=str(critical))
        self.stat_cards["high"].configure(text=str(high_))
        self.stat_cards["active"].configure(text=str(active_))
        self.stat_cards["mitigated"].configure(text=str(mitig_))
        if not HAS_MPL: return
        type_ctr = Counter(r.risk_type for r in records)
        self.ax_left.clear(); self._style_ax(self.ax_left)
        self.ax_left.set_title("Розподiл за типом ризику", color=C["text_muted"], fontsize=9)
        if type_ctr:
            lbls = list(type_ctr.keys()); vals = list(type_ctr.values())
            clrs = [RISK_COLORS.get(l, C["text_muted"]) for l in lbls]
            _, _, autotexts = self.ax_left.pie(vals, labels=lbls, autopct="%1.0f%%",
                colors=clrs, startangle=90,
                textprops={"color": C["text_muted"], "fontsize": 7})
            for at in autotexts: at.set_fontsize(7); at.set_color("white")
        self.canvas_left.draw()
        self.ax_right.clear(); self._style_ax(self.ax_right)
        self.ax_right.set_title("Розподiл за рiвнем ризику", color=C["text_muted"], fontsize=9)
        level_ctr = {"Низький": 0, "Помiрний": 0, "Високий": 0, "Критичний": 0}
        for r in records:
            if r.score.isdigit(): level_ctr[_score_label(int(r.score))] += 1
        if any(level_ctr.values()):
            lbls = list(level_ctr.keys()); vals = list(level_ctr.values())
            clrs = [C["accent_success"], C["accent_warning"], "#f97316", C["accent_danger"]]
            bars = self.ax_right.bar(lbls, vals, color=clrs, edgecolor="none")
            for bar, val in zip(bars, vals, strict=False):
                if val > 0:
                    self.ax_right.text(bar.get_x() + bar.get_width()/2,
                        bar.get_height()+0.1, str(val), ha="center", va="bottom",
                        color=C["text_muted"], fontsize=8)
            self.ax_right.tick_params(axis="x", labelrotation=15, labelsize=7)
            mx = max(vals); self.ax_right.set_ylim(0, mx*1.2+1 if mx > 0 else 1)
        self.canvas_right.draw()
        self.ax_heat.clear(); self._style_ax(self.ax_heat)
        self.ax_heat.set_title("Матриця ризикiв (Iмовiрнiсть × Вплив)",
                                color=C["text_muted"], fontsize=9)
        matrix = [[0]*5 for _ in range(5)]
        for r in records:
            try:
                prob = _extract_num(r.probability); imp = _extract_num(r.impact)
                if 1 <= prob <= 5 and 1 <= imp <= 5:
                    matrix[5-prob][imp-1] += 1
            except (ValueError, IndexError): pass
        if any(any(row) for row in matrix) and np is not None:
            self.ax_heat.imshow(matrix, cmap=plt.cm.RdYlGn_r, aspect="auto")
            self.ax_heat.set_xticks(range(5)); self.ax_heat.set_yticks(range(5))
            self.ax_heat.set_xticklabels([str(i) for i in range(1, 6)])
            self.ax_heat.set_yticklabels([str(i) for i in range(5, 0, -1)])
            self.ax_heat.set_xlabel("Вплив →",       color=C["text_muted"], fontsize=8)
            self.ax_heat.set_ylabel("Iмовiрнiсть →", color=C["text_muted"], fontsize=8)
            for i in range(5):
                for j in range(5):
                    if matrix[i][j] > 0:
                        self.ax_heat.text(j, i, str(matrix[i][j]),
                            ha="center", va="center", color="white",
                            fontsize=10, weight="bold")
        self.canvas_heat.draw()
        self.stats_tree.delete(*self.stats_tree.get_children())
        all_types = set(RISK_TYPES) | {r.risk_type for r in records}
        for rt in sorted(all_types):
            recs = [r for r in records if r.risk_type == rt]
            cnt  = len(recs)
            if cnt:
                scores    = [int(r.score) for r in recs if r.score.isdigit()]
                avg_score = sum(scores)/len(scores) if scores else 0
                max_score = max(scores) if scores else 0
                act       = sum(1 for r in recs if r.status == "Активний")
                self.stats_tree.insert("", tk.END,
                    values=(rt, cnt, f"{avg_score:.1f}", max_score, act))

    def _clear_charts(self) -> None:
        if not HAS_MPL: return
        for ax in (self.ax_left, self.ax_right, self.ax_heat): ax.clear()
        for cv in (self.canvas_left, self.canvas_right, self.canvas_heat): cv.draw()

    def get_frame(self) -> ttk.Frame:
        return self.frame


# =============================================================================
#  RISK SETTINGS TAB
# =============================================================================

class RiskSettingsTab:
    def __init__(self, parent: tk.Misc) -> None:
        self.frame = ttk.Frame(parent)
        self._build_ui()

    def _build_ui(self) -> None:
        C = COLORS
        self.frame.columnconfigure(0, weight=1); self.frame.rowconfigure(1, weight=1)
        header = tk.Frame(self.frame, bg=C["bg_header"], height=56)
        header.grid(row=0, column=0, sticky="ew"); header.grid_propagate(False)
        tk.Label(header, text="НАЛАШТУВАННЯ РЕЄСТРУ РИЗИКIВ",
                 bg=C["bg_header"], fg=C["text_muted"],
                 font=FONT_TITLE).pack(side="left", padx=20, pady=14)
        content = tk.Frame(self.frame, bg=C["bg_main"])
        content.grid(row=1, column=0, sticky="nsew", padx=40, pady=30)
        content.columnconfigure(0, weight=1)
        for row, (label, value) in enumerate([
            ("Файл даних:", RISK_DATA_FILE),
            ("Версiя:", "2.3 — OR інтегровано в ризик-апетит"),
            ("matplotlib:", "встановлено" if HAS_MPL else "не встановлено"),
            ("pandas:", "встановлено" if pd else "не встановлено"),
        ]):
            f = tk.Frame(content, bg=C["bg_main"])
            f.grid(row=row, column=0, sticky="ew", pady=4)
            tk.Label(f, text=label, bg=C["bg_main"], fg=C["text_muted"],
                     font=FONT_DEFAULT, width=22, anchor="w").pack(side="left")
            tk.Label(f, text=value, bg=C["bg_main"], fg=C["text_primary"],
                     font=FONT_DEFAULT).pack(side="left")

    def get_frame(self) -> ttk.Frame:
        return self.frame


# =============================================================================
#  RISK REGISTER PAGE
# =============================================================================

class RiskRegisterPage(tk.Frame):
    def __init__(self, master: tk.Misc, **kwargs) -> None:
        super().__init__(master, bg=COLORS["bg_main"], **kwargs)
        self.grid_rowconfigure(0, weight=1); self.grid_rowconfigure(1, weight=0)
        self.grid_columnconfigure(0, weight=1)
        self.notebook = ttk.Notebook(self)
        self.notebook.grid(row=0, column=0, sticky="nsew")
        self.analytics_tab = RiskAnalyticsTab(self.notebook)
        self.registry_tab  = RiskRegistryTab(self.notebook,
            on_data_change=self.analytics_tab.update_data)
        self.settings_tab  = RiskSettingsTab(self.notebook)
        self.notebook.add(self.registry_tab.get_frame(),  text="  Реєстр ризикiв  ")
        self.notebook.add(self.analytics_tab.get_frame(), text="  Аналiтика  ")
        self.notebook.add(self.settings_tab.get_frame(),  text="  Налаштування  ")
        statusbar = tk.Frame(self, bg=COLORS["bg_header"], height=22)
        statusbar.grid(row=1, column=0, sticky="ew"); statusbar.grid_propagate(False)
        self._status_lbl = tk.Label(statusbar, text="Готово",
            bg=COLORS["bg_header"], fg=COLORS["text_muted"],
            font=FONT_TINY, padx=10)
        self._status_lbl.pack(side="left", pady=3)
        self._time_lbl = tk.Label(statusbar, text="",
            bg=COLORS["bg_header"], fg=COLORS["text_muted"],
            font=FONT_TINY, padx=10)
        self._time_lbl.pack(side="right", pady=3)
        self._start_clock(); self._schedule_autosave()
        self.after(600, lambda: self.analytics_tab.update_data(
            self.registry_tab.all_records))

    def _start_clock(self) -> None:
        self._time_lbl.configure(text=datetime.now().strftime("%d.%m.%Y  %H:%M:%S"))
        self.after(1000, self._start_clock)

    def _schedule_autosave(self) -> None:
        try:
            self.registry_tab.save()
            self._status_lbl.configure(
                text=f"Автозбережено о {datetime.now().strftime('%H:%M:%S')}")
        except OSError:
            self._status_lbl.configure(text="Помилка автозбереження")
        self.after(30000, self._schedule_autosave)

    def save_before_exit(self) -> None:
        try: self.registry_tab.save()
        except OSError: pass


# OPERATIONAL_INDICATORS (8 OR-індикаторів)
OPERATIONAL_INDICATORS = [
    {
        "code": "OR 1.1",
        "group": "OR1",
        "name": "Рівень збитків від внутрішнього шахрайства",
        "mode": "or_fraud",
        "risk_type": "Внутрішнє шахрайство",
        "desc": (
            "Збитки від дій власного персоналу (крадіжка, привласнення, шахрайство). "
            "Поріг толерантності обирається за рівнем річного доходу підприємства."
        ),
        "fields": [],   # поля задаються динамічно в IndicatorFormFrame
        "incident_cols": ["Дата", "Сума збитків (тис. грн)", "Опис", "Примітка"],
        "incident_keys": ["date", "amount", "description", "note"],
        "thresholds": {
            "Green":  "0 — збитків немає",
            "Yellow": "Сума < порогу толерантності",
            "Orange": "Сума = порогу толерантності",
            "Red":    "Сума > порогу толерантності",
        },
    },
    {
        "code": "OR 2.1",
        "group": "OR2",
        "name": "Рівень збитків від зовнішнього шахрайства",
        "mode": "or_fraud",
        "risk_type": "Зовнішнє шахрайство",
        "desc": (
            "Збитки від дій третіх осіб (шахрайство, кібератаки, підробка документів). "
            "Логіка та поріг толерантності ідентичні OR 1.1."
        ),
        "fields": [],
        "incident_cols": ["Дата", "Сума збитків (тис. грн)", "Опис", "Примітка"],
        "incident_keys": ["date", "amount", "description", "note"],
        "thresholds": {
            "Green":  "0 — збитків немає",
            "Yellow": "Сума < порогу",
            "Orange": "Сума = порогу",
            "Red":    "Сума > порогу",
        },
    },
    {
        "code": "OR 3.1",
        "group": "OR3",
        "name": "Час простою",
        "mode": "or_downtime",
        "risk_type": "Недосконалі процеси",
        "desc": (
            "Частка непродуктивного часу від загального фонду робочого часу за квартал. "
            "Поріг: 10% загального фонду часу."
        ),
        "fields": [
            ("hours_per_day", "Годин на день"),
            ("days_m1",       "Роб. днів М1"),
            ("days_m2",       "Роб. днів М2"),
            ("days_m3",       "Роб. днів М3"),
            ("workers",       "Чисельність (кін. кварталу)"),
        ],
        "incident_cols": ["Дата", "Підрозділ/Цех", "Час простою (год)", "Опис", "Примітка"],
        "incident_keys": ["date", "department", "hours", "description", "note"],
        "thresholds": {
            "Green":  "0 — простоїв не було",
            "Yellow": "Частка < 10%",
            "Orange": "Частка = 10%",
            "Red":    "Частка > 10%",
        },
    },
    {
        "code": "OR 4.1",
        "group": "OR4",
        "name": "Рівень збитків від неправомірних дій щодо управління персоналом",
        "mode": "or_fraud",
        "risk_type": "Управління персоналом",
        "desc": (
            "Виплати на користь (колишніх) працівників поза рамками стандартних нарахувань "
            "(рішення суду, домовленості, штрафи тощо)."
        ),
        "fields": [],
        "incident_cols": ["Дата", "№ справи", "Сума виплат (тис. грн)", "Опис", "Примітка"],
        "incident_keys": ["date", "case_number", "amount", "description", "note"],
        "thresholds": {
            "Green":  "0 — збитків немає",
            "Yellow": "Сума < порогу",
            "Orange": "Сума = порогу",
            "Red":    "Сума > порогу",
        },
    },
    {
        "code": "OR 5.1",
        "group": "OR5",
        "name": "Рівень дебіторської заборгованості",
        "mode": "or_pct_income",
        "risk_type": "Взаємовідносини з контрагентами",
        "desc": (
            "Прострочена дебіторська заборгованість (>90 днів) + штрафи/пеня. "
            "Поріг: 1% чистого доходу підприємства (Ф2, рядок 2000)."
        ),
        "fields": [
            ("net_income", "Чистий дохід (Ф2 р.2000), тис. грн"),
        ],
        "incident_cols": ["Дата", "Боржник", "Сума (тис. грн)", "Опис", "Примітка"],
        "incident_keys": ["date", "counterpart", "amount", "description", "note"],
        "thresholds": {
            "Green":  "0 — заборгованості немає",
            "Yellow": "Сума < 1% доходу",
            "Orange": "Сума = 1% доходу",
            "Red":    "Сума > 1% доходу",
        },
    },
    {
        "code": "OR 7.1",
        "group": "OR7",
        "name": "Рівень збитків від пошкодження активів",
        "mode": "or_pct_cumul",
        "risk_type": "Пошкодження активів",
        "desc": (
            "Збитки від пошкодження матеріальних активів (аварії, стихійні лиха, вандалізм). "
            "Поріг: 1% чистого доходу. Сума рахується наростаючим підсумком з початку року."
        ),
        "fields": [
            ("net_income",     "Чистий дохід (Ф2 р.2000), тис. грн"),
            ("cumul_prev",     "Накопичено з попередніх кварталів (тис. грн)"),
        ],
        "incident_cols": ["Дата", "Сума збитків (тис. грн)", "Опис", "Примітка"],
        "incident_keys": ["date", "amount", "description", "note"],
        "thresholds": {
            "Green":  "0 — збитків немає",
            "Yellow": "Наростаюча сума < 1% доходу",
            "Orange": "Наростаюча сума = 1% доходу",
            "Red":    "Наростаюча сума > 1% доходу",
        },
    },
    {
        "code": "OR 8.1",
        "group": "OR8",
        "name": "Сума позовних вимог",
        "mode": "or_pct_cumul",
        "risk_type": "Юридичні ризики",
        "desc": (
            "Майнові позови до підприємства без остаточного рішення або ті, що будуть оскаржені. "
            "Поріг: 1% чистого доходу. Наростаючий підсумок за рік."
        ),
        "fields": [
            ("net_income",  "Чистий дохід (Ф2 р.2000), тис. грн"),
            ("cumul_prev",  "Накопичено з попередніх кварталів (тис. грн)"),
        ],
        "incident_cols": ["Дата", "№ справи", "Сума позову (тис. грн)", "Опис", "Примітка"],
        "incident_keys": ["date", "case_number", "amount", "description", "note"],
        "thresholds": {
            "Green":  "0 — позовів немає",
            "Yellow": "Наростаюча сума < 1% доходу",
            "Orange": "Наростаюча сума = 1% доходу",
            "Red":    "Наростаюча сума > 1% доходу",
        },
    },
    {
        "code": "OR 10.1",
        "group": "OR10",
        "name": "Рівень неприйнятої замовником продукції / робіт / послуг",
        "mode": "or_pct_income",
        "risk_type": "Технологічний ризик",
        "desc": (
            "Загальна сума продукції/робіт/послуг, не прийнятих замовником через невідповідність. "
            "Поріг: 1% чистого доходу."
        ),
        "fields": [
            ("net_income", "Чистий дохід (Ф2 р.2000), тис. грн"),
        ],
        "incident_cols": ["Дата", "Сума (тис. грн)", "Опис", "Примітка"],
        "incident_keys": ["date", "amount", "description", "note"],
        "thresholds": {
            "Green":  "0 — відмов немає",
            "Yellow": "Сума < 1% доходу",
            "Orange": "Сума = 1% доходу",
            "Red":    "Сума > 1% доходу",
        },
    },
]

# Пороги для fraud-режиму


# =============================================================================
#  OR CONSTANTS, INDICATORS, COMPUTE ENGINE & FORM WIDGETS
# =============================================================================

_OR_FRAUD_THRESHOLDS = [
    ("до 300 млн грн доходу → 300 тис. грн",   300.0),    # тис. грн
    ("300–1000 млн грн доходу → 1 млн грн",  1_000.0),
    ("понад 1000 млн грн → 5 млн грн",        5_000.0),
]


# ─────────────────────────────────────────────────────────────────────────────
#  ЗАМІНА Б: РОЗШИРЕНА compute_indicator_level
#  Вставити замість існуючої функції compute_indicator_level
# ─────────────────────────────────────────────────────────────────────────────

def compute_indicator_level(spec: dict, values: dict[str, str]) -> tuple[str, str]:
    """
    Повертає (level, detail).
    Підтримує оригінальні режими + нові OR-режими.
    """
    mode = spec.get("mode", "")

    # ── OR-режими ────────────────────────────────────────────────────────────

    if mode == "or_fraud":
        incidents = values.get("__incidents__", [])
        total = sum(_safe_num(i.get("amount", "0")) for i in incidents)
        thr_idx = int(_safe_num(values.get("__threshold_idx__", "0")))
        thr     = _OR_FRAUD_THRESHOLDS[min(thr_idx, 2)][1]   # тис. грн
        detail  = f"Сума: {total:,.2f} тис. грн | Поріг: {thr:,.0f} тис. грн"
        if total == 0:
            return "Green",  detail
        elif total < thr:
            return "Yellow", detail
        elif abs(total - thr) < 0.01:
            return "Orange", detail
        else:
            return "Red",    detail

    elif mode == "or_downtime":
        incidents  = values.get("__incidents__", [])
        total_h    = sum(_safe_num(i.get("hours", "0")) for i in incidents)
        h_day   = _safe_num(values.get("hours_per_day", "8"))
        d_m1    = _safe_num(values.get("days_m1", "0"))
        d_m2    = _safe_num(values.get("days_m2", "0"))
        d_m3    = _safe_num(values.get("days_m3", "0"))
        workers = _safe_num(values.get("workers", "0"))
        work_h  = h_day * (d_m1 + d_m2 + d_m3) * workers
        if work_h == 0:
            return "Green", "Загальний час не введено"
        pct    = total_h / work_h * 100
        detail = f"Простій: {total_h:.1f} год / {work_h:.1f} год = {pct:.2f}%"
        if total_h == 0:
            return "Green",  detail
        elif pct < 10:
            return "Yellow", detail
        elif abs(pct - 10) < 0.01:
            return "Orange", detail
        else:
            return "Red",    detail

    elif mode in ("or_pct_income", "or_pct_cumul"):
        incidents = values.get("__incidents__", [])
        total     = sum(_safe_num(i.get("amount", "0")) for i in incidents)
        if mode == "or_pct_cumul":
            total += _safe_num(values.get("cumul_prev", "0"))
        income    = _safe_num(values.get("net_income", "0"))
        if income == 0:
            return "Green", "Чистий дохід не введено"
        thr    = income * 0.01
        pct    = total / income * 100
        detail = f"Сума: {total:,.2f} | 1% доходу: {thr:,.2f} | {pct:.2f}%"
        if total == 0:
            return "Green",  detail
        elif total < thr:
            return "Yellow", detail
        elif abs(total - thr) < 0.01:
            return "Orange", detail
        else:
            return "Red",    detail

    # ── Оригінальні режими (без змін) ────────────────────────────────────────

    if mode == "steps3":
        plan_p  = _safe_num(values.get("plan_profit",  "0"))
        fact_p  = _safe_num(values.get("fact_profit",  "0"))
        plan_e  = _safe_num(values.get("plan_expense", "0"))
        fact_e  = _safe_num(values.get("fact_expense", "0"))
        plan_i  = _safe_num(values.get("plan_income",  "0"))
        fact_i  = _safe_num(values.get("fact_income",  "0"))
        plan_n  = _safe_num(values.get("plan_net",     "0"))
        fact_n  = _safe_num(values.get("fact_net",     "0"))

        failed  = 0
        details = []
        if plan_p != 0:
            dev1 = abs(fact_p - plan_p) / abs(plan_p) * 100
            if not (fact_p >= plan_p or dev1 < 10):
                failed += 1
                details.append(f"Прибуток: відхилення {dev1:.1f}%")
        if plan_e != 0:
            dev2 = (fact_e - plan_e) / abs(plan_e) * 100
            income_grew = (plan_i != 0 and fact_i > plan_i * 1.05)
            if not (dev2 <= 10 or income_grew):
                failed += 1
                details.append(f"Видатки: перевищення {dev2:.1f}%")
        if plan_n != 0:
            dev3 = abs(fact_n - plan_n) / abs(plan_n) * 100
            if not (fact_n >= plan_n or dev3 < 10):
                failed += 1
                details.append(f"Чист. прибуток: відхилення {dev3:.1f}%")
        levels = ["Green", "Yellow", "Orange", "Red"]
        detail = "; ".join(details) if details else "Всі кроки OK"
        return levels[min(failed, 3)], detail

    elif mode == "pct_income":
        field_keys = [f[0] for f in spec.get("fields", [])]
        if "debt" in field_keys and "penalty" in field_keys:
            num = _safe_num(values.get("debt", "0")) + _safe_num(values.get("penalty", "0"))
        elif "loss" in field_keys:
            num = _safe_num(values.get("loss", "0"))
        elif "lawsuits_amount" in field_keys:
            num = _safe_num(values.get("lawsuits_amount", "0"))
        elif "sla_breaches" in field_keys:
            num = _safe_num(values.get("sla_breaches", "0"))
        else:
            num = _safe_num(values.get(field_keys[0] if field_keys else "", "0"))
        denom_key = next((k for k in field_keys if k in ("income", "total_requests")), None)
        denom = _safe_num(values.get(denom_key, "0")) if denom_key else 0.0
        if denom == 0:
            return "—", "Знаменник = 0"
        pct    = num / denom * 100
        detail = f"Частка = {pct:.2f}%"
        if pct < 1:   return "Green",  detail
        elif pct == 1: return "Yellow", detail
        elif pct <= 2: return "Orange", detail
        else:          return "Red",    detail

    elif mode == "ratio":
        field_keys = [f[0] for f in spec.get("fields", [])]
        num   = _safe_num(values.get(field_keys[0] if field_keys else "", "0"))
        denom = _safe_num(values.get(field_keys[1] if len(field_keys) > 1 else "", "0"))
        if denom == 0:
            return "—", "Знаменник = 0"
        ratio  = num / denom
        code   = spec.get("code", "")
        detail = f"Коефіцієнт = {ratio:.3f}"
        if "FR 1.1" in code:
            if ratio > 2.0:    return "Green",  detail
            elif ratio >= 1.5: return "Yellow", detail
            elif ratio >= 1.0: return "Orange", detail
            else:              return "Red",    detail
        elif "FR 1.2" in code:
            pct = ratio * 100
            det = f"ROA = {pct:.2f}%"
            if pct > 5:    return "Green",  det
            elif pct >= 2: return "Yellow", det
            elif pct >= 0: return "Orange", det
            else:          return "Red",    det
        elif "FR 2.1" in code:
            if ratio < 2:   return "Green",  detail
            elif ratio < 3: return "Yellow", detail
            elif ratio < 4: return "Orange", detail
            else:           return "Red",    detail
        return "—", detail

    elif mode in ("count4", "count3"):
        field_keys = [f[0] for f in spec.get("fields", [])]
        if "required" in field_keys:
            impl_key = next((k for k in field_keys if k in ("implemented", "executed")), None)
            required = _safe_num(values.get("required", "0"))
            done     = _safe_num(values.get(impl_key, "0")) if impl_key else 0
            count    = max(0, required - done)
            detail   = f"Не виконано: {int(count)} з {int(required)}"
        else:
            count  = sum(_safe_num(values.get(k, "0")) for k in field_keys)
            detail = f"Кількість: {int(count)}"
        if mode == "count3":
            if count == 0:    return "Green",  detail
            elif count == 1:  return "Yellow", detail
            else:             return "Red",    detail
        else:
            if count == 0:    return "Green",  detail
            elif count == 1:  return "Yellow", detail
            elif count == 2:  return "Orange", detail
            else:             return "Red",    detail

    return "—", "Режим розрахунку не визначено"


def _safe_num(val: object) -> float:
    """Безпечна конвертація у float."""
    try:
        return float(str(val).replace(" ", "").replace(",", ".") or "0")
    except (ValueError, TypeError):
        return 0.0


# ─────────────────────────────────────────────────────────────────────────────
#  ЗАМІНА В: РОЗШИРЕНИЙ IndicatorFormFrame
#  Вставити замість існуючого класу IndicatorFormFrame
# ─────────────────────────────────────────────────────────────────────────────

class IndicatorFormFrame(tk.Frame):
    """
    Розгортна форма введення даних для одного індикатора.
    Підтримує прості числові поля (оригінал) та OR-режими з табличними реєстрами.
    """

    def __init__(
        self,
        parent:       tk.Misc,
        spec:         dict,
        saved_values: dict[str, str],
        saved_notes:  str,
        on_change:    Callable[[str, dict, str, str], None],
        row:          int,
    ) -> None:
        C = COLORS
        super().__init__(parent, bg=C["bg_main"])
        self.spec         = spec
        self._on_change   = on_change
        self._expanded    = False
        self._vars:   dict[str, tk.StringVar] = {}
        self._note_var = tk.StringVar(value=saved_notes)
        # OR-specific
        self._incidents: list[dict] = list(saved_values.get("__incidents__", []))
        self._threshold_var = tk.IntVar(
            value=int(saved_values.get("__threshold_idx__", 0)))

        self._level, self._detail = compute_indicator_level(
            spec, self._build_values_dict(saved_values))

        self.grid(row=row, column=0, sticky="ew", padx=12, pady=(4, 0))
        self.columnconfigure(0, weight=1)
        self._build(saved_values)

    # ── HEADER ────────────────────────────────────────────────────────────

    def _build(self, saved: dict[str, str]) -> None:
        C  = COLORS
        sp = self.spec

        hdr = tk.Frame(self, bg=C["bg_surface"], cursor="hand2")
        hdr.grid(row=0, column=0, sticky="ew")
        hdr.columnconfigure(2, weight=1)

        lvl_color = RA_COLORS.get(self._level, C["border_soft"])
        self._strip = tk.Frame(hdr, bg=lvl_color, width=5)
        self._strip.grid(row=0, column=0, sticky="ns")

        code_lbl = tk.Label(hdr, text=f"  {sp['code']}",
                             bg=C["bg_surface"], fg=C["accent_muted"],
                             font=FONT_SMALL_BOLD, width=9, anchor="w")
        code_lbl.grid(row=0, column=1, padx=(4, 0), pady=6)

        name_lbl = tk.Label(hdr, text=sp["name"],
                             bg=C["bg_surface"], fg=C["text_primary"],
                             font=FONT_BOLD, anchor="w")
        name_lbl.grid(row=0, column=2, sticky="ew", padx=6, pady=6)

        self._badge = tk.Label(
            hdr,
            text=f"  {RA_LABELS.get(self._level, '—')}  ",
            bg=lvl_color, fg="white", font=FONT_SMALL_BOLD, pady=2)
        self._badge.grid(row=0, column=3, padx=8, pady=6)

        self._toggle_btn = make_button(
            hdr, "▶", bg=C["bg_surface"], fg=C["text_muted"],
            activebackground=C["bg_surface_alt"],
            font=FONT_DEFAULT, padx=8, pady=3,
            command=self._toggle)
        self._toggle_btn.grid(row=0, column=4, padx=(0, 6), pady=4)

        for w in (hdr, code_lbl, name_lbl):
            w.bind("<Button-1>", lambda _: self._toggle())

        # ── Body ─────────────────────────────────────────────────────────
        self._body = tk.Frame(self, bg=C["bg_surface_alt"])
        self._body.columnconfigure(0, weight=1)
        self._body.columnconfigure(1, weight=1)

        mode = sp.get("mode", "")
        is_or = mode.startswith("or_")

        # Description
        tk.Label(self._body, text=sp.get("desc", ""),
                 bg=C["bg_surface_alt"], fg=C["text_muted"],
                 font=("Arial", 8, "italic"),
                 justify="left", wraplength=700, anchor="w").grid(
            row=0, column=0, columnspan=2,
            sticky="ew", padx=14, pady=(8, 6))

        body_row = 1

        if is_or:
            body_row = self._build_or_body(saved, mode, body_row)
        else:
            body_row = self._build_standard_body(saved, body_row)

        # ── Result row ────────────────────────────────────────────────────
        res_f = tk.Frame(self._body, bg=C["bg_surface_alt"])
        res_f.grid(row=body_row, column=0, columnspan=2,
                   sticky="ew", padx=14, pady=(10, 4))
        tk.Label(res_f, text="Результат:",
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
        body_row += 1

        # Thresholds legend
        thr_f = tk.Frame(self._body, bg=C["bg_surface_alt"])
        thr_f.grid(row=body_row, column=0, columnspan=2,
                   sticky="ew", padx=14, pady=(2, 4))
        tk.Label(thr_f, text="Шкала: ",
                 bg=C["bg_surface_alt"], fg=C["text_subtle"],
                 font=FONT_TINY).pack(side="left")
        for lvl, txt in sp.get("thresholds", {}).items():
            bg = RA_COLORS.get(lvl, COLORS["border_soft"])
            tk.Label(thr_f, text=f" {txt} ",
                     bg=bg, fg="white", font=FONT_TINY).pack(side="left", padx=2)
        body_row += 1

        # Note
        tk.Label(self._body, text="Коментар / примітка:",
                 bg=C["bg_surface_alt"], fg=C["text_subtle"],
                 font=FONT_SMALL).grid(
            row=body_row, column=0, columnspan=2,
            sticky="w", padx=14, pady=(8, 0))
        body_row += 1
        note_e = make_dark_entry(self._body)
        note_e.grid(row=body_row, column=0, columnspan=2,
                    sticky="ew", padx=14, pady=(2, 10))
        note_e.configure(textvariable=self._note_var)
        self._note_var.trace_add("write", lambda *_: self._recalc())

    # ── OR body builder ───────────────────────────────────────────────────

    def _build_or_body(self, saved: dict, mode: str, start_row: int) -> int:
        C  = COLORS
        sp = self.spec
        r  = start_row

        # ── Fraud: threshold selector ─────────────────────────────────────
        if mode == "or_fraud":
            tk.Label(self._body,
                     text="Поріг толерантності (за рівнем річного доходу):",
                     bg=C["bg_surface_alt"], fg=C["text_muted"],
                     font=FONT_SMALL_BOLD).grid(
                row=r, column=0, columnspan=2, sticky="w", padx=14, pady=(4, 2))
            r += 1
            for i, (label, _) in enumerate(_OR_FRAUD_THRESHOLDS):
                tk.Radiobutton(
                    self._body, text=label,
                    variable=self._threshold_var, value=i,
                    bg=C["bg_surface_alt"], fg=C["text_primary"],
                    activebackground=C["bg_surface_alt"],
                    selectcolor=C["bg_main"],
                    font=FONT_SMALL,
                    command=self._recalc,
                ).grid(row=r, column=0, columnspan=2, sticky="w", padx=14, pady=1)
                r += 1

        # ── Downtime: work-time fields ────────────────────────────────────
        elif mode == "or_downtime":
            fields = sp.get("fields", [])
            inp_f  = tk.Frame(self._body, bg=C["bg_surface_alt"])
            inp_f.grid(row=r, column=0, columnspan=2, sticky="ew", padx=14, pady=4)
            for ci, (key, lbl_t) in enumerate(fields):
                inp_f.columnconfigure(ci, weight=1)
                tk.Label(inp_f, text=lbl_t + ":",
                         bg=C["bg_surface_alt"], fg=C["text_subtle"],
                         font=FONT_TINY).grid(row=0, column=ci, sticky="w",
                                               padx=(0 if ci == 0 else 8, 0), pady=(4, 0))
                var = tk.StringVar(value=saved.get(key, ""))
                self._vars[key] = var
                var.trace_add("write", lambda *_: self._recalc())
                e = make_dark_entry(inp_f, width=9)
                e.grid(row=1, column=ci, padx=(0 if ci == 0 else 8, 0), pady=2, sticky="ew")
                e.configure(textvariable=var)
            r += 1

        # ── Income fields (pct_income / pct_cumul) ────────────────────────
        elif mode in ("or_pct_income", "or_pct_cumul"):
            fields = sp.get("fields", [])
            inp_f  = tk.Frame(self._body, bg=C["bg_surface_alt"])
            inp_f.grid(row=r, column=0, columnspan=2, sticky="ew", padx=14, pady=4)
            for ci, (key, lbl_t) in enumerate(fields):
                inp_f.columnconfigure(ci, weight=1)
                tk.Label(inp_f, text=lbl_t + ":",
                         bg=C["bg_surface_alt"], fg=C["text_subtle"],
                         font=FONT_SMALL).grid(row=0, column=ci, sticky="w",
                                                padx=(0 if ci == 0 else 12, 0))
                var = tk.StringVar(value=saved.get(key, ""))
                self._vars[key] = var
                var.trace_add("write", lambda *_: self._recalc())
                e = make_dark_entry(inp_f, width=18)
                e.grid(row=1, column=ci, padx=(0 if ci == 0 else 12, 0),
                       pady=(2, 8), sticky="ew")
                e.configure(textvariable=var)
            r += 1

        # ── Incident table ────────────────────────────────────────────────
        inc_cols = sp.get("incident_cols", [])
        inc_keys = sp.get("incident_keys", [])

        tk.Label(self._body,
                 text="Реєстр інцидентів / подій за квартал:",
                 bg=C["bg_surface_alt"], fg=C["text_muted"],
                 font=FONT_SMALL_BOLD).grid(
            row=r, column=0, columnspan=2, sticky="w", padx=14, pady=(10, 2))
        r += 1

        inc_container = tk.Frame(self._body, bg=C["bg_surface_alt"])
        inc_container.grid(row=r, column=0, columnspan=2, sticky="ew", padx=14, pady=0)
        inc_container.columnconfigure(0, weight=1)
        r += 1

        self._inc_table = _IncidentTableWidget(
            inc_container,
            cols=inc_cols,
            keys=inc_keys,
            data=self._incidents,
            on_change=self._on_incidents_change,
        )
        self._inc_table.grid(row=0, column=0, sticky="ew")

        return r

    # ── Standard body (original fields) ──────────────────────────────────

    def _build_standard_body(self, saved: dict, start_row: int) -> int:
        sp  = self.spec
        fields = sp.get("fields", [])
        if not fields:
            return start_row

        max_fi_row = ((len(fields) - 1) // 2 + 1) * 2 + start_row

        for fi, (key, label) in enumerate(fields):
            actual_row = (fi // 2) + start_row
            actual_col = fi % 2

            tk.Label(self._body, text=label + ":",
                     bg=COLORS["bg_surface_alt"], fg=COLORS["text_subtle"],
                     font=FONT_SMALL, anchor="w").grid(
                row=actual_row * 2 - 1, column=actual_col,
                sticky="w", padx=(14, 4), pady=(6, 0))

            var = tk.StringVar(value=saved.get(key, ""))
            self._vars[key] = var

            ent = make_dark_entry(self._body)
            ent.grid(row=actual_row * 2, column=actual_col,
                     sticky="ew", padx=(14, 8), pady=(2, 0))
            ent.configure(textvariable=var)
            var.trace_add("write", lambda *_: self._recalc())

        return max_fi_row + 1

    # ── Toggle ────────────────────────────────────────────────────────────

    def _toggle(self) -> None:
        self._expanded = not self._expanded
        self._toggle_btn.configure(text="▼" if self._expanded else "▶")
        if self._expanded:
            self._body.grid(row=1, column=0, sticky="ew")
        else:
            self._body.grid_remove()

    # ── Recalc ────────────────────────────────────────────────────────────

    def _on_incidents_change(self) -> None:
        if hasattr(self, "_inc_table"):
            self._incidents = self._inc_table.get_data()
        self._recalc()

    def _build_values_dict(self, saved: dict) -> dict:
        """Збирає словник values для compute_indicator_level."""
        vals = {k: v.get() if isinstance(v, tk.StringVar) else str(v)
                for k, v in self._vars.items()}
        vals.update({k: v for k, v in saved.items()
                     if k not in vals and not k.startswith("__")})
        vals["__incidents__"]    = self._incidents
        vals["__threshold_idx__"] = str(self._threshold_var.get())
        return vals

    def _recalc(self) -> None:
        vals          = self._build_values_dict({})
        level, detail = compute_indicator_level(self.spec, vals)
        self._level   = level
        self._detail  = detail

        bg      = RA_COLORS.get(level, COLORS["border_soft"])
        lbl_txt = f"  {RA_LABELS.get(level, '—')}  "
        self._badge.configure(text=lbl_txt, bg=bg)
        self._result_badge.configure(text=lbl_txt, bg=bg)
        self._detail_lbl.configure(text=detail)
        try:
            self._strip.configure(bg=bg)
        except tk.TclError:
            pass

        self._on_change(self.spec["code"], vals, level, self._note_var.get())

    # ── Public getters ────────────────────────────────────────────────────

    def get_values(self) -> dict[str, str]:
        vals = {k: v.get() for k, v in self._vars.items()}
        vals["__incidents__"]     = self._incidents
        vals["__threshold_idx__"] = str(self._threshold_var.get())
        return vals

    def get_level(self) -> str:
        return self._level

    def get_note(self) -> str:
        return self._note_var.get()


# ─────────────────────────────────────────────────────────────────────────────
#  ДОПОМІЖНИЙ КЛАС: таблиця інцидентів (вбудована у форму)
# ─────────────────────────────────────────────────────────────────────────────

class _IncidentTableWidget(tk.Frame):
    """
    Вбудована таблиця рядків для реєстру інцидентів.
    Використовується виключно всередині IndicatorFormFrame.
    """

    def __init__(
        self,
        parent:    tk.Misc,
        cols:      list[str],
        keys:      list[str],
        data:      list[dict],
        on_change: Callable[[], None],
    ) -> None:
        C = COLORS
        super().__init__(parent, bg=C["bg_main"])
        self.columnconfigure(0, weight=1)
        self._cols      = cols
        self._keys      = keys
        self._data:     list[dict] = [dict(d) for d in data]
        self._on_change = on_change
        self._row_vars: list[dict[str, tk.StringVar]] = []
        self._row_frames: list[tk.Frame] = []

        self._build_header()
        self._rebuild_rows()

    def _build_header(self) -> None:
        C   = COLORS
        hdr = tk.Frame(self, bg=C["bg_surface"])
        hdr.grid(row=0, column=0, sticky="ew")
        for ci, col in enumerate(self._cols):
            hdr.columnconfigure(ci, weight=1)
            tk.Label(hdr, text=col, bg=C["bg_surface"],
                     fg=C["text_subtle"],
                     font=("Arial", 7, "bold")).grid(
                row=0, column=ci, padx=(8, 0), pady=4, sticky="w")
        tk.Label(hdr, text="", bg=C["bg_surface"], width=3).grid(
            row=0, column=len(self._cols))

    def _rebuild_rows(self) -> None:
        for f in self._row_frames:
            f.destroy()
        self._row_frames.clear()
        self._row_vars.clear()

        for i, row_data in enumerate(self._data):
            self._add_row(i, row_data)

        self._rebuild_add_btn()

    def _add_row(self, i: int, row_data: dict) -> None:
        C   = COLORS
        bg  = C["row_even"] if i % 2 == 0 else C["row_odd"]
        rf  = tk.Frame(self, bg=bg)
        rf.grid(row=i + 1, column=0, sticky="ew")
        self._row_frames.append(rf)

        vars_: dict[str, tk.StringVar] = {}
        for ci, key in enumerate(self._keys):
            val = row_data.get(key, "")
            var = tk.StringVar(value=str(val) if val else "")
            vars_[key] = var
            var.trace_add(
                "write",
                lambda *_, idx=i, k=key: self._cell_change(idx, k))
            rf.columnconfigure(ci, weight=1)
            e = tk.Entry(rf, textvariable=var,
                         bg=C["bg_input"], fg=C["text_primary"],
                         relief="flat", bd=1, highlightthickness=1,
                         highlightbackground=C["border_soft"],
                         font=("Arial", 8))
            e.grid(row=0, column=ci, padx=(4, 0), pady=2, sticky="ew")
        self._row_vars.append(vars_)

        tk.Button(rf, text="✕", bg=bg, fg=C["accent_danger"],
                  relief="flat", bd=0, cursor="hand2",
                  font=("Arial", 9), padx=4,
                  command=lambda idx=i: self._delete_row(idx)).grid(
            row=0, column=len(self._keys), padx=2)

    def _rebuild_add_btn(self) -> None:
        # remove old btn
        for w in self.winfo_children():
            if isinstance(w, tk.Button) and w.cget("text") == "+ Рядок":
                w.destroy()
        add_row = len(self._data) + 1
        tk.Button(self, text="+ Рядок",
                  bg=COLORS["bg_surface_alt"], fg=COLORS["text_muted"],
                  relief="flat", bd=0, cursor="hand2",
                  font=("Arial", 8), pady=3,
                  command=self._add_empty).grid(
            row=add_row + 1, column=0, sticky="w", padx=8, pady=4)

    def _cell_change(self, idx: int, key: str) -> None:
        if idx >= len(self._data) or idx >= len(self._row_vars):
            return
        val = self._row_vars[idx][key].get()
        if key in ("amount", "hours"):
            try:
                self._data[idx][key] = float(val.replace(",", ".") or "0")
            except ValueError:
                self._data[idx][key] = 0.0
        else:
            self._data[idx][key] = val
        self._on_change()

    def _delete_row(self, idx: int) -> None:
        if idx < len(self._data):
            self._data.pop(idx)
        self._rebuild_rows()
        self._on_change()

    def _add_empty(self) -> None:
        import uuid as _uuid
        self._data.append({"_id": str(_uuid.uuid4())[:6]})
        self._rebuild_rows()
        self._on_change()

    def get_data(self) -> list[dict]:
        return self._data


# ─────────────────────────────────────────────────────────────────────────────
#  ЗАМІНА Д: метод _maybe_build_or_report для RiskDirectionFrame
#  Додати як метод класу RiskDirectionFrame
# ─────────────────────────────────────────────────────────────────────────────

def _rdf_maybe_build_or_report(self, nb: ttk.Notebook, tab: tk.Frame) -> None:
    """Будує вкладку «Звіт ОР» для напрямку operational."""
    idx = nb.index(nb.select())
    # Tab 0 = Введення, 1 = Статистика, 2 = Звіт ОР
    if idx != 2:
        return

    C = COLORS
    for w in tab.winfo_children():
        w.destroy()
    tab.rowconfigure(0, weight=1)
    tab.columnconfigure(0, weight=1)

    canvas = tk.Canvas(tab, bg=C["bg_main"], highlightthickness=0)
    sb     = ttk.Scrollbar(tab, orient="vertical", command=canvas.yview)
    canvas.configure(yscrollcommand=sb.set)
    canvas.grid(row=0, column=0, sticky="nsew")
    sb.grid(row=0, column=1, sticky="ns")

    body = tk.Frame(canvas, bg=C["bg_main"])
    body.columnconfigure(0, weight=1)
    cw   = canvas.create_window((0, 0), window=body, anchor="nw")

    def _conf(_: object) -> None:
        canvas.configure(scrollregion=canvas.bbox("all"))
        canvas.itemconfig(cw, width=canvas.winfo_width())

    body.bind("<Configure>", _conf)
    canvas.bind("<Configure>", _conf)
    _scroll_mgr.attach(canvas)

    # ── Header ──────────────────────────────────────────────────────────
    ent = self._ent_var.get().strip() if hasattr(self, "_ent_var") else "—"
    per = self._period_var.get()     if hasattr(self, "_period_var") else "—"

    hf = tk.Frame(body, bg=C["bg_header"])
    hf.grid(row=0, column=0, sticky="ew")
    tk.Frame(hf, bg=C["accent_warning"], width=5).grid(row=0, column=0, sticky="ns")
    tf = tk.Frame(hf, bg=C["bg_header"])
    tf.grid(row=0, column=1, sticky="ew", padx=14, pady=10)
    tk.Label(tf, text="ЗВІТ ОР — Реєстр операційних ризиків",
             bg=C["bg_header"], fg=C["text_primary"],
             font=FONT_HEADING).pack(anchor="w")
    tk.Label(tf, text=f"Підприємство: {ent}  |  Квартал: {per}",
             bg=C["bg_header"], fg=C["text_muted"],
             font=FONT_DEFAULT).pack(anchor="w")

    # ── Collect current values from open indicator frames ────────────────
    ind_data: dict[str, dict] = {}
    if hasattr(self, "_ind_frames"):
        for frm in self._ind_frames:
            code = frm.spec.get("code", "")
            ind_data[code] = {
                "values": frm.get_values(),
                "level":  frm.get_level(),
                "notes":  frm.get_note(),
            }

    # Also pull saved data for unsaved frames
    dir_data_saved = (self._all_data
                      .get("operational", {})
                      .get(per, {})
                      .get(ent, {})) if hasattr(self, "_all_data") else {}

    # ── Table header ─────────────────────────────────────────────────────
    col_defs = [
        ("Код",                    7),
        ("Назва індикатора",       30),
        ("Тип ризику",             16),
        ("Поріг толерантності",    16),
        ("Факт (тис. грн / %)",    15),
        ("% до доходу",             9),
        ("Рівень",                  8),
    ]

    hdr_row = tk.Frame(body, bg=C["bg_surface_alt"])
    hdr_row.grid(row=1, column=0, sticky="ew")
    for ci, (lbl, w) in enumerate(col_defs):
        hdr_row.columnconfigure(ci, weight=1 if ci == 1 else 0)
        tk.Label(hdr_row, text=lbl, bg=C["bg_surface_alt"],
                 fg=C["text_subtle"], font=("Arial", 7, "bold"),
                 width=w, anchor="w").grid(
            row=0, column=ci, padx=(6, 0), pady=6)

    # ── Rows ─────────────────────────────────────────────────────────────
    total_fact = 0.0
    net_income = 0.0

    for ri, spec in enumerate(OPERATIONAL_INDICATORS):
        code    = spec["code"].replace(" ", "")
        mode    = spec.get("mode", "")
        # prefer live frame data, fallback to saved
        frame_d = ind_data.get(spec["code"], {})
        vals    = frame_d.get("values", dir_data_saved.get(code, {}).get("values", {}))
        level_s = frame_d.get("level")
        if not level_s:
            level_s, _ = compute_indicator_level(spec, vals)

        level_i = {"Green": 0, "Yellow": 1, "Orange": 2, "Red": 3}.get(level_s, 0)
        color_h = [C["accent_success"], C["accent_warning"], "#F97316",
                   C["accent_danger"]][level_i]
        level_sym = ["🟢 0", "🟡 1", "🟠 2", "🔴 3"][level_i]

        incidents = vals.get("__incidents__", [])
        fact      = sum(_safe_num(i.get("amount", "0")) for i in incidents)

        if mode == "or_fraud":
            thr_idx = int(_safe_num(vals.get("__threshold_idx__", "0")))
            thr_k   = _OR_FRAUD_THRESHOLDS[min(thr_idx, 2)][1]
            thr_str = f"{thr_k:,.0f} тис. грн"
            inc_val = fact
            inc_str = f"{fact:,.2f} тис. грн"
            pct_str = f"{fact/thr_k*100:.1f}%" if thr_k else "—"
        elif mode == "or_downtime":
            h_day  = _safe_num(vals.get("hours_per_day", "8"))
            total_h = sum(_safe_num(i.get("hours", "0")) for i in incidents)
            work_h  = h_day * (
                _safe_num(vals.get("days_m1", "0")) +
                _safe_num(vals.get("days_m2", "0")) +
                _safe_num(vals.get("days_m3", "0"))
            ) * _safe_num(vals.get("workers", "0"))
            thr_str = "10% роб. часу"
            inc_val = total_h
            inc_str = f"{total_h:.1f} год"
            pct_str = f"{total_h/work_h*100:.1f}%" if work_h else "—"
            fact    = 0.0  # not monetary
        else:
            inc  = _safe_num(vals.get("net_income", "0"))
            cumul = _safe_num(vals.get("cumul_prev", "0"))
            total = fact + cumul
            thr_k = inc * 0.01
            thr_str = f"1% доходу = {thr_k:,.0f} тис. грн" if inc else "1% доходу"
            inc_val = total
            inc_str = f"{total:,.2f} тис. грн"
            pct_str = f"{total/inc*100:.2f}%" if inc else "—"
            if inc:
                net_income = inc
            total_fact += total

        bg = C["row_even"] if ri % 2 == 0 else C["row_odd"]
        rf = tk.Frame(body, bg=bg)
        rf.grid(row=ri + 2, column=0, sticky="ew")

        cells = [
            spec["code"],
            spec["name"],
            spec.get("risk_type", ""),
            thr_str,
            inc_str,
            pct_str,
        ]
        for ci, (val, (_, w)) in enumerate(zip(cells, col_defs)):
            tk.Label(rf, text=val, bg=bg, fg=C["text_primary"],
                     font=("Arial", 7), width=w, anchor="w").grid(
                row=0, column=ci, padx=(6, 0), pady=5)
        tk.Label(rf, text=f"  {level_sym}  ",
                 bg=color_h, fg="white", font=("Arial", 8, "bold")).grid(
            row=0, column=len(col_defs) - 1, padx=4, pady=3)

    sep_row = len(OPERATIONAL_INDICATORS) + 2
    tk.Frame(body, bg=C["border_soft"], height=2).grid(
        row=sep_row, column=0, sticky="ew")

    # ── Aggregate 5% row ─────────────────────────────────────────────────
    agg_pct = (total_fact / net_income * 100) if net_income else 0.0
    if   agg_pct == 0:       agg_col = C["accent_success"]; agg_txt = "🟢 В межах апетиту"
    elif agg_pct < 5:        agg_col = C["accent_warning"]; agg_txt = "🟡 Наближається до ліміту"
    elif abs(agg_pct-5)<0.1: agg_col = "#F97316";           agg_txt = "🟠 На межі апетиту"
    else:                    agg_col = C["accent_danger"];  agg_txt = "🔴 Перевищено апетит"

    agg_f = tk.Frame(body, bg=C["bg_surface"])
    agg_f.grid(row=sep_row + 1, column=0, sticky="ew")
    agg_f.columnconfigure(0, weight=1)
    tk.Label(
        agg_f,
        text=(
            "Узагальнений ризик-апетит до операційних ризиків — "
            "до 5% чистого доходу підприємства"
        ),
        bg=C["bg_surface"], fg=C["text_muted"],
        font=FONT_SMALL_BOLD,
    ).grid(row=0, column=0, padx=14, pady=8, sticky="w")

    tk.Label(
        agg_f,
        text=(
            f"  Факт: {agg_pct:.2f}%  |  {agg_txt}  "
            if net_income else "  Введіть чистий дохід для розрахунку  "
        ),
        bg=agg_col, fg="white", font=FONT_BOLD,
    ).grid(row=0, column=1, padx=14, pady=8)

    # ── Export CSV ────────────────────────────────────────────────────────
    import csv as _csv

    def _export() -> None:
        path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV", "*.csv")],
            title="Зберегти Звіт ОР")
        if not path:
            return
        try:
            with open(path, "w", newline="", encoding="utf-8-sig") as f:
                w = _csv.writer(f)
                w.writerow(["Підприємство", ent, "Квартал", per])
                w.writerow([])
                w.writerow(["Код", "Назва", "Тип ризику",
                             "Поріг", "Факт", "% доходу", "Рівень"])
                for spec in OPERATIONAL_INDICATORS:
                    code  = spec["code"].replace(" ", "")
                    vals  = ind_data.get(spec["code"], {}).get(
                        "values",
                        dir_data_saved.get(code, {}).get("values", {}))
                    lv, _ = compute_indicator_level(spec, vals)
                    incidents = vals.get("__incidents__", [])
                    fact_ = sum(_safe_num(i.get("amount", "0")) for i in incidents)
                    w.writerow([spec["code"], spec["name"], spec.get("risk_type",""),
                                "", f"{fact_:.2f}", "", lv])
                w.writerow([])
                w.writerow(["Узагальнений ризик-апетит",
                             f"{agg_pct:.2f}%", agg_txt])
            messagebox.showinfo("Експорт", "CSV збережено")
        except OSError as e:
            messagebox.showerror("Помилка", str(e))

    exp_f = tk.Frame(body, bg=C["bg_main"])
    exp_f.grid(row=sep_row + 2, column=0, sticky="ew", padx=16, pady=(8, 20))
    tk.Button(exp_f, text="Експорт CSV",
              bg=C["bg_surface"], fg=C["text_primary"],
              relief="flat", bd=0, cursor="hand2",
              font=FONT_SMALL, padx=12, pady=4,
              command=_export).pack(side="left")



# =============================================================================
#  COORDINATOR RECORD & PAGE
# =============================================================================

@dataclass
class CoordinatorRecord:
    id:            str = ""
    enterprise:    str = ""
    department:    str = ""
    location:      str = ""
    is_staff_unit: str = "Нi"
    is_concurrent: str = "Нi"
    main_position: str = ""
    full_name:     str = ""
    phone:         str = ""
    appointed_date:str = "—"
    order_number:  str = ""
    has_approval:  str = "Нi"

    def to_list(self) -> list:
        return [self.id, self.enterprise, self.department, self.location,
                self.is_staff_unit, self.is_concurrent, self.main_position,
                self.full_name, self.phone, self.appointed_date,
                self.order_number, self.has_approval]

    @classmethod
    def from_dict(cls, d: dict) -> "CoordinatorRecord":
        valid = {k: str(v) for k, v in d.items() if k in cls.__dataclass_fields__}
        return cls(**valid)

    @classmethod
    def from_list(cls, row: list) -> "CoordinatorRecord":
        r = list(row) + [""] * max(0, 12 - len(row))
        return cls(id=str(r[0]), enterprise=str(r[1]), department=str(r[2]),
                   location=str(r[3]), is_staff_unit=str(r[4]),
                   is_concurrent=str(r[5]), main_position=str(r[6]),
                   full_name=str(r[7]), phone=str(r[8]),
                   appointed_date=str(r[9]), order_number=str(r[10]),
                   has_approval=str(r[11]))


class CoordinatorFormDialog:
    def __init__(self, parent_root: tk.Misc,
                 on_save: Callable[[CoordinatorRecord], None],
                 record: CoordinatorRecord | None = None,
                 default_enterprise: str = "") -> None:
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
        self.win.update_idletasks()
        rx = parent_root.winfo_x(); ry = parent_root.winfo_y()
        rw = parent_root.winfo_width(); rh = parent_root.winfo_height()
        ww, wh = 560, 680
        self.win.geometry(f"{ww}x{wh}+{rx+(rw-ww)//2}+{ry+(rh-wh)//2}")
        self.win.rowconfigure(1, weight=1)
        self.win.columnconfigure(0, weight=1)

        hdr = tk.Frame(self.win, bg=C["bg_header"], height=52)
        hdr.grid(row=0, column=0, sticky="ew")
        hdr.grid_propagate(False)
        hdr.columnconfigure(0, weight=1)
        accent_color = C["accent_warning"] if self.is_edit else C["accent"]
        tk.Frame(hdr, bg=accent_color, width=4).grid(row=0, column=0, sticky="ns")
        tk.Label(hdr, text=title_text,
                 bg=C["bg_header"], fg=C["text_primary"],
                 font=FONT_HEADING).grid(row=0, column=0, padx=20, pady=14, sticky="w")

        canvas = tk.Canvas(self.win, bg=C["bg_main"], highlightthickness=0)
        sb = ttk.Scrollbar(self.win, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=sb.set)
        canvas.grid(row=1, column=0, sticky="nsew")
        sb.grid(row=1, column=1, sticky="ns")
        body = tk.Frame(canvas, bg=C["bg_main"])
        cw   = canvas.create_window((0, 0), window=body, anchor="nw")

        def _conf(_):
            canvas.configure(scrollregion=canvas.bbox("all"))
            canvas.itemconfig(cw, width=canvas.winfo_width())

        body.bind("<Configure>", _conf)
        canvas.bind("<Configure>", _conf)
        _scroll_mgr.attach(canvas)
        body.columnconfigure(0, weight=1)
        body.columnconfigure(1, weight=1)
        rec = self.record or CoordinatorRecord()

        def section(txt, r, span=2):
            f = tk.Frame(body, bg=C["bg_main"])
            f.grid(row=r, column=0, columnspan=span,
                   sticky="ew", padx=12, pady=(14, 4))
            tk.Frame(f, bg=accent_color, width=3, height=14).pack(side="left")
            tk.Label(f, text=txt, bg=C["bg_main"], fg=accent_color,
                     font=FONT_BOLD).pack(side="left", padx=8)
            return r + 1

        def lbl_entry(txt, row, col, default="", width=None):
            tk.Label(body, text=txt, bg=C["bg_main"],
                     fg=C["text_subtle"], font=FONT_SMALL).grid(
                row=row, column=col, sticky="w", padx=(14, 4), pady=(6, 0))
            kw = dict(accent=accent_color)
            if width: kw["width"] = width
            e = make_dark_entry(body, **kw)
            e.grid(row=row+1, column=col, sticky="ew", padx=(14, 8), pady=(2, 0))
            if default: e.insert(0, default)
            return e

        def lbl_combo(txt, row, col, values, current):
            tk.Label(body, text=txt, bg=C["bg_main"],
                     fg=C["text_subtle"], font=FONT_SMALL).grid(
                row=row, column=col, sticky="w", padx=(14, 4), pady=(6, 0))
            cb = make_dark_combo(body, values=values)
            cb.grid(row=row+1, column=col, sticky="ew", padx=(14, 8), pady=(2, 0))
            cb.set(current if current in values else values[0])
            return cb

        row = 0
        row = section("Підприємство та підрозділ", row)
        tk.Label(body, text="Назва підприємства:", bg=C["bg_main"],
                 fg=C["text_subtle"], font=FONT_SMALL).grid(
            row=row, column=0, columnspan=2, sticky="w", padx=(14, 4), pady=(6, 0))
        self.e_enterprise = make_dark_entry(body, accent=accent_color)
        self.e_enterprise.grid(row=row+1, column=0, columnspan=2,
                                sticky="ew", padx=(14, 14), pady=(2, 0))
        self.e_enterprise.insert(0, rec.enterprise or default_enterprise)
        row += 2
        self.e_department = lbl_entry("Назва управління/підрозділу:", row, 0,
                                       default=rec.department)
        self.e_location   = lbl_entry("Місцезнаходження:", row, 1,
                                       default=rec.location)
        row += 2

        row = section("Посада та ПІБ", row)
        tk.Label(body, text="Назва основної посади:", bg=C["bg_main"],
                 fg=C["text_subtle"], font=FONT_SMALL).grid(
            row=row, column=0, columnspan=2, sticky="w", padx=(14, 4), pady=(6, 0))
        self.e_main_position = make_dark_entry(body, accent=accent_color)
        self.e_main_position.grid(row=row+1, column=0, columnspan=2,
                                   sticky="ew", padx=(14, 14), pady=(2, 0))
        self.e_main_position.insert(0, rec.main_position); row += 2
        tk.Label(body, text="ПІБ:", bg=C["bg_main"],
                 fg=C["text_subtle"], font=FONT_SMALL).grid(
            row=row, column=0, columnspan=2, sticky="w", padx=(14, 4), pady=(6, 0))
        self.e_full_name = make_dark_entry(body, accent=accent_color)
        self.e_full_name.grid(row=row+1, column=0, columnspan=2,
                               sticky="ew", padx=(14, 14), pady=(2, 0))
        self.e_full_name.insert(0, rec.full_name); row += 2

        row = section("Контакти та призначення", row)
        self.e_phone = lbl_entry("Номер телефону:", row, 0, default=rec.phone)
        tk.Label(body, text="Дата прийняття на посаду:", bg=C["bg_main"],
                 fg=C["text_subtle"], font=FONT_SMALL).grid(
            row=row, column=1, sticky="w", padx=(14, 4), pady=(6, 0))
        self.e_appointed = make_dark_entry(body, accent=accent_color)
        self.e_appointed.grid(row=row+1, column=1, sticky="ew",
                               padx=(14, 14), pady=(2, 0))
        if rec.appointed_date and rec.appointed_date != "—":
            self.e_appointed.insert(0, rec.appointed_date)
        else:
            add_placeholder(self.e_appointed, "дд.мм.рррр")
        row += 2
        tk.Label(body, text="Номер наказу:", bg=C["bg_main"],
                 fg=C["text_subtle"], font=FONT_SMALL).grid(
            row=row, column=0, columnspan=2, sticky="w", padx=(14, 4), pady=(6, 0))
        self.e_order = make_dark_entry(body, accent=accent_color)
        self.e_order.grid(row=row+1, column=0, columnspan=2,
                           sticky="ew", padx=(14, 14), pady=(2, 0))
        self.e_order.insert(0, rec.order_number); row += 2

        row = section("Відмітки та погодження", row)
        self.cb_staff  = lbl_combo("Штатна одиниця:", row, 0,
                                    ["Так", "Нi"], rec.is_staff_unit)
        self.cb_concur = lbl_combo("Виконання обов'язків за сумісництвом:", row, 1,
                                    ["Так", "Нi"], rec.is_concurrent)
        row += 2
        tk.Label(body, text="Наявність погодження:", bg=C["bg_main"],
                 fg=C["text_subtle"], font=FONT_SMALL).grid(
            row=row, column=0, columnspan=2, sticky="w", padx=(14, 4), pady=(6, 0))
        appr_f = tk.Frame(body, bg=C["bg_main"])
        appr_f.grid(row=row+1, column=0, columnspan=2, sticky="w", padx=14, pady=(2, 0))
        self._approval_var = tk.StringVar(value=rec.has_approval)
        def _update_dot(*_):
            val = self._approval_var.get()
            self._approval_dot.configure(
                bg=C["accent_success"] if val == "Так" else C["accent_danger"])
        for val in ("Так", "Нi"):
            tk.Radiobutton(appr_f, text=val, variable=self._approval_var, value=val,
                           bg=C["bg_main"], fg=C["text_primary"],
                           activebackground=C["bg_main"],
                           activeforeground=C["text_primary"],
                           selectcolor=C["bg_surface"],
                           font=FONT_DEFAULT, command=_update_dot,
                           ).pack(side="left", padx=(0, 16))
        self._approval_dot = tk.Label(appr_f, text="  ●  ",
            bg=C["accent_success"] if rec.has_approval == "Так" else C["accent_danger"],
            fg="white", font=FONT_BOLD)
        self._approval_dot.pack(side="left", padx=6)
        row += 2
        tk.Frame(body, bg=C["bg_main"], height=16).grid(row=row, column=0, columnspan=2)

        btn_bar = tk.Frame(self.win, bg=C["bg_header"], height=46)
        btn_bar.grid(row=2, column=0, columnspan=2, sticky="ew")
        btn_bar.grid_propagate(False)
        make_button(btn_bar, "Скасувати", bg=C["bg_surface"], fg=C["text_muted"],
                    activebackground=C["bg_surface_alt"], activeforeground=C["text_primary"],
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
            messagebox.showwarning("Обов'язкові поля",
                "Заповніть назву підприємства та ПІБ координатора",
                parent=self.win); return
        appointed = self.e_appointed.get().strip()
        if appointed in ("дд.мм.рррр", ""):
            appointed = "—"
        elif not is_valid_date(appointed):
            messagebox.showwarning("Помилка дати",
                "Неправильний формат дати (очікується дд.мм.рррр)",
                parent=self.win); return
        rec_id = self.record.id if self.record else str(uuid.uuid4())[:8]
        result = CoordinatorRecord(
            id=rec_id, enterprise=enterprise,
            department=self.e_department.get().strip(),
            location=self.e_location.get().strip(),
            is_staff_unit=self.cb_staff.get(),
            is_concurrent=self.cb_concur.get(),
            main_position=self.e_main_position.get().strip(),
            full_name=full_name, phone=self.e_phone.get().strip(),
            appointed_date=appointed, order_number=self.e_order.get().strip(),
            has_approval=self._approval_var.get())
        self.on_save(result)
        self.win.destroy()


class EnterpriseCard:
    def __init__(self, parent: tk.Misc, enterprise: str,
                 records: list[CoordinatorRecord],
                 on_edit: Callable[[CoordinatorRecord], None],
                 on_delete: Callable[[CoordinatorRecord], None],
                 on_add: Callable[[str], None], row: int) -> None:
        C = COLORS
        self.parent     = parent
        self.enterprise = enterprise
        self.records    = records
        self.on_edit    = on_edit
        self.on_delete  = on_delete
        self.on_add     = on_add
        self._expanded  = True

        self.hdr = tk.Frame(parent, bg=C["bg_surface_alt"], cursor="hand2")
        self.hdr.grid(row=row, column=0, sticky="ew", padx=12, pady=(8, 0))
        self.hdr.columnconfigure(1, weight=1)
        row += 1
        tk.Frame(self.hdr, bg=C["accent"], width=5).grid(row=0, column=0, sticky="ns")
        name_f = tk.Frame(self.hdr, bg=C["bg_surface_alt"])
        name_f.grid(row=0, column=1, sticky="ew", padx=14, pady=8)
        tk.Label(name_f, text=enterprise, bg=C["bg_surface_alt"],
                 fg=C["text_primary"], font=("Arial", 10, "bold")).pack(side="left")
        tk.Label(name_f, text=f"  {len(records)} координ.",
                 bg=C["bg_surface_alt"], fg=C["accent_muted"],
                 font=FONT_SMALL).pack(side="left", padx=6)
        btn_area = tk.Frame(self.hdr, bg=C["bg_surface_alt"])
        btn_area.grid(row=0, column=2, padx=8, pady=4)
        make_button(btn_area, "+ Додати", bg=C["accent"], fg="white",
                    activebackground=C["accent_soft"],
                    font=FONT_SMALL_BOLD, padx=10, pady=3,
                    command=lambda: on_add(enterprise)).pack(side="left", padx=(0, 4))
        self.toggle_btn = make_button(btn_area, "▼",
                    bg=C["bg_surface_alt"], fg=C["text_muted"],
                    activebackground=C["bg_surface"],
                    font=FONT_DEFAULT, padx=8, pady=3,
                    command=self._toggle)
        self.toggle_btn.pack(side="left")

        self.body = tk.Frame(parent, bg=C["bg_main"])
        self.body.grid(row=row, column=0, sticky="ew", padx=12, pady=(0, 4))
        self.body.columnconfigure(0, weight=1)
        self._build_table()

    def _build_table(self) -> None:
        C = COLORS
        for w in self.body.winfo_children(): w.destroy()
        if not self.records:
            tk.Label(self.body,
                     text="  Координаторів немає. Натисніть «+ Додати».",
                     bg=C["bg_main"], fg=C["text_subtle"],
                     font=("Arial", 8, "italic")).pack(anchor="w", padx=16, pady=8)
            return
        cols = [
            ("ПІБ", 220, "full_name"), ("Підрозділ", 170, "department"),
            ("Телефон", 110, "phone"), ("Посада", 180, "main_position"),
            ("Дата прийняття", 100, "appointed_date"), ("Наказ №", 80, "order_number"),
            ("Штатна", 68, "is_staff_unit"), ("Сумісн.", 68, "is_concurrent"),
            ("Погодження", 90, "has_approval"),
        ]
        th = tk.Frame(self.body, bg=C["bg_surface"])
        th.pack(fill="x")
        for label, w, _ in cols:
            tk.Label(th, text=label, bg=C["bg_surface"], fg=C["text_subtle"],
                     font=FONT_TINY, width=w//7, anchor="w").pack(
                side="left", padx=(8, 0), pady=4)
        tk.Label(th, text="Дії", bg=C["bg_surface"], fg=C["text_subtle"],
                 font=FONT_TINY, width=8).pack(side="right", padx=8, pady=4)
        for i, rec in enumerate(self.records):
            row_bg = C["row_even"] if i % 2 == 0 else C["row_odd"]
            rf = tk.Frame(self.body, bg=row_bg)
            rf.pack(fill="x")
            for _, w, attr in cols:
                val = getattr(rec, attr, "—") or "—"
                if attr == "has_approval":
                    dot_color = (C["accent_success"] if val == "Так" else C["accent_danger"])
                    df = tk.Frame(rf, bg=row_bg, width=w//7*7, height=24)
                    df.pack_propagate(False); df.pack(side="left", padx=(8, 0))
                    tk.Label(df, text="  ●", bg=row_bg, fg=dot_color,
                             font=("Arial", 11)).pack(side="left")
                    tk.Label(df, text=val, bg=row_bg, fg=dot_color,
                             font=FONT_TINY).pack(side="left", padx=2)
                elif attr in ("is_staff_unit", "is_concurrent"):
                    color = C["accent_success"] if val == "Так" else C["text_subtle"]
                    tk.Label(rf, text=val, bg=row_bg, fg=color,
                             font=FONT_SMALL, width=w//7, anchor="w").pack(
                        side="left", padx=(8, 0), pady=3)
                else:
                    display = (val[:22] + "…") if len(val) > 22 else val
                    tk.Label(rf, text=display, bg=row_bg,
                             fg=C["text_primary"], font=FONT_SMALL,
                             width=w//7, anchor="w").pack(
                        side="left", padx=(8, 0), pady=3)
            act = tk.Frame(rf, bg=row_bg); act.pack(side="right", padx=8, pady=2)
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
        self.records = records; self._build_table()

    def _toggle(self) -> None:
        self._expanded = not self._expanded
        self.toggle_btn.configure(text="▼" if self._expanded else "▶")
        if self._expanded: self.body.grid()
        else: self.body.grid_remove()


class RiskCoordinatorsPage(tk.Frame):
    DATA_FILE = COORDS_DATA_FILE
    _EXPORT_HEADERS = [
        "ID", "Підприємство", "Управління/Підрозділ", "Місцезнаходження",
        "Штатна одиниця", "Суміщення", "Основна посада",
        "ПІБ", "Телефон", "Дата прийняття", "Наказ №", "Погодження",
    ]

    def __init__(self, master: tk.Misc, **kwargs) -> None:
        super().__init__(master, bg=COLORS["bg_main"], **kwargs)
        self.grid_rowconfigure(1, weight=1); self.grid_columnconfigure(0, weight=1)
        self._records: list[CoordinatorRecord] = []
        self._cards:   dict[str, EnterpriseCard] = {}
        self._filtered: list[CoordinatorRecord] = []
        self._build_toolbar(); self._build_scroll_area()
        self._build_statusbar(); self._load_data(); self._schedule_autosave()

    def _build_toolbar(self) -> None:
        C = COLORS
        tb = tk.Frame(self, bg=C["bg_header"], height=56)
        tb.grid(row=0, column=0, sticky="ew")
        tb.columnconfigure(2, weight=1); tb.grid_propagate(False)
        tk.Label(tb, text="РИЗИК КООРДИНАТОРИ",
                 bg=C["bg_header"], fg=C["accent_muted"],
                 font=FONT_TITLE).grid(row=0, column=0, padx=20, pady=14, sticky="w")
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
        make_button(sf, "✕", bg=C["bg_header"], fg=C["text_muted"],
                    activebackground=C["bg_surface"],
                    font=FONT_DEFAULT, padx=6, pady=1,
                    command=lambda: self._search_var.set("")).pack(side="left", padx=2)
        ff = tk.Frame(tb, bg=C["bg_header"])
        ff.grid(row=0, column=2, sticky="w")
        tk.Label(ff, text="Погодження:", bg=C["bg_header"],
                 fg=C["text_muted"], font=FONT_SMALL).pack(side="left", padx=(0, 5))
        self._filter_appr = make_dark_combo(ff, values=["Всі", "Так", "Нi"], width=8)
        self._filter_appr.set("Всі"); self._filter_appr.pack(side="left")
        self._filter_appr.bind("<<ComboboxSelected>>", lambda _: self._apply_filter())
        bf = tk.Frame(tb, bg=C["bg_header"])
        bf.grid(row=0, column=3, sticky="e", padx=12)
        make_button(bf, "+ Новий координатор", bg=C["accent"], fg="white",
                    activebackground=C["accent_soft"],
                    font=FONT_BOLD, padx=14, pady=4,
                    command=self._add_new).pack(side="left", padx=(0, 6))
        make_button(bf, "Експорт CSV", bg=C["bg_surface"], fg=C["text_primary"],
                    activebackground=C["bg_surface_alt"],
                    font=FONT_SMALL, padx=10, pady=4,
                    command=self._export_csv).pack(side="left", padx=(0, 4))
        if pd:
            make_button(bf, "Експорт Excel", bg=C["accent_success"], fg="white",
                        activebackground="#16a34a",
                        font=FONT_SMALL, padx=10, pady=4,
                        command=self._export_excel).pack(side="left", padx=(0, 4))
        make_button(bf, "Імпорт JSON", bg=C["bg_surface"], fg=C["text_primary"],
                    activebackground=C["bg_surface_alt"],
                    font=FONT_SMALL, padx=10, pady=4,
                    command=self._import_json).pack(side="left")

    def _build_scroll_area(self) -> None:
        C = COLORS
        outer = tk.Frame(self, bg=C["bg_main"])
        outer.grid(row=1, column=0, sticky="nsew")
        outer.rowconfigure(0, weight=1); outer.columnconfigure(0, weight=1)
        self._canvas = tk.Canvas(outer, bg=C["bg_main"], highlightthickness=0)
        sb = ttk.Scrollbar(outer, orient="vertical", command=self._canvas.yview)
        self._canvas.configure(yscrollcommand=sb.set)
        self._canvas.grid(row=0, column=0, sticky="nsew")
        sb.grid(row=0, column=1, sticky="ns")
        self._scroll_frame = tk.Frame(self._canvas, bg=C["bg_main"])
        self._scroll_frame.columnconfigure(0, weight=1)
        self._cw = self._canvas.create_window(
            (0, 0), window=self._scroll_frame, anchor="nw")

        def _conf(_):
            self._canvas.configure(scrollregion=self._canvas.bbox("all"))
            self._canvas.itemconfig(self._cw, width=self._canvas.winfo_width())

        self._scroll_frame.bind("<Configure>", _conf)
        self._canvas.bind("<Configure>", _conf)
        _scroll_mgr.attach(self._canvas)

    def _build_statusbar(self) -> None:
        C = COLORS
        sb = tk.Frame(self, bg=C["bg_header"], height=24)
        sb.grid(row=2, column=0, sticky="ew"); sb.grid_propagate(False)
        self._status_lbl = tk.Label(sb, text="Готово",
            bg=C["bg_header"], fg=C["text_muted"], font=FONT_TINY, padx=10)
        self._status_lbl.pack(side="left", pady=3)
        self._stats_lbl = tk.Label(sb, text="",
            bg=C["bg_header"], fg=C["text_subtle"], font=FONT_TINY, padx=10)
        self._stats_lbl.pack(side="right", pady=3)

    def _load_data(self) -> None:
        self._records.clear()
        if not os.path.exists(self.DATA_FILE):
            self._rebuild_cards(); return
        try:
            with open(self.DATA_FILE, "r", encoding="utf-8") as f:
                raw = json.load(f)
            if not isinstance(raw, list): raise ValueError("Очікується список")
            for item in raw:
                if isinstance(item, dict):
                    self._records.append(CoordinatorRecord.from_dict(item))
                elif isinstance(item, (list, tuple)):
                    self._records.append(CoordinatorRecord.from_list(list(item)))
        except (json.JSONDecodeError, ValueError, OSError) as e:
            messagebox.showerror("Помилка завантаження координаторів", str(e))
        self._filtered = list(self._records)
        self._rebuild_cards(); self._update_stats()

    def _save_data(self) -> None:
        try:
            with open(self.DATA_FILE, "w", encoding="utf-8") as f:
                json.dump([asdict(r) for r in self._records],
                          f, ensure_ascii=False, indent=2)
        except OSError as e:
            messagebox.showerror("Помилка збереження", str(e))

    def save(self) -> None: self._save_data()
    def save_before_exit(self) -> None: self._save_data()

    def _schedule_autosave(self) -> None:
        try:
            self._save_data()
            self._status_lbl.configure(
                text=f"Автозбережено о {datetime.now().strftime('%H:%M:%S')}")
        except OSError:
            self._status_lbl.configure(text="Помилка автозбереження")
        self.after(30_000, self._schedule_autosave)

    def _rebuild_cards(self) -> None:
        for w in self._scroll_frame.winfo_children(): w.destroy()
        self._cards.clear()
        groups: dict[str, list[CoordinatorRecord]] = {}
        for rec in self._filtered:
            groups.setdefault(rec.enterprise, []).append(rec)
        if not groups:
            C = COLORS
            tk.Label(self._scroll_frame,
                     text="Координаторів не знайдено.\nНатисніть «+ Новий координатор».",
                     bg=C["bg_main"], fg=C["text_subtle"],
                     font=("Arial", 11), justify="center").grid(row=0, column=0, pady=60)
            self._update_stats(); return
        row = 0
        for enterprise, recs in groups.items():
            card = EnterpriseCard(parent=self._scroll_frame, enterprise=enterprise,
                                   records=recs, on_edit=self._edit_record,
                                   on_delete=self._delete_record,
                                   on_add=self._add_for_enterprise, row=row)
            self._cards[enterprise] = card; row += 2
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

    def _apply_filter(self) -> None:
        q = self._search_var.get().strip().lower()
        appr = self._filter_appr.get()
        self._filtered = [
            r for r in self._records
            if (not q or q in (r.enterprise + r.full_name + r.department +
                r.location + r.phone + r.main_position + r.order_number).lower())
            and (appr == "Всі" or r.has_approval == appr)
        ]
        self._rebuild_cards()

    def _add_new(self) -> None:
        CoordinatorFormDialog(parent_root=self.winfo_toplevel(),
                               on_save=self._on_form_save)

    def _add_for_enterprise(self, enterprise: str) -> None:
        CoordinatorFormDialog(parent_root=self.winfo_toplevel(),
                               on_save=self._on_form_save,
                               default_enterprise=enterprise)

    def _on_form_save(self, rec: CoordinatorRecord) -> None:
        for i, r in enumerate(self._records):
            if r.id == rec.id:
                self._records[i] = rec; break
        else:
            self._records.append(rec)
        self._save_data(); self._apply_filter()
        _show_toast(self, "Збережено")

    def _edit_record(self, rec: CoordinatorRecord) -> None:
        CoordinatorFormDialog(parent_root=self.winfo_toplevel(),
                               on_save=self._on_form_save, record=rec)

    def _delete_record(self, rec: CoordinatorRecord) -> None:
        if not messagebox.askyesno("Підтвердження",
            f"Видалити координатора «{rec.full_name}»?\nПідприємство: {rec.enterprise}",
            parent=self.winfo_toplevel()): return
        self._records = [r for r in self._records if r.id != rec.id]
        self._save_data(); self._apply_filter()
        _show_toast(self, "Видалено")

    def _export_csv(self) -> None:
        if not self._records:
            messagebox.showinfo("Експорт", "Немає записів"); return
        path = filedialog.asksaveasfilename(defaultextension=".csv",
            filetypes=[("CSV файли", "*.csv")], title="Зберегти координаторів як CSV")
        if not path: return
        try:
            with open(path, "w", newline="", encoding="utf-8-sig") as f:
                w = csv.writer(f); w.writerow(self._EXPORT_HEADERS)
                for r in self._records: w.writerow(r.to_list())
            _show_toast(self, "CSV збережено")
        except OSError as e: messagebox.showerror("Помилка", str(e))

    def _export_excel(self) -> None:
        if not pd:
            messagebox.showwarning("Excel", "Встановіть pandas та openpyxl"); return
        if not self._records:
            messagebox.showinfo("Експорт", "Немає записів"); return
        path = filedialog.asksaveasfilename(defaultextension=".xlsx",
            filetypes=[("Excel файли", "*.xlsx")],
            title="Зберегти координаторів як Excel")
        if not path: return
        try:
            rows = [r.to_list() for r in self._records]
            df   = pd.DataFrame(rows, columns=self._EXPORT_HEADERS)
            with pd.ExcelWriter(path, engine="openpyxl") as writer:
                df.to_excel(writer, index=False, sheet_name="Координатори")
                ws = writer.sheets["Координатори"]
                for col_cells in ws.columns:
                    mx = max(len(str(c.value or "")) for c in col_cells)
                    ws.column_dimensions[col_cells[0].column_letter].width = min(mx+4, 50)
            _show_toast(self, "Excel збережено")
        except (OSError, Exception) as e: messagebox.showerror("Помилка", str(e))

    def _import_json(self) -> None:
        path = filedialog.askopenfilename(filetypes=[("JSON файли", "*.json")],
                                          title="Імпорт координаторів з JSON")
        if not path: return
        try:
            with open(path, "r", encoding="utf-8") as f: data = json.load(f)
            if not isinstance(data, list):
                raise ValueError("Файл повинен містити список записів")
            added = 0; existing_ids = {r.id for r in self._records}
            for item in data:
                if isinstance(item, dict): rec = CoordinatorRecord.from_dict(item)
                elif isinstance(item, (list, tuple)):
                    rec = CoordinatorRecord.from_list(list(item))
                else: continue
                if not rec.id or rec.id in existing_ids:
                    rec = CoordinatorRecord(**{**asdict(rec),
                                               "id": str(uuid.uuid4())[:8]})
                self._records.append(rec); existing_ids.add(rec.id); added += 1
            self._save_data(); self._apply_filter()
            _show_toast(self, f"Імпортовано: {added}")
        except (json.JSONDecodeError, ValueError, OSError) as e:
            messagebox.showerror("Помилка імпорту", str(e))


# =============================================================================
#  RISK APPETITE PAGE — з інтегрованими OR-індикаторами
# =============================================================================

RA_COLORS = {
    "Green":  COLORS["accent_success"],
    "Yellow": COLORS["accent_warning"],
    "Orange": "#F97316",
    "Red":    COLORS["accent_danger"],
    "—":      COLORS["border_soft"],
}
RA_LABELS = {
    "Green":  "Зелений — норма",
    "Yellow": "Жовтий — увага",
    "Orange": "Помаранчевий — попередження",
    "Red":    "Червоний — порушення",
    "—":      "—",
}

DIRECTION_INDICATORS: dict[str, list[dict]] = {
    "strategic": [
        {
            "code": "SR 1.1",
            "name": "Виконання стратегічних показників (план/факт)",
            "mode": "steps3",
            "desc": "Відповідність фактичних фінансових показників плановим (прибуток, видатки, чистий прибуток).",
            "fields": [
                ("plan_profit",  "Плановий прибуток"),
                ("fact_profit",  "Фактичний прибуток"),
                ("plan_expense", "Планові видатки"),
                ("fact_expense", "Фактичні видатки"),
                ("plan_income",  "Плановий дохід"),
                ("fact_income",  "Фактичний дохід"),
                ("plan_net",     "Плановий чистий прибуток"),
                ("fact_net",     "Фактичний чистий прибуток"),
            ],
            "thresholds": {
                "Green": "Всі 3 показники в нормі",
                "Yellow": "1 показник з відхиленням",
                "Orange": "2 показники з відхиленням",
                "Red": "Всі 3 показники з відхиленням",
            },
        },
        {
            "code": "SR 2.1",
            "name": "Реалізація стратегічних проектів та ініціатив",
            "mode": "count4",
            "desc": "Кількість стратегічних проектів, що не виконуються за планом або скасовані.",
            "fields": [
                ("required", "Заплановано проектів"),
                ("implemented", "Виконується за планом"),
            ],
            "thresholds": {
                "Green": "Всі проекти в нормі",
                "Yellow": "1 проект з відхиленням",
                "Orange": "2 проекти з відхиленням",
                "Red": "3+ проекти з відхиленням",
            },
        },
        {
            "code": "SR 3.1",
            "name": "Рівень репутаційних ризиків",
            "mode": "count3",
            "desc": "Кількість суттєвих репутаційних інцидентів за звітний квартал.",
            "fields": [
                ("incidents", "Кількість інцидентів"),
            ],
            "thresholds": {
                "Green": "0 інцидентів",
                "Yellow": "1 інцидент",
                "Red": "2+ інциденти",
            },
        },
    ],
    "operational": OPERATIONAL_INDICATORS,
    "financial": [
        {
            "code": "FR 1.1",
            "name": "Коефіцієнт поточної ліквідності",
            "mode": "ratio",
            "desc": "Відношення поточних активів до поточних зобов'язань. Норма > 2.0.",
            "fields": [
                ("current_assets",      "Поточні активи"),
                ("current_liabilities", "Поточні зобов'язання"),
            ],
            "thresholds": {
                "Green":  "Коеф. > 2.0",
                "Yellow": "Коеф. 1.5–2.0",
                "Orange": "Коеф. 1.0–1.5",
                "Red":    "Коеф. < 1.0",
            },
        },
        {
            "code": "FR 1.2",
            "name": "Рентабельність активів (ROA)",
            "mode": "ratio",
            "desc": "Чистий прибуток / Загальні активи × 100%. Норма > 5%.",
            "fields": [
                ("net_profit",   "Чистий прибуток"),
                ("total_assets", "Загальні активи"),
            ],
            "thresholds": {
                "Green":  "ROA > 5%",
                "Yellow": "ROA 2–5%",
                "Orange": "ROA 0–2%",
                "Red":    "ROA < 0%",
            },
        },
        {
            "code": "FR 2.1",
            "name": "Рівень боргового навантаження",
            "mode": "ratio",
            "desc": "Співвідношення загального боргу до EBITDA. Норма < 2.0.",
            "fields": [
                ("total_debt", "Загальний борг"),
                ("ebitda",     "EBITDA"),
            ],
            "thresholds": {
                "Green":  "Борг/EBITDA < 2",
                "Yellow": "Борг/EBITDA 2–3",
                "Orange": "Борг/EBITDA 3–4",
                "Red":    "Борг/EBITDA > 4",
            },
        },
        {
            "code": "FR 3.1",
            "name": "Прострочена дебіторська заборгованість",
            "mode": "pct_income",
            "desc": "Частка простроченої дебіторської заборгованості (> 90 днів) від виручки. Поріг 1%.",
            "fields": [
                ("debt",    "Прострочена заборгованість (> 90 днів)"),
                ("penalty", "Штрафи та пеня"),
                ("income",  "Виручка від реалізації"),
            ],
            "thresholds": {
                "Green":  "Частка < 1%",
                "Yellow": "Частка = 1%",
                "Orange": "Частка 1–2%",
                "Red":    "Частка > 2%",
            },
        },
    ],
    "compliance": [
        {
            "code": "CR 1.1",
            "name": "Виконання регуляторних вимог",
            "mode": "count4",
            "desc": "Кількість порушень регуляторних вимог або невиконаних приписів.",
            "fields": [
                ("required",    "Вимог / приписів"),
                ("executed",    "Виконано"),
            ],
            "thresholds": {
                "Green":  "0 порушень",
                "Yellow": "1 порушення",
                "Orange": "2 порушення",
                "Red":    "3+ порушень",
            },
        },
        {
            "code": "CR 2.1",
            "name": "Рівень збитків від судових рішень",
            "mode": "pct_income",
            "desc": "Збитки за рішенням суду / штрафи від виручки. Поріг 1%.",
            "fields": [
                ("loss",   "Збитки за рішенням суду / штрафи"),
                ("income", "Виручка від реалізації"),
            ],
            "thresholds": {
                "Green":  "Частка < 1%",
                "Yellow": "Частка = 1%",
                "Orange": "Частка 1–2%",
                "Red":    "Частка > 2%",
            },
        },
        {
            "code": "CR 3.1",
            "name": "SLA та якість обслуговування",
            "mode": "pct_income",
            "desc": "Частка порушень SLA від загальної кількості запитів. Поріг 1%.",
            "fields": [
                ("sla_breaches",  "Кількість порушень SLA"),
                ("total_requests","Загальна кількість запитів"),
            ],
            "thresholds": {
                "Green":  "Частка < 1%",
                "Yellow": "Частка = 1%",
                "Orange": "Частка 1–2%",
                "Red":    "Частка > 2%",
            },
        },
        {
            "code": "CR 4.1",
            "name": "Рівень корупційних ризиків",
            "mode": "count3",
            "desc": "Кількість виявлених корупційних інцидентів або конфліктів інтересів.",
            "fields": [
                ("incidents", "Кількість інцидентів"),
            ],
            "thresholds": {
                "Green":  "0 інцидентів",
                "Yellow": "1 інцидент",
                "Red":    "2+ інциденти",
            },
        },
    ],
}

DIRECTIONS = [
    {
        "key":   "strategic",
        "title": "Стратегічні ризики",
        "icon":  "🎯",
        "color": COLORS["accent"],
        "desc":  "Моніторинг відхилень від стратегічних цілей та планових показників.",
    },
    {
        "key":   "operational",
        "title": "Операційні ризики",
        "icon":  "⚙",
        "color": COLORS["accent_warning"],
        "desc":  "8 індикаторів OR з реєстрами інцидентів. Апетит: до 5% чистого доходу.",
    },
    {
        "key":   "financial",
        "title": "Фінансові ризики",
        "icon":  "💰",
        "color": COLORS["accent_success"],
        "desc":  "Ліквідність, рентабельність, боргове навантаження, дебіторська заборгованість.",
    },
    {
        "key":   "compliance",
        "title": "Комплаєнс-ризики",
        "icon":  "⚖",
        "color": "#a855f7",
        "desc":  "Регуляторні вимоги, судові рішення, SLA, корупційні ризики.",
    },
]


class RiskDirectionFrame(tk.Frame):
    """Розгортна картка одного напрямку ризик-апетиту."""

    def __init__(self, parent: tk.Misc, direction: dict,
                 saved_data: dict,
                 on_change: Callable[[str, str, dict, str, str], None],
                 row: int) -> None:
        C = COLORS
        super().__init__(parent, bg=C["bg_main"])
        self.grid(row=row, column=0, sticky="ew", pady=(0, 8))
        self.columnconfigure(0, weight=1)
        self._direction  = direction
        self._saved_data  = saved_data
        self._on_change   = on_change
        self._expanded    = False
        self._ind_frames: list[IndicatorFormFrame] = []
        self._ind_data:   dict[str, dict]          = {}
        self._build()

    def _build(self) -> None:
        C   = COLORS
        dir = self._direction
        indicators = DIRECTION_INDICATORS.get(dir["key"], [])
        saved      = self._saved_data

        # ── Overall level ────────────────────────────────────────────────
        levels_found = []
        for spec in indicators:
            code  = spec["code"].replace(" ", "")
            vals  = saved.get(code, {}).get("values", {})
            lv, _ = compute_indicator_level(spec, vals)
            if lv in RA_COLORS:
                levels_found.append(lv)
        level_order = {"Green": 0, "Yellow": 1, "Orange": 2, "Red": 3, "—": -1}
        worst = max(levels_found, key=lambda l: level_order.get(l, -1)) if levels_found else "—"

        # ── Header ───────────────────────────────────────────────────────
        hdr = tk.Frame(self, bg=C["bg_header"], cursor="hand2")
        hdr.grid(row=0, column=0, sticky="ew")
        hdr.columnconfigure(2, weight=1)
        tk.Frame(hdr, bg=dir["color"], width=5).grid(row=0, column=0, sticky="ns")
        icon_lbl = tk.Label(hdr, text=dir["icon"],
                             bg=C["bg_header"], fg=dir["color"],
                             font=("Arial", 18), width=3)
        icon_lbl.grid(row=0, column=1, padx=(10, 0), pady=10)
        title_f = tk.Frame(hdr, bg=C["bg_header"])
        title_f.grid(row=0, column=2, sticky="ew", padx=12, pady=6)
        tk.Label(title_f, text=dir["title"],
                 bg=C["bg_header"], fg=C["text_primary"],
                 font=("Arial", 11, "bold")).pack(anchor="w")
        tk.Label(title_f, text=dir["desc"],
                 bg=C["bg_header"], fg=C["text_muted"],
                 font=FONT_SMALL).pack(anchor="w")
        self._hdr_badge = tk.Label(hdr,
                                    text=f"  {RA_LABELS.get(worst, '—')}  ",
                                    bg=RA_COLORS.get(worst, C["border_soft"]),
                                    fg="white", font=FONT_SMALL_BOLD, pady=3)
        self._hdr_badge.grid(row=0, column=3, padx=8, pady=14)
        ind_count = tk.Label(hdr,
                              text=f"  {len(indicators)} індик.  ",
                              bg=C["bg_surface"], fg=C["text_muted"],
                              font=FONT_SMALL)
        ind_count.grid(row=0, column=4, padx=(0, 4), pady=14)
        self._toggle_btn = make_button(hdr, "▶",
                                        bg=C["bg_header"], fg=C["text_muted"],
                                        activebackground=C["bg_surface"],
                                        font=FONT_DEFAULT, padx=10, pady=5,
                                        command=self._toggle)
        self._toggle_btn.grid(row=0, column=5, padx=(0, 8), pady=8)
        for w in (hdr, icon_lbl, title_f):
            w.bind("<Button-1>", lambda _: self._toggle())

        # ── Body (notebook) ───────────────────────────────────────────────
        self._body = tk.Frame(self, bg=C["bg_main"])
        self._body.columnconfigure(0, weight=1)
        nb = ttk.Notebook(self._body)
        nb.grid(row=0, column=0, sticky="ew")

        # Tab 1: Введення даних
        tab_input = tk.Frame(nb, bg=C["bg_main"])
        tab_input.columnconfigure(0, weight=1)
        nb.add(tab_input, text="  Введення даних  ")

        self._ind_frames.clear()
        for ri, spec in enumerate(indicators):
            code     = spec["code"].replace(" ", "")
            s_values = saved.get(code, {}).get("values", {})
            s_notes  = saved.get(code, {}).get("notes", "")
            frm = IndicatorFormFrame(
                parent=tab_input, spec=spec,
                saved_values=s_values, saved_notes=s_notes,
                on_change=self._on_indicator_change, row=ri)
            self._ind_frames.append(frm)
            self._ind_data[spec["code"]] = {
                "values": frm.get_values(),
                "level":  frm.get_level(),
                "notes":  frm.get_note(),
            }

        save_btn_f = tk.Frame(tab_input, bg=C["bg_main"])
        save_btn_f.grid(row=len(indicators), column=0,
                        sticky="ew", padx=12, pady=(10, 14))
        make_button(save_btn_f, "💾  Зберегти всі зміни",
                    bg=dir["color"], fg="white",
                    activebackground=dir["color"],
                    font=FONT_BOLD, padx=20, pady=6,
                    command=self._save_all).pack(side="left")
        self._save_lbl = tk.Label(save_btn_f, text="",
                                   bg=C["bg_main"], fg=C["accent_success"],
                                   font=FONT_SMALL)
        self._save_lbl.pack(side="left", padx=10)

        # Tab 2: Статистика
        tab_stat = tk.Frame(nb, bg=C["bg_main"])
        tab_stat.columnconfigure(0, weight=1)
        nb.add(tab_stat, text="  Статистика  ")
        self._build_stat_tab(tab_stat, indicators, saved)

        # Tab 3: Звіт ОР — тільки для операційного напрямку
        if dir["key"] == "operational":
            self._tab_report = tk.Frame(nb, bg=C["bg_main"])
            self._tab_report.rowconfigure(0, weight=1)
            self._tab_report.columnconfigure(0, weight=1)
            nb.add(self._tab_report, text="  Звіт ОР  ")
            nb.bind("<<NotebookTabChanged>>",
                    lambda e, nb_=nb: self._maybe_rebuild_or_report(nb_))

    def _build_stat_tab(self, parent: tk.Frame,
                         indicators: list[dict], saved: dict) -> None:
        C = COLORS
        parent.columnconfigure(0, weight=1)
        tk.Label(parent, text="Зведена таблиця індикаторів",
                 bg=C["bg_main"], fg=C["text_muted"],
                 font=FONT_BOLD).grid(row=0, column=0, sticky="w",
                                      padx=14, pady=(12, 6))
        hdr = tk.Frame(parent, bg=C["bg_surface"])
        hdr.grid(row=1, column=0, sticky="ew", padx=14)
        col_widths = [8, 38, 18, 14, 28]
        col_hdrs   = ["Код", "Назва", "Режим", "Рівень", "Деталь"]
        for ci, (h, w) in enumerate(zip(col_hdrs, col_widths)):
            tk.Label(hdr, text=h, bg=C["bg_surface"], fg=C["text_subtle"],
                     font=FONT_SMALL_BOLD, width=w, anchor="w").grid(
                row=0, column=ci, padx=(8, 0), pady=5)

        level_order = {"Green": 0, "Yellow": 1, "Orange": 2, "Red": 3, "—": -1}
        worst_level = "—"
        for ri, spec in enumerate(indicators):
            code  = spec["code"].replace(" ", "")
            vals  = saved.get(code, {}).get("values", {})
            lv, detail = compute_indicator_level(spec, vals)
            if level_order.get(lv, -1) > level_order.get(worst_level, -1):
                worst_level = lv
            bg = C["row_even"] if ri % 2 == 0 else C["row_odd"]
            rf = tk.Frame(parent, bg=bg)
            rf.grid(row=ri + 2, column=0, sticky="ew", padx=14)
            cells = [spec["code"], spec["name"][:38], spec.get("mode","—"), "", detail[:28]]
            for ci, (txt, w) in enumerate(zip(cells, col_widths)):
                if ci == 3:
                    tk.Label(rf, text=f"  {RA_LABELS.get(lv,'—')[:12]}  ",
                             bg=RA_COLORS.get(lv, C["border_soft"]),
                             fg="white", font=FONT_TINY,
                             width=w, anchor="w").grid(
                        row=0, column=ci, padx=(8, 0), pady=4)
                else:
                    tk.Label(rf, text=txt, bg=bg, fg=C["text_primary"],
                             font=FONT_SMALL, width=w, anchor="w").grid(
                        row=0, column=ci, padx=(8, 0), pady=4)

        sep_row = len(indicators) + 2
        tk.Frame(parent, bg=C["border_soft"], height=2).grid(
            row=sep_row, column=0, sticky="ew", padx=14, pady=4)
        agg_f = tk.Frame(parent, bg=C["bg_surface"])
        agg_f.grid(row=sep_row + 1, column=0, sticky="ew", padx=14, pady=(0, 14))
        agg_f.columnconfigure(0, weight=1)
        tk.Label(agg_f, text=f"Загальний рівень напрямку:",
                 bg=C["bg_surface"], fg=C["text_muted"],
                 font=FONT_SMALL_BOLD).grid(row=0, column=0, padx=12, pady=8, sticky="w")
        tk.Label(agg_f,
                 text=f"  {RA_LABELS.get(worst_level,'—')}  ",
                 bg=RA_COLORS.get(worst_level, C["border_soft"]),
                 fg="white", font=FONT_BOLD).grid(row=0, column=1, padx=12, pady=8)

    def _maybe_rebuild_or_report(self, nb: ttk.Notebook) -> None:
        """Будує Звіт ОР лише коли активна відповідна вкладка."""
        try:
            idx = nb.index(nb.select())
        except tk.TclError:
            return
        if idx != 2:
            return
        if not hasattr(self, "_tab_report"):
            return
        for w in self._tab_report.winfo_children():
            w.destroy()
        self._build_or_report(self._tab_report)

    def _build_or_report(self, container: tk.Frame) -> None:
        C = COLORS
        container.rowconfigure(0, weight=1)
        container.columnconfigure(0, weight=1)
        canvas = tk.Canvas(container, bg=C["bg_main"], highlightthickness=0)
        sb = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=sb.set)
        canvas.grid(row=0, column=0, sticky="nsew")
        sb.grid(row=0, column=1, sticky="ns")
        body = tk.Frame(canvas, bg=C["bg_main"])
        body.columnconfigure(0, weight=1)
        cw = canvas.create_window((0, 0), window=body, anchor="nw")
        def _conf(_):
            canvas.configure(scrollregion=canvas.bbox("all"))
            canvas.itemconfig(cw, width=canvas.winfo_width())
        body.bind("<Configure>", _conf)
        canvas.bind("<Configure>", _conf)
        _scroll_mgr.attach(canvas)

        # Header
        hf = tk.Frame(body, bg=C["bg_header"])
        hf.grid(row=0, column=0, sticky="ew")
        tk.Frame(hf, bg=C["accent_warning"], width=5).grid(row=0, column=0, sticky="ns")
        tf = tk.Frame(hf, bg=C["bg_header"])
        tf.grid(row=0, column=1, sticky="ew", padx=14, pady=10)
        tk.Label(tf, text="ЗВІТ ОР — Реєстр операційних ризиків",
                 bg=C["bg_header"], fg=C["text_primary"],
                 font=FONT_HEADING).pack(anchor="w")

        # Column definitions
        col_defs = [
            ("Код",                  7),
            ("Назва індикатора",    30),
            ("Тип ризику",          14),
            ("Поріг толерантності", 16),
            ("Факт / показник",     14),
            ("% до доходу",          9),
            ("Рівень",               8),
        ]
        hdr_row = tk.Frame(body, bg=C["bg_surface_alt"])
        hdr_row.grid(row=1, column=0, sticky="ew")
        for ci, (lbl, w) in enumerate(col_defs):
            hdr_row.columnconfigure(ci, weight=1 if ci == 1 else 0)
            tk.Label(hdr_row, text=lbl, bg=C["bg_surface_alt"],
                     fg=C["text_subtle"], font=("Arial", 7, "bold"),
                     width=w, anchor="w").grid(row=0, column=ci, padx=(6, 0), pady=6)

        # Aggregate across frames
        total_monetary = 0.0
        net_income_ref = 0.0

        for ri, spec in enumerate(OPERATIONAL_INDICATORS):
            mode  = spec.get("mode", "")
            frm   = self._ind_frames[ri] if ri < len(self._ind_frames) else None
            vals  = frm.get_values() if frm else {}
            level_s, detail = compute_indicator_level(spec, vals)

            incidents = vals.get("__incidents__", [])
            fact_amt  = sum(_safe_num(inc.get("amount", "0")) for inc in incidents)

            if mode == "or_fraud":
                thr_idx = int(_safe_num(vals.get("__threshold_idx__", "0")))
                thr_val = _OR_FRAUD_THRESHOLDS[min(thr_idx, 2)][1]
                thr_str = f"{thr_val:,.0f} тис. грн"
                fact_str = f"{fact_amt:,.2f} тис. грн"
                pct_str  = f"{fact_amt / thr_val * 100:.1f}%" if thr_val else "—"
                total_monetary += fact_amt
            elif mode == "or_downtime":
                h_day  = _safe_num(vals.get("hours_per_day", "8"))
                total_h = sum(_safe_num(inc.get("hours", "0")) for inc in incidents)
                work_h  = h_day * (
                    _safe_num(vals.get("days_m1", "0")) +
                    _safe_num(vals.get("days_m2", "0")) +
                    _safe_num(vals.get("days_m3", "0"))
                ) * _safe_num(vals.get("workers", "0"))
                thr_str  = "10% роб. часу"
                fact_str = f"{total_h:.1f} год"
                pct_str  = f"{total_h / work_h * 100:.1f}%" if work_h else "—"
            else:
                inc_val   = _safe_num(vals.get("net_income", "0"))
                cumul_prev = _safe_num(vals.get("cumul_prev", "0"))
                total_val = fact_amt + (cumul_prev if "cumul" in mode else 0)
                thr_k     = inc_val * 0.01
                thr_str   = f"1% доходу = {thr_k:,.0f}" if inc_val else "1% доходу"
                fact_str  = f"{total_val:,.2f} тис. грн"
                pct_str   = f"{total_val / inc_val * 100:.2f}%" if inc_val else "—"
                total_monetary += total_val
                if inc_val: net_income_ref = inc_val

            level_i   = {"Green": 0, "Yellow": 1, "Orange": 2, "Red": 3}.get(level_s, 0)
            color_h   = [C["accent_success"], C["accent_warning"],
                         "#F97316", C["accent_danger"]][level_i]
            level_sym = ["🟢 0", "🟡 1", "🟠 2", "🔴 3"][level_i]

            bg = C["row_even"] if ri % 2 == 0 else C["row_odd"]
            rf = tk.Frame(body, bg=bg)
            rf.grid(row=ri + 2, column=0, sticky="ew")
            cells = [spec["code"], spec["name"], spec.get("risk_type",""),
                     thr_str, fact_str, pct_str]
            for ci, (txt, (_, w)) in enumerate(zip(cells, col_defs)):
                tk.Label(rf, text=txt, bg=bg, fg=C["text_primary"],
                         font=("Arial", 7), width=w, anchor="w").grid(
                    row=0, column=ci, padx=(6, 0), pady=5)
            tk.Label(rf, text=f"  {level_sym}  ",
                     bg=color_h, fg="white",
                     font=("Arial", 8, "bold")).grid(
                row=0, column=len(col_defs) - 1, padx=4, pady=3)

        # Aggregate footer
        sep_row = len(OPERATIONAL_INDICATORS) + 2
        tk.Frame(body, bg=C["border_soft"], height=2).grid(
            row=sep_row, column=0, sticky="ew")

        agg_pct = (total_monetary / net_income_ref * 100) if net_income_ref else 0.0
        if   agg_pct == 0:            agg_col = C["accent_success"]; agg_txt = "🟢 В межах апетиту"
        elif agg_pct < 5:             agg_col = C["accent_warning"]; agg_txt = "🟡 Наближається до ліміту"
        elif abs(agg_pct - 5) < 0.1:  agg_col = "#F97316";           agg_txt = "🟠 На межі апетиту"
        else:                          agg_col = C["accent_danger"];  agg_txt = "🔴 Перевищено апетит"

        agg_f = tk.Frame(body, bg=C["bg_surface"])
        agg_f.grid(row=sep_row + 1, column=0, sticky="ew")
        agg_f.columnconfigure(0, weight=1)
        tk.Label(agg_f,
                 text="Узагальнений ризик-апетит до операційних ризиків — до 5% чистого доходу",
                 bg=C["bg_surface"], fg=C["text_muted"],
                 font=FONT_SMALL_BOLD).grid(row=0, column=0, padx=14, pady=8, sticky="w")
        tk.Label(agg_f,
                 text=(f"  Факт: {agg_pct:.2f}%  |  {agg_txt}  "
                       if net_income_ref else "  Введіть чистий дохід для розрахунку  "),
                 bg=agg_col, fg="white", font=FONT_BOLD).grid(
            row=0, column=1, padx=14, pady=8)

        # CSV export
        def _export_csv():
            path = filedialog.asksaveasfilename(defaultextension=".csv",
                filetypes=[("CSV", "*.csv")], title="Зберегти Звіт ОР")
            if not path: return
            try:
                with open(path, "w", newline="", encoding="utf-8-sig") as f:
                    import csv as _csv
                    w = _csv.writer(f)
                    w.writerow(["Код", "Назва", "Тип ризику",
                                "Поріг", "Факт", "% доходу", "Рівень"])
                    for spec in OPERATIONAL_INDICATORS:
                        ri2   = OPERATIONAL_INDICATORS.index(spec)
                        frm2  = self._ind_frames[ri2] if ri2 < len(self._ind_frames) else None
                        vals2 = frm2.get_values() if frm2 else {}
                        lv2, _ = compute_indicator_level(spec, vals2)
                        facts2 = sum(_safe_num(i.get("amount","0"))
                                     for i in vals2.get("__incidents__",[]))
                        w.writerow([spec["code"], spec["name"],
                                    spec.get("risk_type",""), "", f"{facts2:.2f}", "", lv2])
                    w.writerow([])
                    w.writerow(["Узагальнений апетит", f"{agg_pct:.2f}%", agg_txt])
                _show_toast(container, "CSV збережено")
            except OSError as e:
                messagebox.showerror("Помилка", str(e))

        exp_f = tk.Frame(body, bg=C["bg_main"])
        exp_f.grid(row=sep_row + 2, column=0, sticky="ew", padx=16, pady=(8, 20))
        make_button(exp_f, "Експорт CSV",
                    bg=C["bg_surface"], fg=C["text_primary"],
                    activebackground=C["bg_surface_alt"],
                    font=FONT_SMALL, padx=12, pady=4,
                    command=_export_csv).pack(side="left")

    def _on_indicator_change(self, code: str, values: dict,
                              level: str, notes: str) -> None:
        self._ind_data[code] = {"values": values, "level": level, "notes": notes}
        levels = [d["level"] for d in self._ind_data.values() if d["level"] in RA_COLORS]
        level_order = {"Green": 0, "Yellow": 1, "Orange": 2, "Red": 3, "—": -1}
        worst = max(levels, key=lambda l: level_order.get(l, -1)) if levels else "—"
        bg = RA_COLORS.get(worst, COLORS["border_soft"])
        self._hdr_badge.configure(text=f"  {RA_LABELS.get(worst,'—')}  ", bg=bg)
        self._on_change(self._direction["key"], code, values, level, notes)

    def _save_all(self) -> None:
        for frm in self._ind_frames:
            self._on_indicator_change(frm.spec["code"],
                                       frm.get_values(),
                                       frm.get_level(),
                                       frm.get_note())
        self._save_lbl.configure(
            text=f"Збережено о {datetime.now().strftime('%H:%M:%S')}")
        self.after(3000, lambda: self._save_lbl.configure(text="") if self._save_lbl.winfo_exists() else None)

    def _toggle(self) -> None:
        self._expanded = not self._expanded
        self._toggle_btn.configure(text="▼" if self._expanded else "▶")
        if self._expanded:
            self._body.grid(row=1, column=0, sticky="ew")
        else:
            self._body.grid_remove()


class RiskAppetitePage(tk.Frame):
    """Головна сторінка Ризик-Апетит з 4 напрямками."""

    def __init__(self, master: tk.Misc, **kwargs) -> None:
        C = COLORS
        super().__init__(master, bg=C["bg_main"], **kwargs)
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)
        self._data: dict = {}
        self._dir_frames: list[RiskDirectionFrame] = []
        self._load_data()
        self._build_toolbar()
        self._build_scroll_area()
        self._build_directions()
        self._schedule_autosave()

    def _load_data(self) -> None:
        if not os.path.exists(APPETITE_FILE):
            return
        try:
            with open(APPETITE_FILE, "r", encoding="utf-8") as f:
                self._data = json.load(f)
        except (json.JSONDecodeError, OSError):
            self._data = {}

    def _save_data(self) -> None:
        try:
            with open(APPETITE_FILE, "w", encoding="utf-8") as f:
                json.dump(self._data, f, ensure_ascii=False, indent=2)
        except OSError as e:
            messagebox.showerror("Помилка збереження", str(e))

    def _build_toolbar(self) -> None:
        C = COLORS
        tb = tk.Frame(self, bg=C["bg_header"], height=58)
        tb.grid(row=0, column=0, sticky="ew")
        tb.grid_propagate(False)
        tb.columnconfigure(1, weight=1)
        tk.Frame(tb, bg=C["accent_warning"], width=4).grid(row=0, column=0, sticky="ns")
        title_f = tk.Frame(tb, bg=C["bg_header"])
        title_f.grid(row=0, column=1, sticky="ew", padx=16, pady=8)
        tk.Label(title_f, text="РИЗИК-АПЕТИТ",
                 bg=C["bg_header"], fg=C["accent_muted"],
                 font=FONT_TITLE).pack(side="left")
        tk.Label(title_f,
                 text="  4 напрямки  |  OR-індикатори інтегровані  |  Квартальний моніторинг",
                 bg=C["bg_header"], fg=C["text_subtle"],
                 font=FONT_SMALL).pack(side="left", padx=12)
        bf = tk.Frame(tb, bg=C["bg_header"])
        bf.grid(row=0, column=2, sticky="e", padx=12)
        make_button(bf, "💾 Зберегти все",
                    bg=C["accent_warning"], fg="white",
                    activebackground="#d97706",
                    font=FONT_BOLD, padx=14, pady=5,
                    command=self._save_all).pack(side="left", padx=4)
        make_button(bf, "Розгорнути всі",
                    bg=C["bg_surface"], fg=C["text_primary"],
                    activebackground=C["bg_surface_alt"],
                    font=FONT_SMALL, padx=10, pady=5,
                    command=self._expand_all).pack(side="left", padx=4)
        make_button(bf, "Згорнути всі",
                    bg=C["bg_surface"], fg=C["text_primary"],
                    activebackground=C["bg_surface_alt"],
                    font=FONT_SMALL, padx=10, pady=5,
                    command=self._collapse_all).pack(side="left", padx=(0, 4))

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
        def _conf(_):
            self._canvas.configure(scrollregion=self._canvas.bbox("all"))
            self._canvas.itemconfig(self._cw, width=self._canvas.winfo_width())
        self._scroll_frame.bind("<Configure>", _conf)
        self._canvas.bind("<Configure>", _conf)
        _scroll_mgr.attach(self._canvas)

    def _build_directions(self) -> None:
        self._dir_frames.clear()
        for ri, direction in enumerate(DIRECTIONS):
            dir_key    = direction["key"]
            saved_data = self._data.get(dir_key, {})
            frm = RiskDirectionFrame(
                parent=self._scroll_frame,
                direction=direction,
                saved_data=saved_data,
                on_change=self._on_indicator_change,
                row=ri)
            self._dir_frames.append(frm)

    def _on_indicator_change(self, dir_key: str, code: str,
                              values: dict, level: str, notes: str) -> None:
        safe_code = code.replace(" ", "")
        if dir_key not in self._data:
            self._data[dir_key] = {}
        self._data[dir_key][safe_code] = {
            "values": {k: (v if not isinstance(v, list)
                           else [dict(i) for i in v])
                       for k, v in values.items()},
            "level":  level,
            "notes":  notes,
        }
        self._save_data()

    def _save_all(self) -> None:
        for frm in self._dir_frames:
            frm._save_all()
        self._save_data()
        _show_toast(self, "Всі зміни збережено")

    def _expand_all(self) -> None:
        for frm in self._dir_frames:
            if not frm._expanded:
                frm._toggle()

    def _collapse_all(self) -> None:
        for frm in self._dir_frames:
            if frm._expanded:
                frm._toggle()

    def _schedule_autosave(self) -> None:
        self._save_data()
        self.after(30000, self._schedule_autosave)

    def save_before_exit(self) -> None:
        self._save_data()


# =============================================================================
#  ATLAS APP — головний клас програми
# =============================================================================

PageKey: TypeAlias = Literal[
    "events", "risks", "coordinators", "appetite"
]

class AtlasApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("ATLAS — Система управління ризиками")
        self.geometry("1380x820")
        self.minsize(1100, 640)
        apply_dark_style(self)
        self.configure(bg=COLORS["bg_main"])
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self._current_page_key: PageKey | None = None
        self._pages: dict[str, tk.Frame] = {}
        self._build_layout()
        self._on_nav_click("events")

    def _build_layout(self) -> None:
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)
        self._build_sidebar()
        self.content_frame = tk.Frame(self, bg=COLORS["bg_main"])
        self.content_frame.grid(row=0, column=1, sticky="nsew")
        self.content_frame.rowconfigure(0, weight=1)
        self.content_frame.columnconfigure(0, weight=1)

    def _build_sidebar(self) -> None:
        C = COLORS
        sidebar = tk.Frame(self, bg=C["bg_sidebar"], width=220)
        sidebar.grid(row=0, column=0, sticky="nsew")
        sidebar.grid_propagate(False)
        sidebar.columnconfigure(0, weight=1)

        # Logo
        logo_f = tk.Frame(sidebar, bg=C["bg_sidebar"], height=70)
        logo_f.grid(row=0, column=0, sticky="ew")
        logo_f.grid_propagate(False)
        tk.Frame(logo_f, bg=C["accent_warning"], height=3).pack(fill="x", side="bottom")
        lf = tk.Frame(logo_f, bg=C["bg_sidebar"])
        lf.pack(expand=True, fill="both")
        tk.Label(lf, text="ATLAS",
                 bg=C["bg_sidebar"], fg=C["text_primary"],
                 font=("Arial", 20, "bold")).pack(side="left", padx=18, pady=14)
        tk.Label(lf, text="GRC",
                 bg=C["accent_warning"], fg="white",
                 font=("Arial", 9, "bold"),
                 padx=5, pady=2).pack(side="left", pady=14)

        # Nav items
        menu_items = [
            ("events",       "📋",  "Суттєві події"),
            ("risks",        "⚠",   "Реєстр ризиків"),
            ("coordinators", "👥",  "Координатори"),
            ("appetite",     "🎚",  "Ризик-апетит"),
        ]
        self._nav_btns: dict[str, tk.Button] = {}
        for ri, (key, icon, label) in enumerate(menu_items, start=1):
            btn_f = tk.Frame(sidebar, bg=C["bg_sidebar"], cursor="hand2")
            btn_f.grid(row=ri, column=0, sticky="ew", pady=(2, 0))
            btn_f.columnconfigure(1, weight=1)
            self._indicator = tk.Frame(btn_f, bg=C["bg_sidebar"], width=4)
            self._indicator.grid(row=0, column=0, sticky="ns")
            icon_l = tk.Label(btn_f, text=icon,
                               bg=C["bg_sidebar"], fg=C["text_muted"],
                               font=("Arial", 13), width=3)
            icon_l.grid(row=0, column=1, padx=(6, 0), pady=10)
            lbl_l = tk.Label(btn_f, text=label,
                              bg=C["bg_sidebar"], fg=C["text_muted"],
                              font=FONT_DEFAULT, anchor="w")
            lbl_l.grid(row=0, column=2, sticky="ew", padx=(4, 8), pady=10)
            btn = tk.Button(btn_f, text="", bg=C["bg_sidebar"],
                            relief="flat", bd=0, cursor="hand2",
                            command=lambda k=key: self._on_nav_click(k))
            btn.place(x=0, y=0, relwidth=1, relheight=1)
            self._nav_btns[key] = (btn_f, self._indicator, icon_l, lbl_l)

            def _enter(e, bf=btn_f, il=icon_l, ll=lbl_l):
                if self._current_page_key != key:
                    bf.configure(bg=C["bg_surface"])
                    il.configure(bg=C["bg_surface"])
                    ll.configure(bg=C["bg_surface"])
            def _leave(e, bf=btn_f, il=icon_l, ll=lbl_l):
                if self._current_page_key != key:
                    bf.configure(bg=C["bg_sidebar"])
                    il.configure(bg=C["bg_sidebar"])
                    ll.configure(bg=C["bg_sidebar"])
            for w in (btn_f, icon_l, lbl_l, btn):
                w.bind("<Enter>", _enter)
                w.bind("<Leave>", _leave)

        # Footer
        footer = tk.Frame(sidebar, bg=C["bg_sidebar"])
        footer.grid(row=len(menu_items) + 2, column=0, sticky="sew", pady=16)
        sidebar.rowconfigure(len(menu_items) + 2, weight=1)
        tk.Label(footer, text="v2.3 — OR інтегровано",
                 bg=C["bg_sidebar"], fg=C["text_subtle"],
                 font=FONT_TINY).pack(padx=16, anchor="w")
        tk.Label(footer, text="© 2025 ATLAS GRC",
                 bg=C["bg_sidebar"], fg=C["text_subtle"],
                 font=FONT_TINY).pack(padx=16, anchor="w")

    def _on_nav_click(self, key: PageKey) -> None:
        C = COLORS
        # Deactivate old
        if self._current_page_key and self._current_page_key in self._nav_btns:
            bf, ind, il, ll = self._nav_btns[self._current_page_key]
            bf.configure(bg=C["bg_sidebar"])
            ind.configure(bg=C["bg_sidebar"])
            il.configure(bg=C["bg_sidebar"], fg=C["text_muted"])
            ll.configure(bg=C["bg_sidebar"], fg=C["text_muted"])

        self._current_page_key = key

        # Activate new
        if key in self._nav_btns:
            bf, ind, il, ll = self._nav_btns[key]
            bf.configure(bg=C["bg_surface"])
            ind.configure(bg=C["accent_warning"])
            il.configure(bg=C["bg_surface"], fg=C["text_primary"])
            ll.configure(bg=C["bg_surface"], fg=C["text_primary"],
                         font=FONT_BOLD)

        # Hide all pages
        for page in self._pages.values():
            page.grid_remove()

        # Show / create page
        if key not in self._pages:
            if key == "events":
                page = MaterialEventsPage(self.content_frame)
            elif key == "risks":
                page = RiskRegisterPage(self.content_frame)
            elif key == "coordinators":
                page = RiskCoordinatorsPage(self.content_frame)
            elif key == "appetite":
                page = RiskAppetitePage(self.content_frame)
            else:
                return
            page.grid(row=0, column=0, sticky="nsew")
            self._pages[key] = page

        self._pages[key].grid(row=0, column=0, sticky="nsew")

    def _on_close(self) -> None:
        for key, page in self._pages.items():
            if hasattr(page, "save_before_exit"):
                try:
                    page.save_before_exit()
                except Exception:
                    pass
        self.destroy()


# =============================================================================
#  ENTRY POINT
# =============================================================================

if __name__ == "__main__":
    app = AtlasApp()
    app.mainloop()
