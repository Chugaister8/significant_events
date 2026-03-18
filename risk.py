# risk_register.py
# =============================================================================
#  РЕЄСТР РИЗИКІВ — ATLAS модуль
#  Структура запису (tuple, 16 полів):
#    [0]  id             — №  (str, "001")
#    [1]  entity         — Підприємство
#    [2]  risk_name      — Назва ризику
#    [3]  risk_category  — Категорія ризику (з RISK_CATEGORIES)
#    [4]  risk_type      — Тип ризику (з RISK_TYPES)
#    [5]  probability    — Імовірність (str "1"–"5")
#    [6]  impact         — Вплив (str "1"–"5")
#    [7]  risk_score     — Рівень ризику (str, prob*impact)
#    [8]  owner          — Власник ризику
#    [9]  controls       — Заходи контролю (текст)
#    [10] residual_risk  — Залишковий ризик (str "1"–"25")
#    [11] date_identified— Дата виявлення (дд.мм.рррр)
#    [12] review_date    — Дата перегляду (дд.мм.рррр)
#    [13] priority       — Пріоритет
#    [14] status         — Статус
#    [15] description    — Детальний опис
# =============================================================================

from __future__ import annotations

from typing import Callable
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
#  КОНСТАНТИ МОДУЛЯ
# =============================================================================

RISK_DATA_FILE = "risk_register.json"

# Імпортуємо палітру з основного модуля, або визначаємо локально
try:
    from __main__ import COLORS, RISK_COLORS, RISK_TYPES
except ImportError:
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

PROBABILITY_LEVELS = ["1 — Мiнiмальна", "2 — Низька", "3 — Середня", "4 — Висока", "5 — Критична"]
IMPACT_LEVELS     = ["1 — Незначний",   "2 — Малий",  "3 — Помiрний","4 — Суттєвий","5 — Катастрофiчний"]

# Кольори рівнів ризику (score = probability * impact)
def _score_color(score: int) -> str:
    if score <= 4:
        return COLORS["accent_success"]   # зелений
    elif score <= 9:
        return COLORS["accent_warning"]   # жовтий
    elif score <= 16:
        return "#f97316"                  # помаранчевий
    else:
        return COLORS["accent_danger"]    # червоний

def _score_label(score: int) -> str:
    if score <= 4:   return "Низький"
    elif score <= 9: return "Помiрний"
    elif score <= 16:return "Високий"
    else:            return "Критичний"


# =============================================================================
#  ХЕЛПЕРИ
# =============================================================================

def _is_valid_date(s: str) -> bool:
    if not s or s in ("дд.мм.рррр", ""):
        return True
    if not re.match(r"^\d{2}\.\d{2}\.\d{4}$", s):
        return False
    try:
        datetime.strptime(s, "%d.%m.%Y")
        return True
    except ValueError:
        return False


def _make_dark_text(parent: tk.Misc, **kwargs) -> tk.Text:
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


def _add_placeholder(entry: tk.Entry, text: str) -> None:
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
    """Витягує цифру з рядку типу '3 — Середня' -> 3."""
    try:
        return int(str(val).split()[0])
    except (ValueError, IndexError):
        try:
            return int(str(val))
        except ValueError:
            return 1


# =============================================================================
#  ДЕТАЛЬНЕ ВІКНО РИЗИКУ
# =============================================================================

class RiskDetailWindow:
    """Спливаюче вікно для перегляду та редагування запису ризику."""

    # Структура запису: 16 полів (індекси 0–15)
    RECORD_LEN = 16

    def __init__(
        self,
        parent_root: tk.Misc,
        record: tuple,
        all_records: list[tuple],
        save_callback:   Callable[[str, tuple], None],
        delete_callback: Callable[[str], None],
        toast_callback:  Callable[[str], None],
    ) -> None:
        self.parent_root    = parent_root
        self.record         = list(record)
        self.all_records    = all_records
        self.save_callback  = save_callback
        self.delete_callback = delete_callback
        self.toast_callback = toast_callback
        self.is_edit_mode   = False
        self._build_window()

    # ------------------------------------------------------------------
    def _build_window(self) -> None:
        C = COLORS
        self.win = tk.Toplevel(self.parent_root)
        self.win.title(f"Ризик #{self.record[0]}  —  {self.record[1]}")
        self.win.geometry("820x740")
        self.win.minsize(660, 520)
        self.win.configure(bg=C["bg_main"])
        self.win.grab_set()

        self.win.update_idletasks()
        rx = self.parent_root.winfo_x()
        ry = self.parent_root.winfo_y()
        rw = self.parent_root.winfo_width()
        rh = self.parent_root.winfo_height()
        ww, wh = 820, 740
        self.win.geometry(f"{ww}x{wh}+{rx + (rw - ww)//2}+{ry + (rh - wh)//2}")

        self.win.rowconfigure(1, weight=1)
        self.win.columnconfigure(0, weight=1)

        # ---- Header ----
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
        tk.Frame(header, bg=strip_color, width=4).grid(row=0, column=0, sticky="ns")

        title_frame = tk.Frame(header, bg=C["bg_header"])
        title_frame.grid(row=0, column=1, sticky="ew", padx=16, pady=10)

        self.lbl_title = tk.Label(
            title_frame,
            text=f"Ризик #{self.record[0]}",
            bg=C["bg_header"], fg=COLORS["accent_muted"],
            font=("Arial", 11, "bold"),
        )
        self.lbl_title.pack(anchor="w")
        self.lbl_subtitle = tk.Label(
            title_frame,
            text=self.record[1] if len(self.record) > 1 else "",
            bg=C["bg_header"], fg=COLORS["text_muted"],
            font=("Arial", 9),
        )
        self.lbl_subtitle.pack(anchor="w")

        status_val = self.record[14] if len(self.record) > 14 else "—"
        status_colors = {
            "Активний":    COLORS["accent_danger"],
            "Монiторинг":  COLORS["accent_warning"],
            "Мiтигований": COLORS["accent"],
            "Закрито":     COLORS["text_muted"],
        }
        sc = status_colors.get(status_val, COLORS["text_muted"])
        self.lbl_status_badge = tk.Label(
            header, text=f"  {status_val}  ",
            bg=sc, fg="white", font=("Arial", 8, "bold"), pady=3,
        )
        self.lbl_status_badge.grid(row=0, column=2, padx=16, pady=18)

        # Score badge
        score_lbl_text = f"  {_score_label(score_val)} ({score_val})  "
        self.lbl_score_badge = tk.Label(
            header, text=score_lbl_text,
            bg=strip_color, fg="white", font=("Arial", 8, "bold"), pady=3,
        )
        self.lbl_score_badge.grid(row=0, column=3, padx=(0, 16), pady=18)

        # ---- Scroll content ----
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

        # ---- Bottom buttons ----
        btn_bar = tk.Frame(self.win, bg=C["bg_header"], height=46)
        btn_bar.grid(row=2, column=0, columnspan=2, sticky="ew")
        btn_bar.grid_propagate(False)
        btn_bar.columnconfigure(0, weight=1)

        left_btns = tk.Frame(btn_bar, bg=C["bg_header"])
        left_btns.pack(side="left", padx=12, pady=8)

        self.btn_edit = tk.Button(
            left_btns, text="Редагувати",
            bg=COLORS["accent_warning"], fg=COLORS["bg_main"],
            activebackground="#d97706", activeforeground="white",
            relief="flat", bd=0, cursor="hand2",
            font=("Arial", 9, "bold"), padx=14, pady=4,
            command=self._toggle_edit_mode,
        )
        self.btn_edit.pack(side="left", padx=(0, 8))

        self.btn_save = tk.Button(
            left_btns, text="Зберегти змiни",
            bg=COLORS["accent_success"], fg="white",
            activebackground="#16a34a", activeforeground="white",
            relief="flat", bd=0, cursor="hand2",
            font=("Arial", 9, "bold"), padx=14, pady=4,
            command=self._save_changes,
        )
        self.btn_save.pack_forget()

        self.btn_cancel_edit = tk.Button(
            left_btns, text="Скасувати",
            bg=COLORS["bg_surface"], fg=COLORS["text_muted"],
            activebackground=COLORS["bg_surface_alt"],
            activeforeground=COLORS["text_primary"],
            relief="flat", bd=0, cursor="hand2",
            font=("Arial", 9), padx=12, pady=4,
            command=self._cancel_edit,
        )
        self.btn_cancel_edit.pack_forget()

        right_btns = tk.Frame(btn_bar, bg=C["bg_header"])
        right_btns.pack(side="right", padx=12, pady=8)

        tk.Button(
            right_btns, text="Видалити",
            bg=COLORS["accent_danger"], fg="white",
            activebackground="#dc2626", activeforeground="white",
            relief="flat", bd=0, cursor="hand2",
            font=("Arial", 9, "bold"), padx=14, pady=4,
            command=self._delete_record,
        ).pack(side="right", padx=(8, 0))

        tk.Button(
            right_btns, text="Закрити",
            bg=COLORS["bg_surface"], fg=COLORS["text_primary"],
            activebackground=COLORS["bg_surface_alt"],
            activeforeground=COLORS["text_primary"],
            relief="flat", bd=0, cursor="hand2",
            font=("Arial", 9), padx=12, pady=4,
            command=self.win.destroy,
        ).pack(side="right")

    # ------------------------------------------------------------------
    def _section_label(self, parent: tk.Misc, text: str, row: int) -> None:
        f = tk.Frame(parent, bg=COLORS["bg_main"])
        f.grid(row=row, column=0, columnspan=2, sticky="ew", padx=8, pady=(14, 4))
        tk.Frame(f, bg=COLORS["accent_soft"], width=2, height=16).pack(side="left")
        tk.Label(f, text=text, bg=COLORS["bg_main"], fg=COLORS["accent_muted"],
                 font=("Arial", 9, "bold")).pack(side="left", padx=8)

    def _info_row(self, parent: tk.Misc, label: str, value: str,
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
        tk.Label(
            cell, text=value or "—",
            bg=C["bg_surface"],
            fg=value_color if value_color else C["text_primary"],
            font=("Arial", 9 if not value_color else 10,
                  "normal" if not value_color else "bold"),
            wraplength=260, justify="left",
        ).grid(row=1, column=0, sticky="w", pady=(2, 0))

    def _text_block(self, parent: tk.Misc, label: str, value: str, row: int) -> None:
        C = COLORS
        cell = tk.Frame(parent, bg=C["bg_surface"], padx=10, pady=6)
        cell.grid(row=row, column=0, columnspan=2, sticky="nsew", padx=8, pady=3)
        cell.columnconfigure(0, weight=1)
        tk.Label(cell, text=label, bg=C["bg_surface"], fg=C["text_subtle"],
                 font=("Arial", 7, "bold")).grid(row=0, column=0, sticky="w")
        t = _make_dark_text(cell, height=3, wrap="word", state="normal")
        t.insert("1.0", value or "—")
        t.configure(state="disabled")
        t.grid(row=1, column=0, sticky="ew", pady=(4, 0))

    # ------------------------------------------------------------------
    def _build_view_content(self) -> None:
        C = COLORS
        r = self.record
        for w in self.content.winfo_children():
            w.destroy()
        row = 0

        self._section_label(self.content, "Iнформацiя про пiдприємство", row); row += 1

        self._info_row(self.content, "Пiдприємство",
                       r[1] if len(r) > 1 else "—", row, 0)
        # Пріоритет
        priority_val = r[13] if len(r) > 13 else "—"
        priority_colors = {
            "Критичний": COLORS["accent_danger"],
            "Високий":   COLORS["accent_warning"],
            "Середнiй":  COLORS["accent"],
            "Низький":   COLORS["accent_success"],
        }
        self._info_row(self.content, "Прiоритет", priority_val, row, 1,
                       value_color=priority_colors.get(priority_val))
        row += 1

        self._section_label(self.content, "Опис ризику", row); row += 1

        self._info_row(self.content, "Назва ризику",
                       r[2] if len(r) > 2 else "—", row, 0)

        risk_type_val = r[4] if len(r) > 4 else "—"
        self._info_row(self.content, "Тип ризику", risk_type_val, row, 1,
                       value_color=RISK_COLORS.get(risk_type_val))
        row += 1

        self._info_row(self.content, "Категорiя ризику",
                       r[3] if len(r) > 3 else "—", row, 0)
        self._info_row(self.content, "Власник ризику",
                       r[8] if len(r) > 8 else "—", row, 1)
        row += 1

        self._section_label(self.content, "Оцiнка ризику", row); row += 1

        prob_raw = r[5] if len(r) > 5 else "—"
        imp_raw  = r[6] if len(r) > 6 else "—"
        score_raw = r[7] if len(r) > 7 else "0"
        try:
            score_int = int(score_raw)
        except (ValueError, TypeError):
            score_int = 0
        res_raw = r[10] if len(r) > 10 else "—"

        self._info_row(self.content, "Iмовiрнiсть", prob_raw, row, 0)
        self._info_row(self.content, "Вплив", imp_raw, row, 1)
        row += 1

        # Score cell (full width)
        score_cell = tk.Frame(self.content, bg=C["bg_surface"], padx=10, pady=8)
        score_cell.grid(row=row, column=0, columnspan=2, sticky="nsew", padx=8, pady=3)
        score_cell.columnconfigure(0, weight=1)
        tk.Label(score_cell, text="Рiвень ризику (Score = Iмовiрнiсть x Вплив)",
                 bg=C["bg_surface"], fg=C["text_subtle"],
                 font=("Arial", 7, "bold")).grid(row=0, column=0, sticky="w")
        sc_color = _score_color(score_int)
        tk.Label(
            score_cell,
            text=f"{score_int}  —  {_score_label(score_int)}",
            bg=C["bg_surface"], fg=sc_color,
            font=("Arial", 14, "bold"),
        ).grid(row=1, column=0, sticky="w", pady=(2, 0))
        row += 1

        self._info_row(self.content, "Залишковий ризик",
                       res_raw, row, 0,
                       value_color=_score_color(int(res_raw) if str(res_raw).isdigit() else 0))
        status_val = r[14] if len(r) > 14 else "—"
        status_colors = {
            "Активний":    COLORS["accent_danger"],
            "Монiторинг":  COLORS["accent_warning"],
            "Мiтигований": COLORS["accent"],
            "Закрито":     COLORS["text_muted"],
        }
        self._info_row(self.content, "Статус", status_val, row, 1,
                       value_color=status_colors.get(status_val))
        row += 1

        self._section_label(self.content, "Дати", row); row += 1
        self._info_row(self.content, "Дата виявлення",
                       r[11] if len(r) > 11 else "—", row, 0)
        self._info_row(self.content, "Дата перегляду",
                       r[12] if len(r) > 12 else "—", row, 1)
        row += 1

        self._section_label(self.content, "Деталi", row); row += 1
        self._text_block(self.content, "Заходи контролю",
                         r[9] if len(r) > 9 else "—", row)
        row += 1
        self._text_block(self.content, "Детальний опис ризику",
                         r[15] if len(r) > 15 else "—", row)
        row += 1
        tk.Frame(self.content, bg=C["bg_main"], height=12).grid(
            row=row, column=0, columnspan=2)

    # ------------------------------------------------------------------
    def _build_edit_content(self) -> None:
        C = COLORS
        rec = self.record
        for w in self.content.winfo_children():
            w.destroy()
        self.content.columnconfigure(0, weight=1)

        def section(txt: str, row_idx: int) -> int:
            f = tk.Frame(self.content, bg=C["bg_main"])
            f.grid(row=row_idx, column=0, sticky="ew", padx=8, pady=(14, 4))
            tk.Frame(f, bg=COLORS["accent_warning"], width=2, height=16).pack(side="left")
            tk.Label(f, text=txt, bg=C["bg_main"], fg=COLORS["accent_warning"],
                     font=("Arial", 9, "bold")).pack(side="left", padx=8)
            return row_idx + 1

        def lbl(text: str, row_idx: int) -> None:
            tk.Label(self.content, text=text, bg=C["bg_main"],
                     fg=C["text_subtle"], font=("Arial", 8)).grid(
                row=row_idx, column=0, sticky="w", padx=10, pady=(6, 0))

        def make_entry(**kw) -> tk.Entry:
            return tk.Entry(
                self.content,
                bg=C["bg_input"], fg=C["text_primary"],
                insertbackground=C["text_primary"],
                relief="flat", bd=2, highlightthickness=1,
                highlightbackground=C["border_soft"],
                highlightcolor=COLORS["accent_warning"],
                font=("Arial", 9), **kw,
            )

        def make_combo(values: list[str] | None = None, **kw) -> ttk.Combobox:
            return ttk.Combobox(self.content, values=values or [],
                                state="readonly", font=("Arial", 9), **kw)

        row = 0
        row = section("Пiдприємство та ризик", row)

        lbl("Скорочена назва пiдприємства:", row); row += 1
        self.e_entity = make_entry()
        self.e_entity.grid(row=row, column=0, sticky="ew", padx=10, pady=(2, 0))
        self.e_entity.insert(0, rec[1] if len(rec) > 1 else "")
        row += 1

        lbl("Назва ризику:", row); row += 1
        self.e_risk_name = make_entry()
        self.e_risk_name.grid(row=row, column=0, sticky="ew", padx=10, pady=(2, 0))
        self.e_risk_name.insert(0, rec[2] if len(rec) > 2 else "")
        row += 1

        lbl("Категорiя ризику:", row); row += 1
        self.e_category = make_combo(values=RISK_CATEGORIES)
        self.e_category.grid(row=row, column=0, sticky="ew", padx=10, pady=(2, 0))
        self.e_category.set(rec[3] if len(rec) > 3 else "")
        row += 1

        lbl("Тип ризику:", row); row += 1
        self.e_risk_type = make_combo(values=RISK_TYPES)
        self.e_risk_type.grid(row=row, column=0, sticky="ew", padx=10, pady=(2, 0))
        self.e_risk_type.set(rec[4] if len(rec) > 4 and rec[4] != "—" else "")
        row += 1

        lbl("Власник ризику:", row); row += 1
        self.e_owner = make_entry()
        self.e_owner.grid(row=row, column=0, sticky="ew", padx=10, pady=(2, 0))
        self.e_owner.insert(0, rec[8] if len(rec) > 8 else "")
        row += 1

        row = section("Оцiнка ризику", row)

        score_frame = tk.Frame(self.content, bg=C["bg_main"])
        score_frame.grid(row=row, column=0, sticky="ew", padx=10, pady=4)
        score_frame.columnconfigure((0, 1, 2), weight=1)
        row += 1

        for col_i, (lbl_t, attr, vals, val_idx) in enumerate([
            ("Iмовiрнiсть:",  "e_prob",   PROBABILITY_LEVELS, 5),
            ("Вплив:",         "e_impact", IMPACT_LEVELS,      6),
        ]):
            tk.Label(score_frame, text=lbl_t, bg=C["bg_main"],
                     fg=C["text_subtle"], font=("Arial", 8)).grid(
                row=0, column=col_i, sticky="w",
                padx=(0 if col_i == 0 else 16, 0))
            combo = ttk.Combobox(score_frame, values=vals,
                                 state="readonly", font=("Arial", 9), width=22)
            combo.grid(row=1, column=col_i, sticky="ew",
                       padx=(0 if col_i == 0 else 16, 0), pady=2)
            cur = rec[val_idx] if len(rec) > val_idx else ""
            # спробуємо підібрати за першим символом
            matched = next((v for v in vals if v.startswith(str(cur)[:1])), "")
            combo.set(cur if cur in vals else matched)
            setattr(self, attr, combo)

        # live score display
        self.lbl_live_score = tk.Label(
            score_frame, text="Score: —",
            bg=C["bg_main"], fg=C["text_muted"], font=("Arial", 11, "bold"),
        )
        self.lbl_live_score.grid(row=1, column=2, padx=20)

        def _update_score(_: object = None) -> None:
            try:
                p = _extract_num(self.e_prob.get())
                i = _extract_num(self.e_impact.get())
                s = p * i
                col = _score_color(s)
                self.lbl_live_score.configure(
                    text=f"Score: {s}  ({_score_label(s)})", fg=col)
            except Exception:
                self.lbl_live_score.configure(text="Score: —", fg=C["text_muted"])

        self.e_prob.bind("<<ComboboxSelected>>",   _update_score)
        self.e_impact.bind("<<ComboboxSelected>>", _update_score)
        _update_score()

        lbl("Залишковий ризик (1–25):", row); row += 1
        self.e_residual = make_entry()
        self.e_residual.grid(row=row, column=0, sticky="w", padx=10, pady=(2, 0))
        self.e_residual.insert(0, rec[10] if len(rec) > 10 else "")
        row += 1

        row = section("Дати", row)
        date_frame = tk.Frame(self.content, bg=C["bg_main"])
        date_frame.grid(row=row, column=0, sticky="w", padx=10, pady=4)
        row += 1

        for col_i, (lbl_t, attr, val_idx) in enumerate([
            ("Дата виявлення:", "e_date_id",  11),
            ("Дата перегляду:", "e_date_rev", 12),
        ]):
            tk.Label(date_frame, text=lbl_t, bg=C["bg_main"],
                     fg=C["text_subtle"], font=("Arial", 8)).grid(
                row=0, column=col_i,
                padx=(0 if col_i == 0 else 20, 0), sticky="w")
            e = tk.Entry(
                date_frame, bg=C["bg_input"], fg=C["text_primary"],
                insertbackground=C["text_primary"],
                relief="flat", bd=2, highlightthickness=1,
                highlightbackground=C["border_soft"],
                highlightcolor=COLORS["accent_warning"],
                font=("Arial", 9), width=14,
            )
            e.grid(row=1, column=col_i,
                   padx=(0 if col_i == 0 else 20, 0), pady=2)
            val = rec[val_idx] if len(rec) > val_idx else ""
            if val and val != "—":
                e.insert(0, val)
            else:
                _add_placeholder(e, "дд.мм.рррр")
            setattr(self, attr, e)

        row = section("Прiоритет та статус", row)
        pri_stat_frame = tk.Frame(self.content, bg=C["bg_main"])
        pri_stat_frame.grid(row=row, column=0, sticky="w", padx=10, pady=4)
        row += 1

        tk.Label(pri_stat_frame, text="Прiоритет:", bg=C["bg_main"],
                 fg=C["text_subtle"], font=("Arial", 8)).grid(row=0, column=0, sticky="w")
        self.e_priority = ttk.Combobox(
            pri_stat_frame,
            values=["Критичний", "Високий", "Середнiй", "Низький"],
            state="readonly", font=("Arial", 9), width=14,
        )
        self.e_priority.grid(row=1, column=0, pady=(2, 0), padx=(0, 20))
        self.e_priority.set(rec[13] if len(rec) > 13 else "Середнiй")

        tk.Label(pri_stat_frame, text="Статус:", bg=C["bg_main"],
                 fg=C["text_subtle"], font=("Arial", 8)).grid(row=0, column=1, sticky="w")
        self.e_status = ttk.Combobox(
            pri_stat_frame,
            values=["Активний", "Монiторинг", "Мiтигований", "Закрито"],
            state="readonly", font=("Arial", 9), width=14,
        )
        self.e_status.grid(row=1, column=1, pady=(2, 0))
        self.e_status.set(rec[14] if len(rec) > 14 else "Активний")

        row = section("Деталi", row)

        lbl("Заходи контролю:", row); row += 1
        self.e_controls = _make_dark_text(self.content, height=3, wrap="word")
        self.e_controls.grid(row=row, column=0, sticky="ew", padx=10, pady=(2, 0))
        if len(rec) > 9 and rec[9] and rec[9] != "—":
            self.e_controls.insert("1.0", rec[9])
        row += 1

        lbl("Детальний опис ризику:", row); row += 1
        self.e_description = _make_dark_text(self.content, height=4, wrap="word")
        self.e_description.grid(row=row, column=0, sticky="ew", padx=10, pady=(2, 0))
        if len(rec) > 15 and rec[15] and rec[15] != "—":
            self.e_description.insert("1.0", rec[15])
        row += 1

        tk.Frame(self.content, bg=C["bg_main"], height=16).grid(row=row, column=0)

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
        for val, label in [(date_id, "дати виявлення"), (date_rev, "дати перегляду")]:
            if val in ("дд.мм.рррр", ""):
                continue
            if not _is_valid_date(val):
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
        prob   = _extract_num(prob_str)
        impact = _extract_num(impact_str)
        score  = prob * impact

        try:
            residual = int(self.e_residual.get().strip() or "0")
        except ValueError:
            residual = 0

        old_id = self.record[0]
        new_record = (
            self.record[0],
            entity,
            risk_name,
            self.e_category.get().strip() or "—",
            self.e_risk_type.get().strip() or "—",
            prob_str or "—",
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
        status_val = new_record[14]
        status_colors = {
            "Активний":    COLORS["accent_danger"],
            "Монiторинг":  COLORS["accent_warning"],
            "Мiтигований": COLORS["accent"],
            "Закрито":     COLORS["text_muted"],
        }
        sc = status_colors.get(status_val, COLORS["text_muted"])
        self.lbl_status_badge.configure(text=f"  {status_val}  ", bg=sc)
        sc2 = _score_color(score)
        self.lbl_score_badge.configure(
            text=f"  {_score_label(score)} ({score})  ", bg=sc2)
        self._cancel_edit()
        self.toast_callback("Запис оновлено")

    def _delete_record(self) -> None:
        idx_str = self.record[0]
        if not messagebox.askyesno(
            "Пiдтвердження",
            f"Видалити ризик #{idx_str}?\nЦю дiю не можна скасувати.",
            parent=self.win,
        ):
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
        self.all_records: list[tuple] = []
        self._build_ui()
        self._load_data()

    # ------------------------------------------------------------------
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
            header, text="РЕЄСТР РИЗИКIВ",
            bg=C["bg_header"], fg=COLORS["accent_muted"],
            font=("Arial", 13, "bold"),
        ).grid(row=0, column=0, padx=20, pady=14, sticky="w")

        search_frame = tk.Frame(header, bg=C["bg_header"])
        search_frame.grid(row=0, column=1, sticky="e", padx=20)

        tk.Label(search_frame, text="Пошук:", bg=C["bg_header"],
                 fg=C["text_muted"], font=("Arial", 8)).pack(side="left", padx=(0, 6))

        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *_: self._apply_filter())
        ent_search = tk.Entry(
            search_frame, textvariable=self.search_var,
            bg=C["bg_input"], fg=C["text_primary"],
            insertbackground=C["text_primary"],
            relief="flat", bd=2, font=("Arial", 9), width=34,
        )
        ent_search.pack(side="left", padx=(0, 8), ipady=2)
        tk.Button(
            search_frame, text="Скинути",
            bg=C["bg_surface"], fg=C["text_muted"],
            activebackground=C["bg_surface_alt"],
            activeforeground=C["text_primary"],
            relief="flat", bd=0, font=("Arial", 8),
            cursor="hand2", padx=8, pady=2,
            command=self._reset_filter,
        ).pack(side="left")

        paned = ttk.PanedWindow(self.frame, orient="horizontal")
        paned.grid(row=1, column=0, sticky="nsew")
        left_wrap  = ttk.Frame(paned)
        right_wrap = ttk.Frame(paned)
        paned.add(left_wrap,  weight=4)
        paned.add(right_wrap, weight=7)

        self._build_form(left_wrap)
        self._build_table(right_wrap)

    # ------------------------------------------------------------------
    def _build_form(self, container: tk.Misc) -> None:
        C = COLORS
        container.columnconfigure(0, weight=1)
        container.rowconfigure(0, weight=1)

        canvas   = tk.Canvas(container, bg=C["bg_main"], highlightthickness=0)
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")

        form = tk.Frame(canvas, bg=C["bg_main"])
        fw   = canvas.create_window((0, 0), window=form, anchor="nw")

        def on_conf(_: object) -> None:
            canvas.configure(scrollregion=canvas.bbox("all"))
            canvas.itemconfig(fw, width=canvas.winfo_width())

        form.bind("<Configure>", on_conf)
        canvas.bind("<Configure>", on_conf)
        canvas.bind_all("<MouseWheel>",
                        lambda e: canvas.yview_scroll(int(-1*(e.delta/120)), "units"))
        form.columnconfigure(0, weight=1)

        def section(txt: str, r: int) -> int:
            f = tk.Frame(form, bg=C["bg_main"])
            f.grid(row=r, column=0, sticky="ew", padx=16, pady=(16, 4))
            tk.Frame(f, bg=COLORS["accent"], width=2, height=16).pack(side="left")
            tk.Label(f, text=txt, bg=C["bg_main"], fg=COLORS["accent"],
                     font=("Arial", 9, "bold")).pack(side="left", padx=8)
            return r + 1

        def make_entry(parent: tk.Misc, **kw) -> tk.Entry:
            return tk.Entry(
                parent, bg=C["bg_input"], fg=C["text_primary"],
                insertbackground=C["text_primary"],
                relief="flat", bd=2, highlightthickness=1,
                highlightbackground=C["border_soft"],
                highlightcolor=C["accent"], font=("Arial", 9), **kw,
            )

        def make_combo(parent: tk.Misc, values=None, **kw) -> ttk.Combobox:
            return ttk.Combobox(parent, values=values or [],
                                state="readonly", font=("Arial", 9), **kw)

        def lbl_field(text: str, r: int, factory) -> tuple:
            tk.Label(form, text=text, bg=C["bg_main"],
                     fg=C["text_subtle"], font=("Arial", 8)).grid(
                row=r, column=0, sticky="w", padx=16, pady=(4, 0))
            w = factory(form)
            w.grid(row=r + 1, column=0, sticky="ew", padx=16, pady=(2, 0))
            return w, r + 2

        row = 0

        # Badge
        badge_f = tk.Frame(form, bg=C["bg_main"])
        badge_f.grid(row=row, column=0, sticky="ew", padx=16, pady=(12, 0))
        tk.Label(badge_f, text="  + НОВИЙ РИЗИК  ",
                 bg=COLORS["accent"], fg="white",
                 font=("Arial", 8, "bold"), pady=3).pack(anchor="w")
        row += 1

        row = section("Iнформацiя про пiдприємство", row)
        self.ent_entity, row = lbl_field(
            "Скорочена назва пiдприємства:", row,
            lambda p: make_entry(p))
        self.ent_owner, row = lbl_field(
            "Власник ризику:", row, lambda p: make_entry(p))

        row = section("Опис ризику", row)
        self.ent_risk_name, row = lbl_field(
            "Назва ризику:", row, lambda p: make_entry(p))
        self.cb_category, row = lbl_field(
            "Категорiя ризику:", row,
            lambda p: make_combo(p, values=RISK_CATEGORIES))
        self.cb_risk_type, row = lbl_field(
            "Тип ризику:", row,
            lambda p: make_combo(p, values=RISK_TYPES))

        row = section("Оцiнка ризику", row)

        score_f = tk.Frame(form, bg=C["bg_main"])
        score_f.grid(row=row, column=0, sticky="ew", padx=16, pady=4)
        score_f.columnconfigure((0, 1), weight=1)
        row += 1

        for col_i, (lbl_t, attr, vals) in enumerate([
            ("Iмовiрнiсть (1–5):", "cb_prob",   PROBABILITY_LEVELS),
            ("Вплив (1–5):",        "cb_impact", IMPACT_LEVELS),
        ]):
            tk.Label(score_f, text=lbl_t, bg=C["bg_main"],
                     fg=C["text_subtle"], font=("Arial", 8)).grid(
                row=0, column=col_i, sticky="w",
                padx=(0 if col_i == 0 else 10, 0))
            combo = make_combo(score_f, values=vals, width=19)
            combo.grid(row=1, column=col_i, sticky="ew",
                       padx=(0 if col_i == 0 else 10, 0), pady=2)
            setattr(self, attr, combo)

        net_f = tk.Frame(form, bg=C["bg_main"])
        net_f.grid(row=row, column=0, sticky="ew", padx=16, pady=(2, 4))
        row += 1
        tk.Label(net_f, text="Рiвень ризику (Score):",
                 bg=C["bg_main"], fg=C["text_muted"],
                 font=("Arial", 8)).pack(side="left")
        self.lbl_score = tk.Label(net_f, text="—",
                                  bg=C["bg_main"], fg=COLORS["accent_success"],
                                  font=("Arial", 13, "bold"))
        self.lbl_score.pack(side="left", padx=10)

        def update_score(_: object = None) -> None:
            try:
                p = _extract_num(self.cb_prob.get())
                i = _extract_num(self.cb_impact.get())
                s = p * i
                c = _score_color(s)
                self.lbl_score.configure(
                    text=f"{s}  ({_score_label(s)})", fg=c)
            except Exception:
                self.lbl_score.configure(text="—", fg=COLORS["text_muted"])

        self.cb_prob.bind("<<ComboboxSelected>>",   update_score)
        self.cb_impact.bind("<<ComboboxSelected>>", update_score)

        tk.Label(form, text="Залишковий ризик (1–25):",
                 bg=C["bg_main"], fg=C["text_subtle"],
                 font=("Arial", 8)).grid(row=row, column=0, sticky="w",
                                         padx=16, pady=(6, 0))
        row += 1
        self.ent_residual = make_entry(form, width=8)
        self.ent_residual.grid(row=row, column=0, sticky="w", padx=16, pady=(2, 0))
        row += 1

        row = section("Дати", row)
        date_f = tk.Frame(form, bg=C["bg_main"])
        date_f.grid(row=row, column=0, sticky="w", padx=16, pady=4)
        row += 1
        for col_i, (lbl_t, attr) in enumerate([
            ("Дата виявлення:",  "ent_date_id"),
            ("Дата перегляду:",  "ent_date_rev"),
        ]):
            tk.Label(date_f, text=lbl_t, bg=C["bg_main"],
                     fg=C["text_subtle"], font=("Arial", 8)).grid(
                row=0, column=col_i,
                padx=(0 if col_i == 0 else 20, 0), sticky="w")
            e = make_entry(date_f, width=14)
            e.grid(row=1, column=col_i,
                   padx=(0 if col_i == 0 else 20, 0), pady=2)
            _add_placeholder(e, "дд.мм.рррр")
            setattr(self, attr, e)

        tk.Label(form, text="Прiоритет:", bg=C["bg_main"],
                 fg=C["text_subtle"], font=("Arial", 8)).grid(
            row=row, column=0, sticky="w", padx=16, pady=(8, 0))
        row += 1
        self.cb_priority = make_combo(
            form, values=["Критичний", "Високий", "Середнiй", "Низький"])
        self.cb_priority.grid(row=row, column=0, sticky="w", padx=16, pady=(2, 0))
        row += 1

        tk.Label(form, text="Статус:", bg=C["bg_main"],
                 fg=C["text_subtle"], font=("Arial", 8)).grid(
            row=row, column=0, sticky="w", padx=16, pady=(8, 0))
        row += 1
        self.cb_status = make_combo(
            form, values=["Активний", "Монiторинг", "Мiтигований", "Закрито"])
        self.cb_status.grid(row=row, column=0, sticky="w", padx=16, pady=(2, 0))
        row += 1

        row = section("Деталi", row)
        for lbl_t, attr, h in [
            ("Заходи контролю:",       "txt_controls",     3),
            ("Детальний опис ризику:", "txt_description",  4),
        ]:
            tk.Label(form, text=lbl_t, bg=C["bg_main"],
                     fg=C["text_subtle"], font=("Arial", 8)).grid(
                row=row, column=0, sticky="w", padx=16, pady=(6, 0))
            row += 1
            t = _make_dark_text(form, height=h, wrap="word")
            t.grid(row=row, column=0, sticky="ew", padx=16, pady=(2, 0))
            setattr(self, attr, t)
            row += 1

        btn_f = tk.Frame(form, bg=C["bg_main"])
        btn_f.grid(row=row, column=0, sticky="ew", padx=16, pady=(16, 20))
        btn_f.columnconfigure((0, 1), weight=1)

        tk.Button(
            btn_f, text="Очистити",
            bg=C["bg_surface"], fg=C["text_muted"],
            activebackground=C["bg_surface_alt"],
            activeforeground=C["text_primary"],
            relief="flat", bd=0, cursor="hand2",
            font=("Arial", 9), padx=14, pady=6,
            command=self._clear_form,
        ).grid(row=0, column=0, padx=4, sticky="ew")

        tk.Button(
            btn_f, text="Додати ризик",
            bg=COLORS["accent"], fg="white",
            activebackground=COLORS["accent_soft"],
            activeforeground="white",
            relief="flat", bd=0, cursor="hand2",
            font=("Arial", 9, "bold"), padx=14, pady=6,
            command=self._add_record,
        ).grid(row=0, column=1, padx=4, sticky="ew")

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
                                  fg=COLORS["accent"], font=("Arial", 8, "bold"))
        self.lbl_count.pack(side="left", pady=8)

        tk.Label(toolbar, text="  |  Тип:", bg=C["bg_surface"],
                 fg=C["text_muted"], font=("Arial", 8)).pack(side="left", pady=8)
        self.filter_type = ttk.Combobox(toolbar, width=16, state="readonly",
                                        values=["Всi"] + RISK_TYPES, font=("Arial", 8))
        self.filter_type.set("Всi")
        self.filter_type.pack(side="left", padx=6, pady=8)
        self.filter_type.bind("<<ComboboxSelected>>", lambda _: self._apply_filter())

        tk.Label(toolbar, text="Статус:", bg=C["bg_surface"],
                 fg=C["text_muted"], font=("Arial", 8)).pack(side="left", pady=8)
        self.filter_status = ttk.Combobox(
            toolbar, width=12, state="readonly",
            values=["Всi", "Активний", "Монiторинг", "Мiтигований", "Закрито"],
            font=("Arial", 8))
        self.filter_status.set("Всi")
        self.filter_status.pack(side="left", padx=6, pady=8)
        self.filter_status.bind("<<ComboboxSelected>>", lambda _: self._apply_filter())

        for txt, cmd, bg in [
            ("Переглянути", self._open_selected_detail, COLORS["accent"]),
            ("Дублювати",   self._duplicate_record,      C["bg_surface_alt"]),
            ("Видалити",    self._delete_selected,        COLORS["accent_danger"]),
        ]:
            tk.Button(
                toolbar, text=txt, bg=bg,
                fg="white" if bg != C["bg_surface_alt"] else C["text_primary"],
                activebackground=bg, activeforeground="white",
                relief="flat", bd=0, cursor="hand2",
                font=("Arial", 8), padx=10, pady=3, command=cmd,
            ).pack(side="right", padx=4, pady=6)

        tree_frame = ttk.Frame(container)
        tree_frame.grid(row=1, column=0, sticky="nsew")
        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)

        cols = ("id", "entity", "risk_name", "category", "risk_type",
                "score", "priority", "status", "owner", "date_id")
        self.tree = ttk.Treeview(tree_frame, columns=cols,
                                  show="headings", selectmode="browse")
        self.tree.grid(row=0, column=0, sticky="nsew")

        headers = {
            "id":        ("№",          46),
            "entity":    ("Пiдприємство", 150),
            "risk_name": ("Назва ризику", 200),
            "category":  ("Категорiя",   110),
            "risk_type": ("Тип ризику",  110),
            "score":     ("Score",        62),
            "priority":  ("Прiоритет",    88),
            "status":    ("Статус",        90),
            "owner":     ("Власник",      130),
            "date_id":   ("Дата виявл.", 100),
        }
        for col, (txt, w) in headers.items():
            self.tree.heading(col, text=txt,
                               command=lambda c=col: self._sort_tree(c))
            self.tree.column(col, width=w, anchor="w")

        sy = ttk.Scrollbar(tree_frame, orient="vertical",   command=self.tree.yview)
        sx = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        sy.grid(row=0, column=1, sticky="ns")
        sx.grid(row=1, column=0, sticky="ew")
        self.tree.configure(yscrollcommand=sy.set, xscrollcommand=sx.set)

        self.tree.tag_configure("even", background=C["row_even"])
        self.tree.tag_configure("odd",  background=C["row_odd"])
        for risk, color in RISK_COLORS.items():
            self.tree.tag_configure(f"risk_{risk}", foreground=color)
        # Score-level tags
        for tag, color in [
            ("score_low",    COLORS["accent_success"]),
            ("score_mod",    COLORS["accent_warning"]),
            ("score_high",   "#f97316"),
            ("score_crit",   COLORS["accent_danger"]),
        ]:
            self.tree.tag_configure(tag, foreground=color)

        self.tree.bind("<<TreeviewSelect>>", self._on_select)
        self.tree.bind("<Double-1>", self._on_double_click)

        tk.Frame(container, bg=C["bg_surface"]).grid(row=2, column=0, sticky="ew")
        tk.Label(
            container,
            text="  Подвiйний клiк по рядку — переглянути / редагувати ризик",
            bg=C["bg_surface"], fg=C["text_muted"],
            font=("Arial", 7, "italic"),
        ).grid(row=2, column=0, sticky="ew", ipadx=8, ipady=4)

        # Quick-view panel
        detail_frame = tk.Frame(container, bg=C["bg_surface"])
        detail_frame.grid(row=3, column=0, sticky="ew")
        detail_frame.columnconfigure((0, 1), weight=1)

        for col_i, (lbl_t, attr) in enumerate([
            ("Заходи контролю",    "det_controls"),
            ("Опис ризику",        "det_desc"),
        ]):
            sub = tk.Frame(detail_frame, bg=C["bg_surface"])
            sub.grid(row=0, column=col_i, sticky="nsew",
                     padx=(12 if col_i == 0 else 4, 4), pady=8)
            sub.columnconfigure(0, weight=1)
            tk.Label(sub, text=lbl_t, bg=C["bg_surface"],
                     fg=C["text_muted"], font=("Arial", 7, "bold")).grid(
                row=0, column=0, sticky="w")
            t = _make_dark_text(sub, height=4, wrap="word", state="disabled")
            t.grid(row=1, column=0, sticky="ew", pady=(2, 0))
            setattr(self, attr, t)

        # Export bar
        exp_bar = tk.Frame(container, bg=C["bg_main"])
        exp_bar.grid(row=4, column=0, sticky="ew", padx=8, pady=6)

        tk.Button(exp_bar, text="Експорт CSV",
                  bg=C["bg_surface"], fg=C["text_primary"],
                  activebackground=C["bg_surface_alt"],
                  activeforeground=C["text_primary"],
                  relief="flat", bd=0, cursor="hand2",
                  font=("Arial", 8), padx=12, pady=4,
                  command=self._export_csv).pack(side="left", padx=(0, 6))

        if pd:
            tk.Button(exp_bar, text="Експорт Excel",
                      bg=COLORS["accent_success"], fg="white",
                      activebackground="#16a34a", activeforeground="white",
                      relief="flat", bd=0, cursor="hand2",
                      font=("Arial", 8), padx=12, pady=4,
                      command=self._export_excel).pack(side="left", padx=(0, 6))

        tk.Button(exp_bar, text="Iмпорт JSON",
                  bg=C["bg_surface"], fg=C["text_primary"],
                  activebackground=C["bg_surface_alt"],
                  activeforeground=C["text_primary"],
                  relief="flat", bd=0, cursor="hand2",
                  font=("Arial", 8), padx=12, pady=4,
                  command=self._import_json).pack(side="left")

    # ------------------------------------------------------------------
    #  Подвійний клік / деталі
    # ------------------------------------------------------------------
    def _on_double_click(self, event: tk.Event) -> None:
        sel = self.tree.selection()
        if not sel:
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
            save_callback=lambda old_id, new_rec:
                self._on_detail_save(iid, old_id, new_rec),
            delete_callback=lambda idx_s:
                self._on_detail_delete(iid, idx_s),
            toast_callback=self._show_toast,
        )

    def _find_record(self, idx_str: str) -> tuple | None:
        for r in self.all_records:
            if (str(r[0]) == idx_str or
                    str(r[0]).lstrip("0") == str(idx_str).lstrip("0")):
                return r
        return None

    def _on_detail_save(self, iid: str, old_id: str, new_record: tuple) -> None:
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
        self.all_records = [r for r in self.all_records if str(r[0]) != idx_str]
        self._recolor_rows()
        self._save_data()
        self._update_count()
        if self.on_data_change:
            self.on_data_change(self.all_records)

    # ------------------------------------------------------------------
    #  Сортування / перефарбовка
    # ------------------------------------------------------------------
    def _sort_tree(self, col: str) -> None:
        data = [(self.tree.set(iid, col), iid)
                for iid in self.tree.get_children("")]
        try:
            data.sort(key=lambda x: float(x[0]) if x[0] not in ("—", "") else 0)
        except ValueError:
            data.sort(key=lambda x: x[0].lower())
        for i, (_, iid) in enumerate(data):
            self.tree.move(iid, "", i)
        self._recolor_rows()

    def _recolor_rows(self) -> None:
        for i, iid in enumerate(self.tree.get_children()):
            risk  = self.tree.set(iid, "risk_type")
            score_str = self.tree.set(iid, "score")
            base_tag = "even" if i % 2 == 0 else "odd"
            tags = [base_tag]
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
    #  Збереження / завантаження
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
                    row = self._normalize_record(row)
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
        data = self._normalize_record(list(data))
        return (
            data[0],   # id
            data[1],   # entity
            data[2],   # risk_name
            data[3],   # category
            data[4],   # risk_type
            data[7],   # score
            data[13],  # priority
            data[14],  # status
            data[8],   # owner
            data[11],  # date_id
        )

    def _insert_tree_row(self, data: tuple) -> str:
        iid = self.tree.insert("", tk.END, values=self._tree_values(data))
        self._recolor_rows()
        return iid

    # ------------------------------------------------------------------
    #  Форма
    # ------------------------------------------------------------------
    def _get_form_data(self) -> tuple | None:
        date_id  = self.ent_date_id.get().strip()
        date_rev = self.ent_date_rev.get().strip()
        for val, label in [
            (date_id,  "дати виявлення"),
            (date_rev, "дати перегляду"),
        ]:
            if val in ("дд.мм.рррр", ""):
                continue
            if not _is_valid_date(val):
                messagebox.showwarning(
                    "Помилка",
                    f"Неправильний формат {label} (очiкується дд.мм.рррр)")
                return None
        date_id  = "" if date_id  == "дд.мм.рррр" else date_id
        date_rev = "" if date_rev == "дд.мм.рррр" else date_rev

        prob_str   = self.cb_prob.get().strip()
        impact_str = self.cb_impact.get().strip()
        prob   = _extract_num(prob_str)
        impact = _extract_num(impact_str)
        score  = prob * impact

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
        for w in [self.ent_entity, self.ent_owner, self.ent_risk_name,
                  self.ent_residual]:
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
            _add_placeholder(e, ph)

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
    #  Видалення / дублювання
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

    # ------------------------------------------------------------------
    #  Фільтри / пошук
    # ------------------------------------------------------------------
    def _apply_filter(self) -> None:
        q      = self.search_var.get().strip().lower()
        r_type = self.filter_type.get()
        status = self.filter_status.get()
        self.tree.delete(*self.tree.get_children())
        for row in self.all_records:
            row_str = " ".join(str(v).lower() for v in row)
            if q and q not in row_str:
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
    #  Вибір рядка
    # ------------------------------------------------------------------
    def _on_select(self, _: object | None = None) -> None:
        sel = self.tree.selection()
        if not sel:
            return
        idx_str = self.tree.set(sel[0], "id")
        rec = self._find_record(idx_str)
        if not rec:
            return
        controls = rec[9]  if len(rec) > 9  else ""
        desc     = rec[15] if len(rec) > 15 else ""
        for widget, text in [
            (self.det_controls, controls),
            (self.det_desc,     desc),
        ]:
            widget.configure(state="normal")
            widget.delete("1.0", tk.END)
            widget.insert("1.0", text or "")
            widget.configure(state="disabled")

    # ------------------------------------------------------------------
    def _update_count(self) -> None:
        self.lbl_count.configure(text=f" {len(self.tree.get_children())}")

    def _show_toast(self, msg: str) -> None:
        toast = tk.Toplevel(self.frame)
        toast.overrideredirect(True)
        toast.attributes("-topmost", True)
        toast.configure(bg=COLORS["accent_success"])
        tk.Label(toast, text=f"  {msg}  ",
                 bg=COLORS["accent_success"], fg="white",
                 font=("Arial", 9, "bold"), pady=6).pack()
        root = self.frame.winfo_toplevel()
        x = root.winfo_x() + root.winfo_width() - 220
        y = root.winfo_y() + root.winfo_height() - 80
        toast.geometry(f"+{x}+{y}")
        toast.after(2000, toast.destroy)

    # ------------------------------------------------------------------
    #  Експорт / імпорт
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
            n = len(self._HEADERS)
            data = [
                r if len(r) == n else list(r) + [""] * (n - len(r))
                for r in self.all_records
            ]
            df = pd.DataFrame(data, columns=self._HEADERS)
            with pd.ExcelWriter(path, engine="openpyxl") as writer:
                df.to_excel(writer, index=False, sheet_name="Реєстр ризикiв")
                ws = writer.sheets["Реєстр ризикiв"]
                for col_cells in ws.columns:
                    max_len = max(len(str(c.value or "")) for c in col_cells)
                    ws.column_dimensions[
                        col_cells[0].column_letter].width = min(max_len + 4, 60)
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
                    row = self._normalize_record(list(row))
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
#  ВКЛАДКА: АНАЛІТИКА РИЗИКІВ
# =============================================================================

class RiskAnalyticsTab:
    """Аналітика реєстру ризиків."""

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
                 bg=C["bg_header"], fg=COLORS["accent_muted"],
                 font=("Arial", 13, "bold")).pack(side="left", padx=20, pady=14)
        tk.Button(header, text="Оновити",
                  bg=COLORS["accent"], fg="white",
                  activebackground=COLORS["accent_soft"],
                  activeforeground="white",
                  relief="flat", bd=0, cursor="hand2",
                  font=("Arial", 9, "bold"), padx=12, pady=4,
                  command=self.refresh).pack(side="right", padx=20, pady=12)

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
        canvas.bind_all("<MouseWheel>",
                        lambda e: canvas.yview_scroll(int(-1*e.delta/120), "units"))
        self.content.columnconfigure(0, weight=1)

        self._build_stat_cards()
        self._build_charts_and_table()

    def _build_stat_cards(self) -> None:
        C = COLORS
        cards_frame = tk.Frame(self.content, bg=C["bg_main"])
        cards_frame.grid(row=0, column=0, sticky="ew", padx=16, pady=(16, 0))
        for i in range(5):
            cards_frame.columnconfigure(i, weight=1)

        self.stat_cards: dict[str, tk.Label] = {}
        defs = [
            ("total",     "Всього ризикiв",       "0", C["accent"]),
            ("critical",  "Критичних (Score >16)", "0", C["accent_danger"]),
            ("high",      "Високих (Score 10–16)", "0", "#f97316"),
            ("active",    "Активних",               "0", C["accent_warning"]),
            ("mitigated", "Мiтигованих",            "0", C["accent_success"]),
        ]
        for col, (key, title, val, color) in enumerate(defs):
            card = tk.Frame(cards_frame, bg=C["bg_surface"], padx=18, pady=12)
            card.grid(row=0, column=col, padx=6, sticky="nsew")
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

            # Pie — розподіл за типом ризику
            self.fig_left = Figure(figsize=(5, 3.5), dpi=90, facecolor=C["bg_surface"])
            self.ax_left  = self.fig_left.add_subplot(111)
            self._style_ax(self.ax_left)
            self.ax_left.set_title("Розподiл за типом ризику",
                                    color=C["text_muted"], fontsize=9)
            frame_l = tk.Frame(charts_row, bg=C["bg_surface"], padx=8, pady=8)
            frame_l.grid(row=0, column=0, padx=(0, 8), sticky="nsew")
            self.canvas_left = FigureCanvasTkAgg(self.fig_left, master=frame_l)
            self.canvas_left.get_tk_widget().pack(fill="both", expand=True)

            # Bar — розподіл за рівнем ризику
            self.fig_right = Figure(figsize=(5, 3.5), dpi=90, facecolor=C["bg_surface"])
            self.ax_right  = self.fig_right.add_subplot(111)
            self._style_ax(self.ax_right)
            self.ax_right.set_title("Розподiл за рiвнем ризику",
                                     color=C["text_muted"], fontsize=9)
            frame_r = tk.Frame(charts_row, bg=C["bg_surface"], padx=8, pady=8)
            frame_r.grid(row=0, column=1, padx=(8, 0), sticky="nsew")
            self.canvas_right = FigureCanvasTkAgg(self.fig_right, master=frame_r)
            self.canvas_right.get_tk_widget().pack(fill="both", expand=True)

            # Heatmap — матриця ризиків (probability x impact)
            self.fig_heat = Figure(figsize=(10, 4), dpi=90, facecolor=C["bg_surface"])
            self.ax_heat  = self.fig_heat.add_subplot(111)
            self._style_ax(self.ax_heat)
            self.ax_heat.set_title("Матриця ризикiв (Iмовiрнiсть x Вплив)",
                                    color=C["text_muted"], fontsize=9)
            frame_h = tk.Frame(self.content, bg=C["bg_surface"], padx=8, pady=8)
            frame_h.grid(row=2, column=0, padx=16, pady=(0, 16), sticky="ew")
            self.canvas_heat = FigureCanvasTkAgg(self.fig_heat, master=frame_h)
            self.canvas_heat.get_tk_widget().pack(fill="both", expand=True)
        else:
            tk.Label(
                self.content,
                text="Встановiть matplotlib для вiдображення графiкiв:\n"
                     "  pip install matplotlib",
                bg=COLORS["bg_main"], fg=COLORS["text_muted"],
                font=("Arial", 10),
            ).grid(row=1, column=0, pady=40)

        # Таблиця статистики
        frame = tk.Frame(self.content, bg=COLORS["bg_surface"], padx=16, pady=12)
        frame.grid(row=3, column=0, padx=16, pady=(0, 16), sticky="ew")
        frame.columnconfigure(0, weight=1)
        tk.Label(frame, text="Деталiзована статистика за типом ризику",
                 bg=COLORS["bg_surface"], fg=COLORS["text_muted"],
                 font=("Arial", 9, "bold")).grid(row=0, column=0, sticky="w",
                                                  pady=(0, 8))
        cols = ("risk_type", "count", "avg_score", "max_score", "active")
        self.stats_tree = ttk.Treeview(frame, columns=cols,
                                        show="headings", height=7)
        for col, hdr, w in [
            ("risk_type",  "Тип ризику",     180),
            ("count",      "Всього",           70),
            ("avg_score",  "Сер. Score",       90),
            ("max_score",  "Макс. Score",      90),
            ("active",     "Активних",         80),
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

    def update_data(self, records: list[tuple]) -> None:
        self.records = records
        self.refresh()

    def refresh(self) -> None:
        if not self.records:
            for key in ["total", "critical", "high", "active", "mitigated"]:
                self.stat_cards[key].configure(text="0")
            if HAS_MPL:
                self._clear_charts()
            self.stats_tree.delete(*self.stats_tree.get_children())
            return

        C = COLORS
        records = self.records

        total = len(records)
        critical_c = sum(
            1 for r in records
            if len(r) > 7 and str(r[7]).isdigit() and int(r[7]) > 16
        )
        high_c = sum(
            1 for r in records
            if len(r) > 7 and str(r[7]).isdigit() and 10 <= int(r[7]) <= 16
        )
        active = sum(
            1 for r in records
            if len(r) > 14 and r[14] == "Активний"
        )
        mitigated = sum(
            1 for r in records
            if len(r) > 14 and r[14] == "Мiтигований"
        )

        self.stat_cards["total"].configure(text=str(total))
        self.stat_cards["critical"].configure(text=str(critical_c))
        self.stat_cards["high"].configure(text=str(high_c))
        self.stat_cards["active"].configure(text=str(active))
        self.stat_cards["mitigated"].configure(text=str(mitigated))

        if not HAS_MPL:
            return

        type_counter = Counter(r[4] for r in records if len(r) > 4)
        
        # Pie по типах ризику
        self.ax_left.clear()
        self._style_ax(self.ax_left)
        self.ax_left.set_title("Розподiл за типом ризику",
                                color=C["text_muted"], fontsize=9)
        if type_counter:
            labels = list(type_counter.keys())
            values = list(type_counter.values())
            colors = [RISK_COLORS.get(l, C["text_muted"]) for l in labels]
            wedges, texts, autotexts = self.ax_left.pie(
                values, labels=labels, autopct="%1.0f%%",
                colors=colors, startangle=90,
                textprops={"color": C["text_muted"], "fontsize": 7},
            )
            for at in autotexts:
                at.set_fontsize(7)
                at.set_color("white")
        else:
            self.ax_left.text(0.5, 0.5, "Немає даних",
                              transform=self.ax_left.transAxes,
                              ha="center", va="center", color=C["text_muted"])
        self.canvas_left.draw()

        # Bar за рівнем ризику
        self.ax_right.clear()
        self._style_ax(self.ax_right)
        self.ax_right.set_title("Розподiл за рiвнем ризику",
                                 color=C["text_muted"], fontsize=9)
        
        level_counter = {"Низький": 0, "Помiрний": 0, "Високий": 0, "Критичний": 0}
        for r in records:
            if len(r) > 7 and str(r[7]).isdigit():
                s = int(r[7])
                level_counter[_score_label(s)] += 1
        
        if any(level_counter.values()):
            labels = list(level_counter.keys())
            values = list(level_counter.values())
            colors_bar = [
                COLORS["accent_success"],
                COLORS["accent_warning"],
                "#f97316",
                COLORS["accent_danger"],
            ]
            bars = self.ax_right.bar(labels, values, color=colors_bar, edgecolor="none")
            for bar, val in zip(bars, values, strict=False):
                if val > 0:
                    self.ax_right.text(
                        bar.get_x() + bar.get_width() / 2,
                        bar.get_height() + 0.1,
                        str(val), ha="center", va="bottom",
                        color=C["text_muted"], fontsize=8,
                    )
            self.ax_right.tick_params(axis="x", labelrotation=15, labelsize=7)
            self.ax_right.set_ylim(0, max(values) * 1.2 + 1 if max(values) > 0 else 1)
        else:
            self.ax_right.text(0.5, 0.5, "Немає даних",
                               transform=self.ax_right.transAxes,
                               ha="center", va="center", color=C["text_muted"])
        self.canvas_right.draw()

        # Heatmap матриці ризиків
        self.ax_heat.clear()
        self._style_ax(self.ax_heat)
        self.ax_heat.set_title("Матриця ризикiв (Iмовiрнiсть x Вплив)",
                                color=C["text_muted"], fontsize=9)
        
        matrix = [[0]*5 for _ in range(5)]
        for r in records:
            if len(r) > 6:
                try:
                    prob = _extract_num(r[5])
                    imp  = _extract_num(r[6])
                    if 1 <= prob <= 5 and 1 <= imp <= 5:
                        matrix[5 - prob][imp - 1] += 1
                except (ValueError, IndexError):
                    pass
        
        if any(any(row) for row in matrix):
            import numpy as np
            cmap = plt.cm.RdYlGn_r
            im = self.ax_heat.imshow(matrix, cmap=cmap, aspect='auto')
            self.ax_heat.set_xticks(range(5))
            self.ax_heat.set_yticks(range(5))
            self.ax_heat.set_xticklabels([f"{i}" for i in range(1, 6)])
            self.ax_heat.set_yticklabels([f"{i}" for i in range(5, 0, -1)])
            self.ax_heat.set_xlabel("Вплив →", color=C["text_muted"], fontsize=8)
            self.ax_heat.set_ylabel("Iмовiрнiсть →", color=C["text_muted"], fontsize=8)
            
            for i in range(5):
                for j in range(5):
                    val = matrix[i][j]
                    if val > 0:
                        self.ax_heat.text(j, i, str(val), ha="center", va="center",
                                          color="white", fontsize=10, weight="bold")
        else:
            self.ax_heat.text(0.5, 0.5, "Немає даних",
                              transform=self.ax_heat.transAxes,
                              ha="center", va="center", color=C["text_muted"])
        self.canvas_heat.draw()

        # Таблиця статистики
        self.stats_tree.delete(*self.stats_tree.get_children())
        all_types = set(RISK_TYPES) | set(r[4] for r in records if len(r) > 4)
        for risk_type in sorted(all_types):
            recs = [r for r in records if len(r) > 4 and r[4] == risk_type]
            cnt = len(recs)
            if cnt:
                scores = [int(r[7]) for r in recs 
                          if len(r) > 7 and str(r[7]).isdigit()]
                avg_score = sum(scores) / len(scores) if scores else 0
                max_score = max(scores) if scores else 0
                active_c = sum(1 for r in recs if len(r) > 14 and r[14] == "Активний")
                self.stats_tree.insert(
                    "", tk.END,
                    values=(risk_type, cnt, f"{avg_score:.1f}", max_score, active_c)
                )

    def _clear_charts(self) -> None:
        if HAS_MPL:
            self.ax_left.clear()
            self.ax_right.clear()
            self.ax_heat.clear()
            self.canvas_left.draw()
            self.canvas_right.draw()
            self.canvas_heat.draw()

    def get_frame(self) -> ttk.Frame:
        return self.frame


# =============================================================================
#  ВКЛАДКА: НАЛАШТУВАННЯ РИЗИКІВ
# =============================================================================

class RiskSettingsTab:
    """Вкладка налаштувань модуля 'Реєстр ризиків'."""

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
            header, text="НАЛАШТУВАННЯ РЕЄСТРУ РИЗИКIВ",
            bg=C["bg_header"], fg=C["text_muted"],
            font=("Arial", 13, "bold"),
        ).pack(side="left", padx=20, pady=14)

        content = tk.Frame(self.frame, bg=C["bg_main"])
        content.grid(row=1, column=0, sticky="nsew", padx=40, pady=30)
        content.columnconfigure(0, weight=1)

        self._row(content, 0, "Файл даних:", RISK_DATA_FILE, C)
        self._row(content, 1, "Версiя:", "1.0 — Реєстр ризикiв", C)
        self._row(content, 2, "matplotlib:",
                  "встановлено" if HAS_MPL else "не встановлено", C)
        self._row(content, 3, "pandas:",
                  "встановлено" if pd else "не встановлено", C)

        tk.Label(
            content, text="Встановлення залежностей:",
            bg=C["bg_main"], fg=C["text_muted"],
            font=("Arial", 8, "bold"),
        ).grid(row=4, column=0, sticky="w", pady=(24, 6))
        tk.Label(
            content, text="  pip install matplotlib pandas openpyxl",
            bg=C["bg_surface"], fg=COLORS["accent_muted"],
            font=("Courier", 9), padx=12, pady=8,
        ).grid(row=5, column=0, sticky="w")

        tk.Label(
            content, text="Структура запису (16 полiв):",
            bg=C["bg_main"], fg=C["text_muted"],
            font=("Arial", 8, "bold"),
        ).grid(row=6, column=0, sticky="w", pady=(24, 6))

        fields_text = """ID, Пiдприємство, Назва ризику, Категорiя, Тип ризику,
Iмовiрнiсть, Вплив, Score, Власник, Заходи контролю,
Залишковий ризик, Дата виявлення, Дата перегляду, Прiоритет, Статус, Опис"""
        
        tk.Label(
            content, text=fields_text,
            bg=C["bg_surface"], fg=C["text_subtle"],
            font=("Arial", 8), justify="left", padx=12, pady=8,
        ).grid(row=7, column=0, sticky="w")

        tk.Label(
            content, text="Пiдказки:",
            bg=C["bg_main"], fg=C["text_muted"],
            font=("Arial", 8, "bold"),
        ).grid(row=8, column=0, sticky="w", pady=(24, 6))

        hints = [
            "Score розраховується автоматично: Iмовiрнiсть × Вплив (1–25)",
            "Рівні ризику: Низький (1–4), Помiрний (5–9), Високий (10–16), Критичний (17–25)",
            "Подвiйний клiк по рядку — вiдкрити детальне вiкно ризику",
            "Матриця ризикiв показує розподiл за ймовiрнiстю та впливом",
        ]
        for i, hint in enumerate(hints):
            f = tk.Frame(content, bg=C["bg_main"])
            f.grid(row=9 + i, column=0, sticky="w", pady=2)
            tk.Frame(f, bg=COLORS["accent_success"], width=4, height=4).pack(
                side="left", padx=(0, 8))
            tk.Label(f, text=hint, bg=C["bg_main"], fg=C["text_subtle"],
                     font=("Arial", 8)).pack(side="left")

    def _row(self, parent: tk.Misc, row: int, label: str, value: str, C: dict) -> None:
        f = tk.Frame(parent, bg=C["bg_main"])
        f.grid(row=row, column=0, sticky="ew", pady=4)
        tk.Label(f, text=label, bg=C["bg_main"], fg=C["text_muted"],
                 font=("Arial", 9), width=22, anchor="w").pack(side="left")
        tk.Label(f, text=value, bg=C["bg_main"], fg=C["text_primary"],
                 font=("Arial", 9)).pack(side="left")

    def get_frame(self) -> ttk.Frame:
        return self.frame


# =============================================================================
#  СТОРІНКА "РЕЄСТР РИЗИКІВ" ДЛЯ ATLAS
# =============================================================================

class RiskRegisterPage(tk.Frame):
    """Сторінка 'Реєстр ризиків' (вкладки + статусбар)."""

    def __init__(self, master: tk.Misc, **kwargs) -> None:
        super().__init__(master, bg=COLORS["bg_main"], **kwargs)

        self.grid_rowconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=0)
        self.grid_columnconfigure(0, weight=1)

        self.notebook = ttk.Notebook(self)
        self.notebook.grid(row=0, column=0, sticky="nsew")

        self.analytics_tab = RiskAnalyticsTab(self.notebook)
        self.registry_tab  = RiskRegistryTab(
            self.notebook, on_data_change=self.analytics_tab.update_data)
        self.settings_tab  = RiskSettingsTab(self.notebook)

        self.notebook.add(self.registry_tab.get_frame(),  text="  Реєстр ризикiв  ")
        self.notebook.add(self.analytics_tab.get_frame(), text="  Аналiтика  ")
        self.notebook.add(self.settings_tab.get_frame(),  text="  Налаштування  ")

        statusbar = tk.Frame(self, bg=COLORS["bg_header"], height=22)
        statusbar.grid(row=1, column=0, sticky="ew")
        statusbar.grid_propagate(False)

        self._status_lbl = tk.Label(
            statusbar, text="Готово", bg=COLORS["bg_header"],
            fg=COLORS["text_muted"], font=("Arial", 7), padx=10,
        )
        self._status_lbl.pack(side="left", pady=3)

        self._time_lbl = tk.Label(
            statusbar, text="", bg=COLORS["bg_header"],
            fg=COLORS["text_muted"], font=("Arial", 7), padx=10,
        )
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
            self.registry_tab._save_data()
            self._status_lbl.configure(
                text=f"Автозбережено о {datetime.now().strftime('%H:%M:%S')}")
        except Exception:
            self._status_lbl.configure(text="Помилка автозбереження")
        self.after(30000, self._schedule_autosave)

    def save_before_exit(self) -> None:
        try:
            self.registry_tab._save_data()
        except Exception:
            pass


# =============================================================================
#  ДЕМО / STANDALONE
# =============================================================================

def main() -> int:
    """Запуск модуля окремо для тестування."""
    root = tk.Tk()
    root.title("ATLAS | Реєстр ризиків")
    root.geometry("1300x800")
    root.configure(bg=COLORS["bg_main"])
    
    try:
        from __main__ import apply_dark_style
        apply_dark_style(root)
    except ImportError:
        pass
    
    page = RiskRegisterPage(root)
    page.pack(fill="both", expand=True)
    
    root.protocol("WM_DELETE_WINDOW", lambda: (page.save_before_exit(), root.destroy()))
    root.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
