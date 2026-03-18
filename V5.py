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
RISK_DATA_FILE = "risk_register.json"

COLORS = {
    "bg_main": "#243640", "bg_sidebar": "#1E2C33", "bg_header": "#1E2C33",
    "bg_surface": "#2E4450", "bg_surface_alt": "#344E5A", "bg_input": "#1E2C33",
    "accent": "#4F46E5", "accent_soft": "#6366F1", "accent_muted": "#818CF8",
    "accent_success": "#10B981", "accent_danger": "#EF4444", "accent_warning": "#F59E0B",
    "text_primary": "#F3F4F6", "text_muted": "#CBD5E1", "text_subtle": "#94A3B8",
    "border_soft": "#3B4F59", "border_strong": "#556871",
    "row_even": "#2A3D47", "row_odd": "#243640", "row_select": "#3B82F6",
}

RISK_COLORS = {
    "Операцiйний": COLORS["accent_warning"],
    "Технiчний": COLORS["accent"],
    "Фiнансовий": COLORS["accent_danger"],
    "Репутацiйний": "#a855f7",
    "Екологiчний": COLORS["accent_success"],
    "Надзвичайна ситуацiя": "#f97316",
}

RISK_TYPES = [
    "Операцiйний", "Технiчний", "Фiнансовий", "Репутацiйний",
    "Екологiчний", "Надзвичайна ситуацiя",
]

EVENT_TYPES = [
    "Вимушений простiй < 24 год", "Вимушений простiй >= 24 год",
    "Зупинка виробництва", "Аварiя обладнання", "Пошкодження майна",
    "Порушення дозволiв", "Крадiжка / диверсiя", "Iнше",
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
#  ХЕЛПЕРИ ТА СТИЛЬ (спільні)
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
    C = COLORS
    return tk.Text(
        parent, bg=C["bg_input"], fg=C["text_primary"], insertbackground=C["text_primary"],
        selectbackground=C["row_select"], relief="flat", bd=1, highlightthickness=1,
        highlightbackground=C["border_soft"], highlightcolor=C["accent"],
        font=("Arial", 9), **kwargs,
    )

def add_placeholder(entry: tk.Entry, text: str) -> None:
    entry.insert(0, text)
    entry.configure(fg=COLORS["text_muted"])

    def on_in(_): 
        if entry.get() == text:
            entry.delete(0, tk.END)
            entry.configure(fg=COLORS["text_primary"])
    def on_out(_):
        if not entry.get():
            entry.insert(0, text)
            entry.configure(fg=COLORS["text_muted"])

    entry.bind("<FocusIn>", on_in)
    entry.bind("<FocusOut>", on_out)

# =============================================================================
#  ІМПОРТУЄМО/ВСТАВЛЯЄМО ОБИДВА МОДУЛІ
# =============================================================================

# === 1. Material Events (ваш оригінальний код) ===
# Вставте сюди весь ваш оригінальний код від class EventDetailWindow до class MaterialEventsPage
# (щоб не робити повідомлення надто довгим, я позначаю місця вставки)

# ... [ВСТАВТЕ ТУТ ВЕСЬ ВАШ ОРИГІНАЛЬНИЙ КОД від EventDetailWindow до MaterialEventsPage включно] ...

# === 2. Risk Register (новий модуль) ===
# Вставте сюди весь наданий вами код risk_register.py 
# (від RISK_CATEGORIES до class RiskRegisterPage)

# ... [ВСТАВТЕ ТУТ ВЕСЬ КОД РЕЄСТРУ РИЗИКІВ] ...

# =============================================================================
#  ATLAS APP (оновлений)
# =============================================================================

PageKey: TypeAlias = Literal[
    "risk_register", "material_events", "risk_appetite",
    "analytics", "reports", "risk_coordinators", "settings",
]

APP_TITLE = "ATLAS"
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
        topbar = tk.Frame(self, bg=COLORS["bg_header"], height=56)
        topbar.grid(row=0, column=0, columnspan=2, sticky="ew")
        topbar.grid_propagate(False)
        topbar.columnconfigure(1, weight=1)

        tk.Label(topbar, text=APP_TITLE, bg=COLORS["bg_header"],
                 fg=COLORS["text_primary"], font=("Arial", 14, "bold")).grid(
            row=0, column=0, padx=20, pady=8, sticky="w")

        tk.Label(topbar, text=self._user_full_name, bg=COLORS["bg_header"],
                 fg=COLORS["text_muted"], font=("Arial", 10)).grid(
            row=0, column=1, padx=8, pady=8, sticky="e")

        tk.Button(topbar, text="🔔", bg=COLORS["bg_surface"], fg=COLORS["text_primary"],
                  activebackground=COLORS["bg_surface_alt"], relief="flat", bd=0,
                  cursor="hand2", font=("Arial", 14), padx=10, pady=4,
                  command=lambda: messagebox.showinfo("Нагадування", "Нагадувань поки немає.")
                  ).grid(row=0, column=2, padx=20, pady=8, sticky="e")

    def _build_sidebar(self) -> None:
        sidebar = tk.Frame(self, bg=COLORS["bg_sidebar"], width=220)
        sidebar.grid(row=1, column=0, sticky="nsw")
        sidebar.grid_propagate(False)

        tk.Label(sidebar, text="Навігація", bg=COLORS["bg_sidebar"],
                 fg=COLORS["text_muted"], font=("Arial", 10, "bold")).grid(
            row=0, column=0, padx=16, pady=(8, 4), sticky="w")

        nav_frame = tk.Frame(sidebar, bg=COLORS["bg_sidebar"])
        nav_frame.grid(row=1, column=0, sticky="nsew", padx=8)

        menu_items: list[tuple[PageKey, str]] = [
            ("risk_register",    "Реєстр ризиків"),
            ("material_events",  "Реєстр суттєвих подій"),
            ("risk_appetite",    "Ризик апетит"),
            ("analytics",        "Аналітика"),
            ("reports",          "Звіти"),
            ("risk_coordinators","Ризик координатори"),
            ("settings",         "Налаштування"),
        ]

        self._nav_buttons: dict[PageKey, tk.Button] = {}
        for i, (key, label) in enumerate(menu_items):
            btn = tk.Button(
                nav_frame, text=label, bg=COLORS["bg_sidebar"], fg=COLORS["text_muted"],
                activebackground=COLORS["bg_surface"], activeforeground=COLORS["text_primary"],
                relief="flat", bd=0, cursor="hand2", anchor="w", font=("Arial", 9),
                padx=24, pady=7,
                command=lambda k=key: self._on_nav_click(k),
            )
            btn.grid(row=i, column=0, sticky="ew", pady=1)
            self._nav_buttons[key] = btn

        tk.Label(sidebar, text=COPYRIGHT_TEXT, bg=COLORS["bg_sidebar"],
                 fg=COLORS["text_subtle"], font=("Arial", 8)).grid(
            row=2, column=0, padx=16, pady=12, sticky="w")

    def _build_content(self) -> None:
        self.content_frame = tk.Frame(self, bg=COLORS["bg_main"])
        self.content_frame.grid(row=1, column=1, sticky="nsew")
        self.content_frame.grid_rowconfigure(0, weight=1)
        self.content_frame.grid_columnconfigure(0, weight=1)

        self._on_nav_click("risk_register")   # стартова сторінка

    def _on_nav_click(self, page_key: PageKey) -> None:
        if self._current_page == page_key:
            return
        self._current_page = page_key

        # Підсвітка кнопки
        for k, btn in self._nav_buttons.items():
            if k == page_key:
                btn.configure(bg=COLORS["bg_surface"], fg=COLORS["text_primary"])
            else:
                btn.configure(bg=COLORS["bg_sidebar"], fg=COLORS["text_muted"])

        # Видаляємо попередню сторінку
        for child in self.content_frame.winfo_children():
            child.grid_forget()

        if page_key not in self._pages:
            if page_key == "material_events":
                page = MaterialEventsPage(self.content_frame)
            elif page_key == "risk_register":
                page = RiskRegisterPage(self.content_frame)
            else:
                page = self._create_placeholder_page(page_key)
            self._pages[page_key] = page

        self._pages[page_key].grid(row=0, column=0, sticky="nsew")

    def _create_placeholder_page(self, page_key: PageKey) -> tk.Frame:
        page = tk.Frame(self.content_frame, bg=COLORS["bg_main"])
        tk.Label(page, text=page_key.replace("_", " ").title(),
                 bg=COLORS["bg_main"], fg=COLORS["text_primary"],
                 font=("Arial", 18, "bold")).pack(pady=60)
        tk.Label(page, text="Модуль в розробці...",
                 bg=COLORS["bg_main"], fg=COLORS["text_muted"],
                 font=("Arial", 11)).pack()
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
    current_user = "Оніщенко Андрій Сергійович"
    app = AtlasApp(user_full_name=current_user)
    app.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
