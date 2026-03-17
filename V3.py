import tkinter as tk
from tkinter import ttk, font, messagebox, filedialog
import json
import os
import re
import csv
from datetime import datetime
from collections import Counter

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

DATA_FILE = "essential_events.json"

COLORS = {
    "bg_dark":       "#0f1117",
    "bg_panel":      "#1a1d27",
    "bg_card":       "#22263a",
    "bg_input":      "#2a2f45",
    "accent":        "#4f8ef7",
    "accent_hover":  "#3a75e0",
    "accent2":       "#00d4aa",
    "danger":        "#ef4444",
    "warning":       "#f59e0b",
    "success":       "#22c55e",
    "text_primary":  "#e8eaf0",
    "text_muted":    "#7b82a0",
    "text_label":    "#a0a8c0",
    "border":        "#2e3350",
    "header_bg":     "#131624",
    "row_even":      "#1e2235",
    "row_odd":       "#1a1d2b",
    "row_select":    "#1d3461",
    "sidebar_bg":    "#10131e",
    "sidebar_hover": "#1e2235",
    "sidebar_active":"#1d3461",
    "navbar_bg":     "#0d1020",
}

RISK_COLORS = {
    "Операційний":          "#f59e0b",
    "Технічний":            "#4f8ef7",
    "Фінансовий":           "#ef4444",
    "Репутаційний":         "#a855f7",
    "Екологічний":          "#22c55e",
    "Надзвичайна ситуація": "#ff6b35",
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


def apply_dark_style(root):
    style = ttk.Style(root)
    style.theme_use("clam")
    C = COLORS

    style.configure(".",
        background=C["bg_dark"], foreground=C["text_primary"],
        fieldbackground=C["bg_input"], troughcolor=C["bg_panel"],
        bordercolor=C["border"], darkcolor=C["bg_panel"],
        lightcolor=C["bg_card"], insertcolor=C["text_primary"],
        selectbackground=C["row_select"], selectforeground=C["text_primary"],
        font=("Segoe UI", 9),
    )
    style.configure("TFrame", background=C["bg_dark"])
    style.configure("Card.TFrame", background=C["bg_card"])
    style.configure("Panel.TFrame", background=C["bg_panel"])
    style.configure("TLabel",
        background=C["bg_dark"], foreground=C["text_primary"], font=("Segoe UI", 9))
    style.configure("Section.TLabel",
        background=C["bg_dark"], foreground=C["accent"], font=("Segoe UI", 10, "bold"))
    style.configure("Muted.TLabel",
        background=C["bg_dark"], foreground=C["text_muted"], font=("Segoe UI", 8))
    style.configure("Net.TLabel",
        background=C["bg_dark"], foreground=C["accent2"], font=("Segoe UI", 11, "bold"))
    style.configure("TEntry",
        fieldbackground=C["bg_input"], foreground=C["text_primary"],
        insertcolor=C["text_primary"], bordercolor=C["border"],
        lightcolor=C["border"], darkcolor=C["border"], font=("Segoe UI", 9))
    style.map("TEntry",
        fieldbackground=[("focus", C["bg_card"])],
        bordercolor=[("focus", C["accent"])])
    style.configure("TCombobox",
        fieldbackground=C["bg_input"], background=C["bg_input"],
        foreground=C["text_primary"], arrowcolor=C["text_muted"],
        selectbackground=C["bg_input"], selectforeground=C["text_primary"],
        font=("Segoe UI", 9))
    style.map("TCombobox",
        fieldbackground=[("readonly", C["bg_input"])],
        selectbackground=[("readonly", C["bg_input"])])
    style.configure("TButton",
        background=C["bg_card"], foreground=C["text_primary"],
        bordercolor=C["border"], lightcolor=C["bg_card"], darkcolor=C["bg_card"],
        relief="flat", padding=(10, 5), font=("Segoe UI", 9))
    style.map("TButton",
        background=[("active", C["bg_input"]), ("pressed", C["border"])],
        relief=[("pressed", "flat")])
    style.configure("Accent.TButton",
        background=C["accent"], foreground="white",
        bordercolor=C["accent"], lightcolor=C["accent"], darkcolor=C["accent_hover"],
        font=("Segoe UI", 9, "bold"), padding=(12, 6))
    style.map("Accent.TButton",
        background=[("active", C["accent_hover"]), ("pressed", C["accent_hover"])])
    style.configure("Danger.TButton",
        background=C["danger"], foreground="white",
        bordercolor=C["danger"], font=("Segoe UI", 9), padding=(10, 5))
    style.map("Danger.TButton", background=[("active", "#dc2626")])
    style.configure("Success.TButton",
        background=C["success"], foreground="white",
        bordercolor=C["success"], font=("Segoe UI", 9), padding=(10, 5))
    style.map("Success.TButton", background=[("active", "#16a34a")])
    style.configure("TNotebook",
        background=C["bg_dark"], bordercolor=C["border"], tabmargins=[0,0,0,0])
    style.configure("TNotebook.Tab",
        background=C["bg_panel"], foreground=C["text_muted"],
        padding=[16, 8], font=("Segoe UI", 9), bordercolor=C["border"])
    style.map("TNotebook.Tab",
        background=[("selected", C["bg_card"]), ("active", C["bg_card"])],
        foreground=[("selected", C["accent"]), ("active", C["text_primary"])],
        expand=[("selected", [1,1,1,0])])
    style.configure("Treeview",
        background=C["row_odd"], foreground=C["text_primary"],
        fieldbackground=C["row_odd"], bordercolor=C["border"],
        font=("Segoe UI", 9), rowheight=26)
    style.configure("Treeview.Heading",
        background=C["bg_panel"], foreground=C["text_muted"],
        bordercolor=C["border"], relief="flat", font=("Segoe UI", 8, "bold"))
    style.map("Treeview",
        background=[("selected", C["row_select"])],
        foreground=[("selected", C["text_primary"])])
    style.map("Treeview.Heading",
        background=[("active", C["bg_card"])],
        foreground=[("active", C["text_primary"])])
    style.configure("Vertical.TScrollbar",
        background=C["bg_panel"], troughcolor=C["bg_dark"],
        arrowcolor=C["text_muted"], bordercolor=C["bg_dark"])
    style.configure("Horizontal.TScrollbar",
        background=C["bg_panel"], troughcolor=C["bg_dark"],
        arrowcolor=C["text_muted"], bordercolor=C["bg_dark"])
    style.configure("TSeparator", background=C["border"])
    style.configure("TPanedwindow", background=C["border"])
    style.configure("TProgressbar",
        troughcolor=C["bg_panel"], background=C["accent"])


def make_dark_text(parent, **kwargs):
    C = COLORS
    return tk.Text(parent,
        bg=C["bg_input"], fg=C["text_primary"],
        insertbackground=C["text_primary"],
        selectbackground=C["row_select"], selectforeground=C["text_primary"],
        relief="flat", bd=1, highlightthickness=1,
        highlightbackground=C["border"], highlightcolor=C["accent"],
        font=("Segoe UI", 9), **kwargs)


def add_placeholder(entry, text):
    entry.insert(0, text)
    entry.configure(foreground=COLORS["text_muted"])
    def on_in(_):
        if entry.get() == text:
            entry.delete(0, tk.END)
            entry.configure(foreground=COLORS["text_primary"])
    def on_out(_):
        if not entry.get():
            entry.insert(0, text)
            entry.configure(foreground=COLORS["text_muted"])
    entry.bind("<FocusIn>", on_in)
    entry.bind("<FocusOut>", on_out)


# ─── NAVBAR ───────────────────────────────────────────────────────────────────
class Navbar(tk.Frame):
    """
    Верхня панель навігації.
    Містить: логотип / назву, поточну сторінку, годинник, статус.
    """
    def __init__(self, parent, app_title="Суттєвi подiї", version="v3.0", **kw):
        C = COLORS
        super().__init__(parent, bg=C["navbar_bg"], height=48, **kw)
        self.pack_propagate(False)
        self.grid_propagate(False)

        # Ліва частина — логотип + версія
        left = tk.Frame(self, bg=C["navbar_bg"])
        left.pack(side="left", padx=0)

        # Акцентна смуга зліва
        tk.Frame(left, bg=C["accent"], width=4).pack(side="left", fill="y")

        brand = tk.Frame(left, bg=C["navbar_bg"])
        brand.pack(side="left", padx=(12, 0))

        tk.Label(brand, text=app_title,
                 bg=C["navbar_bg"], fg=C["text_primary"],
                 font=("Segoe UI", 11, "bold")).pack(anchor="w")
        tk.Label(brand, text=version,
                 bg=C["navbar_bg"], fg=C["text_muted"],
                 font=("Segoe UI", 7)).pack(anchor="w")

        # Центральна частина — назва поточної сторінки
        self._page_lbl = tk.Label(self, text="",
                                   bg=C["navbar_bg"], fg=C["accent"],
                                   font=("Segoe UI", 9, "bold"))
        self._page_lbl.pack(side="left", padx=24)

        # Права частина — статус + час
        right = tk.Frame(self, bg=C["navbar_bg"])
        right.pack(side="right", padx=14)

        self._time_lbl = tk.Label(right, text="",
                                   bg=C["navbar_bg"], fg=C["text_muted"],
                                   font=("Segoe UI", 8))
        self._time_lbl.pack(side="right", padx=(12, 0))

        tk.Frame(right, bg=C["border"], width=1).pack(side="right", fill="y", pady=10)

        self._status_lbl = tk.Label(right, text="Готово",
                                     bg=C["navbar_bg"], fg=C["text_muted"],
                                     font=("Segoe UI", 8))
        self._status_lbl.pack(side="right", padx=(0, 12))

        # Нижня межа navbar
        tk.Frame(self, bg=C["border"], height=1).pack(side="bottom", fill="x")

        self._tick()

    def _tick(self):
        self._time_lbl.configure(
            text=datetime.now().strftime("%d.%m.%Y  %H:%M:%S"))
        self.after(1000, self._tick)

    def set_page(self, name: str):
        self._page_lbl.configure(text=name)

    def set_status(self, msg: str, color: str | None = None):
        c = color or COLORS["text_muted"]
        self._status_lbl.configure(text=msg, fg=c)


# ─── SIDEBAR ─────────────────────────────────────────────────────────────────
class Sidebar(tk.Frame):
    """
    Бічна панель навігації.
    nav_items: список словників
        { "key": str, "label": str, "icon": str (ASCII/text symbol),
          "badge_color": str (optional) }
    on_select(key): callback при виборі пункту
    """
    SIDEBAR_W = 190

    def __init__(self, parent, nav_items: list, on_select, **kw):
        C = COLORS
        super().__init__(parent, bg=C["sidebar_bg"], width=self.SIDEBAR_W, **kw)
        self.pack_propagate(False)
        self.grid_propagate(False)

        self._on_select = on_select
        self._buttons   = {}
        self._active    = None
        self._nav_items = nav_items
        self._C         = C

        self._build(nav_items)

    def _build(self, nav_items):
        C = self._C

        # Логотип / заголовок бокової панелі
        header = tk.Frame(self, bg=C["sidebar_bg"])
        header.pack(fill="x", pady=(12, 4))
        tk.Label(header, text="НАВIГАЦIЯ",
                 bg=C["sidebar_bg"], fg=C["text_muted"],
                 font=("Segoe UI", 7, "bold")).pack(padx=16, anchor="w")

        # Роздільник
        tk.Frame(self, bg=C["border"], height=1).pack(fill="x", padx=12, pady=(0, 8))

        # Пункти навігації
        self._nav_frame = tk.Frame(self, bg=C["sidebar_bg"])
        self._nav_frame.pack(fill="both", expand=True)

        for item in nav_items:
            self._add_item(item)

        # Нижня частина сайдбару
        tk.Frame(self, bg=C["border"], height=1).pack(fill="x", padx=12, pady=8)
        bottom = tk.Frame(self, bg=C["sidebar_bg"])
        bottom.pack(fill="x", padx=12, pady=(0, 12))
        tk.Label(bottom, text="Реєстр суттєвих подiй",
                 bg=C["sidebar_bg"], fg=C["text_muted"],
                 font=("Segoe UI", 7), wraplength=160, justify="left").pack(anchor="w")

        # Права межа сайдбару
        tk.Frame(self, bg=C["border"], width=1).pack(side="right", fill="y")

    def _add_item(self, item):
        C = self._C
        key   = item["key"]
        label = item["label"]
        icon  = item.get("icon", ">")
        bc    = item.get("badge_color", None)

        btn_frame = tk.Frame(self._nav_frame, bg=C["sidebar_bg"], cursor="hand2")
        btn_frame.pack(fill="x", padx=8, pady=2)

        # Акцентна смуга (видима при активному стані)
        accent_bar = tk.Frame(btn_frame, bg=C["sidebar_bg"], width=3)
        accent_bar.pack(side="left", fill="y")

        inner = tk.Frame(btn_frame, bg=C["sidebar_bg"], padx=10, pady=9)
        inner.pack(side="left", fill="both", expand=True)
        inner.columnconfigure(1, weight=1)

        icon_lbl = tk.Label(inner, text=icon,
                             bg=C["sidebar_bg"], fg=C["text_muted"],
                             font=("Segoe UI", 10, "bold"), width=2)
        icon_lbl.grid(row=0, column=0, sticky="w")

        text_lbl = tk.Label(inner, text=label,
                             bg=C["sidebar_bg"], fg=C["text_muted"],
                             font=("Segoe UI", 9), anchor="w")
        text_lbl.grid(row=0, column=1, sticky="ew", padx=(6, 0))

        if bc:
            dot = tk.Frame(inner, bg=bc, width=6, height=6)
            dot.grid(row=0, column=2, padx=(4, 0))

        widgets = [btn_frame, inner, icon_lbl, text_lbl]

        def on_enter(_, wl=widgets, ab=accent_bar, k=key):
            if k != self._active:
                for w in wl:
                    w.configure(bg=C["sidebar_hover"])
                ab.configure(bg=C["accent"], width=3)

        def on_leave(_, wl=widgets, ab=accent_bar, k=key):
            if k != self._active:
                for w in wl:
                    w.configure(bg=C["sidebar_bg"])
                ab.configure(bg=C["sidebar_bg"])

        def on_click(_, k=key):
            self.select(k)

        for w in widgets:
            w.bind("<Enter>", on_enter)
            w.bind("<Leave>", on_leave)
            w.bind("<Button-1>", on_click)

        self._buttons[key] = {
            "frame":      btn_frame,
            "inner":      inner,
            "icon_lbl":   icon_lbl,
            "text_lbl":   text_lbl,
            "accent_bar": accent_bar,
            "all":        widgets,
        }

    def select(self, key: str):
        C = self._C

        # Скинути попередній активний
        if self._active and self._active in self._buttons:
            prev = self._buttons[self._active]
            for w in prev["all"]:
                w.configure(bg=C["sidebar_bg"])
            prev["accent_bar"].configure(bg=C["sidebar_bg"])
            prev["icon_lbl"].configure(fg=C["text_muted"])
            prev["text_lbl"].configure(fg=C["text_muted"], font=("Segoe UI", 9))

        self._active = key

        if key in self._buttons:
            cur = self._buttons[key]
            for w in cur["all"]:
                w.configure(bg=C["sidebar_active"])
            cur["accent_bar"].configure(bg=C["accent"])
            cur["icon_lbl"].configure(fg=C["accent"],  bg=C["sidebar_active"])
            cur["text_lbl"].configure(fg=C["text_primary"],
                                      bg=C["sidebar_active"],
                                      font=("Segoe UI", 9, "bold"))
            cur["inner"].configure(bg=C["sidebar_active"])

        self._on_select(key)

    def add_module(self, item: dict):
        """Динамічно додати новий модуль до сайдбару."""
        self._nav_items.append(item)
        self._add_item(item)


# ─── КОНТЕЙНЕР СТОРІНОК ───────────────────────────────────────────────────────
class PageContainer(tk.Frame):
    """
    Тримає всі фрейми-сторінки та показує лише одну за раз.
    """
    def __init__(self, parent, **kw):
        C = COLORS
        super().__init__(parent, bg=C["bg_dark"], **kw)
        self._pages: dict[str, tk.Frame] = {}
        self._current: str | None = None

    def register(self, key: str, frame: tk.Frame):
        """Зареєструвати фрейм під ключем."""
        self._pages[key] = frame
        frame.place(in_=self, x=0, y=0, relwidth=1, relheight=1)
        frame.lower()

    def show(self, key: str):
        """Показати сторінку за ключем."""
        if self._current == key:
            return
        if self._current and self._current in self._pages:
            self._pages[self._current].lower()
        if key in self._pages:
            self._pages[key].lift()
            self._current = key


# ─── ВІКНО ПЕРЕГЛЯДУ / РЕДАГУВАННЯ ЗАПИСУ ────────────────────────────────────
class EventDetailWindow:
    def __init__(self, parent_root, record, all_records,
                 save_callback, delete_callback, toast_callback):
        self.parent_root    = parent_root
        self.record         = list(record)
        self.all_records    = all_records
        self.save_callback  = save_callback
        self.delete_callback= delete_callback
        self.toast_callback = toast_callback
        self.is_edit_mode   = False
        self._build_window()

    def _build_window(self):
        C = COLORS
        self.win = tk.Toplevel(self.parent_root)
        self.win.title(f"Подiя #{self.record[0]}  —  {self.record[1]}")
        self.win.geometry("780x700")
        self.win.minsize(640, 500)
        self.win.configure(bg=C["bg_dark"])
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

        header = tk.Frame(self.win, bg=C["header_bg"], height=60)
        header.grid(row=0, column=0, sticky="ew")
        header.columnconfigure(1, weight=1)
        header.grid_propagate(False)

        risk_color = RISK_COLORS.get(self.record[4], C["accent"])
        tk.Frame(header, bg=risk_color, width=5).grid(row=0, column=0, sticky="ns")

        title_frame = tk.Frame(header, bg=C["header_bg"])
        title_frame.grid(row=0, column=1, sticky="ew", padx=16, pady=8)

        self.lbl_title = tk.Label(title_frame,
            text=f"Запис #{self.record[0]}",
            bg=C["header_bg"], fg=C["accent"], font=("Segoe UI", 12, "bold"))
        self.lbl_title.pack(anchor="w")

        self.lbl_subtitle = tk.Label(title_frame,
            text=self.record[1],
            bg=C["header_bg"], fg=C["text_muted"], font=("Segoe UI", 9))
        self.lbl_subtitle.pack(anchor="w")

        status_val = self.record[10] if len(self.record) > 10 else "—"
        status_colors = {
            "Вiдкрито":  C["danger"],
            "В обробцi": C["warning"],
            "Закрито":   C["text_muted"],
            "Вирiшено":  C["success"],
        }
        sc = status_colors.get(status_val, C["text_muted"])
        self.lbl_status_badge = tk.Label(header,
            text=f"  {status_val}  ",
            bg=sc, fg="white", font=("Segoe UI", 8, "bold"), pady=3)
        self.lbl_status_badge.grid(row=0, column=2, padx=16, pady=18)

        canvas = tk.Canvas(self.win, bg=C["bg_dark"], highlightthickness=0)
        sb = ttk.Scrollbar(self.win, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=sb.set)
        canvas.grid(row=1, column=0, sticky="nsew")
        sb.grid(row=1, column=1, sticky="ns")

        self.content = tk.Frame(canvas, bg=C["bg_dark"])
        self._cw = canvas.create_window((0, 0), window=self.content, anchor="nw")

        def on_conf(e):
            canvas.configure(scrollregion=canvas.bbox("all"))
            canvas.itemconfig(self._cw, width=canvas.winfo_width())

        self.content.bind("<Configure>", on_conf)
        canvas.bind("<Configure>", on_conf)
        canvas.bind_all("<MouseWheel>",
            lambda e: canvas.yview_scroll(int(-1 * e.delta / 120), "units"))

        self.content.columnconfigure(0, weight=1)
        self.content.columnconfigure(1, weight=1)
        self._build_view_content()

        btn_bar = tk.Frame(self.win, bg=C["bg_panel"], height=52)
        btn_bar.grid(row=2, column=0, columnspan=2, sticky="ew")
        btn_bar.grid_propagate(False)
        btn_bar.columnconfigure(0, weight=1)

        left_btns = tk.Frame(btn_bar, bg=C["bg_panel"])
        left_btns.pack(side="left", padx=12, pady=10)

        self.btn_edit = tk.Button(left_btns, text="Редагувати",
            bg=C["warning"], fg="white",
            activebackground="#d97706", activeforeground="white",
            relief="flat", bd=0, cursor="hand2",
            font=("Segoe UI", 9, "bold"), padx=16, pady=6,
            command=self._toggle_edit_mode)
        self.btn_edit.pack(side="left", padx=(0, 8))

        self.btn_save = tk.Button(left_btns, text="Зберегти змiни",
            bg=C["success"], fg="white",
            activebackground="#16a34a", activeforeground="white",
            relief="flat", bd=0, cursor="hand2",
            font=("Segoe UI", 9, "bold"), padx=16, pady=6,
            command=self._save_changes)
        self.btn_save.pack(side="left", padx=(0, 8))
        self.btn_save.pack_forget()

        self.btn_cancel_edit = tk.Button(left_btns, text="Скасувати",
            bg=C["bg_card"], fg=C["text_muted"],
            activebackground=C["bg_input"], activeforeground=C["text_primary"],
            relief="flat", bd=0, cursor="hand2",
            font=("Segoe UI", 9), padx=14, pady=6,
            command=self._cancel_edit)
        self.btn_cancel_edit.pack(side="left")
        self.btn_cancel_edit.pack_forget()

        right_btns = tk.Frame(btn_bar, bg=C["bg_panel"])
        right_btns.pack(side="right", padx=12, pady=10)

        tk.Button(right_btns, text="Видалити",
            bg=C["danger"], fg="white",
            activebackground="#dc2626", activeforeground="white",
            relief="flat", bd=0, cursor="hand2",
            font=("Segoe UI", 9, "bold"), padx=16, pady=6,
            command=self._delete_record).pack(side="right", padx=(8, 0))

        tk.Button(right_btns, text="Закрити",
            bg=C["bg_card"], fg=C["text_primary"],
            activebackground=C["bg_input"], activeforeground=C["text_primary"],
            relief="flat", bd=0, cursor="hand2",
            font=("Segoe UI", 9), padx=14, pady=6,
            command=self.win.destroy).pack(side="right")

    def _section_label(self, parent, text, row, col=0, colspan=2):
        C = COLORS
        f = tk.Frame(parent, bg=C["bg_dark"])
        f.grid(row=row, column=col, columnspan=colspan, sticky="ew",
               padx=16, pady=(14, 4))
        tk.Frame(f, bg=C["accent"], width=3, height=16).pack(side="left")
        tk.Label(f, text=text, bg=C["bg_dark"], fg=C["accent"],
                 font=("Segoe UI", 9, "bold")).pack(side="left", padx=8)

    def _info_row(self, parent, label, value, row, col=0):
        C = COLORS
        cell = tk.Frame(parent, bg=C["bg_card"], padx=12, pady=8)
        cell.grid(row=row, column=col, sticky="nsew",
                  padx=(16 if col == 0 else 6, 6 if col == 0 else 16), pady=3)
        cell.columnconfigure(0, weight=1)
        tk.Label(cell, text=label, bg=C["bg_card"], fg=C["text_muted"],
                 font=("Segoe UI", 7, "bold")).grid(row=0, column=0, sticky="w")
        tk.Label(cell, text=value or "—", bg=C["bg_card"], fg=C["text_primary"],
                 font=("Segoe UI", 9), wraplength=280,
                 justify="left").grid(row=1, column=0, sticky="w", pady=(2, 0))

    def _text_block(self, parent, label, value, row, col=0, colspan=1):
        C = COLORS
        cell = tk.Frame(parent, bg=C["bg_card"], padx=12, pady=8)
        padx_l = 16 if col == 0 else 6
        padx_r = 6  if (col == 0 and colspan == 1) else 16
        cell.grid(row=row, column=col, columnspan=colspan, sticky="nsew",
                  padx=(padx_l, padx_r), pady=3)
        cell.columnconfigure(0, weight=1)
        tk.Label(cell, text=label, bg=C["bg_card"], fg=C["text_muted"],
                 font=("Segoe UI", 7, "bold")).grid(row=0, column=0, sticky="w")
        t = make_dark_text(cell, height=3, wrap="word", state="normal")
        t.insert("1.0", value or "—")
        t.configure(state="disabled")
        t.grid(row=1, column=0, sticky="ew", pady=(4, 0))

    def _build_view_content(self):
        C = COLORS
        rec = self.record
        for w in self.content.winfo_children():
            w.destroy()
        r   = rec
        row = 0

        self._section_label(self.content, "Iнформацiя про пiдприємство", row)
        row += 1
        self._info_row(self.content, "Пiдприємство", r[1] if len(r) > 1 else "—", row, 0)
        priority_val = r[9] if len(r) > 9 else "—"
        priority_colors = {
            "Критичний": COLORS["danger"], "Високий": COLORS["warning"],
            "Середнiй":  COLORS["accent"], "Низький": COLORS["success"],
        }
        p_color = priority_colors.get(priority_val, COLORS["text_primary"])
        cell_p  = tk.Frame(self.content, bg=C["bg_card"], padx=12, pady=8)
        cell_p.grid(row=row, column=1, sticky="nsew", padx=(6, 16), pady=3)
        cell_p.columnconfigure(0, weight=1)
        tk.Label(cell_p, text="Прiоритет", bg=C["bg_card"], fg=C["text_muted"],
                 font=("Segoe UI", 7, "bold")).grid(row=0, column=0, sticky="w")
        tk.Label(cell_p, text=priority_val, bg=C["bg_card"], fg=p_color,
                 font=("Segoe UI", 10, "bold")).grid(row=1, column=0, sticky="w", pady=(2, 0))
        row += 1

        self._section_label(self.content, "Опис подiї / ризику", row)
        row += 1
        self._info_row(self.content, "Назва подiї", r[2] if len(r) > 2 else "—", row, 0)
        risk_val = r[4] if len(r) > 4 else "—"
        rc       = RISK_COLORS.get(risk_val, C["text_primary"])
        cell_r   = tk.Frame(self.content, bg=C["bg_card"], padx=12, pady=8)
        cell_r.grid(row=row, column=1, sticky="nsew", padx=(6, 16), pady=3)
        cell_r.columnconfigure(0, weight=1)
        tk.Label(cell_r, text="Тип ризику", bg=C["bg_card"], fg=C["text_muted"],
                 font=("Segoe UI", 7, "bold")).grid(row=0, column=0, sticky="w")
        tk.Label(cell_r, text=risk_val, bg=C["bg_card"], fg=rc,
                 font=("Segoe UI", 10, "bold")).grid(row=1, column=0, sticky="w", pady=(2, 0))
        row += 1

        self._info_row(self.content, "Статус", r[10] if len(r) > 10 else "—", row, 0)
        self._info_row(self.content, "Задiянi пiдроздiли / особи", r[3] if len(r) > 3 else "—", row, 1)
        row += 1

        self._section_label(self.content, "Дати", row)
        row += 1
        self._info_row(self.content, "Дата подiї",     r[5] if len(r) > 5 else "—", row, 0)
        self._info_row(self.content, "Дата виявлення", r[8] if len(r) > 8 else "—", row, 1)
        row += 1

        self._section_label(self.content, "Деталi подiї", row)
        row += 1
        self._text_block(self.content, "Детальний опис подiї",
                         r[6] if len(r) > 6 else "—", row, 0, colspan=2)
        row += 1
        self._text_block(self.content, "Вжитi заходи",
                         r[7] if len(r) > 7 else "—", row, 0, colspan=2)
        row += 1
        tk.Frame(self.content, bg=C["bg_dark"], height=16).grid(row=row, column=0, columnspan=2)

    def _build_edit_content(self):
        C = COLORS
        rec = self.record
        for w in self.content.winfo_children():
            w.destroy()
        self.content.columnconfigure(0, weight=1)

        def section(txt, row_idx):
            f = tk.Frame(self.content, bg=C["bg_dark"])
            f.grid(row=row_idx, column=0, sticky="ew", padx=16, pady=(14, 4))
            tk.Frame(f, bg=C["warning"], width=3, height=16).pack(side="left")
            tk.Label(f, text=txt, bg=C["bg_dark"], fg=C["warning"],
                     font=("Segoe UI", 9, "bold")).pack(side="left", padx=8)
            return row_idx + 1

        def lbl(text, row_idx):
            tk.Label(self.content, text=text, bg=C["bg_dark"], fg=C["text_label"],
                     font=("Segoe UI", 8)).grid(
                row=row_idx, column=0, sticky="w", padx=16, pady=(6, 0))

        def make_entry(**kw):
            return tk.Entry(self.content,
                bg=C["bg_input"], fg=C["text_primary"],
                insertbackground=C["text_primary"], relief="flat", bd=4,
                highlightthickness=1, highlightbackground=C["border"],
                highlightcolor=C["warning"], font=("Segoe UI", 9), **kw)

        def make_combo(values=None, **kw):
            return ttk.Combobox(self.content, values=values or [],
                                state="readonly", font=("Segoe UI", 9), **kw)

        row = 0
        row = section("Пiдприємство та подiя", row)
        lbl("Скорочена назва пiдприємства:", row); row += 1
        self.e_entity = make_entry()
        self.e_entity.grid(row=row, column=0, sticky="ew", padx=16, pady=(2, 0))
        self.e_entity.insert(0, rec[1] if len(rec) > 1 else "")
        row += 1

        lbl("Назва подiї:", row); row += 1
        self.e_event = make_combo(values=EVENT_TYPES)
        self.e_event.grid(row=row, column=0, sticky="ew", padx=16, pady=(2, 0))
        self.e_event.set(rec[2] if len(rec) > 2 else "")
        row += 1

        lbl("Тип ризику:", row); row += 1
        self.e_risk = make_combo(values=RISK_TYPES)
        self.e_risk.grid(row=row, column=0, sticky="ew", padx=16, pady=(2, 0))
        self.e_risk.set(rec[4] if (len(rec) > 4 and rec[4] != "—") else "")
        row += 1

        lbl("Задiянi пiдроздiли / особи:", row); row += 1
        self.e_involved = make_dark_text(self.content, height=2, wrap="word")
        self.e_involved.grid(row=row, column=0, sticky="ew", padx=16, pady=(2, 0))
        if len(rec) > 3 and rec[3] and rec[3] != "—":
            self.e_involved.insert("1.0", rec[3])
        row += 1

        row = section("Дати", row)
        date_frame = tk.Frame(self.content, bg=C["bg_dark"])
        date_frame.grid(row=row, column=0, sticky="w", padx=16, pady=4)
        row += 1

        for col_i, (lbl_t, attr, val_idx) in enumerate([
            ("Дата виявлення:", "e_detect",    8),
            ("Дата подiї:",     "e_event_date", 5),
        ]):
            tk.Label(date_frame, text=lbl_t, bg=C["bg_dark"],
                     fg=C["text_label"], font=("Segoe UI", 8)).grid(
                row=0, column=col_i, padx=(0 if col_i == 0 else 20, 0), sticky="w")
            e = tk.Entry(date_frame, bg=C["bg_input"], fg=C["text_primary"],
                insertbackground=C["text_primary"], relief="flat", bd=4,
                highlightthickness=1, highlightbackground=C["border"],
                highlightcolor=C["warning"], font=("Segoe UI", 9), width=14)
            e.grid(row=1, column=col_i, padx=(0 if col_i == 0 else 20, 0), pady=2)
            val = rec[val_idx] if len(rec) > val_idx else ""
            if val and val != "—":
                e.insert(0, val)
                e.configure(fg=C["text_primary"])
            else:
                add_placeholder(e, "дд.мм.рррр")
            setattr(self, attr, e)

        row = section("Прiоритет та статус", row)
        pf = tk.Frame(self.content, bg=C["bg_dark"])
        pf.grid(row=row, column=0, sticky="w", padx=16, pady=4)
        row += 1

        tk.Label(pf, text="Прiоритет:", bg=C["bg_dark"],
                 fg=C["text_label"], font=("Segoe UI", 8)).grid(row=0, column=0, sticky="w")
        self.e_priority = ttk.Combobox(pf,
            values=["Критичний", "Високий", "Середнiй", "Низький"],
            state="readonly", font=("Segoe UI", 9), width=14)
        self.e_priority.grid(row=1, column=0, pady=(2, 0), padx=(0, 20))
        self.e_priority.set(rec[9] if len(rec) > 9 else "Середнiй")

        tk.Label(pf, text="Статус:", bg=C["bg_dark"],
                 fg=C["text_label"], font=("Segoe UI", 8)).grid(row=0, column=1, sticky="w")
        self.e_status = ttk.Combobox(pf,
            values=["Вiдкрито", "В обробцi", "Закрито", "Вирiшено"],
            state="readonly", font=("Segoe UI", 9), width=14)
        self.e_status.grid(row=1, column=1, pady=(2, 0))
        self.e_status.set(rec[10] if len(rec) > 10 else "Вiдкрито")

        row = section("Деталi подiї", row)
        lbl("Детальний опис подiї:", row); row += 1
        self.e_description = make_dark_text(self.content, height=4, wrap="word")
        self.e_description.grid(row=row, column=0, sticky="ew", padx=16, pady=(2, 0))
        if len(rec) > 6 and rec[6] and rec[6] != "—":
            self.e_description.insert("1.0", rec[6])
        row += 1

        lbl("Вжитi заходи:", row); row += 1
        self.e_measures = make_dark_text(self.content, height=3, wrap="word")
        self.e_measures.grid(row=row, column=0, sticky="ew", padx=16, pady=(2, 0))
        if len(rec) > 7 and rec[7] and rec[7] != "—":
            self.e_measures.insert("1.0", rec[7])
        row += 1
        tk.Frame(self.content, bg=C["bg_dark"], height=20).grid(row=row, column=0)

    def _toggle_edit_mode(self):
        self.is_edit_mode = True
        self._build_edit_content()
        self.btn_edit.pack_forget()
        self.btn_save.pack(side="left", padx=(0, 8))
        self.btn_cancel_edit.pack(side="left")
        self.lbl_title.configure(
            text=f"Редагування запису #{self.record[0]}", fg=COLORS["warning"])

    def _cancel_edit(self):
        self.is_edit_mode = False
        self._build_view_content()
        self.btn_save.pack_forget()
        self.btn_cancel_edit.pack_forget()
        self.btn_edit.pack(side="left", padx=(0, 8))
        self.lbl_title.configure(
            text=f"Запис #{self.record[0]}", fg=COLORS["accent"])

    def _save_changes(self):
        C = COLORS
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
        old_id = self.record[0]
        new_record = (
            self.record[0], entity, event,
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
        self.save_callback(old_id, new_record)
        self.lbl_subtitle.configure(text=entity)
        status_val = new_record[10]
        status_colors = {
            "Вiдкрито":  C["danger"],  "В обробцi": C["warning"],
            "Закрито":   C["text_muted"], "Вирiшено": C["success"],
        }
        sc = status_colors.get(status_val, C["text_muted"])
        self.lbl_status_badge.configure(text=f"  {status_val}  ", bg=sc)
        self._cancel_edit()
        self.toast_callback("Запис оновлено")

    def _delete_record(self):
        idx_str = self.record[0]
        if not messagebox.askyesno("Пiдтвердження",
                f"Видалити запис #{idx_str}?\nЦю дiю не можна скасувати.",
                parent=self.win):
            return
        self.delete_callback(idx_str)
        self.toast_callback("Запис видалено")
        self.win.destroy()


# ─── ВКЛАДКА: РЕЄСТР ─────────────────────────────────────────────────────────
class RegistryTab:
    def __init__(self, parent, on_data_change=None):
        self.parent        = parent
        self.on_data_change= on_data_change
        self.frame         = ttk.Frame(parent)
        self.all_records   = []
        self._build_ui()
        self._load_data()

    def _build_ui(self):
        C = COLORS
        self.frame.rowconfigure(1, weight=1)
        self.frame.columnconfigure(0, weight=1)

        header = tk.Frame(self.frame, bg=C["header_bg"], height=56)
        header.grid(row=0, column=0, sticky="ew")
        header.columnconfigure(1, weight=1)
        header.grid_propagate(False)

        tk.Label(header, text="РЕЄСТР СУТТЄВИХ ПОДIЙ",
                 bg=C["header_bg"], fg=C["accent"],
                 font=("Segoe UI", 13, "bold")).grid(row=0, column=0, padx=20, pady=14, sticky="w")

        search_frame = tk.Frame(header, bg=C["header_bg"])
        search_frame.grid(row=0, column=1, sticky="e", padx=20)

        tk.Label(search_frame, text="Пошук:", bg=C["header_bg"],
                 fg=C["text_muted"], font=("Segoe UI", 8)).pack(side="left", padx=(0, 6))

        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *_: self._apply_filter())

        ent_search = tk.Entry(search_frame, textvariable=self.search_var,
            bg=C["bg_input"], fg=C["text_primary"],
            insertbackground=C["text_primary"], relief="flat", bd=4,
            font=("Segoe UI", 9), width=34)
        ent_search.pack(side="left", padx=(0, 8), ipady=3)

        tk.Button(search_frame, text="Скинути",
                  bg=C["bg_card"], fg=C["text_muted"],
                  relief="flat", bd=0, font=("Segoe UI", 8), cursor="hand2",
                  command=self._reset_filter).pack(side="left")

        paned = ttk.PanedWindow(self.frame, orient="horizontal")
        paned.grid(row=1, column=0, sticky="nsew")

        left_wrap  = ttk.Frame(paned)
        right_wrap = ttk.Frame(paned)
        paned.add(left_wrap,  weight=4)
        paned.add(right_wrap, weight=7)

        self._build_form(left_wrap)
        self._build_table(right_wrap)

    def _build_form(self, container):
        C = COLORS
        container.columnconfigure(0, weight=1)
        container.rowconfigure(0, weight=1)

        canvas   = tk.Canvas(container, bg=C["bg_dark"], highlightthickness=0)
        scrollbar= ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")
        container.columnconfigure(0, weight=1)
        container.rowconfigure(0, weight=1)

        form = tk.Frame(canvas, bg=C["bg_dark"])
        fw   = canvas.create_window((0, 0), window=form, anchor="nw")

        def on_configure(e):
            canvas.configure(scrollregion=canvas.bbox("all"))
            canvas.itemconfig(fw, width=canvas.winfo_width())
        form.bind("<Configure>", on_configure)
        canvas.bind("<Configure>", on_configure)
        canvas.bind_all("<MouseWheel>",
            lambda e: canvas.yview_scroll(int(-1 * (e.delta / 120)), "units"))

        form.columnconfigure(0, weight=1)

        def section(txt, row_idx):
            f = tk.Frame(form, bg=C["bg_dark"])
            f.grid(row=row_idx, column=0, sticky="ew", padx=16, pady=(16, 4))
            tk.Frame(f, bg=C["accent"], width=3, height=18).pack(side="left")
            tk.Label(f, text=txt, bg=C["bg_dark"], fg=C["accent"],
                     font=("Segoe UI", 9, "bold")).pack(side="left", padx=8)
            return row_idx + 1

        def field(lbl_text, row_idx, widget_factory, **kw):
            tk.Label(form, text=lbl_text, bg=C["bg_dark"], fg=C["text_label"],
                     font=("Segoe UI", 8)).grid(row=row_idx, column=0, sticky="w",
                                                padx=16, pady=(4, 0))
            w = widget_factory(form, **kw)
            w.grid(row=row_idx + 1, column=0, sticky="ew", padx=16, pady=(2, 0))
            return w, row_idx + 2

        def make_entry(parent, **kw):
            return tk.Entry(parent,
                bg=C["bg_input"], fg=C["text_primary"],
                insertbackground=C["text_primary"], relief="flat", bd=4,
                highlightthickness=1, highlightbackground=C["border"],
                highlightcolor=C["accent"], font=("Segoe UI", 9), **kw)

        def make_combo(parent, values=None, **kw):
            return ttk.Combobox(parent, values=values or [],
                                state="readonly", font=("Segoe UI", 9), **kw)

        row = 0
        new_badge = tk.Frame(form, bg=C["bg_dark"])
        new_badge.grid(row=row, column=0, sticky="ew", padx=16, pady=(12, 0))
        tk.Label(new_badge, text="  + НОВИЙ ЗАПИС  ",
                 bg=C["accent"], fg="white",
                 font=("Segoe UI", 8, "bold"), pady=4).pack(anchor="w")
        row += 1

        row = section("Iнформацiя про пiдприємство та особу", row)
        self.ent_entity,   row = field("Скорочена назва пiдприємства:", row, make_entry)
        self.ent_position, row = field("Посада:", row, make_entry)
        self.ent_reporter, row = field("ПIБ особи, що звiтує:", row, make_entry)

        row = section("Опис подiї / ризику", row)
        self.cb_event, row = field("Назва подiї:", row, make_combo, values=EVENT_TYPES)
        self.cb_risk,  row = field("Тип ризику:", row, make_combo, values=RISK_TYPES)

        tk.Label(form, text="Задiянi пiдроздiли / особи:",
                 bg=C["bg_dark"], fg=C["text_label"],
                 font=("Segoe UI", 8)).grid(row=row, column=0, sticky="w", padx=16, pady=(4, 0))
        row += 1
        self.txt_involved = make_dark_text(form, height=2, wrap="word")
        self.txt_involved.grid(row=row, column=0, sticky="ew", padx=16, pady=(2, 0))
        row += 1

        row = section("Фiнансовий вплив (млн грн)", row)
        fin_frame = tk.Frame(form, bg=C["bg_dark"])
        fin_frame.grid(row=row, column=0, sticky="ew", padx=16, pady=4)
        fin_frame.columnconfigure((0, 1, 2, 3), weight=1)
        row += 1

        for col, title in enumerate(["Втрати", "Резерв", "Заплановані втрати", "Відшкодування"]):
            tk.Label(fin_frame, text=title, bg=C["bg_dark"],
                     fg=C["text_muted"], font=("Segoe UI", 7)).grid(
                row=0, column=col, sticky="w", padx=4)

        self.ent_loss    = make_entry(fin_frame, width=11)
        self.ent_reserve = make_entry(fin_frame, width=11)
        self.ent_planned = make_entry(fin_frame, width=11)
        self.ent_refund  = make_entry(fin_frame, width=11)
        for col, e in enumerate([self.ent_loss, self.ent_reserve,
                                  self.ent_planned, self.ent_refund]):
            e.grid(row=1, column=col, padx=4, pady=2, sticky="ew")

        net_frame = tk.Frame(form, bg=C["bg_dark"])
        net_frame.grid(row=row, column=0, sticky="ew", padx=16, pady=(2, 6))
        row += 1
        tk.Label(net_frame, text="Чистий вплив (млн грн):",
                 bg=C["bg_dark"], fg=C["text_muted"],
                 font=("Segoe UI", 8)).pack(side="left")
        self.lbl_net = tk.Label(net_frame, text="0.00",
                                 bg=C["bg_dark"], fg=C["accent2"],
                                 font=("Segoe UI", 13, "bold"))
        self.lbl_net.pack(side="left", padx=10)

        def update_net(*_):
            try:
                loss   = float(self.ent_loss.get().replace(",", ".") or 0)
                refund = float(self.ent_refund.get().replace(",", ".") or 0)
                net    = loss - refund
                col    = COLORS["danger"] if net > 0 else COLORS["success"]
                self.lbl_net.configure(text=f"{net:,.2f}", fg=col)
            except Exception:
                self.lbl_net.configure(text="—", fg=COLORS["text_muted"])

        for e in [self.ent_loss, self.ent_refund]:
            e.bind("<KeyRelease>", update_net)

        row = section("Деталi подiї", row)
        text_fields = [
            ("Вплив на iншi пiдприємства:", "txt_impact", 2),
            ("Нефiнансовий / якiсний вплив:", "txt_qualitative", 2),
            ("Детальний опис подiї:", "txt_description", 4),
            ("Вжитi заходи:", "txt_measures", 3),
        ]
        for lbl_text, attr, height in text_fields:
            tk.Label(form, text=lbl_text, bg=C["bg_dark"],
                     fg=C["text_label"], font=("Segoe UI", 8)).grid(
                row=row, column=0, sticky="w", padx=16, pady=(6, 0))
            row += 1
            t = make_dark_text(form, height=height, wrap="word")
            t.grid(row=row, column=0, sticky="ew", padx=16, pady=(2, 0))
            setattr(self, attr, t)
            row += 1

        row = section("Дати", row)
        date_frame = tk.Frame(form, bg=C["bg_dark"])
        date_frame.grid(row=row, column=0, sticky="w", padx=16, pady=4)
        row += 1

        for col, (lbl2, attr) in enumerate([
            ("Дата виявлення:", "ent_detect"),
            ("Дата подiї:",     "ent_event_date"),
        ]):
            tk.Label(date_frame, text=lbl2, bg=C["bg_dark"],
                     fg=C["text_label"], font=("Segoe UI", 8)).grid(
                row=0, column=col, padx=(0 if col == 0 else 20, 0), sticky="w")
            e = make_entry(date_frame, width=14)
            e.grid(row=1, column=col, padx=(0 if col == 0 else 20, 0), pady=2)
            add_placeholder(e, "дд.мм.рррр")
            setattr(self, attr, e)

        tk.Label(form, text="Прiоритет:", bg=C["bg_dark"],
                 fg=C["text_label"], font=("Segoe UI", 8)).grid(
            row=row, column=0, sticky="w", padx=16, pady=(8, 0))
        row += 1
        self.cb_priority = make_combo(form, values=["Критичний", "Високий", "Середнiй", "Низький"])
        self.cb_priority.grid(row=row, column=0, sticky="w", padx=16, pady=(2, 0))
        row += 1

        tk.Label(form, text="Статус:", bg=C["bg_dark"],
                 fg=C["text_label"], font=("Segoe UI", 8)).grid(
            row=row, column=0, sticky="w", padx=16, pady=(8, 0))
        row += 1
        self.cb_status = make_combo(form, values=["Вiдкрито", "В обробцi", "Закрито", "Вирiшено"])
        self.cb_status.grid(row=row, column=0, sticky="w", padx=16, pady=(2, 0))
        row += 1

        btn_frame = tk.Frame(form, bg=C["bg_dark"])
        btn_frame.grid(row=row, column=0, sticky="ew", padx=16, pady=(16, 20))
        row += 1
        btn_frame.columnconfigure((0, 1), weight=1)

        tk.Button(btn_frame, text="Очистити",
                  bg=C["bg_card"], fg=C["text_muted"],
                  activebackground=C["bg_input"], activeforeground=C["text_primary"],
                  relief="flat", bd=0, cursor="hand2",
                  font=("Segoe UI", 9), padx=14, pady=6,
                  command=self._clear_form).grid(row=0, column=0, padx=4, sticky="ew")

        tk.Button(btn_frame, text="Додати запис",
                  bg=C["accent"], fg="white",
                  activebackground=C["accent_hover"], activeforeground="white",
                  relief="flat", bd=0, cursor="hand2",
                  font=("Segoe UI", 9, "bold"), padx=14, pady=6,
                  command=self._add_record).grid(row=0, column=1, padx=4, sticky="ew")

    def _build_table(self, container):
        C = COLORS
        container.rowconfigure(1, weight=1)
        container.columnconfigure(0, weight=1)

        toolbar = tk.Frame(container, bg=C["bg_panel"], height=42)
        toolbar.grid(row=0, column=0, sticky="ew")
        toolbar.grid_propagate(False)

        tk.Label(toolbar, text="Записи", bg=C["bg_panel"],
                 fg=C["text_muted"], font=("Segoe UI", 8)).pack(side="left", padx=12, pady=8)

        self.lbl_count = tk.Label(toolbar, text="0", bg=C["bg_panel"],
                                   fg=C["accent"], font=("Segoe UI", 8, "bold"))
        self.lbl_count.pack(side="left", pady=8)

        tk.Label(toolbar, text="  |  Ризик:", bg=C["bg_panel"],
                 fg=C["text_muted"], font=("Segoe UI", 8)).pack(side="left", pady=8)
        self.filter_risk = ttk.Combobox(toolbar, width=16, state="readonly",
                                        values=["Всi"] + RISK_TYPES, font=("Segoe UI", 8))
        self.filter_risk.set("Всi")
        self.filter_risk.pack(side="left", padx=6, pady=8)
        self.filter_risk.bind("<<ComboboxSelected>>", lambda _: self._apply_filter())

        tk.Label(toolbar, text="Статус:", bg=C["bg_panel"],
                 fg=C["text_muted"], font=("Segoe UI", 8)).pack(side="left", pady=8)
        self.filter_status = ttk.Combobox(toolbar, width=12, state="readonly",
            values=["Всi", "Вiдкрито", "В обробцi", "Закрито", "Вирiшено"],
            font=("Segoe UI", 8))
        self.filter_status.set("Всi")
        self.filter_status.pack(side="left", padx=6, pady=8)
        self.filter_status.bind("<<ComboboxSelected>>", lambda _: self._apply_filter())

        for txt, cmd, bg in [
            ("Переглянути", self._open_selected_detail, C["accent"]),
            ("Дублювати",   self._duplicate_record, C["bg_card"]),
            ("Видалити",    self._delete_selected, C["danger"]),
        ]:
            tk.Button(toolbar, text=txt, bg=bg,
                      fg="white" if bg != C["bg_card"] else C["text_primary"],
                      activebackground=bg, activeforeground="white",
                      relief="flat", bd=0, cursor="hand2",
                      font=("Segoe UI", 8), padx=10, pady=4,
                      command=cmd).pack(side="right", padx=4, pady=6)

        tree_frame = ttk.Frame(container)
        tree_frame.grid(row=1, column=0, sticky="nsew")
        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)

        cols = ("id", "entity", "event", "risk", "priority",
                "status", "date", "involved", "desc", "measures")
        self.tree = ttk.Treeview(tree_frame, columns=cols, show="headings", selectmode="browse")
        self.tree.grid(row=0, column=0, sticky="nsew")

        headers = {
            "id":       ("No", 46),
            "entity":   ("Пiдприємство", 155),
            "event":    ("Назва подiї", 185),
            "risk":     ("Тип ризику", 105),
            "priority": ("Прiоритет", 90),
            "status":   ("Статус", 90),
            "date":     ("Дата подiї", 90),
            "involved": ("Задiянi", 130),
            "desc":     ("Опис", 200),
            "measures": ("Заходи", 200),
        }
        for col, (txt, w) in headers.items():
            self.tree.heading(col, text=txt, command=lambda c=col: self._sort_tree(c))
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

        self.tree.bind("<<TreeviewSelect>>", self._on_select)
        self.tree.bind("<Double-1>", self._on_double_click)

        hint_frame = tk.Frame(container, bg=C["bg_panel"])
        hint_frame.grid(row=2, column=0, sticky="ew")
        tk.Label(hint_frame,
                 text="  Подвiйний клiк по рядку — переглянути / редагувати запис",
                 bg=C["bg_panel"], fg=C["text_muted"],
                 font=("Segoe UI", 7, "italic")).pack(side="left", padx=8, pady=5)

        detail_frame = tk.Frame(container, bg=C["bg_panel"])
        detail_frame.grid(row=3, column=0, sticky="ew")
        detail_frame.columnconfigure((0, 1), weight=1)

        for col, (lbl2, attr) in enumerate([
            ("Опис подiї",   "det_desc"),
            ("Вжитi заходи", "det_measures"),
        ]):
            sub = tk.Frame(detail_frame, bg=C["bg_panel"])
            sub.grid(row=0, column=col, sticky="nsew",
                     padx=(12 if col == 0 else 4, 4), pady=8)
            sub.columnconfigure(0, weight=1)
            tk.Label(sub, text=lbl2, bg=C["bg_panel"],
                     fg=C["text_muted"], font=("Segoe UI", 7, "bold")).grid(
                row=0, column=0, sticky="w")
            t = make_dark_text(sub, height=4, wrap="word", state="disabled")
            t.grid(row=1, column=0, sticky="ew", pady=(2, 0))
            setattr(self, attr, t)

        exp_bar = tk.Frame(container, bg=C["bg_dark"])
        exp_bar.grid(row=4, column=0, sticky="ew", padx=8, pady=6)

        tk.Button(exp_bar, text="Експорт CSV",
                  bg=C["bg_card"], fg=C["text_primary"],
                  activebackground=C["bg_input"], activeforeground=C["text_primary"],
                  relief="flat", bd=0, cursor="hand2",
                  font=("Segoe UI", 8), padx=12, pady=5,
                  command=self._export_csv).pack(side="left", padx=(0, 6))

        if pd:
            tk.Button(exp_bar, text="Експорт Excel",
                      bg=C["success"], fg="white",
                      activebackground="#16a34a", activeforeground="white",
                      relief="flat", bd=0, cursor="hand2",
                      font=("Segoe UI", 8), padx=12, pady=5,
                      command=self._export_excel).pack(side="left", padx=(0, 6))

        tk.Button(exp_bar, text="Iмпорт JSON",
                  bg=C["bg_card"], fg=C["text_primary"],
                  activebackground=C["bg_input"], activeforeground=C["text_primary"],
                  relief="flat", bd=0, cursor="hand2",
                  font=("Segoe UI", 8), padx=12, pady=5,
                  command=self._import_json).pack(side="left")

    def _on_double_click(self, event):
        sel = self.tree.selection()
        if not sel:
            return
        if self.tree.identify("region", event.x, event.y) != "cell":
            return
        self._open_selected_detail()

    def _open_selected_detail(self):
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
            parent_root     = self.frame.winfo_toplevel(),
            record          = rec,
            all_records     = self.all_records,
            save_callback   = lambda old_id, new_rec: self._on_detail_save(iid, old_id, new_rec),
            delete_callback = lambda idx_s: self._on_detail_delete(iid, idx_s),
            toast_callback  = self._show_toast,
        )

    def _find_record(self, idx_str):
        for r in self.all_records:
            if (str(r[0]) == idx_str or
                    str(r[0]).lstrip("0") == str(idx_str).lstrip("0")):
                return r
        return None

    def _on_detail_save(self, iid, old_id, new_record):
        for i, r in enumerate(self.all_records):
            if (str(r[0]) == str(old_id) or
                    str(r[0]).lstrip("0") == str(old_id).lstrip("0")):
                self.all_records[i] = new_record
                break
        try:
            self.tree.item(iid, values=(
                new_record[0], new_record[1], new_record[2], new_record[4],
                new_record[9], new_record[10], new_record[5], new_record[3],
                new_record[6], new_record[7]))
        except tk.TclError:
            pass
        self._recolor_rows()
        self._save_data()
        if self.on_data_change:
            self.on_data_change(self.all_records)

    def _on_detail_delete(self, iid, idx_str):
        try:
            self.tree.delete(iid)
        except tk.TclError:
            pass
        self.all_records = [r for r in self.all_records if str(r[0]) != str(idx_str)]
        self._recolor_rows()
        self._save_data()
        self._update_count()
        if self.on_data_change:
            self.on_data_change(self.all_records)

    def _sort_tree(self, col):
        data = [(self.tree.set(iid, col), iid) for iid in self.tree.get_children("")]
        try:
            data.sort(key=lambda x: float(x[0].replace(",", "")) if x[0] not in ("—", "") else 0)
        except ValueError:
            data.sort(key=lambda x: x[0].lower())
        for i, (_, iid) in enumerate(data):
            self.tree.move(iid, "", i)
        self._recolor_rows()

    def _recolor_rows(self):
        for i, iid in enumerate(self.tree.get_children()):
            risk     = self.tree.set(iid, "risk")
            base_tag = "even" if i % 2 == 0 else "odd"
            tags = [base_tag]
            if risk in RISK_COLORS:
                tags.append(f"risk_{risk}")
            self.tree.item(iid, tags=tags)

    def _load_data(self):
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

    def _save_data(self):
        try:
            with open(DATA_FILE, "w", encoding="utf-8") as f:
                json.dump(self.all_records, f, ensure_ascii=False, indent=2)
        except Exception as e:
            messagebox.showerror("Помилка збереження", str(e))

    def _insert_tree_row(self, data):
        if len(data) < 11:
            data = list(data) + ["—"] * (11 - len(data))
        iid = self.tree.insert("", tk.END, values=(
            data[0], data[1], data[2], data[4], data[9], data[10],
            data[5], data[3], data[6], data[7]))
        self._recolor_rows()
        return iid

    def _get_form_data(self):
        detect  = self.ent_detect.get().strip()
        event_d = self.ent_event_date.get().strip()
        for val, label in [(detect, "дати виявлення"), (event_d, "дати подiї")]:
            if val in ("дд.мм.рррр", ""):
                continue
            if not is_valid_date(val):
                messagebox.showwarning("Помилка",
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
            self.cb_status.get().strip() or "Вiдкрито",
        )

    def _clear_form(self, silent=False):
        for w in [self.ent_entity, self.ent_position, self.ent_reporter,
                  self.ent_loss, self.ent_reserve, self.ent_planned, self.ent_refund]:
            w.delete(0, tk.END)
            w.configure(fg=COLORS["text_primary"])
        for w in [self.cb_event, self.cb_risk, self.cb_priority, self.cb_status]:
            w.set("")
        for w in [self.txt_involved, self.txt_impact, self.txt_qualitative,
                  self.txt_description, self.txt_measures]:
            w.delete("1.0", tk.END)
        self.lbl_net.configure(text="0.00", fg=COLORS["accent2"])
        for e, ph in [(self.ent_detect, "дд.мм.рррр"), (self.ent_event_date, "дд.мм.рррр")]:
            e.delete(0, tk.END)
            add_placeholder(e, ph)

    def _add_record(self):
        data = self._get_form_data()
        if not data:
            return
        if not data[1] or not data[2]:
            messagebox.showwarning("Обов'язковi поля",
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

    def _delete_selected(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Видалення", "Оберiть запис для видалення")
            return
        iid     = sel[0]
        idx_str = self.tree.set(iid, "id")
        if not messagebox.askyesno("Пiдтвердження",
                f"Видалити запис #{idx_str}?\nЦю дiю не можна скасувати."):
            return
        self.tree.delete(iid)
        self.all_records = [r for r in self.all_records if str(r[0]) != idx_str]
        self._recolor_rows()
        self._save_data()
        self._update_count()
        if self.on_data_change:
            self.on_data_change(self.all_records)
        self._show_toast("Запис видалено")

    def _duplicate_record(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Дублювання", "Оберiть запис для дублювання")
            return
        idx_str = self.tree.set(sel[0], "id")
        rec     = self._find_record(idx_str)
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

    def _apply_filter(self):
        q      = self.search_var.get().strip().lower()
        risk   = self.filter_risk.get()
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

    def _reset_filter(self):
        self.search_var.set("")
        self.filter_risk.set("Всi")
        self.filter_status.set("Всi")
        self.tree.delete(*self.tree.get_children())
        for row in self.all_records:
            self._insert_tree_row(row)
        self._update_count()

    def _on_select(self, _=None):
        sel = self.tree.selection()
        if not sel:
            return
        desc     = self.tree.set(sel[0], "desc")
        measures = self.tree.set(sel[0], "measures")
        for widget, text in [(self.det_desc, desc), (self.det_measures, measures)]:
            widget.configure(state="normal")
            widget.delete("1.0", tk.END)
            widget.insert("1.0", text)
            widget.configure(state="disabled")

    def _update_count(self):
        self.lbl_count.configure(text=f" {len(self.tree.get_children())}")

    def _show_toast(self, msg):
        C = COLORS
        toast = tk.Toplevel(self.frame)
        toast.overrideredirect(True)
        toast.attributes("-topmost", True)
        toast.configure(bg=C["success"])
        tk.Label(toast, text=f"  {msg}  ", bg=C["success"], fg="white",
                 font=("Segoe UI", 9, "bold"), pady=6).pack()
        root = self.frame.winfo_toplevel()
        x = root.winfo_x() + root.winfo_width()  - 220
        y = root.winfo_y() + root.winfo_height() - 80
        toast.geometry(f"+{x}+{y}")
        toast.after(2000, toast.destroy)

    def _export_csv(self):
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
            headers = ["ID", "Пiдприємство", "Назва подiї", "Задiянi особи",
                       "Тип ризику", "Дата подiї", "Опис", "Заходи",
                       "Дата виявлення", "Прiоритет", "Статус"]
            with open(path, "w", newline="", encoding="utf-8-sig") as f:
                w = csv.writer(f)
                w.writerow(headers)
                for row in self.all_records:
                    w.writerow(row)
            self._show_toast("CSV збережено")
        except Exception as e:
            messagebox.showerror("Помилка", str(e))

    def _export_excel(self):
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
            headers = ["ID", "Пiдприємство", "Назва подiї", "Задiянi особи",
                       "Тип ризику", "Дата подiї", "Опис", "Заходи",
                       "Дата виявлення", "Прiоритет", "Статус"]
            data = [r if len(r) == 11 else list(r) + [""] * (11 - len(r))
                    for r in self.all_records]
            df = pd.DataFrame(data, columns=headers)
            with pd.ExcelWriter(path, engine="openpyxl") as writer:
                df.to_excel(writer, index=False, sheet_name="Реєстр")
                ws = writer.sheets["Реєстр"]
                for col_cells in ws.columns:
                    max_len = max(len(str(c.value or "")) for c in col_cells)
                    ws.column_dimensions[col_cells[0].column_letter].width = min(max_len + 4, 60)
            self._show_toast("Excel збережено")
        except Exception as e:
            messagebox.showerror("Помилка", str(e))

    def _import_json(self):
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

    def get_frame(self):
        return self.frame


# ─── ВКЛАДКА: АНАЛІТИКА ──────────────────────────────────────────────────────
class AnalyticsTab:
    def __init__(self, parent):
        self.frame   = ttk.Frame(parent)
        self.records = []
        self._build_ui()

    def _build_ui(self):
        C = COLORS
        self.frame.rowconfigure(1, weight=1)
        self.frame.columnconfigure(0, weight=1)

        header = tk.Frame(self.frame, bg=C["header_bg"], height=56)
        header.grid(row=0, column=0, sticky="ew")
        header.grid_propagate(False)
        tk.Label(header, text="АНАЛIТИКА ТА ЗВIТИ",
                 bg=C["header_bg"], fg=C["accent2"],
                 font=("Segoe UI", 13, "bold")).pack(side="left", padx=20, pady=14)
        tk.Button(header, text="Оновити",
                  bg=C["accent2"], fg=C["bg_dark"],
                  activebackground="#00b894", activeforeground=C["bg_dark"],
                  relief="flat", bd=0, cursor="hand2",
                  font=("Segoe UI", 9, "bold"), padx=12, pady=4,
                  command=self.refresh).pack(side="right", padx=20, pady=12)

        canvas = tk.Canvas(self.frame, bg=C["bg_dark"], highlightthickness=0)
        sb     = ttk.Scrollbar(self.frame, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=sb.set)
        canvas.grid(row=1, column=0, sticky="nsew")
        sb.grid(row=1, column=1, sticky="ns")

        self.content = tk.Frame(canvas, bg=C["bg_dark"])
        self._cw     = canvas.create_window((0, 0), window=self.content, anchor="nw")

        def on_conf(e):
            canvas.configure(scrollregion=canvas.bbox("all"))
            canvas.itemconfig(self._cw, width=canvas.winfo_width())

        self.content.bind("<Configure>", on_conf)
        canvas.bind("<Configure>", on_conf)
        canvas.bind_all("<MouseWheel>",
            lambda e: canvas.yview_scroll(int(-1 * e.delta / 120), "units"))

        self.content.columnconfigure(0, weight=1)
        self.canvas_ref = canvas
        self._build_stat_cards()
        self._build_charts_placeholder()

    def _build_stat_cards(self):
        C = COLORS
        cf = tk.Frame(self.content, bg=C["bg_dark"])
        cf.grid(row=0, column=0, sticky="ew", padx=16, pady=(16, 0))
        for i in range(4):
            cf.columnconfigure(i, weight=1)

        self.stat_cards = {}
        defs = [
            ("total",    "Всього записiв", "0", C["accent"]),
            ("open",     "Вiдкрито",       "0", C["danger"]),
            ("critical", "Критичних",      "0", C["warning"]),
            ("closed",   "Закрито",        "0", C["success"]),
        ]
        for col, (key, title, val, color) in enumerate(defs):
            card = tk.Frame(cf, bg=C["bg_card"], padx=20, pady=14)
            card.grid(row=0, column=col, padx=6, sticky="nsew")
            tk.Frame(card, bg=color, height=3).pack(fill="x")
            tk.Label(card, text=title, bg=C["bg_card"],
                     fg=C["text_muted"], font=("Segoe UI", 8)).pack(anchor="w", pady=(8, 2))
            lbl = tk.Label(card, text=val, bg=C["bg_card"],
                           fg=color, font=("Segoe UI", 26, "bold"))
            lbl.pack(anchor="w")
            self.stat_cards[key] = lbl

    def _build_charts_placeholder(self):
        C = COLORS
        if HAS_MPL:
            cr = tk.Frame(self.content, bg=C["bg_dark"])
            cr.grid(row=1, column=0, sticky="ew", padx=16, pady=16)
            cr.columnconfigure((0, 1), weight=1)

            self.fig_left = Figure(figsize=(5, 3.5), dpi=90, facecolor=C["bg_card"])
            self.ax_left  = self.fig_left.add_subplot(111)
            self._style_ax(self.ax_left)
            self.ax_left.set_title("Розподiл за типом ризику",
                                   color=C["text_muted"], fontsize=9)
            fl = tk.Frame(cr, bg=C["bg_card"], padx=8, pady=8)
            fl.grid(row=0, column=0, padx=(0, 8), sticky="nsew")
            self.canvas_left = FigureCanvasTkAgg(self.fig_left, master=fl)
            self.canvas_left.get_tk_widget().pack(fill="both", expand=True)

            self.fig_right = Figure(figsize=(5, 3.5), dpi=90, facecolor=C["bg_card"])
            self.ax_right  = self.fig_right.add_subplot(111)
            self._style_ax(self.ax_right)
            self.ax_right.set_title("Записи за статусом",
                                    color=C["text_muted"], fontsize=9)
            fr = tk.Frame(cr, bg=C["bg_card"], padx=8, pady=8)
            fr.grid(row=0, column=1, padx=(8, 0), sticky="nsew")
            self.canvas_right = FigureCanvasTkAgg(self.fig_right, master=fr)
            self.canvas_right.get_tk_widget().pack(fill="both", expand=True)

            self.fig_bottom = Figure(figsize=(10, 3), dpi=90, facecolor=C["bg_card"])
            self.ax_bottom  = self.fig_bottom.add_subplot(111)
            self._style_ax(self.ax_bottom)
            self.ax_bottom.set_title("Топ-5 пiдприємств за кiлькiстю подiй",
                                     color=C["text_muted"], fontsize=9)
            fb = tk.Frame(self.content, bg=C["bg_card"], padx=8, pady=8)
            fb.grid(row=2, column=0, padx=16, pady=(0, 16), sticky="ew")
            self.canvas_bottom = FigureCanvasTkAgg(self.fig_bottom, master=fb)
            self.canvas_bottom.get_tk_widget().pack(fill="both", expand=True)
        else:
            tk.Label(self.content,
                     text="Встановiть matplotlib для вiдображення графiкiв:\n  pip install matplotlib",
                     bg=COLORS["bg_dark"], fg=COLORS["text_muted"],
                     font=("Segoe UI", 10)).grid(row=1, column=0, pady=40)

        self._build_stats_table()

    def _style_ax(self, ax):
        C = COLORS
        ax.set_facecolor(C["bg_card"])
        ax.tick_params(colors=C["text_muted"], labelsize=8)
        for spine in ax.spines.values():
            spine.set_edgecolor(C["border"])

    def _build_stats_table(self):
        C = COLORS
        frame = tk.Frame(self.content, bg=C["bg_card"], padx=16, pady=12)
        frame.grid(row=3, column=0, padx=16, pady=(0, 16), sticky="ew")
        frame.columnconfigure(0, weight=1)
        tk.Label(frame, text="Деталiзована статистика за типом ризику",
                 bg=C["bg_card"], fg=C["text_muted"],
                 font=("Segoe UI", 9, "bold")).grid(row=0, column=0, sticky="w", pady=(0, 8))
        cols = ("risk", "count", "open", "closed")
        self.stats_tree = ttk.Treeview(frame, columns=cols, show="headings", height=7)
        for col, hdr, w in [
            ("risk",   "Тип ризику", 200),
            ("count",  "Всього",     80),
            ("open",   "Вiдкрито",   80),
            ("closed", "Закрито",    80),
        ]:
            self.stats_tree.heading(col, text=hdr)
            self.stats_tree.column(col, width=w, anchor="w")
        self.stats_tree.grid(row=1, column=0, sticky="ew")

    def update_data(self, records):
        self.records = records
        self.refresh()

    def refresh(self):
        C       = COLORS
        records = self.records

        total    = len(records)
        open_c   = sum(1 for r in records if len(r) > 10 and r[10] in ("Вiдкрито", "В обробцi"))
        critical = sum(1 for r in records if len(r) > 9  and r[9]  == "Критичний")
        closed   = sum(1 for r in records if len(r) > 10 and r[10] in ("Закрито", "Вирiшено"))

        self.stat_cards["total"].configure(text=str(total))
        self.stat_cards["open"].configure(text=str(open_c))
        self.stat_cards["critical"].configure(text=str(critical))
        self.stat_cards["closed"].configure(text=str(closed))

        if not HAS_MPL:
            return

        risk_counter   = Counter(r[4]  for r in records if len(r) > 4)
        status_counter = Counter(r[10] for r in records if len(r) > 10)
        entity_counter = Counter(r[1]  for r in records if len(r) > 1 and r[1])

        # Кругова діаграма ризиків
        self.ax_left.clear()
        self._style_ax(self.ax_left)
        self.ax_left.set_title("Розподiл за типом ризику", color=C["text_muted"], fontsize=9)
        if risk_counter:
            labels = list(risk_counter.keys())
            values = list(risk_counter.values())
            colors = [RISK_COLORS.get(l, C["text_muted"]) for l in labels]
            wedges, texts, autotexts = self.ax_left.pie(
                values, labels=labels, autopct="%1.0f%%",
                colors=colors, startangle=90,
                textprops={"color": C["text_muted"], "fontsize": 7})
            for at in autotexts:
                at.set_fontsize(7)
                at.set_color("white")
        else:
            self.ax_left.text(0.5, 0.5, "Немає даних",
                transform=self.ax_left.transAxes,
                ha="center", va="center", color=C["text_muted"])
        self.canvas_left.draw()

        # Стовпчаста діаграма статусів
        self.ax_right.clear()
        self._style_ax(self.ax_right)
        self.ax_right.set_title("Записи за статусом", color=C["text_muted"], fontsize=9)
        if status_counter:
            s_labels = list(status_counter.keys())
            s_values = list(status_counter.values())
            s_colors = [C["danger"], C["warning"], C["success"], C["accent2"]][:len(s_labels)]
            bars = self.ax_right.bar(s_labels, s_values, color=s_colors, edgecolor="none")
            for bar, val in zip(bars, s_values):
                self.ax_right.text(
                    bar.get_x() + bar.get_width() / 2,
                    bar.get_height() + 0.1,
                    str(val), ha="center", va="bottom",
                    color=C["text_muted"], fontsize=8)
            self.ax_right.tick_params(axis="x", labelrotation=15, labelsize=7)
            self.ax_right.set_ylim(0, max(s_values) * 1.2 + 1)
        else:
            self.ax_right.text(0.5, 0.5, "Немає даних",
                transform=self.ax_right.transAxes,
                ha="center", va="center", color=C["text_muted"])
        self.canvas_right.draw()

        # Топ-5 підприємств
        self.ax_bottom.clear()
        self._style_ax(self.ax_bottom)
        self.ax_bottom.set_title("Топ-5 пiдприємств за кiлькiстю подiй",
                                  color=C["text_muted"], fontsize=9)
        top5 = entity_counter.most_common(5)
        if top5:
            ent_labels = [e[0][:20] for e in top5]
            ent_values = [e[1]      for e in top5]
            bars = self.ax_bottom.barh(ent_labels, ent_values,
                                       color=C["accent"], edgecolor="none")
            for bar, val in zip(bars, ent_values):
                self.ax_bottom.text(
                    bar.get_width() + 0.05,
                    bar.get_y() + bar.get_height() / 2,
                    str(val), ha="left", va="center",
                    color=C["text_muted"], fontsize=8)
            self.ax_bottom.tick_params(axis="y", labelsize=8, colors=C["text_primary"])
        else:
            self.ax_bottom.text(0.5, 0.5, "Немає даних",
                transform=self.ax_bottom.transAxes,
                ha="center", va="center", color=C["text_muted"])
        self.canvas_bottom.draw()

        # Таблиця статистики
        self.stats_tree.delete(*self.stats_tree.get_children())
        all_risks = set(RISK_TYPES) | set(r[4] for r in records if len(r) > 4)
        for risk in sorted(all_risks):
            recs  = [r for r in records if len(r) > 4 and r[4] == risk]
            cnt   = len(recs)
            open_ = sum(1 for r in recs if len(r) > 10 and r[10] in ("Вiдкрито", "В обробцi"))
            cl    = sum(1 for r in recs if len(r) > 10 and r[10] in ("Закрито", "Вирiшено"))
            if cnt:
                self.stats_tree.insert("", tk.END, values=(risk, cnt, open_, cl))

    def get_frame(self):
        return self.frame


# ─── ВКЛАДКА: НАЛАШТУВАННЯ ───────────────────────────────────────────────────
class SettingsTab:
    def __init__(self, parent):
        self.frame = ttk.Frame(parent)
        self._build_ui()

    def _build_ui(self):
        C = COLORS
        self.frame.columnconfigure(0, weight=1)

        header = tk.Frame(self.frame, bg=C["header_bg"], height=56)
        header.grid(row=0, column=0, sticky="ew")
        header.grid_propagate(False)
        tk.Label(header, text="НАЛАШТУВАННЯ",
                 bg=C["header_bg"], fg=C["text_muted"],
                 font=("Segoe UI", 13, "bold")).pack(side="left", padx=20, pady=14)

        content = tk.Frame(self.frame, bg=C["bg_dark"])
        content.grid(row=1, column=0, sticky="nsew", padx=40, pady=30)
        content.columnconfigure(0, weight=1)

        self._row(content, 0, "Файл даних:", DATA_FILE, C)
        self._row(content, 1, "Версiя:", "3.0 — Navbar + Sidebar", C)
        self._row(content, 2, "matplotlib:", "встановлено" if HAS_MPL else "не встановлено", C)
        self._row(content, 3, "pandas:",     "встановлено" if pd      else "не встановлено", C)

        tk.Label(content, text="Встановлення залежностей:",
                 bg=C["bg_dark"], fg=C["text_muted"],
                 font=("Segoe UI", 8, "bold")).grid(row=4, column=0, sticky="w", pady=(24, 6))
        tk.Label(content,
                 text="  pip install matplotlib pandas openpyxl",
                 bg=C["bg_input"], fg=C["accent"],
                 font=("Courier", 9), padx=12, pady=8).grid(row=5, column=0, sticky="w")

        tk.Label(content, text="Як додати новий модуль:",
                 bg=C["bg_dark"], fg=C["text_muted"],
                 font=("Segoe UI", 8, "bold")).grid(row=6, column=0, sticky="w", pady=(24, 6))

        hints = [
            "1. Створiть клас MyModule з атрибутом self.frame = ttk.Frame(...)",
            "2. У main() додайте запис у NAV_ITEMS: {key, label, icon}",
            "3. Зареєструйте фрейм: container.register('mykey', module.get_frame())",
            "4. Сайдбар автоматично вiдобразить новий пункт навiгацiї",
        ]
        for i, hint in enumerate(hints):
            f = tk.Frame(content, bg=C["bg_dark"])
            f.grid(row=7 + i, column=0, sticky="w", pady=3)
            tk.Frame(f, bg=C["accent2"], width=4, height=4).pack(side="left", padx=(0, 8))
            tk.Label(f, text=hint, bg=C["bg_dark"], fg=C["text_label"],
                     font=("Segoe UI", 8)).pack(side="left")

    def _row(self, parent, row, label, value, C):
        f = tk.Frame(parent, bg=C["bg_dark"])
        f.grid(row=row, column=0, sticky="ew", pady=4)
        tk.Label(f, text=label, bg=C["bg_dark"], fg=C["text_muted"],
                 font=("Segoe UI", 9), width=22, anchor="w").pack(side="left")
        tk.Label(f, text=value, bg=C["bg_dark"], fg=C["text_primary"],
                 font=("Segoe UI", 9)).pack(side="left")

    def get_frame(self):
        return self.frame


# ─── ГОЛОВНЕ ВІКНО ───────────────────────────────────────────────────────────
def main():
    root = tk.Tk()
    root.title("Суттєвi подiї — Реєстр v3.0")
    root.geometry("1440x900")
    root.minsize(1100, 700)
    root.configure(bg=COLORS["bg_dark"])

    try:
        root.iconbitmap("icon.ico")
    except Exception:
        pass

    apply_dark_style(root)
    root.option_add("*Font", "\"Segoe UI\" 9")

    # ── Структура вікна ───────────────────────────────────────────────────────
    # root
    #   navbar          (top,  fill=x)
    #   body            (fill=both, expand=True)
    #     sidebar       (left, fill=y)
    #     container     (fill=both, expand=True)
    #   statusbar       (bottom, fill=x)

    # Navbar
    navbar = Navbar(root, app_title="Суттєвi подiї", version="v3.0")
    navbar.pack(side="top", fill="x")

    # Body (sidebar + content)
    body = tk.Frame(root, bg=COLORS["bg_dark"])
    body.pack(side="top", fill="both", expand=True)

    # ── Пункти навігації ──────────────────────────────────────────────────────
    # Щоб додати новий модуль — просто додайте сюди рядок і зареєструйте фрейм
    NAV_ITEMS = [
        {"key": "registry",  "label": "Реєстр подiй",  "icon": "[R]", "badge_color": COLORS["accent"]},
        {"key": "analytics", "label": "Аналiтика",      "icon": "[A]", "badge_color": COLORS["accent2"]},
        {"key": "settings",  "label": "Налаштування",   "icon": "[S]"},
        # Приклад майбутнього модуля:
        # {"key": "reports", "label": "Звiти",          "icon": "[Z]", "badge_color": COLORS["warning"]},
    ]

    # PageContainer — тримає всі фрейми
    container = PageContainer(body)

    def on_nav_select(key: str):
        container.show(key)
        label_map = {
            "registry":  "Реєстр подiй",
            "analytics": "Аналiтика та звiти",
            "settings":  "Налаштування",
        }
        navbar.set_page(label_map.get(key, key))

    # Sidebar
    sidebar = Sidebar(body, nav_items=NAV_ITEMS, on_select=on_nav_select)
    sidebar.pack(side="left", fill="y")

    container.pack(side="left", fill="both", expand=True)

    # ── Ініціалізація модулів ─────────────────────────────────────────────────
    analytics = AnalyticsTab(container)
    registry  = RegistryTab(container, on_data_change=analytics.update_data)
    settings  = SettingsTab(container)

    # Реєстрація фреймів у контейнері
    container.register("registry",  registry.get_frame())
    container.register("analytics", analytics.get_frame())
    container.register("settings",  settings.get_frame())

    # ── Статус-бар ────────────────────────────────────────────────────────────
    statusbar = tk.Frame(root, bg=COLORS["header_bg"], height=24)
    statusbar.pack(side="bottom", fill="x")
    statusbar.pack_propagate(False)

    status_lbl = tk.Label(statusbar, text="Готово",
                           bg=COLORS["header_bg"], fg=COLORS["text_muted"],
                           font=("Segoe UI", 7), padx=12)
    status_lbl.pack(side="left", pady=4)

    records_lbl = tk.Label(statusbar, text="",
                            bg=COLORS["header_bg"], fg=COLORS["text_muted"],
                            font=("Segoe UI", 7), padx=12)
    records_lbl.pack(side="left", pady=4)

    # Синхронізація лічильника записів у статус-барі
    original_on_data_change = analytics.update_data
    def on_data_change_wrapped(records):
        original_on_data_change(records)
        records_lbl.configure(text=f"|  Записiв у БД: {len(records)}")
    registry.on_data_change = on_data_change_wrapped

    # Автозбереження
    def autosave():
        registry._save_data()
        status_lbl.configure(
            text=f"Автозбережено о {datetime.now().strftime('%H:%M:%S')}")
        navbar.set_status(
            f"Автозбережено {datetime.now().strftime('%H:%M')}",
            color=COLORS["success"])
        root.after(30000, autosave)

    root.after(30000, autosave)

    def on_close():
        registry._save_data()
        root.destroy()

    root.protocol("WM_DELETE_WINDOW", on_close)

    # Показати стартову сторінку
    sidebar.select("registry")

    # Оновити аналітику після завантаження
    root.after(500, lambda: analytics.update_data(registry.all_records))

    root.mainloop()


if __name__ == "__main__":
    main()
