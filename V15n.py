"""
╔══════════════════════════════════════════════════════════════════╗
║         CRM / ERP / GRC  —  Desktop Starter Template            ║
║         Python 3.10+  |  Tkinter  |  Hi-Tech Light/Dark UI      ║
╚══════════════════════════════════════════════════════════════════╝

Структура проєкту:
  crm_erp_grc_template.py   ← цей файл (точка входу)

Архітектура:
  App                 — кореневе вікно + менеджер тем
  ThemeManager        — усі кольори / шрифти для LIGHT & DARK
  TopNavBar           — горизонтальна навігація + лого + дії
  SideBar             — вертикальне меню з підменю (акордеон)
  ContentArea         — головна робоча зона (frames-менеджер)
  StatusBar           — нижня смуга статусу
  BaseView            — базовий клас для будь-якого модуля
  DashboardView       — стартова панель
  ListTableView       — шаблон списку / таблиці
  FormView            — шаблон форми
  SettingsView        — налаштування + перемикач теми
  ModalDialog         — базовий модальний діалог
  NotificationBanner  — спливаючі повідомлення
  SearchBar           — глобальний пошук
  Tooltip             — підказки

Щоб додати новий модуль:
  1. Успадкуйте BaseView
  2. Додайте запис у MENU_STRUCTURE
  3. Зареєструйте у ContentArea.register()
"""

import tkinter as tk
from tkinter import ttk, messagebox, font as tkfont
import time
import threading
import platform
import sys
from datetime import datetime
from typing import Optional, Callable, Dict, Any


# ═══════════════════════════════════════════════════════════════
#  КОНСТАНТИ
# ═══════════════════════════════════════════════════════════════

APP_NAME    = "NexusFlow"          # змінте на назву вашої системи
APP_VERSION = "1.0.0"
APP_TAGLINE = "CRM · ERP · GRC"

LOGO_MODE   = "text_icon"          # "text" | "icon" | "text_icon"

MENU_STRUCTURE = [
    {
        "id":    "dashboard",
        "label": "Dashboard",
        "icon":  "⊞",
        "view":  "DashboardView",
    },
    {
        "id":    "crm",
        "label": "CRM",
        "icon":  "◈",
        "children": [
            {"id": "contacts",    "label": "Контакти",    "icon": "◉", "view": "ListTableView"},
            {"id": "leads",       "label": "Ліди",        "icon": "◎", "view": "ListTableView"},
            {"id": "deals",       "label": "Угоди",       "icon": "◇", "view": "ListTableView"},
            {"id": "activities",  "label": "Активності",  "icon": "◆", "view": "ListTableView"},
        ],
    },
    {
        "id":    "erp",
        "label": "ERP",
        "icon":  "◫",
        "children": [
            {"id": "finance",     "label": "Фінанси",     "icon": "◈", "view": "ListTableView"},
            {"id": "inventory",   "label": "Склад",       "icon": "▣",  "view": "ListTableView"},
            {"id": "hr",          "label": "HR",          "icon": "◉", "view": "ListTableView"},
            {"id": "projects",    "label": "Проєкти",     "icon": "◇", "view": "ListTableView"},
        ],
    },
    {
        "id":    "grc",
        "label": "GRC",
        "icon":  "◬",
        "children": [
            {"id": "risks",       "label": "Ризики",      "icon": "△",  "view": "ListTableView"},
            {"id": "compliance",  "label": "Комплаєнс",   "icon": "☑",  "view": "ListTableView"},
            {"id": "audit",       "label": "Аудит",       "icon": "◎", "view": "ListTableView"},
            {"id": "policies",    "label": "Політики",    "icon": "▤",  "view": "ListTableView"},
        ],
    },
    {
        "id":    "analytics",
        "label": "Аналітика",
        "icon":  "◈",
        "children": [
            {"id": "reports",     "label": "Звіти",       "icon": "▦",  "view": "ListTableView"},
            {"id": "kpi",         "label": "KPI",         "icon": "◆", "view": "DashboardView"},
            {"id": "charts",      "label": "Графіки",     "icon": "◈", "view": "DashboardView"},
        ],
    },
    {
        "id":    "settings",
        "label": "Налаштування",
        "icon":  "⚙",
        "view":  "SettingsView",
    },
]


# ═══════════════════════════════════════════════════════════════
#  МЕНЕДЖЕР ТЕМ
# ═══════════════════════════════════════════════════════════════

class ThemeManager:
    """
    Централізоване сховище кольорів, шрифтів і розмірів.
    Всі компоненти звертаються тільки сюди.
    """

    THEMES: Dict[str, Dict[str, Any]] = {

        # ────── LIGHT ──────────────────────────────────────────
        "light": {
            # Backgrounds
            "bg_app":          "#F0F2F5",
            "bg_sidebar":      "#FFFFFF",
            "bg_sidebar_item": "#FFFFFF",
            "bg_sidebar_hover":"#EEF2FF",
            "bg_sidebar_active":"#4F46E5",
            "bg_topbar":       "#FFFFFF",
            "bg_content":      "#F0F2F5",
            "bg_card":         "#FFFFFF",
            "bg_card2":        "#F8FAFC",
            "bg_input":        "#F8FAFC",
            "bg_input_focus":  "#FFFFFF",
            "bg_table_row":    "#FFFFFF",
            "bg_table_alt":    "#F8FAFC",
            "bg_table_hover":  "#EEF2FF",
            "bg_table_header": "#F1F5F9",
            "bg_modal":        "#FFFFFF",
            "bg_overlay":      "#00000044",
            "bg_tooltip":      "#1E293B",
            "bg_statusbar":    "#FFFFFF",
            "bg_badge":        "#EEF2FF",
            "bg_btn_primary":  "#4F46E5",
            "bg_btn_secondary":"#F1F5F9",
            "bg_btn_danger":   "#FEE2E2",
            "bg_btn_success":  "#D1FAE5",
            "bg_separator":    "#E2E8F0",
            "bg_scrollbar":    "#E2E8F0",
            "bg_scrollbar_thumb": "#CBD5E1",

            # Text
            "fg_primary":      "#0F172A",
            "fg_secondary":    "#475569",
            "fg_muted":        "#94A3B8",
            "fg_sidebar_item": "#475569",
            "fg_sidebar_active":"#FFFFFF",
            "fg_topbar":       "#0F172A",
            "fg_label":        "#334155",
            "fg_value":        "#0F172A",
            "fg_table_header": "#334155",
            "fg_table_cell":   "#0F172A",
            "fg_tooltip":      "#F8FAFC",
            "fg_badge":        "#4F46E5",
            "fg_btn_primary":  "#FFFFFF",
            "fg_btn_secondary":"#334155",
            "fg_btn_danger":   "#DC2626",
            "fg_btn_success":  "#059669",
            "fg_statusbar":    "#64748B",
            "fg_logo":         "#0F172A",
            "fg_logo_accent":  "#4F46E5",

            # Accent / Brand
            "accent":          "#4F46E5",      # Indigo-600
            "accent_hover":    "#4338CA",
            "accent_light":    "#EEF2FF",
            "accent2":         "#06B6D4",      # Cyan-500
            "accent3":         "#10B981",      # Emerald-500
            "danger":          "#EF4444",
            "warning":         "#F59E0B",
            "success":         "#10B981",
            "info":            "#06B6D4",

            # Border
            "border":          "#E2E8F0",
            "border_focus":    "#4F46E5",
            "border_card":     "#E2E8F0",
            "border_input":    "#CBD5E1",

            # Misc
            "shadow":          "#00000015",
            "radius":          8,
        },

        # ────── DARK ───────────────────────────────────────────
        "dark": {
            # Backgrounds
            "bg_app":          "#0B0F19",
            "bg_sidebar":      "#111827",
            "bg_sidebar_item": "#111827",
            "bg_sidebar_hover":"#1E2A4A",
            "bg_sidebar_active":"#4F46E5",
            "bg_topbar":       "#111827",
            "bg_content":      "#0B0F19",
            "bg_card":         "#111827",
            "bg_card2":        "#1A2235",
            "bg_input":        "#1A2235",
            "bg_input_focus":  "#1E293B",
            "bg_table_row":    "#111827",
            "bg_table_alt":    "#141E30",
            "bg_table_hover":  "#1E2A4A",
            "bg_table_header": "#1A2235",
            "bg_modal":        "#1A2235",
            "bg_overlay":      "#00000077",
            "bg_tooltip":      "#F8FAFC",
            "bg_statusbar":    "#111827",
            "bg_badge":        "#1E2A4A",
            "bg_btn_primary":  "#4F46E5",
            "bg_btn_secondary":"#1A2235",
            "bg_btn_danger":   "#450A0A",
            "bg_btn_success":  "#052E16",
            "bg_separator":    "#1E293B",
            "bg_scrollbar":    "#1A2235",
            "bg_scrollbar_thumb": "#334155",

            # Text
            "fg_primary":      "#F1F5F9",
            "fg_secondary":    "#94A3B8",
            "fg_muted":        "#475569",
            "fg_sidebar_item": "#94A3B8",
            "fg_sidebar_active":"#FFFFFF",
            "fg_topbar":       "#F1F5F9",
            "fg_label":        "#94A3B8",
            "fg_value":        "#F1F5F9",
            "fg_table_header": "#94A3B8",
            "fg_table_cell":   "#E2E8F0",
            "fg_tooltip":      "#0F172A",
            "fg_badge":        "#818CF8",
            "fg_btn_primary":  "#FFFFFF",
            "fg_btn_secondary":"#CBD5E1",
            "fg_btn_danger":   "#FCA5A5",
            "fg_btn_success":  "#6EE7B7",
            "fg_statusbar":    "#475569",
            "fg_logo":         "#F1F5F9",
            "fg_logo_accent":  "#818CF8",

            # Accent / Brand
            "accent":          "#6366F1",
            "accent_hover":    "#818CF8",
            "accent_light":    "#1E2A4A",
            "accent2":         "#22D3EE",
            "accent3":         "#34D399",
            "danger":          "#F87171",
            "warning":         "#FCD34D",
            "success":         "#34D399",
            "info":            "#22D3EE",

            # Border
            "border":          "#1E293B",
            "border_focus":    "#6366F1",
            "border_card":     "#1E293B",
            "border_input":    "#334155",

            # Misc
            "shadow":          "#00000044",
            "radius":          8,
        },
    }

    FONTS: Dict[str, tuple] = {
        "app_title":    ("Segoe UI", 16, "bold"),
        "logo":         ("Segoe UI", 15, "bold"),
        "logo_sub":     ("Segoe UI", 8,  "normal"),
        "nav_item":     ("Segoe UI", 10, "normal"),
        "nav_item_bold":("Segoe UI", 10, "bold"),
        "section_head": ("Segoe UI", 9,  "bold"),
        "body":         ("Segoe UI", 10, "normal"),
        "body_small":   ("Segoe UI", 9,  "normal"),
        "body_bold":    ("Segoe UI", 10, "bold"),
        "h1":           ("Segoe UI", 22, "bold"),
        "h2":           ("Segoe UI", 16, "bold"),
        "h3":           ("Segoe UI", 13, "bold"),
        "card_value":   ("Segoe UI", 28, "bold"),
        "card_label":   ("Segoe UI", 10, "normal"),
        "table_head":   ("Segoe UI", 9,  "bold"),
        "table_cell":   ("Segoe UI", 10, "normal"),
        "btn":          ("Segoe UI", 10, "bold"),
        "btn_sm":       ("Segoe UI", 9,  "normal"),
        "status":       ("Segoe UI", 9,  "normal"),
        "tooltip":      ("Segoe UI", 9,  "normal"),
        "mono":         ("Consolas",  9,  "normal"),
    }

    SIZES = {
        "sidebar_w":         220,
        "sidebar_collapsed":  52,
        "topbar_h":           56,
        "statusbar_h":        26,
        "sidebar_item_h":     38,
        "card_pad":           16,
        "input_h":            34,
        "btn_h":              32,
        "icon_btn":           32,
        "separator":           1,
        "radius":              8,
    }

    def __init__(self, initial: str = "light"):
        self._name = initial
        self._callbacks: list[Callable] = []

    @property
    def name(self) -> str:
        return self._name

    def get(self, key: str, fallback: Any = None) -> Any:
        return self.THEMES[self._name].get(key, fallback)

    def font(self, key: str) -> tuple:
        return self.FONTS.get(key, ("Segoe UI", 10, "normal"))

    def size(self, key: str) -> int:
        return self.SIZES.get(key, 0)

    def switch(self, name: Optional[str] = None):
        if name:
            self._name = name
        else:
            self._name = "dark" if self._name == "light" else "light"
        for cb in self._callbacks:
            cb()

    def subscribe(self, cb: Callable):
        self._callbacks.append(cb)


# ═══════════════════════════════════════════════════════════════
#  TOOLTIP
# ═══════════════════════════════════════════════════════════════

class Tooltip:
    def __init__(self, widget, text: str, theme: ThemeManager):
        self.widget = widget
        self.text   = text
        self.theme  = theme
        self._tw: Optional[tk.Toplevel] = None
        widget.bind("<Enter>", self._show)
        widget.bind("<Leave>", self._hide)

    def _show(self, _=None):
        x = self.widget.winfo_rootx() + self.widget.winfo_width() + 4
        y = self.widget.winfo_rooty() + 4
        self._tw = tk.Toplevel(self.widget)
        self._tw.wm_overrideredirect(True)
        self._tw.wm_geometry(f"+{x}+{y}")
        lbl = tk.Label(
            self._tw, text=self.text,
            bg=self.theme.get("bg_tooltip"),
            fg=self.theme.get("fg_tooltip"),
            font=self.theme.font("tooltip"),
            padx=8, pady=4,
        )
        lbl.pack()

    def _hide(self, _=None):
        if self._tw:
            self._tw.destroy()
            self._tw = None


# ═══════════════════════════════════════════════════════════════
#  NOTIFICATION BANNER
# ═══════════════════════════════════════════════════════════════

class NotificationBanner:
    """Спливаюче повідомлення знизу праворуч."""

    TYPE_COLORS = {
        "success": ("#059669", "#ECFDF5"),
        "error":   ("#DC2626", "#FEF2F2"),
        "warning": ("#D97706", "#FFFBEB"),
        "info":    ("#0284C7", "#EFF6FF"),
    }

    def __init__(self, root: tk.Tk):
        self.root = root
        self._queue: list = []
        self._showing = False

    def show(self, message: str, kind: str = "info", duration: int = 3000):
        self._queue.append((message, kind, duration))
        if not self._showing:
            self._next()

    def _next(self):
        if not self._queue:
            self._showing = False
            return
        self._showing = True
        msg, kind, dur = self._queue.pop(0)
        accent, bg = self.TYPE_COLORS.get(kind, self.TYPE_COLORS["info"])

        popup = tk.Toplevel(self.root)
        popup.overrideredirect(True)
        popup.attributes("-topmost", True)
        popup.configure(bg=accent)

        frame = tk.Frame(popup, bg=bg, padx=14, pady=10)
        frame.pack(fill="both", expand=True, padx=2, pady=2)

        icons = {"success": "✔", "error": "✖", "warning": "⚠", "info": "ℹ"}
        tk.Label(frame, text=icons.get(kind,"ℹ"), bg=bg, fg=accent,
                 font=("Segoe UI", 14, "bold")).pack(side="left", padx=(0,8))
        tk.Label(frame, text=msg, bg=bg, fg="#1E293B",
                 font=("Segoe UI", 10), wraplength=240, justify="left"
                 ).pack(side="left")

        self.root.update_idletasks()
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        popup.update_idletasks()
        pw = popup.winfo_reqwidth()
        ph = popup.winfo_reqheight()
        popup.geometry(f"{pw}x{ph}+{sw - pw - 20}+{sh - ph - 60}")

        popup.after(dur, lambda: self._dismiss(popup))

    def _dismiss(self, popup):
        try:
            popup.destroy()
        except Exception:
            pass
        self.root.after(200, self._next)


# ═══════════════════════════════════════════════════════════════
#  MODAL DIALOG
# ═══════════════════════════════════════════════════════════════

class ModalDialog(tk.Toplevel):
    """
    Базовий модальний діалог.
    Успадковуйте та перевизначте body() і on_ok().
    """

    def __init__(self, parent, title: str, theme: ThemeManager,
                 width: int = 480, height: int = 340):
        super().__init__(parent)
        self.theme  = theme
        self.result = None

        self.title(title)
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()
        self.configure(bg=theme.get("bg_modal"))

        # центрування
        self.update_idletasks()
        px = parent.winfo_rootx() + (parent.winfo_width()  - width)  // 2
        py = parent.winfo_rooty() + (parent.winfo_height() - height) // 2
        self.geometry(f"{width}x{height}+{px}+{py}")

        self._build_chrome(title)
        inner = tk.Frame(self, bg=theme.get("bg_modal"), padx=20, pady=16)
        inner.pack(fill="both", expand=True)
        self.body(inner)
        self._build_buttons()
        self.protocol("WM_DELETE_WINDOW", self.cancel)

    def _build_chrome(self, title: str):
        bar = tk.Frame(self, bg=self.theme.get("accent"), height=4)
        bar.pack(fill="x")
        hdr = tk.Frame(self, bg=self.theme.get("bg_modal"), pady=14, padx=20)
        hdr.pack(fill="x")
        tk.Label(hdr, text=title,
                 bg=self.theme.get("bg_modal"),
                 fg=self.theme.get("fg_primary"),
                 font=self.theme.font("h3")).pack(side="left")
        tk.Button(hdr, text="✕", bg=self.theme.get("bg_modal"),
                  fg=self.theme.get("fg_muted"),
                  relief="flat", bd=0, cursor="hand2",
                  command=self.cancel,
                  font=("Segoe UI", 12)).pack(side="right")
        tk.Frame(self, bg=self.theme.get("border"), height=1).pack(fill="x")

    def _build_buttons(self):
        tk.Frame(self, bg=self.theme.get("border"), height=1).pack(fill="x")
        bar = tk.Frame(self, bg=self.theme.get("bg_modal"), pady=12, padx=20)
        bar.pack(fill="x")
        StyledButton(bar, "Скасувати", self.theme, style="secondary",
                     command=self.cancel).pack(side="right", padx=(6,0))
        StyledButton(bar, "  OK  ", self.theme, style="primary",
                     command=self.ok).pack(side="right")

    def body(self, master):
        """Перевизначте для додавання контенту."""
        tk.Label(master, text="Тіло діалогу",
                 bg=self.theme.get("bg_modal"),
                 fg=self.theme.get("fg_secondary"),
                 font=self.theme.font("body")).pack(pady=30)

    def ok(self):
        self.on_ok()
        self.destroy()

    def cancel(self):
        self.result = None
        self.destroy()

    def on_ok(self):
        """Перевизначте для обробки підтвердження."""
        self.result = True


# ═══════════════════════════════════════════════════════════════
#  STYLED WIDGETS
# ═══════════════════════════════════════════════════════════════

class StyledButton(tk.Button):
    STYLES = {
        "primary":   ("bg_btn_primary",   "fg_btn_primary"),
        "secondary": ("bg_btn_secondary", "fg_btn_secondary"),
        "danger":    ("bg_btn_danger",    "fg_btn_danger"),
        "success":   ("bg_btn_success",   "fg_btn_success"),
        "ghost":     ("bg_card",          "fg_secondary"),
    }

    def __init__(self, parent, text: str, theme: ThemeManager,
                 style: str = "primary", command=None, icon: str = "", **kw):
        bg_key, fg_key = self.STYLES.get(style, self.STYLES["primary"])
        label = f"{icon}  {text}" if icon else text
        super().__init__(
            parent, text=label,
            bg=theme.get(bg_key), fg=theme.get(fg_key),
            font=theme.font("btn"),
            relief="flat", bd=0, cursor="hand2",
            activebackground=theme.get("accent_hover"),
            activeforeground="#FFFFFF",
            padx=14, pady=0,
            height=1,
            command=command, **kw,
        )
        self._theme   = theme
        self._style   = style
        self._bg_key  = bg_key
        self._fg_key  = fg_key
        self.bind("<Enter>", self._on_enter)
        self.bind("<Leave>", self._on_leave)

    def _on_enter(self, _=None):
        if self._style == "primary":
            self.configure(bg=self._theme.get("accent_hover"))

    def _on_leave(self, _=None):
        self.configure(bg=self._theme.get(self._bg_key))


class StyledEntry(tk.Frame):
    """Текстове поле з підписом та рамкою."""

    def __init__(self, parent, label: str, theme: ThemeManager,
                 placeholder: str = "", **kw):
        super().__init__(parent, bg=theme.get("bg_card"))
        self.theme = theme
        self._var = tk.StringVar()

        tk.Label(self, text=label,
                 bg=theme.get("bg_card"),
                 fg=theme.get("fg_label"),
                 font=theme.font("body_small")).pack(anchor="w", pady=(0,2))

        border = tk.Frame(self, bg=theme.get("border_input"),
                          padx=1, pady=1)
        border.pack(fill="x")

        inner = tk.Frame(border, bg=theme.get("bg_input"))
        inner.pack(fill="x")

        self._entry = tk.Entry(
            inner, textvariable=self._var,
            bg=theme.get("bg_input"),
            fg=theme.get("fg_value"),
            insertbackground=theme.get("accent"),
            relief="flat", bd=4,
            font=theme.font("body"), **kw,
        )
        self._entry.pack(fill="x")

        if placeholder:
            self._entry.insert(0, placeholder)
            self._entry.configure(fg=theme.get("fg_muted"))
            self._entry.bind("<FocusIn>",  lambda _: self._clear_ph(placeholder))
            self._entry.bind("<FocusOut>", lambda _: self._restore_ph(placeholder))

        self._entry.bind("<FocusIn>",  lambda _: border.configure(
            bg=theme.get("border_focus")), add="+")
        self._entry.bind("<FocusOut>", lambda _: border.configure(
            bg=theme.get("border_input")), add="+")

    def _clear_ph(self, ph):
        if self._entry.get() == ph:
            self._entry.delete(0, "end")
            self._entry.configure(fg=self.theme.get("fg_value"))

    def _restore_ph(self, ph):
        if not self._entry.get():
            self._entry.insert(0, ph)
            self._entry.configure(fg=self.theme.get("fg_muted"))

    def get(self) -> str:
        return self._var.get()

    def set(self, value: str):
        self._var.set(value)


class KpiCard(tk.Frame):
    """Картка KPI для Dashboard."""

    def __init__(self, parent, title: str, value: str, delta: str,
                 delta_positive: bool, icon: str, theme: ThemeManager):
        super().__init__(parent,
                         bg=theme.get("bg_card"),
                         relief="flat", bd=0)
        self.configure(padx=0, pady=0)

        # верхня акцентна смуга
        accent_bar = tk.Frame(self, bg=theme.get("accent"), height=3)
        accent_bar.pack(fill="x")

        body = tk.Frame(self, bg=theme.get("bg_card"), padx=18, pady=16)
        body.pack(fill="both", expand=True)

        # іконка + заголовок
        top = tk.Frame(body, bg=theme.get("bg_card"))
        top.pack(fill="x")
        tk.Label(top, text=icon, bg=theme.get("bg_card"),
                 fg=theme.get("accent"), font=("Segoe UI", 20)).pack(side="left")
        tk.Label(top, text=title, bg=theme.get("bg_card"),
                 fg=theme.get("fg_secondary"),
                 font=theme.font("card_label")).pack(side="left", padx=(8,0))

        # значення
        tk.Label(body, text=value, bg=theme.get("bg_card"),
                 fg=theme.get("fg_primary"),
                 font=theme.font("card_value")).pack(anchor="w", pady=(10,2))

        # дельта
        d_color = theme.get("success") if delta_positive else theme.get("danger")
        d_arrow = "▲" if delta_positive else "▼"
        tk.Label(body, text=f"{d_arrow} {delta}",
                 bg=theme.get("bg_card"), fg=d_color,
                 font=theme.font("body_small")).pack(anchor="w")


# ═══════════════════════════════════════════════════════════════
#  TOP NAV BAR
# ═══════════════════════════════════════════════════════════════

class TopNavBar(tk.Frame):
    """
    Горизонтальна навігація: лого | пошук | дії | профіль
    """

    def __init__(self, parent, theme: ThemeManager,
                 on_toggle_sidebar: Callable,
                 on_toggle_theme:   Callable,
                 on_search:         Callable,
                 notifier:          "NotificationBanner"):
        super().__init__(parent,
                         bg=theme.get("bg_topbar"),
                         height=theme.size("topbar_h"))
        self.pack_propagate(False)

        self.theme    = theme
        self.notifier = notifier

        # ── Ліво: бургер + лого ─────────────────────────────
        left = tk.Frame(self, bg=theme.get("bg_topbar"))
        left.pack(side="left", fill="y", padx=(10,0))

        self._burger = tk.Button(
            left, text="☰",
            bg=theme.get("bg_topbar"),
            fg=theme.get("fg_topbar"),
            font=("Segoe UI", 16),
            relief="flat", bd=0, cursor="hand2",
            command=on_toggle_sidebar,
        )
        self._burger.pack(side="left", padx=(4,8))

        self._build_logo(left)

        # ── Центр: пошук ───────────────────────────────────
        center = tk.Frame(self, bg=theme.get("bg_topbar"))
        center.pack(side="left", fill="y", expand=True, padx=20)

        search_wrap = tk.Frame(center,
                               bg=theme.get("bg_input"),
                               highlightbackground=theme.get("border"),
                               highlightthickness=1)
        search_wrap.pack(side="left", ipady=2, padx=0, pady=10, fill="x",
                         expand=True)

        tk.Label(search_wrap, text="⌕",
                 bg=theme.get("bg_input"),
                 fg=theme.get("fg_muted"),
                 font=("Segoe UI", 12)).pack(side="left", padx=(8,2))

        self._search_var = tk.StringVar()
        self._search_entry = tk.Entry(
            search_wrap,
            textvariable=self._search_var,
            bg=theme.get("bg_input"),
            fg=theme.get("fg_primary"),
            insertbackground=theme.get("accent"),
            relief="flat", bd=2,
            font=theme.font("body"),
        )
        self._search_entry.pack(side="left", fill="x", expand=True)
        self._search_entry.insert(0, "Пошук…")
        self._search_entry.configure(fg=theme.get("fg_muted"))
        self._search_entry.bind("<FocusIn>",  self._search_focus_in)
        self._search_entry.bind("<FocusOut>", self._search_focus_out)
        self._search_entry.bind("<Return>",
                                lambda _: on_search(self._search_var.get()))

        # ── Право: кнопки дій ──────────────────────────────
        right = tk.Frame(self, bg=theme.get("bg_topbar"))
        right.pack(side="right", fill="y", padx=(0,12))

        # Перемикач теми
        self._theme_btn = tk.Button(
            right, text="◑",
            bg=theme.get("bg_topbar"), fg=theme.get("fg_topbar"),
            font=("Segoe UI", 14), relief="flat", bd=0, cursor="hand2",
            command=on_toggle_theme,
        )
        self._theme_btn.pack(side="left", padx=4)
        Tooltip(self._theme_btn, "Переключити тему", theme)

        # Сповіщення
        notif_btn = tk.Button(
            right, text="🔔",
            bg=theme.get("bg_topbar"), fg=theme.get("fg_topbar"),
            font=("Segoe UI", 12), relief="flat", bd=0, cursor="hand2",
            command=lambda: notifier.show("Немає нових сповіщень", "info"),
        )
        notif_btn.pack(side="left", padx=4)
        Tooltip(notif_btn, "Сповіщення", theme)

        # Роздільник
        tk.Frame(right, bg=theme.get("border"), width=1).pack(
            side="left", fill="y", padx=8, pady=10)

        # Профіль
        profile = tk.Frame(right, bg=theme.get("bg_topbar"), cursor="hand2")
        profile.pack(side="left")
        av = tk.Frame(profile, bg=theme.get("accent"),
                      width=30, height=30)
        av.pack_propagate(False)
        av.pack(side="left")
        tk.Label(av, text="A", bg=theme.get("accent"),
                 fg="#FFF", font=("Segoe UI", 10, "bold")).place(
            relx=.5, rely=.5, anchor="center")
        tk.Label(profile, text="Admin",
                 bg=theme.get("bg_topbar"),
                 fg=theme.get("fg_topbar"),
                 font=theme.font("body_bold")).pack(side="left", padx=(6,0))

        # Нижня розділювальна лінія зберігається як атрибут.
        # App пакує її відразу після self.pack() — так уникаємо
        # TclError "window isn't packed" при використанні after=self.
        self._bottom_sep = tk.Frame(parent, bg=theme.get("border"), height=1)

    def _build_logo(self, parent):
        logo_frame = tk.Frame(parent, bg=self.theme.get("bg_topbar"))
        logo_frame.pack(side="left", padx=(0, 16))

        if LOGO_MODE in ("icon", "text_icon"):
            icon_box = tk.Frame(logo_frame,
                                bg=self.theme.get("accent"),
                                width=28, height=28)
            icon_box.pack_propagate(False)
            icon_box.pack(side="left")
            tk.Label(icon_box, text="N",
                     bg=self.theme.get("accent"),
                     fg="#FFFFFF",
                     font=("Segoe UI", 13, "bold")).place(
                relx=.5, rely=.5, anchor="center")

        if LOGO_MODE in ("text", "text_icon"):
            txt_frame = tk.Frame(logo_frame, bg=self.theme.get("bg_topbar"))
            txt_frame.pack(side="left", padx=(6, 0))
            tk.Label(txt_frame, text=APP_NAME,
                     bg=self.theme.get("bg_topbar"),
                     fg=self.theme.get("fg_logo"),
                     font=self.theme.font("logo")).pack(anchor="w")
            tk.Label(txt_frame, text=APP_TAGLINE,
                     bg=self.theme.get("bg_topbar"),
                     fg=self.theme.get("fg_logo_accent"),
                     font=self.theme.font("logo_sub")).pack(anchor="w")

    def _search_focus_in(self, _=None):
        if self._search_entry.get() == "Пошук…":
            self._search_entry.delete(0, "end")
            self._search_entry.configure(fg=self.theme.get("fg_value"))

    def _search_focus_out(self, _=None):
        if not self._search_entry.get():
            self._search_entry.insert(0, "Пошук…")
            self._search_entry.configure(fg=self.theme.get("fg_muted"))

    def apply_theme(self):
        """Викликається при зміні теми."""
        t = self.theme
        self.configure(bg=t.get("bg_topbar"))
        # Повне перебудування — спрощений підхід для шаблону


# ═══════════════════════════════════════════════════════════════
#  SIDE BAR
# ═══════════════════════════════════════════════════════════════

class SideBar(tk.Frame):
    """
    Бічне меню акордеон з підменю.
    Підтримує collapsed / expanded стани.
    """

    def __init__(self, parent, theme: ThemeManager,
                 on_navigate: Callable,
                 menu: list):
        self._theme       = theme
        self._collapsed   = False
        self._on_navigate = on_navigate
        self._menu        = menu
        self._active_id   = "dashboard"
        self._open_groups: set = set()
        self._item_widgets: Dict[str, tk.Widget] = {}

        super().__init__(parent,
                         bg=theme.get("bg_sidebar"),
                         width=theme.size("sidebar_w"))
        self.pack_propagate(False)

        self._scrollable_canvas = tk.Canvas(
            self, bg=theme.get("bg_sidebar"),
            highlightthickness=0, bd=0)
        self._scrollable_canvas.pack(fill="both", expand=True)

        self._inner = tk.Frame(self._scrollable_canvas,
                               bg=theme.get("bg_sidebar"))
        self._scrollable_canvas.create_window(
            (0, 0), window=self._inner, anchor="nw",
            tags="inner_win")
        self._inner.bind("<Configure>", self._on_inner_configure)

        self._build_menu()
        self._build_bottom()

    def _on_inner_configure(self, _=None):
        self._scrollable_canvas.configure(
            scrollregion=self._scrollable_canvas.bbox("all"))
        self._scrollable_canvas.itemconfigure(
            "inner_win", width=self.winfo_width())

    def _build_menu(self):
        # Верхній відступ
        tk.Frame(self._inner, bg=self._theme.get("bg_sidebar"),
                 height=10).pack(fill="x")

        for item in self._menu:
            self._build_item(item)

    def _build_item(self, item: dict, depth: int = 0):
        has_children = bool(item.get("children"))
        icon  = item.get("icon", "·")
        label = item.get("label", "")
        iid   = item.get("id", "")

        row = tk.Frame(self._inner, bg=self._theme.get("bg_sidebar"),
                       cursor="hand2")
        row.pack(fill="x", pady=(1,0))

        is_active = (iid == self._active_id)
        row_bg = self._theme.get("bg_sidebar_active") if is_active \
                 else self._theme.get("bg_sidebar_item")
        row_fg = self._theme.get("fg_sidebar_active") if is_active \
                 else self._theme.get("fg_sidebar_item")

        row.configure(bg=row_bg)

        # Accent bar для активного
        if is_active:
            acc = tk.Frame(row, bg="#FFFFFF", width=3)
            acc.pack(side="left", fill="y")
        else:
            tk.Frame(row, bg=row_bg, width=3).pack(side="left", fill="y")

        # Відступ для вкладеності
        if depth > 0:
            tk.Frame(row, bg=row_bg, width=depth*18).pack(side="left")

        # Іконка
        icon_lbl = tk.Label(row, text=icon, bg=row_bg, fg=row_fg,
                             font=("Segoe UI", 13), width=2)
        icon_lbl.pack(side="left", padx=(4,6), pady=8)

        # Текст
        text_lbl = tk.Label(row, text=label, bg=row_bg, fg=row_fg,
                             font=self._theme.font(
                                 "nav_item_bold" if is_active else "nav_item"),
                             anchor="w")
        text_lbl.pack(side="left", fill="x", expand=True)

        # Стрілка для груп
        if has_children:
            arrow_var = tk.StringVar(value="▾" if iid in self._open_groups else "›")
            arrow_lbl = tk.Label(row, textvariable=arrow_var,
                                  bg=row_bg, fg=row_fg,
                                  font=("Segoe UI", 11))
            arrow_lbl.pack(side="right", padx=10)
        else:
            arrow_var = None
            arrow_lbl = None

        self._item_widgets[iid] = row

        # Клік
        def on_click(event=None, _iid=iid, _has=has_children,
                     _avar=arrow_var):
            if _has:
                self._toggle_group(_iid, _avar)
            else:
                self._navigate(_iid)

        for w in [row, icon_lbl, text_lbl]:
            w.bind("<Button-1>", on_click)
        if arrow_lbl:
            arrow_lbl.bind("<Button-1>", on_click)

        # Hover
        def on_enter(e=None, _row=row, _iid=iid):
            if _iid != self._active_id:
                _row.configure(bg=self._theme.get("bg_sidebar_hover"))
                for child in _row.winfo_children():
                    try:
                        child.configure(bg=self._theme.get("bg_sidebar_hover"))
                    except Exception:
                        pass

        def on_leave(e=None, _row=row, _iid=iid):
            if _iid != self._active_id:
                _row.configure(bg=self._theme.get("bg_sidebar_item"))
                for child in _row.winfo_children():
                    try:
                        child.configure(bg=self._theme.get("bg_sidebar_item"))
                    except Exception:
                        pass

        row.bind("<Enter>", on_enter)
        row.bind("<Leave>", on_leave)

        # Дочірні пункти
        if has_children:
            children_frame = tk.Frame(self._inner,
                                      bg=self._theme.get("bg_sidebar"))
            children_frame.pack(fill="x")
            if iid not in self._open_groups:
                children_frame.pack_forget()

            for child in item.get("children", []):
                self._build_child(child, children_frame)

            # зберігаємо посилання
            row._children_frame  = children_frame
            row._arrow_var       = arrow_var

    def _build_child(self, item: dict, parent: tk.Frame, depth: int = 1):
        icon  = item.get("icon", "·")
        label = item.get("label", "")
        iid   = item.get("id", "")

        row = tk.Frame(parent, bg=self._theme.get("bg_sidebar"),
                       cursor="hand2")
        row.pack(fill="x", pady=(1,0))

        is_active = (iid == self._active_id)
        row_bg = self._theme.get("bg_sidebar_active") if is_active \
                 else self._theme.get("bg_sidebar_item")
        row_fg = self._theme.get("fg_sidebar_active") if is_active \
                 else self._theme.get("fg_sidebar_item")
        row.configure(bg=row_bg)

        tk.Frame(row, bg=row_bg, width=3).pack(side="left", fill="y")
        tk.Frame(row, bg=row_bg, width=depth*20).pack(side="left")

        icon_lbl = tk.Label(row, text=icon, bg=row_bg, fg=row_fg,
                             font=("Segoe UI", 11), width=2)
        icon_lbl.pack(side="left", padx=(2,4), pady=7)

        text_lbl = tk.Label(row, text=label, bg=row_bg, fg=row_fg,
                             font=self._theme.font("nav_item"), anchor="w")
        text_lbl.pack(side="left", fill="x", expand=True)

        self._item_widgets[iid] = row

        def on_click(e=None, _iid=iid):
            self._navigate(_iid)

        for w in [row, icon_lbl, text_lbl]:
            w.bind("<Button-1>", on_click)

        def on_enter(e=None, _row=row, _iid=iid):
            if _iid != self._active_id:
                _row.configure(bg=self._theme.get("bg_sidebar_hover"))
                for c in _row.winfo_children():
                    try:
                        c.configure(bg=self._theme.get("bg_sidebar_hover"))
                    except Exception:
                        pass

        def on_leave(e=None, _row=row, _iid=iid):
            if _iid != self._active_id:
                _row.configure(bg=self._theme.get("bg_sidebar_item"))
                for c in _row.winfo_children():
                    try:
                        c.configure(bg=self._theme.get("bg_sidebar_item"))
                    except Exception:
                        pass

        row.bind("<Enter>", on_enter)
        row.bind("<Leave>", on_leave)

    def _toggle_group(self, iid: str, arrow_var):
        w = self._item_widgets.get(iid)
        if not w:
            return
        cf = getattr(w, "_children_frame", None)
        if cf is None:
            return
        if iid in self._open_groups:
            self._open_groups.discard(iid)
            cf.pack_forget()
            if arrow_var:
                arrow_var.set("›")
        else:
            self._open_groups.add(iid)
            cf.pack(fill="x")
            if arrow_var:
                arrow_var.set("▾")

    def _navigate(self, iid: str):
        self._active_id = iid
        self._on_navigate(iid)

    def _build_bottom(self):
        sep = tk.Frame(self, bg=self._theme.get("border"), height=1)
        sep.pack(fill="x", side="bottom")
        bottom = tk.Frame(self, bg=self._theme.get("bg_sidebar"), pady=8)
        bottom.pack(side="bottom", fill="x")
        tk.Label(bottom, text=f"v{APP_VERSION}",
                 bg=self._theme.get("bg_sidebar"),
                 fg=self._theme.get("fg_muted"),
                 font=self._theme.font("body_small")).pack()

    def toggle_collapse(self):
        self._collapsed = not self._collapsed
        new_w = self._theme.size("sidebar_collapsed") \
                if self._collapsed else self._theme.size("sidebar_w")
        self.configure(width=new_w)


# ═══════════════════════════════════════════════════════════════
#  BASE VIEW
# ═══════════════════════════════════════════════════════════════

class BaseView(tk.Frame):
    """
    Базовий клас для всіх модулів системи.
    Успадковуйте і реалізуйте: build(), refresh(), on_activate().
    """

    VIEW_ID    = "base"
    VIEW_TITLE = "Базовий вид"
    VIEW_ICON  = "◈"

    def __init__(self, parent, theme: ThemeManager,
                 notifier: "NotificationBanner"):
        super().__init__(parent, bg=theme.get("bg_content"))
        self.theme    = theme
        self.notifier = notifier
        self.build()

    # ── Шаблонний заголовок сторінки ──────────────────────────
    def page_header(self, title: str, subtitle: str = "",
                    actions: Optional[list] = None) -> tk.Frame:
        hdr = tk.Frame(self, bg=self.theme.get("bg_content"),
                       pady=20, padx=24)
        hdr.pack(fill="x")

        left = tk.Frame(hdr, bg=self.theme.get("bg_content"))
        left.pack(side="left", fill="x", expand=True)

        tk.Label(left, text=title,
                 bg=self.theme.get("bg_content"),
                 fg=self.theme.get("fg_primary"),
                 font=self.theme.font("h2")).pack(anchor="w")
        if subtitle:
            tk.Label(left, text=subtitle,
                     bg=self.theme.get("bg_content"),
                     fg=self.theme.get("fg_secondary"),
                     font=self.theme.font("body")).pack(anchor="w", pady=(2,0))

        if actions:
            right = tk.Frame(hdr, bg=self.theme.get("bg_content"))
            right.pack(side="right")
            for act in actions:
                StyledButton(
                    right,
                    act.get("label",""),
                    self.theme,
                    style=act.get("style","primary"),
                    icon=act.get("icon",""),
                    command=act.get("command"),
                ).pack(side="left", padx=(6,0))

        tk.Frame(self, bg=self.theme.get("border"), height=1).pack(fill="x")
        return hdr

    def card(self, parent, padx: int = 16, pady: int = 16,
             border: bool = True) -> tk.Frame:
        outer = tk.Frame(parent,
                         bg=self.theme.get("border_card") if border
                         else self.theme.get("bg_content"),
                         padx=1 if border else 0,
                         pady=1 if border else 0)
        inner = tk.Frame(outer, bg=self.theme.get("bg_card"),
                         padx=padx, pady=pady)
        inner.pack(fill="both", expand=True)
        return inner

    def build(self):
        """Перевизначте для побудови UI модуля."""
        tk.Label(self, text="Порожній вид",
                 bg=self.theme.get("bg_content"),
                 fg=self.theme.get("fg_muted"),
                 font=self.theme.font("h3")).pack(expand=True)

    def refresh(self):
        """Оновлення даних."""
        pass

    def on_activate(self):
        """Викликається при показі вида."""
        self.refresh()


# ═══════════════════════════════════════════════════════════════
#  DASHBOARD VIEW
# ═══════════════════════════════════════════════════════════════

class DashboardView(BaseView):
    VIEW_ID    = "dashboard"
    VIEW_TITLE = "Dashboard"
    VIEW_ICON  = "⊞"

    def build(self):
        self.page_header(
            "Панель керування",
            f"Ласкаво просимо! {datetime.now():%A, %d %B %Y}",
            actions=[
                {"label": "Оновити", "icon": "↻",
                 "style": "secondary",
                 "command": self.refresh},
                {"label": "Новий запис", "icon": "+",
                 "style": "primary",
                 "command": lambda: self.notifier.show(
                     "Форма відкрита", "success")},
            ],
        )

        scroll_canvas = tk.Canvas(
            self, bg=self.theme.get("bg_content"),
            highlightthickness=0)
        scrollbar = ttk.Scrollbar(self, orient="vertical",
                                  command=scroll_canvas.yview)
        scroll_canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        scroll_canvas.pack(fill="both", expand=True)

        container = tk.Frame(scroll_canvas,
                             bg=self.theme.get("bg_content"),
                             padx=24, pady=16)
        scroll_canvas.create_window((0,0), window=container, anchor="nw")
        container.bind(
            "<Configure>",
            lambda e: scroll_canvas.configure(
                scrollregion=scroll_canvas.bbox("all")))

        self._build_kpi_row(container)
        self._build_activity_section(container)
        self._build_quick_actions(container)

    def _build_kpi_row(self, parent):
        section = tk.Label(parent, text="КЛЮЧОВІ ПОКАЗНИКИ",
                           bg=self.theme.get("bg_content"),
                           fg=self.theme.get("fg_muted"),
                           font=self.theme.font("section_head"))
        section.pack(anchor="w", pady=(0,10))

        grid = tk.Frame(parent, bg=self.theme.get("bg_content"))
        grid.pack(fill="x", pady=(0,24))
        grid.columnconfigure((0,1,2,3), weight=1, uniform="kpi")

        kpi_data = [
            ("Клієнти",     "1,284", "+12.4% цього місяця", True,  "◉"),
            ("Угоди",       "₴4.2M", "+8.1% до минулого",   True,  "◇"),
            ("Відкриті ліди","342",  "−3 порівняно з учора", False, "◎"),
            ("Ризики",      "17",    "+2 нові цього тижня",  False, "△"),
        ]
        for col, (title, val, delta, pos, icon) in enumerate(kpi_data):
            card = KpiCard(grid, title, val, delta, pos, icon, self.theme)
            card.grid(row=0, column=col, padx=(0,12) if col < 3 else 0,
                      sticky="nsew")

    def _build_activity_section(self, parent):
        tk.Label(parent, text="ОСТАННЯ АКТИВНІСТЬ",
                 bg=self.theme.get("bg_content"),
                 fg=self.theme.get("fg_muted"),
                 font=self.theme.font("section_head")).pack(anchor="w",
                                                             pady=(0,10))

        row_frame = tk.Frame(parent, bg=self.theme.get("bg_content"))
        row_frame.pack(fill="x", pady=(0,24))
        row_frame.columnconfigure(0, weight=3)
        row_frame.columnconfigure(1, weight=2)
        row_frame.grid_propagate(False)
        row_frame.configure(height=240)

        # Таблиця активностей
        left_card = tk.Frame(row_frame,
                             bg=self.theme.get("border_card"),
                             padx=1, pady=1)
        left_card.grid(row=0, column=0, sticky="nsew", padx=(0,12))

        inner_l = tk.Frame(left_card, bg=self.theme.get("bg_card"),
                           padx=16, pady=12)
        inner_l.pack(fill="both", expand=True)

        tk.Label(inner_l, text="Останні транзакції",
                 bg=self.theme.get("bg_card"),
                 fg=self.theme.get("fg_primary"),
                 font=self.theme.font("h3")).pack(anchor="w", pady=(0,10))

        columns = ("№", "Клієнт", "Сума", "Статус", "Дата")
        tree = ttk.Treeview(inner_l, columns=columns,
                             show="headings", height=5)
        self._style_tree(tree)

        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=80, anchor="w")
        tree.column("Клієнт", width=140)

        sample = [
            ("001", "Альфа Корп",   "₴24,500", "✓ Оплачено",  "02.04.26"),
            ("002", "Бета ТОВ",     "₴8,200",  "⧖ В роботі",  "01.04.26"),
            ("003", "Гамма ФОП",    "₴3,100",  "✓ Оплачено",  "31.03.26"),
            ("004", "Дельта ЛТД",   "₴61,000", "✗ Відмова",   "30.03.26"),
            ("005", "Епсілон Груп", "₴12,700", "⧖ В роботі",  "29.03.26"),
        ]
        for row in sample:
            tree.insert("", "end", values=row)

        tree.pack(fill="both", expand=True)

        # Мінімальна дошка статусів
        right_card = tk.Frame(row_frame,
                              bg=self.theme.get("border_card"),
                              padx=1, pady=1)
        right_card.grid(row=0, column=1, sticky="nsew")
        inner_r = tk.Frame(right_card, bg=self.theme.get("bg_card"),
                           padx=16, pady=12)
        inner_r.pack(fill="both", expand=True)

        tk.Label(inner_r, text="Стан системи",
                 bg=self.theme.get("bg_card"),
                 fg=self.theme.get("fg_primary"),
                 font=self.theme.font("h3")).pack(anchor="w", pady=(0,10))

        statuses = [
            ("CRM модуль",    "● Активний",  "success"),
            ("ERP модуль",    "● Активний",  "success"),
            ("GRC модуль",    "⚠ Попередження", "warning"),
            ("Синхронізація", "● Активна",   "success"),
            ("Резервна копія","● Оновлено",  "success"),
        ]
        for name, status, kind in statuses:
            r = tk.Frame(inner_r, bg=self.theme.get("bg_card"))
            r.pack(fill="x", pady=3)
            tk.Label(r, text=name,
                     bg=self.theme.get("bg_card"),
                     fg=self.theme.get("fg_secondary"),
                     font=self.theme.font("body")).pack(side="left")
            color = {"success": self.theme.get("success"),
                     "warning": self.theme.get("warning"),
                     "danger":  self.theme.get("danger")}.get(kind)
            tk.Label(r, text=status,
                     bg=self.theme.get("bg_card"),
                     fg=color,
                     font=self.theme.font("body_small")).pack(side="right")

    def _build_quick_actions(self, parent):
        tk.Label(parent, text="ШВИДКІ ДІЇ",
                 bg=self.theme.get("bg_content"),
                 fg=self.theme.get("fg_muted"),
                 font=self.theme.font("section_head")).pack(anchor="w",
                                                             pady=(0,10))
        qa_frame = tk.Frame(parent, bg=self.theme.get("bg_content"))
        qa_frame.pack(fill="x")

        actions = [
            ("+ Новий клієнт",    "primary",   lambda: self.notifier.show(
                "Форма клієнта відкрита", "info")),
            ("+ Нова угода",      "primary",   lambda: self.notifier.show(
                "Форма угоди відкрита", "info")),
            ("+ Новий ризик",     "secondary", lambda: self.notifier.show(
                "Форма ризику відкрита", "info")),
            ("⬇ Звіт PDF",       "secondary", lambda: self.notifier.show(
                "Звіт генерується…", "success")),
            ("⚙ Налаштування",   "ghost",     lambda: self.notifier.show(
                "Перейдіть до Налаштувань", "info")),
        ]
        for label, style, cmd in actions:
            StyledButton(qa_frame, label, self.theme,
                         style=style, command=cmd).pack(
                side="left", padx=(0,10), ipady=4)

    def _style_tree(self, tree):
        style = ttk.Style()
        t = self.theme
        style.theme_use("default")
        style.configure("Treeview",
                         background=t.get("bg_table_row"),
                         foreground=t.get("fg_table_cell"),
                         fieldbackground=t.get("bg_table_row"),
                         rowheight=30,
                         font=t.font("table_cell"),
                         borderwidth=0)
        style.configure("Treeview.Heading",
                         background=t.get("bg_table_header"),
                         foreground=t.get("fg_table_header"),
                         font=t.font("table_head"),
                         relief="flat")
        style.map("Treeview",
                  background=[("selected", t.get("accent"))],
                  foreground=[("selected", "#FFFFFF")])


# ═══════════════════════════════════════════════════════════════
#  LIST TABLE VIEW
# ═══════════════════════════════════════════════════════════════

class ListTableView(BaseView):
    """
    Шаблон списку / таблиці.
    Параметризуйте через VIEW_ID, VIEW_TITLE, COLUMNS, sample_data().
    """
    VIEW_ID    = "list_table"
    VIEW_TITLE = "Список"
    VIEW_ICON  = "▤"

    COLUMNS = ("ID", "Назва", "Категорія", "Статус", "Дата змін")

    def build(self):
        self.page_header(
            self.VIEW_TITLE,
            "Перегляд та управління записами",
            actions=[
                {"label": "Фільтр",    "icon": "⊟",
                 "style": "secondary",
                 "command": lambda: self.notifier.show("Фільтри", "info")},
                {"label": "Експорт",   "icon": "⬇",
                 "style": "secondary",
                 "command": lambda: self.notifier.show(
                     "Експортується…", "success")},
                {"label": "Додати",    "icon": "+",
                 "style": "primary",
                 "command": self._open_form},
            ],
        )

        toolbar = tk.Frame(self, bg=self.theme.get("bg_content"),
                           padx=24, pady=10)
        toolbar.pack(fill="x")

        search_wrap = tk.Frame(toolbar,
                               bg=self.theme.get("bg_input"),
                               highlightbackground=self.theme.get("border"),
                               highlightthickness=1)
        search_wrap.pack(side="left", ipady=2)
        tk.Label(search_wrap, text="⌕",
                 bg=self.theme.get("bg_input"),
                 fg=self.theme.get("fg_muted"),
                 font=("Segoe UI", 11)).pack(side="left", padx=(6,2))
        self._search_var = tk.StringVar()
        tk.Entry(search_wrap, textvariable=self._search_var,
                 bg=self.theme.get("bg_input"),
                 fg=self.theme.get("fg_primary"),
                 insertbackground=self.theme.get("accent"),
                 relief="flat", bd=4, width=30,
                 font=self.theme.font("body")).pack(side="left")

        tk.Label(toolbar,
                 text=f"Всього: {len(self._sample_data())} записів",
                 bg=self.theme.get("bg_content"),
                 fg=self.theme.get("fg_muted"),
                 font=self.theme.font("body_small")).pack(
            side="right", padx=8)

        table_wrap = tk.Frame(self,
                              bg=self.theme.get("border_card"),
                              padx=1, pady=1)
        table_wrap.pack(fill="both", expand=True, padx=24, pady=(0,16))

        inner = tk.Frame(table_wrap, bg=self.theme.get("bg_card"))
        inner.pack(fill="both", expand=True)

        self._tree = ttk.Treeview(inner, columns=self.COLUMNS,
                                   show="headings")
        vsb = ttk.Scrollbar(inner, orient="vertical",
                             command=self._tree.yview)
        hsb = ttk.Scrollbar(inner, orient="horizontal",
                             command=self._tree.xview)
        self._tree.configure(yscrollcommand=vsb.set,
                              xscrollcommand=hsb.set)

        self._style_tree()

        hsb.pack(side="bottom", fill="x")
        vsb.pack(side="right",  fill="y")
        self._tree.pack(fill="both", expand=True)

        for col in self.COLUMNS:
            self._tree.heading(col, text=col,
                                command=lambda c=col: self._sort(c))
            self._tree.column(col, width=140, minwidth=60)

        self._tree.bind("<Double-1>", self._on_double_click)
        self._tree.bind("<Button-3>", self._on_right_click)

        self.refresh()

    def _sample_data(self) -> list:
        statuses = ["Активний", "На паузі", "Завершено", "Нові"]
        cats     = ["CRM", "ERP", "GRC", "HR", "Фінанси"]
        rows = []
        import random
        random.seed(42)
        for i in range(1, 31):
            rows.append((
                f"{i:03d}",
                f"Запис {i:03d}",
                random.choice(cats),
                random.choice(statuses),
                f"{random.randint(1,28):02d}.0{random.randint(1,4)}.2026",
            ))
        return rows

    def refresh(self):
        for row in self._tree.get_children():
            self._tree.delete(row)
        for i, row in enumerate(self._sample_data()):
            tag = "alt" if i % 2 else "norm"
            self._tree.insert("", "end", values=row, tags=(tag,))
        t = self.theme
        self._tree.tag_configure("norm", background=t.get("bg_table_row"))
        self._tree.tag_configure("alt",  background=t.get("bg_table_alt"))

    def _style_tree(self):
        style = ttk.Style()
        t = self.theme
        style.theme_use("default")
        style.configure("Treeview",
                         background=t.get("bg_table_row"),
                         foreground=t.get("fg_table_cell"),
                         fieldbackground=t.get("bg_table_row"),
                         rowheight=32,
                         font=t.font("table_cell"))
        style.configure("Treeview.Heading",
                         background=t.get("bg_table_header"),
                         foreground=t.get("fg_table_header"),
                         font=t.font("table_head"),
                         relief="flat",
                         padding=(8,6))
        style.map("Treeview",
                  background=[("selected", t.get("accent"))],
                  foreground=[("selected", "#FFFFFF")])

    def _sort(self, col: str):
        self.notifier.show(f"Сортування за: {col}", "info")

    def _on_double_click(self, event):
        item = self._tree.selection()
        if item:
            vals = self._tree.item(item[0], "values")
            self.notifier.show(f"Редагування: {vals[1]}", "info")

    def _on_right_click(self, event):
        iid = self._tree.identify_row(event.y)
        if not iid:
            return
        self._tree.selection_set(iid)
        menu = tk.Menu(self, tearoff=0,
                        bg=self.theme.get("bg_card"),
                        fg=self.theme.get("fg_primary"),
                        relief="flat", bd=0,
                        font=self.theme.font("body"))
        menu.add_command(label="✏ Редагувати",
                          command=self._on_double_click)
        menu.add_command(label="⬇ Деталі",
                          command=lambda: self.notifier.show(
                              "Деталі відкрито", "info"))
        menu.add_separator()
        menu.add_command(label="✖ Видалити",
                          command=lambda: self.notifier.show(
                              "Запис видалено", "error"))
        menu.tk_popup(event.x_root, event.y_root)

    def _open_form(self):
        FormDialog(self.winfo_toplevel(), self.theme, self.notifier)


# ═══════════════════════════════════════════════════════════════
#  FORM DIALOG
# ═══════════════════════════════════════════════════════════════

class FormDialog(ModalDialog):
    """Приклад форми в модальному вікні."""

    def __init__(self, parent, theme, notifier):
        self._notifier = notifier
        super().__init__(parent, "Новий запис", theme, 520, 420)

    def body(self, master):
        tk.Label(master, text="Заповніть деталі запису",
                 bg=self.theme.get("bg_modal"),
                 fg=self.theme.get("fg_secondary"),
                 font=self.theme.font("body")).pack(anchor="w", pady=(0,14))

        f1 = StyledEntry(master, "Назва *", self.theme,
                          placeholder="Введіть назву")
        f1.pack(fill="x", pady=(0,10))

        row2 = tk.Frame(master, bg=self.theme.get("bg_modal"))
        row2.pack(fill="x", pady=(0,10))
        f2 = StyledEntry(row2, "Категорія", self.theme)
        f2.pack(side="left", fill="x", expand=True, padx=(0,8))
        f3 = StyledEntry(row2, "Статус", self.theme)
        f3.pack(side="left", fill="x", expand=True)

        f4 = StyledEntry(master, "Опис", self.theme,
                          placeholder="Додатковий опис…")
        f4.pack(fill="x")

    def on_ok(self):
        self._notifier.show("Запис збережено!", "success")


# ═══════════════════════════════════════════════════════════════
#  SETTINGS VIEW
# ═══════════════════════════════════════════════════════════════

class SettingsView(BaseView):
    VIEW_ID    = "settings"
    VIEW_TITLE = "Налаштування"
    VIEW_ICON  = "⚙"

    def build(self):
        self.page_header("Налаштування", "Конфігурація системи")

        canvas = tk.Canvas(self, bg=self.theme.get("bg_content"),
                           highlightthickness=0)
        sb = ttk.Scrollbar(self, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        canvas.pack(fill="both", expand=True)

        cont = tk.Frame(canvas, bg=self.theme.get("bg_content"),
                        padx=32, pady=20)
        canvas.create_window((0,0), window=cont, anchor="nw")
        cont.bind("<Configure>",
                   lambda e: canvas.configure(
                       scrollregion=canvas.bbox("all")))

        self._build_appearance(cont)
        self._build_general(cont)
        self._build_system(cont)

    def _section_header(self, parent, title: str):
        tk.Frame(parent, bg=self.theme.get("border"),
                 height=1).pack(fill="x", pady=(20,12))
        tk.Label(parent, text=title,
                 bg=self.theme.get("bg_content"),
                 fg=self.theme.get("fg_primary"),
                 font=self.theme.font("h3")).pack(anchor="w")

    def _build_appearance(self, parent):
        self._section_header(parent, "Зовнішній вигляд")

        row = tk.Frame(parent, bg=self.theme.get("bg_content"), pady=12)
        row.pack(fill="x")

        tk.Label(row, text="Тема інтерфейсу",
                 bg=self.theme.get("bg_content"),
                 fg=self.theme.get("fg_secondary"),
                 font=self.theme.font("body_bold"),
                 width=28, anchor="w").pack(side="left")

        theme_var = tk.StringVar(value=self.theme.name)
        for val, lbl in [("light","☀ Світла"), ("dark","◑ Темна")]:
            rb = tk.Radiobutton(
                row, text=lbl, variable=theme_var, value=val,
                bg=self.theme.get("bg_content"),
                fg=self.theme.get("fg_primary"),
                selectcolor=self.theme.get("bg_card"),
                activebackground=self.theme.get("bg_content"),
                font=self.theme.font("body"),
                command=lambda v=val: self.theme.switch(v),
            )
            rb.pack(side="left", padx=12)

    def _build_general(self, parent):
        self._section_header(parent, "Загальні")
        fields = [
            ("Назва організації", APP_NAME),
            ("Версія системи",   APP_VERSION),
            ("Мова інтерфейсу",  "Українська"),
            ("Часовий пояс",     "Europe/Kyiv (UTC+3)"),
        ]
        for label, value in fields:
            row = tk.Frame(parent, bg=self.theme.get("bg_content"), pady=6)
            row.pack(fill="x")
            tk.Label(row, text=label,
                     bg=self.theme.get("bg_content"),
                     fg=self.theme.get("fg_secondary"),
                     font=self.theme.font("body"),
                     width=28, anchor="w").pack(side="left")
            tk.Label(row, text=value,
                     bg=self.theme.get("bg_content"),
                     fg=self.theme.get("fg_primary"),
                     font=self.theme.font("body_bold")).pack(side="left")

    def _build_system(self, parent):
        self._section_header(parent, "Система")
        info = [
            ("Python",    sys.version.split()[0]),
            ("Tkinter",   tk.TkVersion),
            ("ОС",        platform.system() + " " + platform.release()),
            ("Процесор",  platform.processor() or "—"),
        ]
        for label, value in info:
            row = tk.Frame(parent, bg=self.theme.get("bg_content"), pady=6)
            row.pack(fill="x")
            tk.Label(row, text=label,
                     bg=self.theme.get("bg_content"),
                     fg=self.theme.get("fg_secondary"),
                     font=self.theme.font("body"),
                     width=28, anchor="w").pack(side="left")
            tk.Label(row, text=str(value),
                     bg=self.theme.get("bg_content"),
                     fg=self.theme.get("fg_primary"),
                     font=self.theme.font("mono")).pack(side="left")

        btn_row = tk.Frame(parent, bg=self.theme.get("bg_content"), pady=16)
        btn_row.pack(fill="x")
        StyledButton(btn_row, "Зберегти зміни", self.theme,
                     style="primary",
                     command=lambda: self.notifier.show(
                         "Налаштування збережено!", "success")).pack(
            side="left", ipady=4)
        StyledButton(btn_row, "Скинути", self.theme,
                     style="danger",
                     command=lambda: self.notifier.show(
                         "Скинуто до стандартних", "warning")).pack(
            side="left", padx=10, ipady=4)


# ═══════════════════════════════════════════════════════════════
#  CONTENT AREA
# ═══════════════════════════════════════════════════════════════

class ContentArea(tk.Frame):
    """
    Керує показом / приховуванням View-фреймів.
    Реєстрація нових модулів: register(id, ViewClass)
    """

    # Відповідність ID → клас вида
    VIEW_REGISTRY: Dict[str, type] = {
        "DashboardView": DashboardView,
        "ListTableView": ListTableView,
        "SettingsView":  SettingsView,
    }

    def __init__(self, parent, theme: ThemeManager,
                 notifier: "NotificationBanner"):
        super().__init__(parent, bg=theme.get("bg_content"))
        self.theme    = theme
        self.notifier = notifier
        self._views:  Dict[str, BaseView] = {}
        self._current: Optional[str] = None

        # Показуємо стартову сторінку
        self.show("dashboard", "DashboardView")

    def register(self, view_id: str, cls: type):
        """Реєстрація нового View-класу."""
        self.VIEW_REGISTRY[cls.__name__] = cls

    def show(self, view_id: str, view_class_name: str = "ListTableView"):
        # Приховуємо поточний вид
        if self._current and self._current in self._views:
            self._views[self._current].pack_forget()

        # Якщо ще не побудовано — будуємо
        if view_id not in self._views:
            cls = self.VIEW_REGISTRY.get(view_class_name, ListTableView)
            v = cls(self, self.theme, self.notifier)
            v.VIEW_ID    = view_id
            v.VIEW_TITLE = view_class_name.replace("View","")
            self._views[view_id] = v

        self._views[view_id].pack(fill="both", expand=True)
        self._views[view_id].on_activate()
        self._current = view_id


# ═══════════════════════════════════════════════════════════════
#  STATUS BAR
# ═══════════════════════════════════════════════════════════════

class StatusBar(tk.Frame):
    def __init__(self, parent, theme: ThemeManager):
        super().__init__(parent,
                         bg=theme.get("bg_statusbar"),
                         height=theme.size("statusbar_h"))
        self.pack_propagate(False)
        # Роздільник зберігаємо як атрибут — App пакує його
        # одразу ПЕРЕД self.pack(), коли self вже в layout.
        self._top_sep = tk.Frame(parent, bg=theme.get("border"), height=1)

        self._status_var = tk.StringVar(value="Готово")
        self._time_var   = tk.StringVar()

        tk.Label(self,
                 textvariable=self._status_var,
                 bg=theme.get("bg_statusbar"),
                 fg=theme.get("fg_statusbar"),
                 font=theme.font("status"),
                 padx=12).pack(side="left", fill="y")

        # Індикатор з'єднання
        tk.Label(self,
                 text="● З'єднання активне",
                 bg=theme.get("bg_statusbar"),
                 fg=theme.get("success"),
                 font=theme.font("status")).pack(side="left", padx=20)

        tk.Label(self,
                 textvariable=self._time_var,
                 bg=theme.get("bg_statusbar"),
                 fg=theme.get("fg_statusbar"),
                 font=theme.font("status"),
                 padx=12).pack(side="right", fill="y")

        self._tick()

    def set_status(self, msg: str):
        self._status_var.set(msg)

    def _tick(self):
        self._time_var.set(datetime.now().strftime("%H:%M:%S  %d.%m.%Y"))
        self.after(1000, self._tick)


# ═══════════════════════════════════════════════════════════════
#  APPLICATION
# ═══════════════════════════════════════════════════════════════

class App(tk.Tk):
    """
    Кореневе вікно програми.
    Ініціалізує всі компоненти та пов'язує їх разом.
    """

    def __init__(self):
        super().__init__()

        self.title(f"{APP_NAME} — {APP_TAGLINE}  v{APP_VERSION}")
        self.geometry("1280x780")
        self.minsize(900, 600)

        # ── Іконка вікна ─────────────────────────────────────
        # Якщо є .ico файл: self.iconbitmap("icon.ico")
        # Або растрова іконка:
        try:
            icon = tk.PhotoImage(width=32, height=32)
            self.iconphoto(True, icon)
        except Exception:
            pass

        # ── Тема ─────────────────────────────────────────────
        self.theme = ThemeManager("light")
        self.theme.subscribe(self._apply_theme)

        # ── Стилі ttk ────────────────────────────────────────
        self._configure_ttk()

        # ── Фон ──────────────────────────────────────────────
        self.configure(bg=self.theme.get("bg_app"))

        # ── Компоненти ────────────────────────────────────────
        self.notifier = NotificationBanner(self)

        # TopBar
        self._topbar = TopNavBar(
            self, self.theme,
            on_toggle_sidebar = self._toggle_sidebar,
            on_toggle_theme   = self.theme.switch,
            on_search         = self._on_search,
            notifier          = self.notifier,
        )
        self._topbar.pack(fill="x", side="top")
        self._topbar._bottom_sep.pack(fill="x", side="top")

        # Головний контейнер (Sidebar + Content)
        self._main = tk.Frame(self, bg=self.theme.get("bg_app"))
        self._main.pack(fill="both", expand=True, side="top")

        # Sidebar
        self._sidebar = SideBar(
            self._main, self.theme,
            on_navigate = self._navigate,
            menu        = MENU_STRUCTURE,
        )
        self._sidebar.pack(side="left", fill="y")

        # Вертикальний роздільник
        tk.Frame(self._main,
                 bg=self.theme.get("border"), width=1).pack(
            side="left", fill="y")

        # Content
        self._content = ContentArea(self._main, self.theme, self.notifier)
        self._content.pack(side="left", fill="both", expand=True)

        # StatusBar
        self._statusbar = StatusBar(self, self.theme)
        self._statusbar._top_sep.pack(fill="x", side="bottom")
        self._statusbar.pack(fill="x", side="bottom")

        # ── Гарячі клавіші ────────────────────────────────────
        self.bind("<Control-d>", lambda _: self._navigate("dashboard"))
        self.bind("<Control-q>", lambda _: self.quit())
        self.bind("<F5>",        lambda _: self._refresh_current())
        self.bind("<Control-t>", lambda _: self.theme.switch())

        # Привітальне сповіщення
        self.after(800, lambda: self.notifier.show(
            f"Ласкаво просимо до {APP_NAME}!", "success"))

    def _configure_ttk(self):
        style = ttk.Style(self)
        style.theme_use("default")
        t = self.theme
        style.configure("TScrollbar",
                          background=t.get("bg_scrollbar"),
                          troughcolor=t.get("bg_scrollbar"),
                          arrowcolor=t.get("fg_muted"),
                          relief="flat")
        style.map("TScrollbar",
                   background=[("active", t.get("bg_scrollbar_thumb"))])

    def _toggle_sidebar(self):
        self._sidebar.toggle_collapse()

    def _navigate(self, view_id: str):
        cls_name = self._find_view_class(view_id, MENU_STRUCTURE)
        self._content.show(view_id, cls_name)
        self._statusbar.set_status(
            f"Відкрито: {view_id.replace('_',' ').title()}")

    def _find_view_class(self, view_id: str, menu: list) -> str:
        for item in menu:
            if item.get("id") == view_id:
                return item.get("view", "ListTableView")
            for child in item.get("children", []):
                if child.get("id") == view_id:
                    return child.get("view", "ListTableView")
        return "ListTableView"

    def _on_search(self, query: str):
        if query and query != "Пошук…":
            self.notifier.show(f"Пошук: «{query}»", "info")
            self._statusbar.set_status(f"Пошук: {query}")

    def _refresh_current(self):
        self.notifier.show("Оновлення даних…", "info")

    def _apply_theme(self):
        """
        Перебудова UI в тому ж самому вікні — без destroy/mainloop.
        Знищуємо всі дочірні віджети та будуємо їх заново.
        """
        # Запам'ятовуємо поточний активний розділ (якщо є)
        current_view_id = getattr(self._content, "_current", "dashboard")

        # Знищуємо всі дочірні віджети кореневого вікна
        for widget in self.winfo_children():
            widget.destroy()

        # Перебудовуємо фон
        self.configure(bg=self.theme.get("bg_app"))

        # Перебудовуємо ttk-стилі
        self._configure_ttk()

        # Новий notifier (старий вже знищено разом з Toplevel-ами)
        self.notifier = NotificationBanner(self)

        # TopBar
        self._topbar = TopNavBar(
            self, self.theme,
            on_toggle_sidebar=self._toggle_sidebar,
            on_toggle_theme=self.theme.switch,
            on_search=self._on_search,
            notifier=self.notifier,
        )
        self._topbar.pack(fill="x", side="top")
        self._topbar._bottom_sep.pack(fill="x", side="top")

        # Головний контейнер
        self._main = tk.Frame(self, bg=self.theme.get("bg_app"))
        self._main.pack(fill="both", expand=True, side="top")

        # Sidebar
        self._sidebar = SideBar(
            self._main, self.theme,
            on_navigate=self._navigate,
            menu=MENU_STRUCTURE,
        )
        self._sidebar.pack(side="left", fill="y")

        # Вертикальний роздільник
        tk.Frame(self._main,
                 bg=self.theme.get("border"), width=1).pack(
            side="left", fill="y")

        # Content — відновлюємо останній активний вид
        self._content = ContentArea(self._main, self.theme, self.notifier)
        self._content.pack(side="left", fill="both", expand=True)

        if current_view_id and current_view_id != "dashboard":
            cls_name = self._find_view_class(current_view_id, MENU_STRUCTURE)
            self._content.show(current_view_id, cls_name)

        # StatusBar
        self._statusbar = StatusBar(self, self.theme)
        self._statusbar._top_sep.pack(fill="x", side="bottom")
        self._statusbar.pack(fill="x", side="bottom")

        # Повідомлення про зміну теми
        theme_label = "Темна" if self.theme.name == "dark" else "Світла"
        self.after(100, lambda: self.notifier.show(
            f"Тема змінена: {theme_label}", "info"))

    def run(self):
        self.mainloop()


# ═══════════════════════════════════════════════════════════════
#  ТОЧКА ВХОДУ
# ═══════════════════════════════════════════════════════════════

if __name__ == "__main__":
    # Windows HiDPI
    if platform.system() == "Windows":
        try:
            from ctypes import windll
            windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass

    app = App()
    app.run()
