"""
╔══════════════════════════════════════════════════════════════════════╗
║  NEXUS  —  Universal CRM · ERP · GRC  Desktop Template  v2.0        ║
║  Python 3.10+  |  Tkinter  |  Windows-first  |  Precision Dark UI   ║
╠══════════════════════════════════════════════════════════════════════╣
║  Дизайн: Precision Dark — Bloomberg × Linear × Figma                ║
║  Теми:   DARK (за замовчуванням)  |  LIGHT                          ║
║  Архітектура: Shell → Workspace → Module (BaseModule)               ║
╚══════════════════════════════════════════════════════════════════════╝

ШВИДКИЙ СТАРТ
─────────────
  python nexus_v2.py

ЯК ДОДАТИ МОДУЛЬ
─────────────────
  1. class MyModule(BaseModule):
         MODULE_ID    = "my_id"
         MODULE_TITLE = "Назва"
         MODULE_ICON  = "◈"
         def build(self): ...

  2. REGISTRY["my_id"] = MyModule

  3. NAVIGATION — додати запис у список

ГАРЯЧІ КЛАВІШІ
──────────────
  Ctrl+1..9   → перехід до модуля
  Ctrl+T      → перемикач теми
  Ctrl+N      → нова форма
  Ctrl+F      → фокус пошуку
  Ctrl+W      → закрити вкладку
  Escape      → закрити діалог / зняти виділення
  F5          → оновити
"""

from __future__ import annotations
import tkinter as tk
from tkinter import ttk
import platform, sys, time, random
from datetime import datetime, timedelta
from typing import Optional, Callable, Any


# ══════════════════════════════════════════════════════
#  КОНФІГ
# ══════════════════════════════════════════════════════

APP = dict(
    name      = "NEXUS",
    subtitle  = "Business Intelligence Platform",
    version   = "2.0.0",
    org       = "Acme Corp",
    user      = "Олексій М.",
    user_role = "Адміністратор",
    user_init = "ОМ",
    default_theme = "dark",
)

# ══════════════════════════════════════════════════════
#  ТЕМИ
# ══════════════════════════════════════════════════════

class T:
    """
    Namespace кольорів — завжди звертайтесь через T.c["key"].
    Після виклику T.apply(name) усі компоненти що підписані
    через T.on_change отримають callback.
    """

    PALETTES = {

      "dark": {
        # ── base ───────────────────────────────
        "bg0":          "#080C14",   # найглибший фон
        "bg1":          "#0D1117",   # фон вікна / sidebar
        "bg2":          "#131B27",   # фон карток
        "bg3":          "#1A2335",   # hover / підсвітка
        "bg4":          "#212D40",   # активний елемент (не акцент)
        "bg_input":     "#0D1117",
        "bg_topbar":    "#0D1117",
        "bg_modal":     "#131B27",
        "bg_overlay":   "#00000099",

        # ── border ─────────────────────────────
        "b0":           "#1A2335",   # тонка межа
        "b1":           "#243045",   # середня межа
        "b2":           "#2E3D55",   # помітна межа
        "b_focus":      "#00D4FF",   # focus ring

        # ── text ───────────────────────────────
        "t0":           "#F0F6FF",   # заголовки / значення
        "t1":           "#8899BB",   # основний текст
        "t2":           "#4A5F80",   # мʼютований текст
        "t3":           "#2A3850",   # найтихіший текст
        "t_inv":        "#080C14",   # текст на акценті

        # ── accent ─────────────────────────────
        "a":            "#00D4FF",   # основний акцент — cyan
        "a_dim":        "#003D4D",   # дуже тихий акцент
        "a_glow":       "#00D4FF33", # glow-ефект
        "a2":           "#7C6EFF",   # другорядний акцент — violet
        "a2_dim":       "#1A1640",

        # ── semantic ───────────────────────────
        "green":        "#00E5A0",
        "green_dim":    "#002E22",
        "red":          "#FF4466",
        "red_dim":      "#2E0A14",
        "amber":        "#FFB020",
        "amber_dim":    "#2E1E00",
        "blue":         "#4499FF",
        "blue_dim":     "#0A1F3A",

        # ── chart bars (5 кольорів) ────────────
        "chart": ["#00D4FF","#7C6EFF","#00E5A0","#FFB020","#FF4466"],

        # ── scrollbar ──────────────────────────
        "sb_bg":        "#0D1117",
        "sb_thumb":     "#1A2335",
        "sb_active":    "#243045",
      },

      "light": {
        "bg0":          "#EEF2F8",
        "bg1":          "#F8FAFD",
        "bg2":          "#FFFFFF",
        "bg3":          "#EEF2F8",
        "bg4":          "#E2E8F5",
        "bg_input":     "#F8FAFD",
        "bg_topbar":    "#FFFFFF",
        "bg_modal":     "#FFFFFF",
        "bg_overlay":   "#00000044",

        "b0":           "#E2E8F5",
        "b1":           "#D0D9EE",
        "b2":           "#B8C5DE",
        "b_focus":      "#0077AA",

        "t0":           "#0A1628",
        "t1":           "#2A3F60",
        "t2":           "#6B7FA0",
        "t3":           "#A0B0CC",
        "t_inv":        "#FFFFFF",

        "a":            "#0077AA",
        "a_dim":        "#E0F4FF",
        "a_glow":       "#0077AA22",
        "a2":           "#5B4FCC",
        "a2_dim":       "#EEF0FF",

        "green":        "#008A5C",
        "green_dim":    "#E6FBF3",
        "red":          "#CC1133",
        "red_dim":      "#FFF0F2",
        "amber":        "#A05500",
        "amber_dim":    "#FFF8E6",
        "blue":         "#0055CC",
        "blue_dim":     "#EEF4FF",

        "chart": ["#0077AA","#5B4FCC","#008A5C","#A05500","#CC1133"],

        "sb_bg":        "#F0F4FA",
        "sb_thumb":     "#D0D9EE",
        "sb_active":    "#B8C5DE",
      },
    }

    FONTS = {
        "ui":      ("Segoe UI",  10, "normal"),
        "ui_b":    ("Segoe UI",  10, "bold"),
        "ui_sm":   ("Segoe UI",   9, "normal"),
        "ui_xs":   ("Segoe UI",   8, "normal"),
        "cap":     ("Segoe UI",   8, "bold"),    # CAPS labels
        "h1":      ("Segoe UI",  20, "bold"),
        "h2":      ("Segoe UI",  15, "bold"),
        "h3":      ("Segoe UI",  12, "bold"),
        "h4":      ("Segoe UI",  10, "bold"),
        "num":     ("Segoe UI",  24, "bold"),    # великі числа
        "num_sm":  ("Segoe UI",  16, "bold"),
        "mono":    ("Consolas",   9, "normal"),
        "logo":    ("Segoe UI",  13, "bold"),
        "tab":     ("Segoe UI",  10, "normal"),
        "tab_a":   ("Segoe UI",  10, "bold"),
    }

    SZ = {
        "topbar_h":  48,
        "sidebar_w": 210,
        "sidebar_c": 44,   # collapsed
        "statusbar_h": 24,
        "row_h":     34,
        "row_h_sm":  28,
        "card_r":     4,
        "input_h":   32,
        "btn_h":     30,
        "btn_h_sm":  26,
    }

    _name: str = "dark"
    c:     dict = {}
    _subs: list = []

    @classmethod
    def apply(cls, name: str):
        cls._name = name
        cls.c = cls.PALETTES[name]
        for cb in cls._subs:
            try: cb()
            except Exception: pass

    @classmethod
    def sub(cls, cb: Callable):
        cls._subs.append(cb)

    @classmethod
    def unsub(cls, cb: Callable):
        if cb in cls._subs: cls._subs.remove(cb)

    @classmethod
    def toggle(cls):
        cls.apply("light" if cls._name == "dark" else "dark")

    @classmethod
    def f(cls, key: str) -> tuple:
        return cls.FONTS.get(key, cls.FONTS["ui"])

    @classmethod
    def sz(cls, key: str) -> int:
        return cls.SZ.get(key, 0)

# ── ініціалізація ──────────────────────────────────────
T.apply(APP["default_theme"])


# ══════════════════════════════════════════════════════
#  УТИЛІТИ
# ══════════════════════════════════════════════════════

def px(n: int) -> int:
    """HiDPI-заглушка, можна розширити."""
    return n

def clamp(v, lo, hi):
    return max(lo, min(hi, v))

def humanize(n: int) -> str:
    if n >= 1_000_000: return f"{n/1_000_000:.1f}M"
    if n >= 1_000:     return f"{n/1_000:.1f}K"
    return str(n)


# ══════════════════════════════════════════════════════
#  БАЗОВІ ВІДЖЕТИ
# ══════════════════════════════════════════════════════

class Frame(tk.Frame):
    """tk.Frame з bg за замовчуванням = bg2."""
    def __init__(self, parent, bg: str | None = None, **kw):
        super().__init__(parent, bg=bg or T.c["bg2"],
                         highlightthickness=0, bd=0, **kw)


class Label(tk.Label):
    def __init__(self, parent, text="", style="ui",
                 color: str | None = None, bg: str | None = None, **kw):
        c = color or T.c["t1"]
        b = bg    or (parent.cget("bg") if hasattr(parent,"cget") else T.c["bg2"])
        super().__init__(parent, text=text, font=T.f(style),
                         fg=c, bg=b, **kw)


class Sep(tk.Frame):
    """Горизонтальний або вертикальний роздільник."""
    def __init__(self, parent, orient="h", **kw):
        if orient == "h":
            super().__init__(parent, bg=T.c["b0"], height=1, **kw)
        else:
            super().__init__(parent, bg=T.c["b0"], width=1, **kw)


# ══════════════════════════════════════════════════════
#  TOAST / NOTIFICATION
# ══════════════════════════════════════════════════════

class Toast:
    _queue: list = []
    _busy:  bool = False
    _root:  tk.Tk | None = None

    ICONS = {"ok": "✔", "err": "✖", "warn": "⚠", "info": "ℹ"}
    COLORS = {
        "ok":   ("green", "green_dim"),
        "err":  ("red",   "red_dim"),
        "warn": ("amber", "amber_dim"),
        "info": ("a",     "a_dim"),
    }

    @classmethod
    def init(cls, root: tk.Tk):
        cls._root = root

    @classmethod
    def show(cls, msg: str, kind: str = "info", ms: int = 3200):
        cls._queue.append((msg, kind, ms))
        if not cls._busy:
            cls._next()

    @classmethod
    def _next(cls):
        if not cls._queue or not cls._root:
            cls._busy = False
            return
        cls._busy = True
        msg, kind, ms = cls._queue.pop(0)
        fg_key, bg_key = cls.COLORS.get(kind, cls.COLORS["info"])
        fg = T.c[fg_key]
        bg = T.c[bg_key]

        w = tk.Toplevel(cls._root)
        w.overrideredirect(True)
        w.attributes("-topmost", True)
        w.configure(bg=fg)

        inner = tk.Frame(w, bg=bg, padx=14, pady=9)
        inner.pack(fill="both", expand=True, padx=1, pady=1)

        tk.Label(inner, text=cls.ICONS.get(kind,"ℹ"),
                 bg=bg, fg=fg, font=("Segoe UI",13,"bold")).pack(side="left",padx=(0,8))
        tk.Label(inner, text=msg, bg=bg, fg=T.c["t0"],
                 font=T.f("ui"), wraplength=280).pack(side="left")

        cls._root.update_idletasks()
        sw = cls._root.winfo_screenwidth()
        sh = cls._root.winfo_screenheight()
        w.update_idletasks()
        ww = w.winfo_reqwidth()
        wh = w.winfo_reqheight()
        w.geometry(f"{ww}x{wh}+{sw-ww-20}+{sh-wh-56}")

        w.after(ms, lambda: cls._dismiss(w))

    @classmethod
    def _dismiss(cls, w):
        try: w.destroy()
        except: pass
        cls._root.after(150, cls._next)


# ══════════════════════════════════════════════════════
#  TOOLTIP
# ══════════════════════════════════════════════════════

class Tip:
    def __init__(self, widget: tk.Widget, text: str):
        self._w = widget
        self._t = text
        self._win: tk.Toplevel | None = None
        widget.bind("<Enter>", self._show, add="+")
        widget.bind("<Leave>", self._hide, add="+")

    def _show(self, e=None):
        if not self._t: return
        x = self._w.winfo_rootx() + self._w.winfo_width() + 6
        y = self._w.winfo_rooty() + 4
        self._win = tk.Toplevel(self._w)
        self._win.overrideredirect(True)
        self._win.attributes("-topmost", True)
        self._win.configure(bg=T.c["b2"])
        inner = tk.Frame(self._win, bg=T.c["bg3"], padx=8, pady=4)
        inner.pack(fill="both", expand=True, padx=1, pady=1)
        tk.Label(inner, text=self._t, bg=T.c["bg3"],
                 fg=T.c["t0"], font=T.f("ui_sm")).pack()
        self._win.geometry(f"+{x}+{y}")

    def _hide(self, e=None):
        if self._win:
            try: self._win.destroy()
            except: pass
            self._win = None


# ══════════════════════════════════════════════════════
#  STYLED BUTTON
# ══════════════════════════════════════════════════════

class Btn(tk.Frame):
    """
    Кнопка з чіткими станами hover/press/disabled.
    style: "primary" | "ghost" | "danger" | "success" | "outline"
    size:  "md" | "sm" | "lg"
    """

    _STYLES = {
        "primary": lambda: (T.c["a"],     T.c["t_inv"], T.c["a_dim"],   T.c["a"]),
        "ghost":   lambda: (T.c["bg3"],   T.c["t1"],    T.c["bg4"],     T.c["b1"]),
        "danger":  lambda: (T.c["red"],   T.c["t_inv"], T.c["red_dim"], T.c["red"]),
        "success": lambda: (T.c["green"], T.c["t_inv"], T.c["green_dim"],T.c["green"]),
        "outline": lambda: (T.c["bg2"],   T.c["t1"],    T.c["bg3"],     T.c["b1"]),
        "accent2": lambda: (T.c["a2"],    T.c["t_inv"], T.c["a2_dim"],  T.c["a2"]),
    }
    _PAD = {"lg":(14,7), "md":(11,5), "sm":(8,3)}

    def __init__(self, parent, text: str = "", style: str = "ghost",
                 icon: str = "", size: str = "md",
                 command: Callable | None = None,
                 tip: str = "", **kw):
        super().__init__(parent, bg=parent.cget("bg"),
                         cursor="hand2")
        self._cmd   = command
        self._style = style
        self._label_text = f"{icon}  {text}".strip() if icon else text
        px_h, px_v = self._PAD.get(size, self._PAD["md"])
        font_key = "ui_b" if style == "primary" else "ui"

        bg, fg, hbg, hborder = self._STYLES.get(style, self._STYLES["ghost"])()

        self._btn_frame = tk.Frame(self,
            bg=bg, cursor="hand2",
            highlightthickness=1,
            highlightbackground=hborder)
        self._btn_frame.pack(fill="both", expand=True)

        self._lbl = tk.Label(self._btn_frame,
            text=self._label_text,
            bg=bg, fg=fg,
            font=T.f(font_key),
            padx=px_h, pady=px_v,
            cursor="hand2")
        self._lbl.pack()

        self._bg  = bg
        self._fg  = fg
        self._hbg = hbg
        self._hborder = hborder

        for w in (self._btn_frame, self._lbl):
            w.bind("<Enter>",   self._enter)
            w.bind("<Leave>",   self._leave)
            w.bind("<Button-1>",self._press)
            w.bind("<ButtonRelease-1>", self._release)

        if tip: Tip(self, tip)

    def _enter(self, _=None):
        self._btn_frame.configure(bg=self._hbg)
        self._lbl.configure(bg=self._hbg)
        self._btn_frame.configure(highlightbackground=T.c["a"])

    def _leave(self, _=None):
        self._btn_frame.configure(bg=self._bg)
        self._lbl.configure(bg=self._bg)
        self._btn_frame.configure(highlightbackground=self._hborder)

    def _press(self, _=None):
        self._btn_frame.configure(highlightbackground=T.c["a"])

    def _release(self, e=None):
        self._leave()
        if self._cmd:
            self._cmd()

    def configure_text(self, text: str):
        self._lbl.configure(text=text)


# ══════════════════════════════════════════════════════
#  INPUT FIELD
# ══════════════════════════════════════════════════════

class Input(tk.Frame):
    """
    Поле вводу з підписом, placeholder, іконкою, валідацією.
    """

    def __init__(self, parent, label: str = "", placeholder: str = "",
                 icon: str = "", width: int = 0,
                 required: bool = False, **kw):
        super().__init__(parent, bg=parent.cget("bg"))
        self._ph = placeholder
        self._ph_active = False
        self._var = tk.StringVar()

        if label:
            lf = tk.Frame(self, bg=self.cget("bg"))
            lf.pack(fill="x", pady=(0, 3))
            tk.Label(lf, text=label,
                     bg=self.cget("bg"), fg=T.c["t2"],
                     font=T.f("cap")).pack(side="left")
            if required:
                tk.Label(lf, text=" *", bg=self.cget("bg"),
                         fg=T.c["red"], font=T.f("cap")).pack(side="left")

        # border wrapper
        self._border = tk.Frame(self,
            bg=T.c["b1"],
            highlightthickness=0)
        self._border.pack(fill="x")

        inner = tk.Frame(self._border, bg=T.c["bg_input"])
        inner.pack(fill="x", padx=1, pady=1)

        if icon:
            tk.Label(inner, text=icon,
                     bg=T.c["bg_input"], fg=T.c["t2"],
                     font=("Segoe UI",11),
                     padx=6).pack(side="left")

        w_kw = {"width": width} if width else {}
        self._entry = tk.Entry(inner,
            textvariable=self._var,
            bg=T.c["bg_input"], fg=T.c["t0"],
            insertbackground=T.c["a"],
            relief="flat", bd=5,
            font=T.f("ui"),
            **w_kw)
        self._entry.pack(side="left", fill="x", expand=True)

        if placeholder:
            self._set_ph()

        self._entry.bind("<FocusIn>",  self._on_focus)
        self._entry.bind("<FocusOut>", self._on_blur)

    def _set_ph(self):
        self._entry.insert(0, self._ph)
        self._entry.configure(fg=T.c["t3"])
        self._ph_active = True

    def _on_focus(self, _=None):
        self._border.configure(bg=T.c["b_focus"])
        if self._ph_active:
            self._entry.delete(0, "end")
            self._entry.configure(fg=T.c["t0"])
            self._ph_active = False

    def _on_blur(self, _=None):
        self._border.configure(bg=T.c["b1"])
        if not self._entry.get() and self._ph:
            self._set_ph()

    def get(self) -> str:
        if self._ph_active: return ""
        return self._var.get()

    def set(self, v: str):
        self._ph_active = False
        self._entry.delete(0, "end")
        self._entry.insert(0, v)
        self._entry.configure(fg=T.c["t0"])

    def focus(self):
        self._entry.focus_set()

    def bind_key(self, seq: str, cb: Callable):
        self._entry.bind(seq, cb)


# ══════════════════════════════════════════════════════
#  COMBOBOX
# ══════════════════════════════════════════════════════

class Combo(tk.Frame):
    def __init__(self, parent, label: str = "",
                 values: list | None = None, **kw):
        super().__init__(parent, bg=parent.cget("bg"))
        self._values = values or []
        if label:
            tk.Label(self, text=label, bg=self.cget("bg"),
                     fg=T.c["t2"], font=T.f("cap")).pack(anchor="w", pady=(0,3))

        style = ttk.Style()
        style.theme_use("default")
        style.configure("N.TCombobox",
            fieldbackground=T.c["bg_input"],
            background=T.c["bg3"],
            foreground=T.c["t0"],
            arrowcolor=T.c["t2"],
            borderwidth=0,
            relief="flat",
            padding=(6,4))
        style.map("N.TCombobox",
            fieldbackground=[("readonly", T.c["bg_input"])],
            selectbackground=[("readonly", T.c["a_dim"])],
            selectforeground=[("readonly", T.c["t0"])])

        self._var = tk.StringVar()
        self._cb  = ttk.Combobox(self, textvariable=self._var,
            values=self._values, state="readonly",
            font=T.f("ui"), style="N.TCombobox", **kw)
        self._cb.pack(fill="x")
        if self._values:
            self._cb.current(0)

    def get(self) -> str:
        return self._var.get()

    def set(self, v: str):
        self._var.set(v)


# ══════════════════════════════════════════════════════
#  MODAL DIALOG
# ══════════════════════════════════════════════════════

class Dialog(tk.Toplevel):
    """
    Базовий модальний діалог.
    Успадковуйте → перевизначте body() та on_confirm().
    """

    def __init__(self, parent: tk.Widget, title: str,
                 w: int = 520, h: int = 380):
        super().__init__(parent)
        self.title(title)
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()
        self.configure(bg=T.c["bg0"])
        self.result = None

        # центрування
        self.update_idletasks()
        rx = parent.winfo_rootx()
        ry = parent.winfo_rooty()
        rw = parent.winfo_width()
        rh = parent.winfo_height()
        x  = rx + (rw - w) // 2
        y  = ry + (rh - h) // 2
        self.geometry(f"{w}x{h}+{x}+{y}")

        self._build(title)
        self.bind("<Escape>", lambda _: self.cancel())

    def _build(self, title: str):
        # top accent bar
        tk.Frame(self, bg=T.c["a"], height=2).pack(fill="x")

        # header
        hdr = tk.Frame(self, bg=T.c["bg_modal"], pady=14, padx=20)
        hdr.pack(fill="x")
        tk.Label(hdr, text=title, bg=T.c["bg_modal"],
                 fg=T.c["t0"], font=T.f("h3")).pack(side="left")
        tk.Button(hdr, text="✕",
                  bg=T.c["bg_modal"], fg=T.c["t2"],
                  relief="flat", bd=0, cursor="hand2",
                  font=("Segoe UI",11),
                  activebackground=T.c["bg3"],
                  activeforeground=T.c["t0"],
                  command=self.cancel).pack(side="right")

        tk.Frame(self, bg=T.c["b0"], height=1).pack(fill="x")

        # body area
        self._body_frame = tk.Frame(self, bg=T.c["bg_modal"],
                                    padx=20, pady=18)
        self._body_frame.pack(fill="both", expand=True)
        self.body(self._body_frame)

        # footer
        tk.Frame(self, bg=T.c["b0"], height=1).pack(fill="x")
        foot = tk.Frame(self, bg=T.c["bg_modal"], padx=20, pady=12)
        foot.pack(fill="x")
        Btn(foot, "Підтвердити", style="primary",
            command=self._ok).pack(side="right", padx=(6,0))
        Btn(foot, "Скасувати",   style="outline",
            command=self.cancel).pack(side="right")

    def body(self, master: tk.Frame):
        """Перевизначте."""
        Label(master, "Тіло діалогу", bg=T.c["bg_modal"]).pack(pady=30)

    def _ok(self):
        if self.on_confirm():
            self.destroy()

    def on_confirm(self) -> bool:
        """Повертає True → закрити, False → залишити відкритим."""
        self.result = True
        return True

    def cancel(self):
        self.result = None
        self.destroy()


# ══════════════════════════════════════════════════════
#  DATA TABLE (Treeview wrapper)
# ══════════════════════════════════════════════════════

class Table(tk.Frame):
    """
    Повноцінна таблиця:
    - сортування кліком на заголовок
    - чередування рядків
    - контекстне меню
    - multi-select
    - рядок кількості
    """

    def __init__(self, parent,
                 columns: list[dict],
                 data:    list[tuple] | None = None,
                 on_row_double: Callable | None = None,
                 context_menu:  list[dict] | None = None,
                 **kw):
        """
        columns = [{"id":"name","label":"Ім'я","w":160,"anchor":"w"}, ...]
        data    = list of tuples matching column order
        context_menu = [{"label":"Редагувати","cmd": fn}, ...]
        """
        super().__init__(parent, bg=T.c["bg2"])
        self._cols    = columns
        self._data    = data or []
        self._dbl_cb  = on_row_double
        self._ctx_cfg = context_menu or []
        self._sort_col: str | None = None
        self._sort_asc: bool = True

        self._apply_style()
        self._build()
        if data: self.load(data)

    # ── style ──────────────────────────────────────────
    def _apply_style(self):
        s = ttk.Style()
        s.theme_use("default")
        s.configure("N.Treeview",
            background=T.c["bg2"],
            foreground=T.c["t1"],
            fieldbackground=T.c["bg2"],
            rowheight=T.sz("row_h"),
            font=T.f("ui"),
            borderwidth=0,
            relief="flat")
        s.configure("N.Treeview.Heading",
            background=T.c["bg3"],
            foreground=T.c["t2"],
            font=T.f("cap"),
            relief="flat",
            padding=(8, 6))
        s.map("N.Treeview",
            background=[("selected", T.c["a_dim"])],
            foreground=[("selected", T.c["a"])])
        s.map("N.Treeview.Heading",
            background=[("active", T.c["bg4"])])

    # ── build ──────────────────────────────────────────
    def _build(self):
        col_ids = [c["id"] for c in self._cols]

        # scrollbars
        vsb = ttk.Scrollbar(self, orient="vertical")
        hsb = ttk.Scrollbar(self, orient="horizontal")

        self._tree = ttk.Treeview(self,
            columns=col_ids, show="headings",
            style="N.Treeview",
            yscrollcommand=vsb.set,
            xscrollcommand=hsb.set,
            selectmode="extended")

        vsb.configure(command=self._tree.yview)
        hsb.configure(command=self._tree.xview)

        # headings
        for col in self._cols:
            cid = col["id"]
            self._tree.heading(cid,
                text=col.get("label", cid),
                anchor=col.get("anchor", "w"),
                command=lambda c=cid: self._sort(c))
            self._tree.column(cid,
                width=col.get("w", 120),
                minwidth=col.get("min_w", 60),
                anchor=col.get("anchor", "w"),
                stretch=col.get("stretch", True))

        vsb.pack(side="right",  fill="y")
        hsb.pack(side="bottom", fill="x")
        self._tree.pack(fill="both", expand=True)

        # bindings
        if self._dbl_cb:
            self._tree.bind("<Double-1>", self._on_dbl)
        self._tree.bind("<Button-3>", self._on_ctx)

        # count bar
        self._count_var = tk.StringVar(value="")
        tk.Label(self, textvariable=self._count_var,
                 bg=T.c["bg3"], fg=T.c["t2"],
                 font=T.f("ui_xs"),
                 padx=10, pady=3).pack(fill="x")

    # ── data ───────────────────────────────────────────
    def load(self, data: list[tuple], clear: bool = True):
        if clear:
            for iid in self._tree.get_children():
                self._tree.delete(iid)
        self._data = data
        for i, row in enumerate(data):
            tag = "alt" if i % 2 else "norm"
            self._tree.insert("", "end", values=row, tags=(tag,))
        self._tree.tag_configure("norm", background=T.c["bg2"])
        self._tree.tag_configure("alt",  background=T.c["bg1"])
        self._update_count()

    def _update_count(self):
        total = len(self._tree.get_children())
        sel   = len(self._tree.selection())
        if sel:
            self._count_var.set(f"Вибрано: {sel}  /  Всього: {total}")
        else:
            self._count_var.set(f"Всього: {total} записів")

    def selected_values(self) -> list[tuple]:
        return [self._tree.item(i, "values")
                for i in self._tree.selection()]

    # ── sort ───────────────────────────────────────────
    def _sort(self, col: str):
        col_ids = [c["id"] for c in self._cols]
        idx = col_ids.index(col)
        if self._sort_col == col:
            self._sort_asc = not self._sort_asc
        else:
            self._sort_col = col
            self._sort_asc = True

        items = [(self._tree.set(i, col), i)
                 for i in self._tree.get_children()]
        items.sort(key=lambda x: x[0], reverse=not self._sort_asc)
        for n, (_, iid) in enumerate(items):
            self._tree.move(iid, "", n)
            tag = "alt" if n % 2 else "norm"
            self._tree.item(iid, tags=(tag,))

        arrow = " ▲" if self._sort_asc else " ▼"
        for c in self._cols:
            lbl = c.get("label", c["id"])
            suffix = arrow if c["id"] == col else ""
            self._tree.heading(c["id"], text=lbl + suffix)

    # ── events ─────────────────────────────────────────
    def _on_dbl(self, e=None):
        sel = self._tree.selection()
        if sel and self._dbl_cb:
            self._dbl_cb(self._tree.item(sel[0], "values"))

    def _on_ctx(self, e):
        iid = self._tree.identify_row(e.y)
        if not iid: return
        if iid not in self._tree.selection():
            self._tree.selection_set(iid)
        self._update_count()

        if not self._ctx_cfg: return
        m = tk.Menu(self, tearoff=0,
            bg=T.c["bg3"], fg=T.c["t0"],
            activebackground=T.c["bg4"],
            activeforeground=T.c["a"],
            relief="flat", bd=0,
            font=T.f("ui"))
        for item in self._ctx_cfg:
            if item == "sep":
                m.add_separator()
            else:
                vals = self._tree.item(iid, "values")
                m.add_command(label=item["label"],
                    command=lambda fn=item["cmd"], v=vals: fn(v))
        m.tk_popup(e.x_root, e.y_root)

    def bind_select(self, cb: Callable):
        self._tree.bind("<<TreeviewSelect>>",
                        lambda _: (self._update_count(), cb()))


# ══════════════════════════════════════════════════════
#  MINI CHART (Canvas bar chart)
# ══════════════════════════════════════════════════════

class BarChart(tk.Canvas):
    """Простий canvas-бар-чарт без залежностей."""

    def __init__(self, parent, data: list[tuple],
                 w: int = 280, h: int = 120):
        """data = [("Січ", 42), ("Лют", 85), ...]"""
        super().__init__(parent, bg=T.c["bg2"],
                         highlightthickness=0, bd=0)
        self.configure(width=w, height=h)
        self._data = data
        self._w    = w
        self._h    = h
        self.bind("<Configure>", lambda _: self.render())
        self.after(50, self.render)

    def render(self, data: list[tuple] | None = None):
        if data: self._data = data
        try:
            self.delete("all")
        except tk.TclError:
            return
        if not self._data: return

        W = self.winfo_width()  or self._w
        H = self.winfo_height() or self._h
        if W <= 1: return
        pad_l, pad_r, pad_t, pad_b = 32, 12, 10, 24

        vals   = [v for _, v in self._data]
        maxv   = max(vals) if vals else 1
        colors = T.c["chart"]
        n      = len(self._data)
        bw     = (W - pad_l - pad_r) / n * 0.6
        gap    = (W - pad_l - pad_r) / n

        ch = H - pad_t - pad_b

        # grid lines
        for i in range(4):
            y = pad_t + ch * i / 3
            self.create_line(pad_l, y, W - pad_r, y,
                             fill=T.c["b0"], width=1)
            val = maxv * (3 - i) / 3
            self.create_text(pad_l - 4, y,
                             text=f"{val:.0f}",
                             anchor="e", fill=T.c["t3"],
                             font=T.f("ui_xs"))

        # bars
        for i, (label, value) in enumerate(self._data):
            x0 = pad_l + i * gap + gap * 0.2
            bh = (value / maxv) * ch if maxv else 0
            y0 = H - pad_b - bh
            y1 = H - pad_b
            color = colors[i % len(colors)]

            # bar shadow
            self.create_rectangle(x0+2, y0+2, x0+bw+2, y1+2,
                                   fill=T.c["bg0"], outline="")
            # bar
            self.create_rectangle(x0, y0, x0+bw, y1,
                                   fill=color, outline="")
            # value on top
            self.create_text(x0 + bw/2, y0 - 4,
                             text=str(int(value)),
                             fill=color, font=T.f("ui_xs"),
                             anchor="s")
            # label
            self.create_text(x0 + bw/2, H - pad_b + 8,
                             text=label, fill=T.c["t2"],
                             font=T.f("ui_xs"), anchor="n")


class LineChart(tk.Canvas):
    """Canvas line chart."""

    def __init__(self, parent, data: list[tuple],
                 w: int = 300, h: int = 100,
                 color: str | None = None):
        super().__init__(parent, bg=T.c["bg2"],
                         highlightthickness=0, bd=0)
        self.configure(width=w, height=h)
        self._data  = data
        self._color = color or T.c["a"]
        self._w     = w
        self._h     = h
        self.bind("<Configure>", lambda _: self.render())
        self.after(50, self.render)

    def render(self, data: list[tuple] | None = None):
        if data: self._data = data
        try:
            self.delete("all")
        except tk.TclError:
            return
        if len(self._data) < 2: return
        W = self.winfo_width()  or self._w
        H = self.winfo_height() or self._h
        if W <= 1: return
        pad = 8
        vals = [v for _, v in self._data]
        minv, maxv = min(vals), max(vals)
        rng = (maxv - minv) or 1
        n   = len(self._data)
        pts = []
        for i, (_, v) in enumerate(self._data):
            x = pad + i * (W - 2*pad) / (n - 1)
            y = H - pad - (v - minv) / rng * (H - 2*pad)
            pts.append((x, y))

        # fill под кривой
        poly = [pad, H-pad] + [c for pt in pts for c in pt] + [W-pad, H-pad]
        self.create_polygon(poly, fill=T.c["a_glow"], outline="")

        # линия
        for i in range(len(pts)-1):
            self.create_line(*pts[i], *pts[i+1],
                             fill=self._color, width=2, smooth=True)
        # dots
        for x, y in pts:
            self.create_oval(x-3, y-3, x+3, y+3,
                             fill=self._color, outline=T.c["bg2"], width=1)


# ══════════════════════════════════════════════════════
#  KPI CARD
# ══════════════════════════════════════════════════════

class KpiCard(tk.Frame):
    def __init__(self, parent, title: str, value: str,
                 delta: str = "", positive: bool = True,
                 icon: str = "", accent: str | None = None,
                 sparkline: list | None = None):
        acc = accent or T.c["a"]
        super().__init__(parent,
            bg=T.c["bg2"],
            highlightthickness=1,
            highlightbackground=T.c["b1"])

        # top accent stripe
        tk.Frame(self, bg=acc, height=2).pack(fill="x")

        body = tk.Frame(self, bg=T.c["bg2"], padx=16, pady=14)
        body.pack(fill="both", expand=True)

        # icon + title
        top_row = tk.Frame(body, bg=T.c["bg2"])
        top_row.pack(fill="x")
        if icon:
            tk.Label(top_row, text=icon, bg=T.c["bg2"],
                     fg=acc, font=("Segoe UI",15)).pack(side="left")
        tk.Label(top_row, text=title, bg=T.c["bg2"],
                 fg=T.c["t2"], font=T.f("cap"),
                 padx=6 if icon else 0).pack(side="left")

        # value
        tk.Label(body, text=value, bg=T.c["bg2"],
                 fg=T.c["t0"], font=T.f("num")).pack(anchor="w", pady=(8,0))

        # delta
        if delta:
            d_col = T.c["green"] if positive else T.c["red"]
            d_arr = "▲" if positive else "▼"
            tk.Label(body, text=f"{d_arr} {delta}",
                     bg=T.c["bg2"], fg=d_col,
                     font=T.f("ui_sm")).pack(anchor="w", pady=(2,0))

        # sparkline
        if sparkline:
            pts = [(str(i), v) for i, v in enumerate(sparkline)]
            LineChart(body, pts, w=160, h=40,
                      color=acc).pack(anchor="w", pady=(8,0))


# ══════════════════════════════════════════════════════
#  TAB BAR
# ══════════════════════════════════════════════════════

class TabBar(tk.Frame):
    """Горизонтальна панель вкладок."""

    def __init__(self, parent, on_change: Callable | None = None):
        super().__init__(parent, bg=T.c["bg1"],
                         highlightthickness=0)
        self._tabs:    dict[str, dict] = {}
        self._active:  str | None = None
        self._on_change = on_change

    def add(self, key: str, label: str, closable: bool = False):
        frm = tk.Frame(self, bg=T.c["bg1"], cursor="hand2")
        frm.pack(side="left")

        lbl = tk.Label(frm, text=label, bg=T.c["bg1"],
                       fg=T.c["t2"], font=T.f("tab"),
                       padx=14, pady=10, cursor="hand2")
        lbl.pack(side="left")

        underline = tk.Frame(frm, bg=T.c["bg1"], height=2)
        underline.pack(fill="x")

        self._tabs[key] = {"frame": frm, "label": lbl,
                            "underline": underline, "text": label}

        for w in (frm, lbl):
            w.bind("<Button-1>", lambda _, k=key: self.select(k))
            w.bind("<Enter>",    lambda _, f=frm: f.configure(bg=T.c["bg3"]) or None)
            w.bind("<Leave>",    lambda _, k=key, f=frm: self._restore_bg(k, f))

        if not self._active:
            self.select(key)

    def _restore_bg(self, key: str, frm: tk.Frame):
        if key == self._active:
            frm.configure(bg=T.c["bg1"])
        else:
            frm.configure(bg=T.c["bg1"])

    def select(self, key: str):
        for k, t in self._tabs.items():
            active = (k == key)
            t["label"].configure(
                fg=T.c["t0"] if active else T.c["t2"],
                font=T.f("tab_a") if active else T.f("tab"))
            t["underline"].configure(bg=T.c["a"] if active else T.c["bg1"])
            t["frame"].configure(bg=T.c["bg1"])
        self._active = key
        if self._on_change:
            self._on_change(key)


# ══════════════════════════════════════════════════════
#  PAGE HEADER
# ══════════════════════════════════════════════════════

class PageHeader(tk.Frame):
    def __init__(self, parent,
                 title: str,
                 breadcrumb: str = "",
                 actions: list[dict] | None = None):
        super().__init__(parent, bg=T.c["bg1"])

        inner = tk.Frame(self, bg=T.c["bg1"], padx=24, pady=16)
        inner.pack(fill="x")

        left = tk.Frame(inner, bg=T.c["bg1"])
        left.pack(side="left", fill="x", expand=True)

        if breadcrumb:
            tk.Label(left, text=breadcrumb, bg=T.c["bg1"],
                     fg=T.c["t3"], font=T.f("ui_xs")).pack(anchor="w")

        tk.Label(left, text=title, bg=T.c["bg1"],
                 fg=T.c["t0"], font=T.f("h2")).pack(anchor="w")

        if actions:
            right = tk.Frame(inner, bg=T.c["bg1"])
            right.pack(side="right")
            for a in reversed(actions):
                Btn(right, a.get("label",""),
                    style=a.get("style","ghost"),
                    icon=a.get("icon",""),
                    command=a.get("cmd"),
                    tip=a.get("tip","")).pack(side="right", padx=(4,0), ipady=1)

        Sep(self).pack(fill="x")


# ══════════════════════════════════════════════════════
#  BASE MODULE
# ══════════════════════════════════════════════════════

class BaseModule(tk.Frame):
    MODULE_ID    = "base"
    MODULE_TITLE = "Модуль"
    MODULE_ICON  = "◈"
    MODULE_GROUP = ""   # "CRM" | "ERP" | "GRC" | ""

    def __init__(self, parent: tk.Widget):
        super().__init__(parent, bg=T.c["bg1"])
        self._built = False

    def activate(self):
        if not self._built:
            self.build()
            self._built = True
        self.refresh()

    def build(self):
        """Побудова UI — перевизначте."""
        Label(self, "Порожній модуль", bg=T.c["bg1"],
              color=T.c["t2"]).pack(expand=True)

    def refresh(self):
        """Оновлення даних — перевизначте."""

    # ── helpers ──────────────────────────────────────
    def scrollable(self, bg: str | None = None) -> tk.Frame:
        """Повертає фрейм всередині scrollable canvas."""
        bg = bg or T.c["bg1"]
        cnv = tk.Canvas(self, bg=bg, highlightthickness=0, bd=0)
        vsb = ttk.Scrollbar(self, orient="vertical",
                            command=cnv.yview)
        cnv.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        cnv.pack(fill="both", expand=True)
        inner = tk.Frame(cnv, bg=bg)
        win   = cnv.create_window((0,0), window=inner, anchor="nw")

        def _resize(e):
            cnv.configure(scrollregion=cnv.bbox("all"))
            cnv.itemconfigure(win, width=cnv.winfo_width())
        inner.bind("<Configure>", _resize)
        cnv.bind("<MouseWheel>",
                 lambda e: cnv.yview_scroll(int(-e.delta/60), "units"))
        return inner

    def card(self, parent: tk.Widget,
             padx: int = 16, pady: int = 14) -> tk.Frame:
        outer = tk.Frame(parent,
            bg=T.c["b1"],
            highlightthickness=0)
        inner = tk.Frame(outer, bg=T.c["bg2"],
                         padx=padx, pady=pady)
        inner.pack(fill="both", expand=True, padx=1, pady=1)
        return inner


# ══════════════════════════════════════════════════════
#  ═══  MODULES  ═══
# ══════════════════════════════════════════════════════

# ── 1. DASHBOARD ──────────────────────────────────────

class Dashboard(BaseModule):
    MODULE_ID    = "dashboard"
    MODULE_TITLE = "Dashboard"
    MODULE_ICON  = "⊞"

    def build(self):
        # header
        PageHeader(self, "Панель керування",
            breadcrumb="NEXUS › Dashboard",
            actions=[
                {"label":"Звіт",  "icon":"↓","style":"outline",
                 "cmd": lambda: Toast.show("Генерація звіту…","info")},
                {"label":"Новий", "icon":"+","style":"primary",
                 "cmd": lambda: QuickAddDialog(self.winfo_toplevel()).grab_set()},
            ]).pack(fill="x")

        scroll = self.scrollable()

        # ── KPI row ─────────────────────────────────
        kpi_frame = tk.Frame(scroll, bg=T.c["bg1"],
                             padx=24, pady=16)
        kpi_frame.pack(fill="x")
        kpi_frame.columnconfigure((0,1,2,3), weight=1, uniform="k")

        spark1 = [42,55,48,70,65,80,77,90]
        spark2 = [100,120,115,140,130,160,150,175]
        spark3 = [8,5,9,7,12,10,14,11]
        spark4 = [3,5,4,7,6,9,8,10]

        cards = [
            ("Клієнти",    "2 847",  "+18% / міс",  True,  "◉", T.c["a"],     spark1),
            ("Дохід",      "₴8.4M",  "+11% / міс",  True,  "◇", T.c["green"], spark2),
            ("Відкр.ліди", "413",    "−12 вчора",   False, "◎", T.c["amber"], spark3),
            ("Ризики",     "24",     "+3 нові",      False, "△", T.c["red"],   spark4),
        ]
        for col, (title,val,delta,pos,icon,acc,sp) in enumerate(cards):
            c = KpiCard(kpi_frame, title, val, delta, pos, icon, acc, sp)
            c.grid(row=0, column=col,
                   padx=(0, 12 if col < 3 else 0),
                   sticky="nsew")

        # ── Charts row ──────────────────────────────
        chart_frame = tk.Frame(scroll, bg=T.c["bg1"], padx=24)
        chart_frame.pack(fill="x", pady=(0,16))
        chart_frame.columnconfigure(0, weight=3)
        chart_frame.columnconfigure(1, weight=2)

        # bar chart card
        bc = self.card(chart_frame, padx=16, pady=14)
        bc.master.grid(row=0, column=0, sticky="nsew", padx=(0,12))
        tk.Label(bc, text="Продажі по місяцях",
                 bg=T.c["bg2"], fg=T.c["t0"],
                 font=T.f("h4")).pack(anchor="w", pady=(0,10))
        months_data = [
            ("Жов","185"),("Лис","240"),("Гру","198"),
            ("Січ","310"),("Лют","275"),("Бер","342"),
        ]
        nums = [(l, int(v)) for l,v in months_data]
        BarChart(bc, nums, h=160).pack(fill="x")

        # pipeline card
        pc = self.card(chart_frame, padx=16, pady=14)
        pc.master.grid(row=0, column=1, sticky="nsew")
        tk.Label(pc, text="Воронка продажів",
                 bg=T.c["bg2"], fg=T.c["t0"],
                 font=T.f("h4")).pack(anchor="w", pady=(0,12))
        stages = [
            ("Лід",       420, T.c["a"]),
            ("Кваліф.",   280, T.c["a2"]),
            ("Пропозиція",160, T.c["green"]),
            ("Угода",      85, T.c["amber"]),
            ("Закрито",    52, T.c["red"]),
        ]
        maxv = stages[0][1]
        for label, val, color in stages:
            r = tk.Frame(pc, bg=T.c["bg2"])
            r.pack(fill="x", pady=3)
            tk.Label(r, text=label, bg=T.c["bg2"],
                     fg=T.c["t1"], font=T.f("ui_sm"),
                     width=11, anchor="w").pack(side="left")
            bar_bg = tk.Frame(r, bg=T.c["bg3"], height=14)
            bar_bg.pack(side="left", fill="x", expand=True, padx=(4,8))
            bar_bg.update_idletasks()
            pct = val / maxv
            # рисуємо через Canvas
            cnv = tk.Canvas(r, bg=T.c["bg2"], height=14, width=120,
                            highlightthickness=0, bd=0)
            cnv.pack(side="left")
            cnv.after(50, lambda c=cnv, p=pct, col=color:
                      c.create_rectangle(0,0, int(c.winfo_width()*p),14,
                                         fill=col, outline=""))
            tk.Label(r, text=str(val), bg=T.c["bg2"],
                     fg=T.c["t2"], font=T.f("ui_sm"),
                     width=4).pack(side="left")

        # ── Activity feed ───────────────────────────
        af = tk.Frame(scroll, bg=T.c["bg1"], padx=24, pady=(0,24))
        af.pack(fill="x")
        af.columnconfigure(0, weight=3)
        af.columnconfigure(1, weight=2)

        # Recent table
        rec = self.card(af)
        rec.master.grid(row=0, column=0, sticky="nsew", padx=(0,12))
        tk.Label(rec, text="Останні угоди",
                 bg=T.c["bg2"], fg=T.c["t0"],
                 font=T.f("h4")).pack(anchor="w", pady=(0,10))
        cols = [
            {"id":"id",      "label":"№",        "w": 50,  "stretch":False},
            {"id":"client",  "label":"Клієнт",   "w":160},
            {"id":"amount",  "label":"Сума",      "w": 90, "anchor":"e"},
            {"id":"status",  "label":"Статус",    "w": 90},
            {"id":"date",    "label":"Дата",      "w": 80},
        ]
        data = [
            ("001","Альфа Корп",  "₴48 500","✔ Закрита","02.04"),
            ("002","Бета ТОВ",    "₴12 200","⧖ Переговори","01.04"),
            ("003","Гамма ФОП",   "₴6 800", "✔ Закрита","31.03"),
            ("004","Дельта ЛТД",  "₴93 000","✖ Втрачена","30.03"),
            ("005","Епсілон Груп","₴21 400","⧖ Переговори","29.03"),
            ("006","Зета LLC",    "₴5 100", "✔ Закрита","28.03"),
        ]
        tbl = Table(rec, cols, data,
            on_row_double=lambda v: Toast.show(f"Угода {v[0]}: {v[1]}","info"),
            context_menu=[
                {"label":"✏  Редагувати",
                 "cmd": lambda v: Toast.show(f"Редагування: {v[1]}","info")},
                {"label":"⬡  Деталі",
                 "cmd": lambda v: Toast.show(f"Деталі: {v[0]}","info")},
                "sep",
                {"label":"✖  Видалити",
                 "cmd": lambda v: Toast.show(f"Видалено: {v[1]}","err")},
            ])
        tbl.pack(fill="both", expand=True)

        # System status
        ss = self.card(af)
        ss.master.grid(row=0, column=1, sticky="nsew")
        tk.Label(ss, text="Стан системи",
                 bg=T.c["bg2"], fg=T.c["t0"],
                 font=T.f("h4")).pack(anchor="w", pady=(0,10))
        statuses = [
            ("CRM модуль",    "● Активний",  "green"),
            ("ERP модуль",    "● Активний",  "green"),
            ("GRC модуль",    "⚠ Увага",     "amber"),
            ("Синхронізація", "● Активна",   "green"),
            ("Резервна копія","● Актуальна", "green"),
            ("API Gateway",   "● Online",    "green"),
            ("БД Postgres",   "● Online",    "green"),
            ("Кеш Redis",     "⚠ 78% CPU",   "amber"),
        ]
        for name, status, kind in statuses:
            r = tk.Frame(ss, bg=T.c["bg2"])
            r.pack(fill="x", pady=3)
            Sep(r, "v").pack(side="left", fill="y", padx=(0,8))
            tk.Label(r, text=name, bg=T.c["bg2"],
                     fg=T.c["t1"], font=T.f("ui"),
                     anchor="w").pack(side="left", fill="x", expand=True)
            tk.Label(r, text=status, bg=T.c["bg2"],
                     fg=T.c[kind], font=T.f("ui_sm")).pack(side="right")


# ── 2. CONTACTS ───────────────────────────────────────

class Contacts(BaseModule):
    MODULE_ID    = "contacts"
    MODULE_TITLE = "Контакти"
    MODULE_ICON  = "◉"
    MODULE_GROUP = "CRM"

    def build(self):
        PageHeader(self, "Контакти",
            breadcrumb="NEXUS › CRM › Контакти",
            actions=[
                {"label":"Імпорт",  "style":"outline","icon":"↑",
                 "cmd": lambda: Toast.show("Імпорт CSV…","info")},
                {"label":"Експорт", "style":"outline","icon":"↓",
                 "cmd": lambda: Toast.show("Експорт…","info")},
                {"label":"+ Контакт","style":"primary",
                 "cmd": lambda: ContactFormDialog(
                     self.winfo_toplevel())},
            ]).pack(fill="x")

        # toolbar
        tb = tk.Frame(self, bg=T.c["bg1"], padx=24, pady=10)
        tb.pack(fill="x")
        self._search = Input(tb, placeholder="⌕  Пошук контактів…",
                             width=32)
        self._search.pack(side="left")
        self._search.bind_key("<Return>", lambda _: self.refresh())

        Btn(tb, "Всі", style="outline",
            command=lambda: self._filter("all")).pack(side="left", padx=(10,4), ipady=1)
        for lbl in ("Клієнти","Партнери","Ліди"):
            Btn(tb, lbl, style="ghost",
                command=lambda l=lbl: Toast.show(f"Фільтр: {l}","info")
                ).pack(side="left", padx=2, ipady=1)

        tk.Label(tb, textvariable=self._count_var(),
                 bg=T.c["bg1"], fg=T.c["t2"],
                 font=T.f("ui_sm")).pack(side="right")

        # table
        wrap = tk.Frame(self, bg=T.c["bg1"], padx=24, pady=(0,20))
        wrap.pack(fill="both", expand=True)

        cols = [
            {"id":"chk",      "label":"",          "w": 28, "stretch":False},
            {"id":"name",     "label":"КОНТАКТ",   "w":180},
            {"id":"company",  "label":"КОМПАНІЯ",  "w":150},
            {"id":"email",    "label":"EMAIL",      "w":180},
            {"id":"phone",    "label":"ТЕЛЕФОН",    "w":120},
            {"id":"type",     "label":"ТИП",        "w": 90},
            {"id":"status",   "label":"СТАТУС",     "w": 90},
            {"id":"created",  "label":"ДОДАНО",     "w": 90},
        ]
        self._table = Table(wrap, cols,
            on_row_double=lambda v: ContactFormDialog(
                self.winfo_toplevel(), v),
            context_menu=[
                {"label":"✏  Редагувати",
                 "cmd": lambda v: ContactFormDialog(
                     self.winfo_toplevel(), v)},
                {"label":"✉  Написати",
                 "cmd": lambda v: Toast.show(
                     f"Email до {v[1]}","info")},
                "sep",
                {"label":"✖  Видалити",
                 "cmd": lambda v: Toast.show(
                     f"Видалено: {v[1]}","err")},
            ])
        self._table.pack(fill="both", expand=True)
        self.refresh()

    def _count_var(self):
        self._cv = tk.StringVar(value="")
        return self._cv

    def _filter(self, f: str):
        self.refresh()

    def refresh(self):
        random.seed(7)
        types   = ["Клієнт","Партнер","Лід","VIP"]
        statuses= ["Активний","Неактивний","Новий"]
        comps   = ["Альфа Корп","Бета ТОВ","Гамма ФОП","Дельта ЛТД",
                   "Омега LLC","Зета Груп","Каппа Inc"]
        names   = ["Олег Петренко","Ірина Коваль","Денис Мельник",
                   "Тетяна Бойко","Андрій Шевченко","Юлія Кравченко",
                   "Максим Лисенко","Олена Павленко","Сергій Марченко",
                   "Ніна Гончар"]
        data = []
        for i, nm in enumerate(names):
            em = nm.lower().replace(" ",".")+"@example.com"
            ph = f"+38(0{random.randint(50,99)}){random.randint(100,999)}-{random.randint(10,99)}-{random.randint(10,99)}"
            data.append(("☐", nm, random.choice(comps), em, ph,
                         random.choice(types), random.choice(statuses),
                         f"{random.randint(1,28):02d}.0{random.randint(1,4)}.26"))
        self._table.load(data)


# ── 3. DEALS ──────────────────────────────────────────

class Deals(BaseModule):
    MODULE_ID    = "deals"
    MODULE_TITLE = "Угоди"
    MODULE_ICON  = "◇"
    MODULE_GROUP = "CRM"

    def build(self):
        PageHeader(self, "Угоди",
            breadcrumb="NEXUS › CRM › Угоди",
            actions=[
                {"label":"+ Угода","style":"primary",
                 "cmd": lambda: Toast.show("Форма угоди","info")},
            ]).pack(fill="x")

        # kanban-like stage summary
        stages_bar = tk.Frame(self, bg=T.c["bg1"], padx=24, pady=12)
        stages_bar.pack(fill="x")
        stages_bar.columnconfigure((0,1,2,3,4), weight=1, uniform="s")

        stage_data = [
            ("Новий", "87", T.c["a"]),
            ("Контакт", "54", T.c["a2"]),
            ("Пропозиція", "31", T.c["green"]),
            ("Переговори", "18", T.c["amber"]),
            ("Закрито", "12", T.c["red"]),
        ]
        for col,(lbl,cnt,clr) in enumerate(stage_data):
            sf = tk.Frame(stages_bar, bg=T.c["bg2"],
                          highlightthickness=1,
                          highlightbackground=T.c["b1"])
            sf.grid(row=0, column=col,
                    padx=(0,8 if col<4 else 0), sticky="nsew")
            tk.Frame(sf, bg=clr, height=3).pack(fill="x")
            inner = tk.Frame(sf, bg=T.c["bg2"], padx=12, pady=8)
            inner.pack()
            tk.Label(inner, text=cnt, bg=T.c["bg2"],
                     fg=clr, font=T.f("num_sm")).pack()
            tk.Label(inner, text=lbl, bg=T.c["bg2"],
                     fg=T.c["t2"], font=T.f("cap")).pack()

        # table
        wrap = tk.Frame(self, bg=T.c["bg1"], padx=24, pady=(8,20))
        wrap.pack(fill="both", expand=True)

        cols = [
            {"id":"id",       "label":"№",         "w": 60, "stretch":False},
            {"id":"name",     "label":"НАЗВА",      "w":200},
            {"id":"client",   "label":"КЛІЄНТ",     "w":150},
            {"id":"amount",   "label":"СУМА",       "w": 90, "anchor":"e"},
            {"id":"stage",    "label":"СТАДІЯ",     "w":110},
            {"id":"owner",    "label":"МЕНЕДЖЕР",   "w":130},
            {"id":"close",    "label":"ЗАКРИТТЯ",   "w": 90},
            {"id":"prob",     "label":"ЙМОВІРН.",   "w": 80, "anchor":"e"},
        ]
        random.seed(42)
        stages_list = ["Новий","Контакт","Пропозиція","Переговори","Закрито"]
        owners = ["О.Мельник","І.Коваль","Д.Петренко","Т.Бойко"]
        clients = ["Альфа","Бета ТОВ","Гамма","Дельта","Омега","Зета","Каппа"]
        data = []
        for i in range(1, 25):
            data.append((
                f"D-{i:04d}",
                f"Проєкт {random.choice(['Alpha','Beta','Gamma','Delta'])} {i}",
                random.choice(clients),
                f"₴{random.randint(5,200)*1000:,}".replace(",","\u202f"),
                random.choice(stages_list),
                random.choice(owners),
                f"{random.randint(1,28):02d}.0{random.randint(4,6)}.26",
                f"{random.randint(20,95)}%",
            ))
        Table(wrap, cols, data,
            on_row_double=lambda v: Toast.show(f"Угода {v[0]}: {v[1]}","info"),
            context_menu=[
                {"label":"✏  Редагувати",
                 "cmd": lambda v: Toast.show(f"Ред. {v[0]}","info")},
                "sep",
                {"label":"✖  Видалити",
                 "cmd": lambda v: Toast.show(f"Видалено {v[0]}","err")},
            ]).pack(fill="both", expand=True)


# ── 4. FINANCE ────────────────────────────────────────

class Finance(BaseModule):
    MODULE_ID    = "finance"
    MODULE_TITLE = "Фінанси"
    MODULE_ICON  = "₴"
    MODULE_GROUP = "ERP"

    def build(self):
        PageHeader(self, "Фінанси",
            breadcrumb="NEXUS › ERP › Фінанси",
            actions=[
                {"label":"Новий рахунок","style":"primary",
                 "cmd": lambda: Toast.show("Форма рахунку","info")},
            ]).pack(fill="x")

        scroll = self.scrollable()
        pad = tk.Frame(scroll, bg=T.c["bg1"], padx=24, pady=16)
        pad.pack(fill="x")
        pad.columnconfigure((0,1,2), weight=1, uniform="f")

        for col,(title,val,delta,pos,acc) in enumerate([
            ("Дохід (квартал)",  "₴24.8M", "+14%", True,  T.c["green"]),
            ("Витрати",          "₴9.2M",  "+6%",  False, T.c["amber"]),
            ("Чистий прибуток",  "₴15.6M", "+18%", True,  T.c["a"]),
        ]):
            KpiCard(pad, title, val, delta, pos, accent=acc
                    ).grid(row=0, column=col,
                           padx=(0,12 if col<2 else 0),
                           sticky="nsew")

        tpad = tk.Frame(scroll, bg=T.c["bg1"], padx=24, pady=(0,24))
        tpad.pack(fill="both", expand=True)
        cols = [
            {"id":"id",      "label":"РАХУНОК", "w": 90},
            {"id":"client",  "label":"КОНТРАГЕНТ","w":160},
            {"id":"amount",  "label":"СУМА",    "w": 90,"anchor":"e"},
            {"id":"vat",     "label":"ПДВ",     "w": 70,"anchor":"e"},
            {"id":"total",   "label":"РАЗОМ",   "w": 90,"anchor":"e"},
            {"id":"status",  "label":"СТАТУС",  "w": 90},
            {"id":"due",     "label":"ТЕРМІН",  "w": 80},
        ]
        random.seed(11)
        data=[]
        for i in range(1,20):
            amt = random.randint(3,80)*1000
            vat = int(amt*0.2)
            st  = random.choice(["✔ Оплачено","⧖ Очікує","✖ Прострочено"])
            data.append((f"INV-{i:04d}",
                random.choice(["Альфа","Бета","Гамма","Дельта","Омега"]),
                f"₴{amt:,}".replace(",","\u202f"),
                f"₴{vat:,}".replace(",","\u202f"),
                f"₴{amt+vat:,}".replace(",","\u202f"),
                st,
                f"{random.randint(1,30):02d}.0{random.randint(4,7)}.26"))
        Table(tpad, cols, data,
            on_row_double=lambda v: Toast.show(f"Рахунок {v[0]}","info"),
            context_menu=[
                {"label":"✏  Деталі",
                 "cmd": lambda v: Toast.show(v[0],"info")},
                {"label":"⬇  PDF",
                 "cmd": lambda v: Toast.show(f"PDF {v[0]}","ok")},
            ]).pack(fill="both", expand=True)


# ── 5. RISKS ──────────────────────────────────────────

class Risks(BaseModule):
    MODULE_ID    = "risks"
    MODULE_TITLE = "Ризики"
    MODULE_ICON  = "△"
    MODULE_GROUP = "GRC"

    def build(self):
        PageHeader(self, "Реєстр ризиків",
            breadcrumb="NEXUS › GRC › Ризики",
            actions=[
                {"label":"Матриця ризиків","style":"outline",
                 "cmd": lambda: RiskMatrixDialog(self.winfo_toplevel())},
                {"label":"+ Ризик","style":"primary",
                 "cmd": lambda: Toast.show("Форма ризику","info")},
            ]).pack(fill="x")

        scroll = self.scrollable()
        kpad = tk.Frame(scroll, bg=T.c["bg1"], padx=24, pady=16)
        kpad.pack(fill="x")
        kpad.columnconfigure((0,1,2,3), weight=1, uniform="r")
        for col,(title,val,clr) in enumerate([
            ("Всього ризиків", "47",  T.c["t0"]),
            ("Критичних",      "5",   T.c["red"]),
            ("Високих",        "12",  T.c["amber"]),
            ("Прийнятних",     "30",  T.c["green"]),
        ]):
            f = tk.Frame(kpad, bg=T.c["bg2"],
                         highlightthickness=1,
                         highlightbackground=T.c["b1"])
            f.grid(row=0,column=col,
                   padx=(0,12 if col<3 else 0), sticky="nsew")
            tk.Frame(f, bg=clr, height=2).pack(fill="x")
            ii = tk.Frame(f, bg=T.c["bg2"], padx=14, pady=10)
            ii.pack()
            tk.Label(ii,text=val,bg=T.c["bg2"],fg=clr,
                     font=T.f("num_sm")).pack()
            tk.Label(ii,text=title,bg=T.c["bg2"],fg=T.c["t2"],
                     font=T.f("cap")).pack()

        tpad = tk.Frame(scroll, bg=T.c["bg1"], padx=24, pady=(0,24))
        tpad.pack(fill="both", expand=True)
        cols = [
            {"id":"id",       "label":"ID",         "w": 70},
            {"id":"name",     "label":"РИЗИК",       "w":220},
            {"id":"category", "label":"КАТЕГОРІЯ",   "w":110},
            {"id":"impact",   "label":"ВПЛИВ",       "w": 80},
            {"id":"prob",     "label":"ЙМОВІРН.",    "w": 90},
            {"id":"level",    "label":"РІВЕНЬ",      "w": 90},
            {"id":"owner",    "label":"ВЛАСНИК",     "w":120},
            {"id":"status",   "label":"СТАТУС",      "w": 90},
        ]
        random.seed(99)
        cats   = ["Операційний","Фінансовий","Комплаєнс","ІТ","Репутаційний"]
        levels = [("Критичний",T.c["red"]),
                  ("Високий",T.c["amber"]),
                  ("Середній",T.c["blue"]),
                  ("Низький",T.c["green"])]
        owners = ["О.Мельник","І.Коваль","Д.Петренко"]
        data=[]
        for i in range(1,25):
            lvl,_  = random.choice(levels)
            data.append((
                f"R-{i:04d}",
                f"Ризик категорії {random.choice(cats)} #{i}",
                random.choice(cats),
                random.choice(["Низький","Середній","Високий","Критичний"]),
                f"{random.randint(10,90)}%",
                lvl,
                random.choice(owners),
                random.choice(["Відкритий","Мітигований","Закритий"]),
            ))
        Table(tpad, cols, data,
            on_row_double=lambda v: Toast.show(f"Ризик {v[0]}","warn"),
            context_menu=[
                {"label":"✏  Редагувати",
                 "cmd": lambda v: Toast.show(v[0],"info")},
                {"label":"✔  Закрити ризик",
                 "cmd": lambda v: Toast.show(f"Закрито {v[0]}","ok")},
            ]).pack(fill="both", expand=True)


# ── 6. SETTINGS ───────────────────────────────────────

class Settings(BaseModule):
    MODULE_ID    = "settings"
    MODULE_TITLE = "Налаштування"
    MODULE_ICON  = "⚙"

    def build(self):
        PageHeader(self, "Налаштування",
            breadcrumb="NEXUS › Налаштування").pack(fill="x")

        scroll = self.scrollable()
        pad    = tk.Frame(scroll, bg=T.c["bg1"], padx=32, pady=20)
        pad.pack(fill="both", expand=True)

        self._sec(pad, "ЗОВНІШНІЙ ВИГЛЯД")
        theme_row = tk.Frame(pad, bg=T.c["bg1"])
        theme_row.pack(fill="x", pady=(8,0))
        tk.Label(theme_row, text="Тема", bg=T.c["bg1"],
                 fg=T.c["t1"], font=T.f("ui_b"),
                 width=24, anchor="w").pack(side="left")
        tv = tk.StringVar(value=T._name)
        for val, lbl in [("dark","◑  Темна"),("light","☀  Світла")]:
            tk.Radiobutton(theme_row, text=lbl,
                variable=tv, value=val,
                bg=T.c["bg1"], fg=T.c["t0"],
                selectcolor=T.c["bg3"],
                activebackground=T.c["bg1"],
                font=T.f("ui"),
                command=lambda v=val: T.apply(v)
            ).pack(side="left", padx=12)

        self._sec(pad, "ЗАГАЛЬНІ")
        for lbl, val in [
            ("Організація",     APP["org"]),
            ("Версія системи",  APP["version"]),
            ("Мова",            "Українська"),
            ("Часовий пояс",    "Europe/Kyiv (UTC+3)"),
        ]:
            r = tk.Frame(pad, bg=T.c["bg1"])
            r.pack(fill="x", pady=6)
            tk.Label(r, text=lbl, bg=T.c["bg1"],
                     fg=T.c["t2"], font=T.f("ui"),
                     width=24, anchor="w").pack(side="left")
            tk.Label(r, text=val, bg=T.c["bg1"],
                     fg=T.c["t0"], font=T.f("ui_b")).pack(side="left")

        self._sec(pad, "СИСТЕМА")
        for lbl, val in [
            ("Python",    sys.version.split()[0]),
            ("Tkinter",   str(tk.TkVersion)),
            ("ОС",        platform.system()+" "+platform.release()),
        ]:
            r = tk.Frame(pad, bg=T.c["bg1"])
            r.pack(fill="x", pady=5)
            tk.Label(r, text=lbl, bg=T.c["bg1"],
                     fg=T.c["t2"], font=T.f("ui"),
                     width=24, anchor="w").pack(side="left")
            tk.Label(r, text=val, bg=T.c["bg1"],
                     fg=T.c["t1"], font=T.f("mono")).pack(side="left")

        tk.Frame(pad, bg=T.c["b0"], height=1).pack(fill="x", pady=20)
        btn_row = tk.Frame(pad, bg=T.c["bg1"])
        btn_row.pack(anchor="w")
        Btn(btn_row, "Зберегти зміни", style="primary",
            command=lambda: Toast.show("Збережено","ok")).pack(
            side="left", ipady=3)
        Btn(btn_row, "Скинути", style="danger",
            command=lambda: Toast.show("Скинуто до стандартних","warn")
            ).pack(side="left", padx=8, ipady=3)

    def _sec(self, parent, title: str):
        tk.Frame(parent, bg=T.c["b0"], height=1).pack(fill="x", pady=(18,6))
        tk.Label(parent, text=title, bg=T.c["bg1"],
                 fg=T.c["t3"], font=T.f("cap")).pack(anchor="w")


# ══════════════════════════════════════════════════════
#  DIALOGS
# ══════════════════════════════════════════════════════

class QuickAddDialog(Dialog):
    def __init__(self, parent):
        super().__init__(parent, "Швидке додавання", 460, 360)

    def body(self, m):
        Combo(m, "Тип запису",
              ["Контакт","Угода","Завдання","Ризик","Рахунок"]
              ).pack(fill="x", pady=(0,12))
        Input(m, "Назва *", required=True).pack(fill="x", pady=(0,12))
        row = tk.Frame(m, bg=m.cget("bg"))
        row.pack(fill="x")
        Input(row, "Відповідальний").pack(side="left", fill="x",
                                          expand=True, padx=(0,8))
        Input(row, "Термін").pack(side="left", fill="x", expand=True)

    def on_confirm(self) -> bool:
        Toast.show("Запис створено","ok")
        return True


class ContactFormDialog(Dialog):
    def __init__(self, parent, values: tuple | None = None):
        title = "Редагувати контакт" if values else "Новий контакт"
        super().__init__(parent, title, 540, 460)
        if values:
            self._prefill(values)

    def body(self, m):
        r1 = tk.Frame(m, bg=m.cget("bg"))
        r1.pack(fill="x", pady=(0,10))
        self._fn = Input(r1, "Імʼя *", required=True)
        self._fn.pack(side="left", fill="x", expand=True, padx=(0,8))
        self._ln = Input(r1, "Прізвище *", required=True)
        self._ln.pack(side="left", fill="x", expand=True)

        self._co = Input(m, "Компанія")
        self._co.pack(fill="x", pady=(0,10))

        r2 = tk.Frame(m, bg=m.cget("bg"))
        r2.pack(fill="x", pady=(0,10))
        self._em = Input(r2, "Email", icon="✉")
        self._em.pack(side="left", fill="x", expand=True, padx=(0,8))
        self._ph = Input(r2, "Телефон", icon="☎")
        self._ph.pack(side="left", fill="x", expand=True)

        r3 = tk.Frame(m, bg=m.cget("bg"))
        r3.pack(fill="x")
        self._tp = Combo(r3, "Тип",
                         ["Клієнт","Партнер","Лід","VIP"])
        self._tp.pack(side="left", fill="x", expand=True, padx=(0,8))
        self._st = Combo(r3, "Статус",
                         ["Активний","Неактивний","Новий"])
        self._st.pack(side="left", fill="x", expand=True)

    def _prefill(self, v: tuple):
        pass  # заповнити поля з tuple

    def on_confirm(self) -> bool:
        if not self._fn.get():
            Toast.show("Вкажіть імʼя","warn")
            return False
        Toast.show("Контакт збережено","ok")
        return True


class RiskMatrixDialog(Dialog):
    """Матриця ризиків 5×5."""
    def __init__(self, parent):
        super().__init__(parent, "Матриця ризиків", 540, 480)

    def body(self, m):
        tk.Label(m, text="Вплив →",
                 bg=m.cget("bg"), fg=T.c["t2"],
                 font=T.f("cap")).pack(anchor="e")

        grid = tk.Frame(m, bg=m.cget("bg"))
        grid.pack(fill="both", expand=True, pady=8)

        labels_x = ["Незначний","Малий","Середній","Великий","Критичний"]
        labels_y = ["Майже неможл.","Малоймовірно","Можливо",
                    "Ймовірно","Майже напевно"]

        colors = [
            [T.c["green"],T.c["green"],T.c["amber"],T.c["amber"],T.c["red"]],
            [T.c["green"],T.c["amber"],T.c["amber"],T.c["red"],  T.c["red"]],
            [T.c["amber"],T.c["amber"],T.c["red"],  T.c["red"],  T.c["red"]],
            [T.c["amber"],T.c["red"],  T.c["red"],  T.c["red"],  T.c["red"]],
            [T.c["red"],  T.c["red"],  T.c["red"],  T.c["red"],  T.c["red"]],
        ]

        # header row
        tk.Label(grid, text="", bg=m.cget("bg"),
                 width=12).grid(row=0, column=0)
        for c, lbl in enumerate(labels_x):
            tk.Label(grid, text=lbl, bg=m.cget("bg"),
                     fg=T.c["t2"], font=T.f("ui_xs"),
                     width=9).grid(row=0, column=c+1)

        for r in range(5):
            tk.Label(grid, text=labels_y[4-r], bg=m.cget("bg"),
                     fg=T.c["t2"], font=T.f("ui_xs"),
                     anchor="e", width=13).grid(row=r+1, column=0, padx=(0,4))
            for c in range(5):
                cell = tk.Frame(grid, bg=colors[4-r][c],
                                width=58, height=36)
                cell.pack_propagate(False)
                cell.grid(row=r+1, column=c+1, padx=2, pady=2)
                score = (5-r) * (c+1)
                tk.Label(cell, text=str(score), bg=colors[4-r][c],
                         fg="#000000AA" if score < 15 else T.c["t0"],
                         font=T.f("ui_b")).place(relx=.5,rely=.5,anchor="center")

    def on_confirm(self) -> bool:
        return True


# ══════════════════════════════════════════════════════
#  REGISTRY
# ══════════════════════════════════════════════════════

REGISTRY: dict[str, type] = {
    "dashboard": Dashboard,
    "contacts":  Contacts,
    "deals":     Deals,
    "finance":   Finance,
    "risks":     Risks,
    "settings":  Settings,
}

# ══════════════════════════════════════════════════════
#  NAVIGATION STRUCTURE
# ══════════════════════════════════════════════════════

NAV = [
    {"id":"dashboard", "label":"Dashboard",   "icon":"⊞"},
    {"label":"CRM", "icon":"◈", "children":[
        {"id":"contacts",    "label":"Контакти",  "icon":"◉"},
        {"id":"deals",       "label":"Угоди",     "icon":"◇"},
        {"id":"leads",       "label":"Ліди",      "icon":"◎"},
        {"id":"activities",  "label":"Активності","icon":"◆"},
    ]},
    {"label":"ERP", "icon":"◫", "children":[
        {"id":"finance",     "label":"Фінанси",   "icon":"₴"},
        {"id":"inventory",   "label":"Склад",     "icon":"▣"},
        {"id":"hr",          "label":"HR",        "icon":"◉"},
        {"id":"projects",    "label":"Проєкти",   "icon":"◇"},
    ]},
    {"label":"GRC", "icon":"◬", "children":[
        {"id":"risks",       "label":"Ризики",    "icon":"△"},
        {"id":"compliance",  "label":"Комплаєнс", "icon":"☑"},
        {"id":"audit",       "label":"Аудит",     "icon":"◎"},
        {"id":"policies",    "label":"Політики",  "icon":"▤"},
    ]},
    {"id":"settings","label":"Налаштування","icon":"⚙"},
]


# ══════════════════════════════════════════════════════
#  TOP BAR
# ══════════════════════════════════════════════════════

class TopBar(tk.Frame):
    def __init__(self, parent,
                 on_burger: Callable,
                 on_search: Callable):
        super().__init__(parent,
            bg=T.c["bg_topbar"],
            height=T.sz("topbar_h"))
        self.pack_propagate(False)

        # left: burger + logo
        left = tk.Frame(self, bg=T.c["bg_topbar"])
        left.pack(side="left", fill="y", padx=(8,0))

        self._burger = tk.Button(left, text="☰",
            bg=T.c["bg_topbar"], fg=T.c["t2"],
            relief="flat", bd=0, cursor="hand2",
            font=("Segoe UI",15),
            activebackground=T.c["bg3"],
            activeforeground=T.c["t0"],
            command=on_burger)
        self._burger.pack(side="left", padx=(4,10))

        # logo
        logo_box = tk.Frame(left, bg=T.c["a"],
                            width=26, height=26)
        logo_box.pack_propagate(False)
        logo_box.pack(side="left")
        tk.Label(logo_box, text="N", bg=T.c["a"],
                 fg=T.c["bg0"], font=("Segoe UI",12,"bold")
                 ).place(relx=.5,rely=.5,anchor="center")

        tk.Label(left, text=APP["name"],
                 bg=T.c["bg_topbar"], fg=T.c["t0"],
                 font=T.f("logo")).pack(side="left", padx=(8,2))
        tk.Label(left, text="v"+APP["version"],
                 bg=T.c["bg_topbar"], fg=T.c["t3"],
                 font=T.f("ui_xs")).pack(side="left", padx=(0,20))

        # center: search
        mid = tk.Frame(self, bg=T.c["bg_topbar"])
        mid.pack(side="left", fill="both", expand=True, padx=10)

        search_wrap = tk.Frame(mid,
            bg=T.c["bg3"],
            highlightthickness=1,
            highlightbackground=T.c["b1"])
        search_wrap.pack(side="left", fill="y", pady=9, ipadx=2)
        tk.Label(search_wrap, text="⌕",
                 bg=T.c["bg3"], fg=T.c["t2"],
                 font=("Segoe UI",11),
                 padx=8).pack(side="left")
        self._sv = tk.StringVar()
        self._se = tk.Entry(search_wrap,
            textvariable=self._sv,
            bg=T.c["bg3"], fg=T.c["t0"],
            insertbackground=T.c["a"],
            relief="flat", bd=0, width=34,
            font=T.f("ui"))
        self._se.pack(side="left", pady=1, padx=(0,8))
        self._se.insert(0,"Пошук…")
        self._se.configure(fg=T.c["t3"])
        self._se.bind("<FocusIn>",  self._sf_in)
        self._se.bind("<FocusOut>", self._sf_out)
        self._se.bind("<Return>",   lambda _: on_search(self._sv.get()))
        tk.Label(search_wrap, text="Ctrl+F",
                 bg=T.c["bg3"], fg=T.c["t3"],
                 font=T.f("ui_xs"), padx=6).pack(side="left")

        # right: actions
        right = tk.Frame(self, bg=T.c["bg_topbar"])
        right.pack(side="right", fill="y", padx=(0,10))

        # новий запис
        Btn(right, "Новий", icon="+", style="primary", size="sm",
            command=lambda: QuickAddDialog(self.winfo_toplevel()),
            tip="Ctrl+N"
            ).pack(side="left", padx=(0,8), ipady=2)

        # тема
        self._tb = tk.Button(right, text="◑",
            bg=T.c["bg_topbar"], fg=T.c["t2"],
            relief="flat", bd=0, cursor="hand2",
            font=("Segoe UI",13),
            activebackground=T.c["bg3"],
            activeforeground=T.c["t0"],
            command=T.toggle)
        self._tb.pack(side="left", padx=4)
        Tip(self._tb, "Ctrl+T — перемикач теми")

        # сповіщення
        nb = tk.Button(right, text="🔔",
            bg=T.c["bg_topbar"], fg=T.c["t2"],
            relief="flat", bd=0, cursor="hand2",
            font=("Segoe UI",11),
            activebackground=T.c["bg3"],
            command=lambda: Toast.show("Нових сповіщень немає","info"))
        nb.pack(side="left", padx=4)

        Sep(right, "v").pack(side="left", fill="y", pady=10, padx=6)

        # profile
        pf = tk.Frame(right, bg=T.c["bg_topbar"], cursor="hand2")
        pf.pack(side="left")
        av = tk.Frame(pf, bg=T.c["a2"], width=28, height=28)
        av.pack_propagate(False)
        av.pack(side="left")
        tk.Label(av, text=APP["user_init"],
                 bg=T.c["a2"], fg="#FFF",
                 font=("Segoe UI",9,"bold")
                 ).place(relx=.5,rely=.5,anchor="center")
        pf_t = tk.Frame(pf, bg=T.c["bg_topbar"])
        pf_t.pack(side="left", padx=(7,0))
        tk.Label(pf_t, text=APP["user"],
                 bg=T.c["bg_topbar"], fg=T.c["t0"],
                 font=T.f("ui_b")).pack(anchor="w")
        tk.Label(pf_t, text=APP["user_role"],
                 bg=T.c["bg_topbar"], fg=T.c["t3"],
                 font=T.f("ui_xs")).pack(anchor="w")

        # separator bottom — повертається як атрибут, App пакує окремо
        self._sep = tk.Frame(parent, bg=T.c["b0"], height=1)

    def _sf_in(self, _=None):
        if self._se.get() == "Пошук…":
            self._se.delete(0,"end")
            self._se.configure(fg=T.c["t0"])

    def _sf_out(self, _=None):
        if not self._se.get():
            self._se.insert(0,"Пошук…")
            self._se.configure(fg=T.c["t3"])

    def focus_search(self):
        self._se.focus_set()
        self._sf_in()


# ══════════════════════════════════════════════════════
#  SIDE BAR
# ══════════════════════════════════════════════════════

class SideBar(tk.Frame):
    def __init__(self, parent,
                 on_nav: Callable,
                 nav:    list):
        super().__init__(parent,
            bg=T.c["bg1"],
            width=T.sz("sidebar_w"))
        self.pack_propagate(False)
        self._on_nav   = on_nav
        self._nav      = nav
        self._active   = "dashboard"
        self._expanded: set = {"CRM"}  # відкриті групи
        self._collapsed = False
        self._widgets:  dict = {}
        self._build()

    # ── build ──────────────────────────────────────────
    def _build(self):
        for w in self.winfo_children():
            w.destroy()
        self._widgets.clear()

        # top padding
        tk.Frame(self, bg=T.c["bg1"], height=8).pack(fill="x")

        for item in self._nav:
            if "children" in item:
                self._add_group(item)
            else:
                self._add_item(item)

        # bottom
        tk.Frame(self, bg=T.c["b0"], height=1).pack(
            fill="x", side="bottom")
        bot = tk.Frame(self, bg=T.c["bg1"], pady=8)
        bot.pack(side="bottom", fill="x")
        tk.Label(bot, text=f"NEXUS  v{APP['version']}",
                 bg=T.c["bg1"], fg=T.c["t3"],
                 font=T.f("ui_xs")).pack()

    def _add_group(self, item: dict):
        grp   = item["label"]
        icon  = item.get("icon","·")
        kids  = item.get("children",[])
        open_ = grp in self._expanded

        # group header
        hdr = tk.Frame(self, bg=T.c["bg1"], cursor="hand2")
        hdr.pack(fill="x", pady=(6,0))
        tk.Label(hdr, text=icon,
                 bg=T.c["bg1"], fg=T.c["t3"],
                 font=("Segoe UI",10),
                 padx=12).pack(side="left")
        tk.Label(hdr, text=grp.upper(),
                 bg=T.c["bg1"], fg=T.c["t3"],
                 font=T.f("cap")).pack(side="left", fill="x", expand=True)
        arr_var = tk.StringVar(value="▾" if open_ else "›")
        arr = tk.Label(hdr, textvariable=arr_var,
                       bg=T.c["bg1"], fg=T.c["t3"],
                       font=T.f("ui_sm"), padx=10)
        arr.pack(side="right")

        # children container
        kids_frame = tk.Frame(self, bg=T.c["bg1"])
        if open_:
            kids_frame.pack(fill="x")

        for child in kids:
            self._add_item(child, kids_frame, depth=1)

        def toggle(_=None, g=grp, kf=kids_frame, av=arr_var):
            if g in self._expanded:
                self._expanded.discard(g)
                kf.pack_forget()
                av.set("›")
            else:
                self._expanded.add(g)
                kf.pack(fill="x")
                av.set("▾")

        for w in (hdr, arr):
            w.bind("<Button-1>", toggle)
            w.bind("<Enter>",
                   lambda _, h=hdr: h.configure(bg=T.c["bg3"]) or
                   [c.configure(bg=T.c["bg3"]) for c in h.winfo_children()])
            w.bind("<Leave>",
                   lambda _, h=hdr: h.configure(bg=T.c["bg1"]) or
                   [c.configure(bg=T.c["bg1"]) for c in h.winfo_children()])

    def _add_item(self, item: dict,
                  parent: tk.Widget | None = None,
                  depth: int = 0):
        parent = parent or self
        iid   = item.get("id","")
        icon  = item.get("icon","·")
        label = item.get("label","")
        is_act = (iid == self._active)

        bg  = T.c["bg4"] if is_act else T.c["bg1"]
        fg  = T.c["t0"]  if is_act else T.c["t1"]

        row = tk.Frame(parent, bg=bg, cursor="hand2")
        row.pack(fill="x", pady=(1,0))

        # active indicator
        ind = tk.Frame(row, bg=T.c["a"] if is_act else bg, width=3)
        ind.pack(side="left", fill="y")

        if depth:
            tk.Frame(row, bg=bg, width=depth*12).pack(side="left")

        tk.Label(row, text=icon, bg=bg, fg=T.c["a"] if is_act else T.c["t3"],
                 font=("Segoe UI",11), padx=8,
                 pady=8).pack(side="left")
        lbl = tk.Label(row, text=label, bg=bg, fg=fg,
                       font=T.f("ui_b" if is_act else "ui"),
                       anchor="w")
        lbl.pack(side="left", fill="x", expand=True, pady=8)

        self._widgets[iid] = {"row": row, "ind": ind, "lbl": lbl}

        def click(_=None, _id=iid):
            self._set_active(_id)
            self._on_nav(_id)

        def enter(_=None, _id=iid, r=row):
            if _id != self._active:
                r.configure(bg=T.c["bg3"])
                for c in r.winfo_children():
                    try: c.configure(bg=T.c["bg3"])
                    except: pass

        def leave(_=None, _id=iid, r=row):
            if _id != self._active:
                r.configure(bg=T.c["bg1"])
                for c in r.winfo_children():
                    try: c.configure(bg=T.c["bg1"])
                    except: pass

        for w in (row, lbl):
            w.bind("<Button-1>", click)
            w.bind("<Enter>",    enter)
            w.bind("<Leave>",    leave)

    def _set_active(self, iid: str):
        old = self._widgets.get(self._active)
        if old:
            old["row"].configure(bg=T.c["bg1"])
            old["ind"].configure(bg=T.c["bg1"])
            old["lbl"].configure(fg=T.c["t1"], font=T.f("ui"),
                                  bg=T.c["bg1"])
            for c in old["row"].winfo_children():
                try: c.configure(bg=T.c["bg1"])
                except: pass

        self._active = iid
        new = self._widgets.get(iid)
        if new:
            new["row"].configure(bg=T.c["bg4"])
            new["ind"].configure(bg=T.c["a"])
            new["lbl"].configure(fg=T.c["t0"], font=T.f("ui_b"),
                                  bg=T.c["bg4"])
            for c in new["row"].winfo_children():
                try: c.configure(bg=T.c["bg4"])
                except: pass
            new["lbl"].configure(bg=T.c["bg4"])

    def toggle_collapse(self):
        self._collapsed = not self._collapsed
        self.configure(width=T.sz("sidebar_c")
                       if self._collapsed else T.sz("sidebar_w"))


# ══════════════════════════════════════════════════════
#  STATUS BAR
# ══════════════════════════════════════════════════════

class StatusBar(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent,
            bg=T.c["bg1"],
            height=T.sz("statusbar_h"))
        self.pack_propagate(False)
        self._sep = tk.Frame(parent, bg=T.c["b0"], height=1)

        self._sv  = tk.StringVar(value="Готово")
        self._tv  = tk.StringVar()

        tk.Label(self, textvariable=self._sv,
                 bg=T.c["bg1"], fg=T.c["t2"],
                 font=T.f("ui_xs"), padx=12).pack(side="left",fill="y")
        tk.Label(self, text="●  З'єднання активне",
                 bg=T.c["bg1"], fg=T.c["green"],
                 font=T.f("ui_xs")).pack(side="left", padx=12)
        tk.Label(self, textvariable=self._tv,
                 bg=T.c["bg1"], fg=T.c["t3"],
                 font=T.f("mono"), padx=12).pack(side="right", fill="y")
        self._tick()

    def msg(self, text: str):
        self._sv.set(text)

    def _tick(self):
        self._tv.set(datetime.now().strftime("%H:%M:%S   %d.%m.%Y"))
        self.after(1000, self._tick)


# ══════════════════════════════════════════════════════
#  WORKSPACE (content router)
# ══════════════════════════════════════════════════════

class Workspace(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent, bg=T.c["bg1"])
        self._cache: dict[str, BaseModule] = {}
        self._current: str | None = None

    def show(self, mid: str):
        if self._current == mid:
            return
        if self._current and self._current in self._cache:
            self._cache[self._current].pack_forget()

        if mid not in self._cache:
            cls = REGISTRY.get(mid, BaseModule)
            inst = cls(self)
            self._cache[mid] = inst

        self._cache[mid].pack(fill="both", expand=True)
        self._cache[mid].activate()
        self._current = mid


# ══════════════════════════════════════════════════════
#  APPLICATION SHELL
# ══════════════════════════════════════════════════════

class Shell(tk.Tk):
    def __init__(self):
        super().__init__()

        # HiDPI Windows
        if platform.system() == "Windows":
            try:
                from ctypes import windll
                windll.shcore.SetProcessDpiAwareness(1)
            except: pass

        self.title(f"{APP['name']}  —  {APP['subtitle']}")
        self.geometry("1320x800")
        self.minsize(960, 620)
        self.configure(bg=T.c["bg0"])

        try:
            img = tk.PhotoImage(width=1, height=1)
            self.iconphoto(True, img)
        except: pass

        Toast.init(self)
        T.sub(self._rebuild)

        self._build()
        self._bind_keys()
        self.after(600, lambda: Toast.show(
            f"Ласкаво просимо, {APP['user']}!","ok"))

    # ── build ─────────────────────────────────────────
    def _build(self):
        self.configure(bg=T.c["bg0"])

        # topbar
        self._topbar = TopBar(self,
            on_burger=self._toggle_sidebar,
            on_search=self._search)
        self._topbar.pack(fill="x", side="top")
        self._topbar._sep.pack(fill="x", side="top")

        # main row
        self._main = tk.Frame(self, bg=T.c["bg0"])
        self._main.pack(fill="both", expand=True, side="top")

        # sidebar
        self._sidebar = SideBar(self._main,
            on_nav=self._navigate, nav=NAV)
        self._sidebar.pack(side="left", fill="y")

        Sep(self._main, "v").pack(side="left", fill="y")

        # workspace
        self._workspace = Workspace(self._main)
        self._workspace.pack(side="left", fill="both", expand=True)

        # statusbar
        self._sb = StatusBar(self)
        self._sb._sep.pack(fill="x", side="bottom")
        self._sb.pack(fill="x", side="bottom")

        # default view
        self._navigate("dashboard")

    # ── rebuild on theme change ────────────────────────
    def _rebuild(self):
        for w in self.winfo_children():
            w.destroy()
        self._build()
        Toast.show("Тема змінена","info")

    # ── navigation ─────────────────────────────────────
    def _navigate(self, mid: str):
        self._workspace.show(mid)
        self._sb.msg(f"Відкрито: {mid}")

    def _toggle_sidebar(self):
        self._sidebar.toggle_collapse()

    def _search(self, q: str):
        if q and q != "Пошук…":
            Toast.show(f"Пошук: «{q}»","info")
            self._sb.msg(f"Пошук: {q}")

    # ── hotkeys ────────────────────────────────────────
    def _bind_keys(self):
        ids = ["dashboard","contacts","deals","finance",
               "risks","settings"]
        for i, mid in enumerate(ids, 1):
            self.bind(f"<Control-Key-{i}>",
                      lambda _, m=mid: self._navigate(m))
        self.bind("<Control-t>",  lambda _: T.toggle())
        self.bind("<Control-T>",  lambda _: T.toggle())
        self.bind("<Control-n>",  lambda _: QuickAddDialog(self))
        self.bind("<Control-N>",  lambda _: QuickAddDialog(self))
        self.bind("<Control-f>",  lambda _: self._topbar.focus_search())
        self.bind("<Control-F>",  lambda _: self._topbar.focus_search())
        self.bind("<F5>",         lambda _: (
            Toast.show("Оновлення…","info"),
            self._workspace._cache.get(
                self._workspace._current,
                BaseModule(self)).refresh()))
        self.bind("<Control-q>",  lambda _: self.quit())

    def run(self):
        self.mainloop()


# ══════════════════════════════════════════════════════
#  ENTRY POINT
# ══════════════════════════════════════════════════════

if __name__ == "__main__":
    Shell().run()
