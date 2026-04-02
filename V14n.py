"""
Archer GRC Platform v3.0
Сучасний десктопний інтерфейс — максимально близький до HTML/CSS
Rounded cards, custom widgets, smooth typography
"""

import tkinter as tk
from tkinter import ttk, messagebox
import sqlite3, os, datetime, random, math

# ══════════════════════════════════════════════════════
#  ДИЗАЙН-ТОКЕНИ
# ══════════════════════════════════════════════════════
T = {
    # Фони
    "bg":       "#0A0A0F",
    "surface":  "#12121A",
    "card":     "#1A1A26",
    "card2":    "#1F1F2E",
    "input":    "#16161F",
    "hover":    "#22222F",
    # Акценти
    "orange":   "#F97316",
    "orange2":  "#FB923C",
    "orange_bg":"#F9731612",
    # Статуси
    "green":    "#22C55E",
    "green_bg": "#22C55E14",
    "yellow":   "#EAB308",
    "yellow_bg":"#EAB30814",
    "red":      "#EF4444",
    "red_bg":   "#EF444414",
    "blue":     "#3B82F6",
    "blue_bg":  "#3B82F614",
    "purple":   "#A855F7",
    # Текст
    "t1":  "#F1F5F9",
    "t2":  "#94A3B8",
    "t3":  "#475569",
    "t4":  "#2D3748",
    # Межі
    "b1":  "#1E2030",
    "b2":  "#2A2D3E",
    "b3":  "#363950",
    # Розміри
    "r":   8,   # border-radius
    "r2":  12,
}

STATUS_MAP = {
    "Відкритий":             (T["red"],    T["red_bg"]),
    "Пом'якшений":           (T["yellow"], T["yellow_bg"]),
    "Закритий":              (T["green"],  T["green_bg"]),
    "Прийнятий":             (T["t3"],     T["card2"]),
    "Активний":              (T["green"],  T["green_bg"]),
    "Неактивний":            (T["t3"],     T["card2"]),
    "Ефективний":            (T["green"],  T["green_bg"]),
    "Частково ефективний":   (T["yellow"], T["yellow_bg"]),
    "Неефективний":          (T["red"],    T["red_bg"]),
    "Не оцінено":            (T["t3"],     T["card2"]),
    "Опублікована":          (T["green"],  T["green_bg"]),
    "Чернетка":              (T["t3"],     T["card2"]),
    "На затвердженні":       (T["yellow"], T["yellow_bg"]),
    "Архівована":            (T["t3"],     T["card2"]),
    "Новий":                 (T["blue"],   T["blue_bg"]),
    "В роботі":              (T["yellow"], T["yellow_bg"]),
    "Вирішений":             (T["green"],  T["green_bg"]),
    "Скасовано":             (T["t3"],     T["card2"]),
    "Заплановано":           (T["blue"],   T["blue_bg"]),
    "В процесі":             (T["yellow"], T["yellow_bg"]),
    "Завершено":             (T["green"],  T["green_bg"]),
    "Відповідає":            (T["green"],  T["green_bg"]),
    "Частково відповідає":   (T["yellow"], T["yellow_bg"]),
    "Не відповідає":         (T["red"],    T["red_bg"]),
    "Відкрита":              (T["red"],    T["red_bg"]),
    "Закрита":               (T["green"],  T["green_bg"]),
    "Низький":               (T["green"],  T["green_bg"]),
    "Середній":              (T["yellow"], T["yellow_bg"]),
    "Високий":               (T["red"],    T["red_bg"]),
    "Критичний":             ("#FF0000",   "#FF000018"),
}

def risk_col(score):
    if score <= 4:  return T["green"]
    if score <= 9:  return T["yellow"]
    if score <= 16: return T["red"]
    return "#FF0000"

# ══════════════════════════════════════════════════════
#  БАЗА ДАНИХ
# ══════════════════════════════════════════════════════
DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "archer_grc.db")

def get_db():
    c = sqlite3.connect(DB_PATH)
    c.row_factory = sqlite3.Row
    c.execute("PRAGMA foreign_keys=ON")
    return c

def init_db():
    db = get_db()
    db.executescript("""
    CREATE TABLE IF NOT EXISTS users(id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT UNIQUE NOT NULL, password TEXT NOT NULL,
        full_name TEXT NOT NULL, role TEXT NOT NULL,
        email TEXT, department TEXT, active INTEGER DEFAULT 1,
        created_at TEXT DEFAULT(datetime('now')));
    CREATE TABLE IF NOT EXISTS risks(id INTEGER PRIMARY KEY AUTOINCREMENT,
        title TEXT NOT NULL, description TEXT, category TEXT,
        likelihood INTEGER DEFAULT 3, impact INTEGER DEFAULT 3,
        risk_score INTEGER GENERATED ALWAYS AS(likelihood*impact) STORED,
        status TEXT DEFAULT 'Відкритий', owner_id INTEGER, department TEXT,
        mitigation_plan TEXT, review_date TEXT,
        created_at TEXT DEFAULT(datetime('now')),
        updated_at TEXT DEFAULT(datetime('now')),
        FOREIGN KEY(owner_id) REFERENCES users(id));
    CREATE TABLE IF NOT EXISTS controls(id INTEGER PRIMARY KEY AUTOINCREMENT,
        title TEXT NOT NULL, description TEXT, type TEXT, frequency TEXT,
        owner_id INTEGER, status TEXT DEFAULT 'Активний',
        effectiveness TEXT DEFAULT 'Не оцінено', last_tested TEXT,
        created_at TEXT DEFAULT(datetime('now')),
        FOREIGN KEY(owner_id) REFERENCES users(id));
    CREATE TABLE IF NOT EXISTS policies(id INTEGER PRIMARY KEY AUTOINCREMENT,
        title TEXT NOT NULL, description TEXT, category TEXT,
        version TEXT DEFAULT '1.0', status TEXT DEFAULT 'Чернетка',
        owner_id INTEGER, review_date TEXT, content TEXT,
        created_at TEXT DEFAULT(datetime('now')),
        FOREIGN KEY(owner_id) REFERENCES users(id));
    CREATE TABLE IF NOT EXISTS suppliers(id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL, category TEXT, contact_name TEXT, contact_email TEXT,
        risk_level TEXT DEFAULT 'Середній', status TEXT DEFAULT 'Активний',
        contract_start TEXT, contract_end TEXT, notes TEXT,
        created_at TEXT DEFAULT(datetime('now')));
    CREATE TABLE IF NOT EXISTS audits(id INTEGER PRIMARY KEY AUTOINCREMENT,
        title TEXT NOT NULL, type TEXT, auditor_id INTEGER, department TEXT,
        start_date TEXT, end_date TEXT, status TEXT DEFAULT 'Заплановано',
        scope TEXT, conclusion TEXT, created_at TEXT DEFAULT(datetime('now')),
        FOREIGN KEY(auditor_id) REFERENCES users(id));
    CREATE TABLE IF NOT EXISTS incidents(id INTEGER PRIMARY KEY AUTOINCREMENT,
        title TEXT NOT NULL, description TEXT, category TEXT,
        severity TEXT DEFAULT 'Середня', status TEXT DEFAULT 'Новий',
        reporter_id INTEGER, owner_id INTEGER, occurred_at TEXT,
        root_cause TEXT, corrective_action TEXT,
        created_at TEXT DEFAULT(datetime('now')),
        FOREIGN KEY(reporter_id) REFERENCES users(id),
        FOREIGN KEY(owner_id) REFERENCES users(id));
    CREATE TABLE IF NOT EXISTS regulations(id INTEGER PRIMARY KEY AUTOINCREMENT,
        title TEXT NOT NULL, authority TEXT, category TEXT, description TEXT,
        effective_date TEXT, compliance_status TEXT DEFAULT 'Не оцінено',
        responsible_id INTEGER, notes TEXT,
        created_at TEXT DEFAULT(datetime('now')),
        FOREIGN KEY(responsible_id) REFERENCES users(id));
    CREATE TABLE IF NOT EXISTS audit_log(id INTEGER PRIMARY KEY AUTOINCREMENT,
        user_id INTEGER, action TEXT, module TEXT, record_id INTEGER,
        details TEXT, created_at TEXT DEFAULT(datetime('now')));
    """)
    db.execute("SELECT id FROM users WHERE username='admin'")
    if not db.execute("SELECT id FROM users WHERE username='admin'").fetchone():
        for row in [
            ('admin','admin123','Системний Адміністратор','Системний адміністратор','admin@co.ua','ІТ'),
            ('risk_officer','pass123','Іваненко Олексій','Власник ризику','risk@co.ua','Ризики'),
            ('auditor','pass123','Петренко Марія','Аудитор','audit@co.ua','Аудит'),
            ('compliance','pass123','Коваль Дмитро','Комплаєнс-офіцер','comp@co.ua','Комплаєнс'),
        ]:
            db.execute("INSERT INTO users(username,password,full_name,role,email,department) VALUES(?,?,?,?,?,?)", row)
        for row in [
            ('Витік персональних даних','Несанкціонований доступ до ПДн','Кіберризик',4,5,'Відкритий',2,'ІТ','Впровадження DLP','2026-06-30'),
            ('Збій ключових ІТ-систем','Відмова банківських систем','Операційний',2,5,"Пом'якшений",2,'ІТ','Резервне копіювання','2026-09-30'),
            ('Недотримання вимог НБУ','Ризик штрафних санкцій','Регуляторний',2,4,'Відкритий',4,'Комплаєнс','Моніторинг нормативів','2026-07-31'),
            ('Шахрайство третіх осіб','Ризик шахрайських дій','Фінансовий',3,4,'Відкритий',2,'Безпека','Due diligence','2026-08-31'),
            ('Кібератака DDoS','Атака на відмову в обслуговуванні','Кіберризик',3,3,'Відкритий',2,'ІТ','Anti-DDoS','2026-05-31'),
            ('Фішинг атаки','Соціальна інженерія','Кіберризик',4,3,'Відкритий',2,'ІТ','Навчання','2026-04-30'),
        ]:
            db.execute("INSERT INTO risks(title,description,category,likelihood,impact,status,owner_id,department,mitigation_plan,review_date) VALUES(?,?,?,?,?,?,?,?,?,?)", row)
        for row in [
            ('Антивірусний захист','Встановлено на всіх ПК','Превентивний','Постійно',2,'Активний','Ефективний','2026-03-15'),
            ('Резервне копіювання','Щоденний backup','Відновлювальний','Щодня',2,'Активний','Ефективний','2026-03-20'),
            ('Розмежування доступу','RBAC для систем','Превентивний','Постійно',4,'Активний','Частково ефективний','2026-02-28'),
            ('SIEM моніторинг','Журналювання аномалій','Детективний','Постійно',2,'Активний','Ефективний','2026-03-01'),
            ('Навчання персоналу','Тренінги з кібербезпеки','Превентивний','Щоквартально',4,'Активний','Не оцінено','2026-01-15'),
        ]:
            db.execute("INSERT INTO controls(title,description,type,frequency,owner_id,status,effectiveness,last_tested) VALUES(?,?,?,?,?,?,?,?)", row)
        for row in [
            ('Політика інформаційної безпеки','Основний документ ІБ','Інформаційна безпека','2.1','Опублікована',4,'2026-12-31'),
            ('Політика управління паролями','Вимоги до паролів','Інформаційна безпека','1.3','Опублікована',4,'2026-06-30'),
            ('Антикорупційна політика','Запобігання корупції','Комплаєнс','1.0','На затвердженні',4,'2027-01-31'),
            ('Політика BYOD','Особисті пристрої','ІТ','1.1','Чернетка',2,'2026-09-30'),
        ]:
            db.execute("INSERT INTO policies(title,description,category,version,status,owner_id,review_date) VALUES(?,?,?,?,?,?,?)", row)
        for row in [
            ('ТОВ "ТехноСофт"','ІТ-послуги','Сидоренко В.В.','v.s@ts.ua','Середній','Активний','2025-01-01','2026-12-31'),
            ('АТ "CloudServ"','Хмарна інфраструктура','Мороз О.П.','o.m@cs.ua','Високий','Активний','2024-07-01','2026-06-30'),
            ('ФОП Бондаренко','Консалтинг','Бондаренко І.С.','i.b@gm.com','Низький','Активний','2026-01-01','2026-12-31'),
        ]:
            db.execute("INSERT INTO suppliers(name,category,contact_name,contact_email,risk_level,status,contract_start,contract_end) VALUES(?,?,?,?,?,?,?,?)", row)
        for row in [
            ('Положення НБУ №95','НБУ','Кіберзахист','Відповідає',4,'2022-01-01'),
            ('Закон про захист ПДн','Верховна Рада','Персональні дані','Частково відповідає',4,'2010-06-01'),
            ('ISO/IEC 27001:2022','ISO','Інформаційна безпека','В процесі',4,'2022-10-25'),
            ('PCI DSS v4.0','PCI SSC','Платіжні системи','Відповідає',4,'2022-03-31'),
            ('GDPR','ЄС','Персональні дані','Не оцінено',4,'2018-05-25'),
        ]:
            db.execute("INSERT INTO regulations(title,authority,category,compliance_status,responsible_id,effective_date) VALUES(?,?,?,?,?,?)", row)
    db.commit(); db.close()


# ══════════════════════════════════════════════════════
#  ПРИМІТИВНІ UI-КОМПОНЕНТИ
# ══════════════════════════════════════════════════════

def rounded_rect(canvas, x1, y1, x2, y2, r=8, **kw):
    """Малює заокруглений прямокутник на Canvas."""
    pts = [
        x1+r, y1,  x2-r, y1,
        x2,   y1,  x2,   y1+r,
        x2,   y2-r,x2,   y2,
        x2-r, y2,  x1+r, y2,
        x1,   y2,  x1,   y2-r,
        x1,   y1+r,x1,   y1,
    ]
    return canvas.create_polygon(pts, smooth=True, **kw)


class RoundedFrame(tk.Canvas):
    """Frame з заокругленими кутами через Canvas."""
    def __init__(self, parent, radius=T["r"], bg=T["card"], border_color=T["b2"],
                 border_width=1, **kw):
        kw.setdefault("highlightthickness", 0)
        kw.setdefault("bd", 0)
        super().__init__(parent, bg=parent.cget("bg"), **kw)
        self._r   = radius
        self._bg  = bg
        self._bc  = border_color
        self._bw  = border_width
        self._rect = None
        self.bind("<Configure>", self._redraw)

    def _redraw(self, _=None):
        self.delete("all")
        w, h = self.winfo_width(), self.winfo_height()
        if w < 2 or h < 2: return
        bw = self._bw
        # shadow effect (subtle)
        for i in range(3, 0, -1):
            rounded_rect(self, i, i, w-i+3, h-i+3, self._r+1,
                         fill="#00000022", outline="")
        # border
        if bw > 0:
            rounded_rect(self, 0, 0, w-1, h-1, self._r,
                         fill=self._bc, outline="")
        # fill
        rounded_rect(self, bw, bw, w-1-bw, h-1-bw, max(1, self._r-bw),
                     fill=self._bg, outline="")

    def _find_inner(self):
        # Return inner frame placed on canvas
        for w in self.winfo_children():
            if isinstance(w, tk.Frame): return w
        f = tk.Frame(self, bg=self._bg)
        self.create_window(T["r"], T["r"], window=f, anchor="nw")
        self.bind("<Configure>", lambda e: (self._redraw(e),
                  self.itemconfig(self.find_withtag("all")[-1] if self.find_withtag("all") else 1,
                                  width=max(1, self.winfo_width()-T["r"]*2))))
        return f


class Pill(tk.Label):
    """Маленький badge з кольором відповідно до статусу."""
    def __init__(self, parent, text, **kw):
        fg, bg = STATUS_MAP.get(text, (T["t2"], T["card2"]))
        super().__init__(parent, text=f"  {text}  ",
                         font=("Segoe UI", 8, "bold"),
                         fg=fg, bg=bg,
                         relief="flat", padx=0, pady=3, **kw)


class ScoreBox(tk.Canvas):
    """Квадратний бокс з оцінкою ризику."""
    def __init__(self, parent, score, size=34, **kw):
        super().__init__(parent, width=size, height=size,
                         highlightthickness=0, bg=parent.cget("bg"), **kw)
        col = risk_col(score)
        rounded_rect(self, 1, 1, size-2, size-2, 5,
                     fill=col+"28", outline=col+"66", width=1)
        self.create_text(size//2, size//2, text=str(score),
                         font=("Segoe UI", 9, "bold"), fill=col)


class GlowButton(tk.Canvas):
    """Сучасна кнопка з hover-ефектом."""
    def __init__(self, parent, text, command=None, style="primary",
                 width=120, height=34, radius=6, **kw):
        super().__init__(parent, width=width, height=height,
                         highlightthickness=0, cursor="hand2",
                         bg=parent.cget("bg"), **kw)
        self._text = text
        self._cmd  = command
        self._w, self._h = width, height
        self._r    = radius
        self._styles = {
            "primary": (T["orange"],   "#000000", T["orange2"]),
            "ghost":   (T["card2"],    T["t2"],   T["hover"]),
            "danger":  (T["red"]+"CC", "#ffffff", T["red"]),
            "outline": ("",            T["orange"],T["hover"]),
        }
        self._style = style
        self._hover = False
        self._draw()
        self.bind("<Enter>",    self._on_enter)
        self.bind("<Leave>",    self._on_leave)
        self.bind("<Button-1>", self._on_click)

    def _draw(self):
        self.delete("all")
        bg, fg, hov = self._styles.get(self._style, self._styles["primary"])
        w, h, r = self._w, self._h, self._r
        fill = hov if self._hover else bg
        if self._style == "outline":
            rounded_rect(self, 0, 0, w-1, h-1, r,
                         fill=fill, outline=T["orange"], width=1)
        else:
            rounded_rect(self, 0, 0, w-1, h-1, r, fill=fill, outline="")
        self.create_text(w//2, h//2, text=self._text,
                         font=("Segoe UI", 9, "bold"),
                         fill=fg if not self._hover or self._style != "outline" else T["orange"])

    def _on_enter(self, _): self._hover=True;  self._draw()
    def _on_leave(self, _): self._hover=False; self._draw()
    def _on_click(self, _):
        if self._cmd: self._cmd()


class ModernEntry(tk.Frame):
    """Entry з підсвіченням при фокусі."""
    def __init__(self, parent, placeholder="", width=200, show=None, **kw):
        super().__init__(parent, bg=T["input"], pady=0)
        self._ph = placeholder
        self._focused = False

        self._border = tk.Frame(self, bg=T["b2"], pady=1)
        self._border.pack(fill="both", expand=True)
        inner = tk.Frame(self._border, bg=T["input"], padx=10)
        inner.pack(fill="both", expand=True)

        self.var = tk.StringVar()
        self.entry = tk.Entry(inner, textvariable=self.var,
                              bg=T["input"], fg=T["t1"],
                              insertbackground=T["orange"],
                              relief="flat", bd=0,
                              font=("Segoe UI", 9),
                              width=width, **kw)
        if show: self.entry.configure(show=show)
        self.entry.pack(fill="both", ipady=7)

        self.entry.bind("<FocusIn>",  self._focus_in)
        self.entry.bind("<FocusOut>", self._focus_out)

    def _focus_in(self, _):
        self._border.configure(bg=T["orange"])
    def _focus_out(self, _):
        self._border.configure(bg=T["b2"])

    def get(self):       return self.var.get()
    def insert(self, i, v): self.entry.insert(i, v)
    def delete(self, a, b): self.entry.delete(a, b)
    def configure(self, **kw): self.entry.configure(**kw)


class ModernCombo(ttk.Combobox):
    def __init__(self, parent, values, **kw):
        kw.setdefault("state", "readonly")
        kw.setdefault("font", ("Segoe UI", 9))
        super().__init__(parent, values=values, **kw)


class ModernText(tk.Frame):
    def __init__(self, parent, height=3, **kw):
        super().__init__(parent, bg=T["b2"], pady=1)
        inner = tk.Frame(self, bg=T["input"], padx=10)
        inner.pack(fill="both", expand=True)
        self.text = tk.Text(inner, height=height, bg=T["input"], fg=T["t1"],
                            insertbackground=T["orange"],
                            relief="flat", bd=0,
                            font=("Segoe UI", 9), wrap="word",
                            selectbackground=T["orange"]+"44", **kw)
        self.text.pack(fill="both", ipady=6)
        self.text.bind("<FocusIn>",  lambda _: self.configure(bg=T["orange"]))
        self.text.bind("<FocusOut>", lambda _: self.configure(bg=T["b2"]))

    def get(self, a="1.0", b="end-1c"): return self.text.get(a, b)
    def insert(self, i, v):             self.text.insert(i, v)


# ══════════════════════════════════════════════════════
#  TTK СТИЛЬ
# ══════════════════════════════════════════════════════

def apply_ttk_style(root):
    s = ttk.Style(root)
    s.theme_use("clam")
    s.configure("Treeview",
        background=T["card"], foreground=T["t1"],
        fieldbackground=T["card"], rowheight=34,
        font=("Segoe UI", 9), borderwidth=0, relief="flat")
    s.configure("Treeview.Heading",
        background=T["card2"], foreground=T["t3"],
        font=("Segoe UI", 8, "bold"), relief="flat",
        borderwidth=0, padding=(8,6))
    s.map("Treeview",
        background=[("selected", T["orange"]+"22")],
        foreground=[("selected", T["orange"])])
    s.map("Treeview.Heading",
        background=[("active", T["card2"])])
    s.configure("TCombobox",
        fieldbackground=T["input"], background=T["input"],
        foreground=T["t1"], arrowcolor=T["t3"],
        borderwidth=0, relief="flat",
        selectbackground=T["input"], selectforeground=T["t1"])
    s.map("TCombobox",
        fieldbackground=[("readonly", T["input"])],
        foreground=[("readonly", T["t1"])],
        selectbackground=[("readonly", T["input"])],
        selectforeground=[("readonly", T["t1"])])
    s.configure("Vertical.TScrollbar",
        background=T["card2"], troughcolor=T["bg"],
        arrowcolor=T["t3"], relief="flat",
        borderwidth=0, width=5)
    s.map("Vertical.TScrollbar",
        background=[("active", T["b3"])])
    s.configure("Horizontal.TScrollbar",
        background=T["card2"], troughcolor=T["bg"],
        arrowcolor=T["t3"], relief="flat",
        borderwidth=0, width=5)
    # Custom notebook
    s.configure("Arc.TNotebook",
        background=T["surface"], borderwidth=0)
    s.configure("Arc.TNotebook.Tab",
        background=T["surface"], foreground=T["t3"],
        font=("Segoe UI", 9), padding=[14, 7])
    s.map("Arc.TNotebook.Tab",
        background=[("selected", T["card"])],
        foreground=[("selected", T["orange"])])


# ══════════════════════════════════════════════════════
#  МІНІ-ГРАФІКИ
# ══════════════════════════════════════════════════════

class SparkLine(tk.Canvas):
    def __init__(self, parent, data, color=T["orange"], height=50, **kw):
        super().__init__(parent, height=height, bg=T["card"],
                         highlightthickness=0, **kw)
        self.data, self.color = data, color
        self.bind("<Configure>", self._draw)

    def _draw(self, _=None):
        self.delete("all")
        w, h = self.winfo_width(), self.winfo_height()
        if w < 2 or not self.data: return
        mn, mx = min(self.data), max(self.data)
        rng = mx - mn or 1
        n   = len(self.data)
        pad = 6
        xs  = [pad + (w-2*pad)*i/max(n-1,1) for i in range(n)]
        ys  = [h-pad-(h-2*pad)*(v-mn)/rng    for v in self.data]
        # gradient fill polygon
        poly = [xs[0], h] + [c for xy in zip(xs, ys) for c in xy] + [xs[-1], h]
        self.create_polygon(poly, fill=self.color+"18", outline="", smooth=True)
        # line
        for i in range(n-1):
            self.create_line(xs[i], ys[i], xs[i+1], ys[i+1],
                             fill=self.color, width=2, smooth=True)
        # last dot glow
        if n:
            x0, y0 = xs[-1], ys[-1]
            for radius, alpha in [(8,"18"),(5,"44"),(3,"FF")]:
                self.create_oval(x0-radius, y0-radius, x0+radius, y0+radius,
                                 fill=self.color+alpha, outline="")


class DonutCanvas(tk.Canvas):
    def __init__(self, parent, segments, size=110, **kw):
        super().__init__(parent, width=size, height=size,
                         bg=T["card"], highlightthickness=0, **kw)
        self.segs, self.size = segments, size
        self.bind("<Configure>", self._draw)

    def _draw(self, _=None):
        self.delete("all")
        w = h = self.winfo_width() or self.size
        cx, cy  = w//2, h//2
        r_out   = min(cx, cy) - 4
        r_in    = int(r_out * 0.6)
        total   = sum(s[1] for s in self.segs) or 1
        angle   = -90.0
        for _, val, color in self.segs:
            sweep = 360.0 * val / total
            if sweep > 0:
                self.create_arc(cx-r_out, cy-r_out, cx+r_out, cy+r_out,
                                start=angle, extent=sweep,
                                fill=color, outline=T["card"], width=2,
                                style="pieslice")
                angle += sweep
        # inner hole
        self.create_oval(cx-r_in, cy-r_in, cx+r_in, cy+r_in,
                         fill=T["card"], outline="")
        total_val = sum(s[1] for s in self.segs)
        self.create_text(cx, cy-7, text=str(total_val),
                         font=("Segoe UI", 14, "bold"), fill=T["t1"])
        self.create_text(cx, cy+8, text="всього",
                         font=("Segoe UI", 7), fill=T["t3"])


class HeatmapCanvas(tk.Canvas):
    def __init__(self, parent, risks, **kw):
        super().__init__(parent, bg=T["card"], highlightthickness=0, **kw)
        self.risks = risks
        self.bind("<Configure>", self._draw)

    def _draw(self, _=None):
        self.delete("all")
        w, h = self.winfo_width(), self.winfo_height()
        if w < 10: return
        PL, PT, PB = 24, 8, 22
        cw = (w - PL - 6) / 5
        ch = (h - PT - PB) / 5
        for ri, like in enumerate(range(5, 0, -1)):
            for ci, imp in enumerate(range(1, 6)):
                score = like * imp
                count = sum(1 for r in self.risks
                            if r.get("likelihood")==like and r.get("impact")==imp)
                col   = risk_col(score)
                x0 = PL + ci*cw + 2
                y0 = PT + ri*ch + 2
                x1, y1 = x0+cw-4, y0+ch-4
                alpha = "33" if not count else "66"
                rounded_rect(self, x0, y0, x1, y1, 4,
                             fill=col+alpha, outline=col+"44", width=1)
                if count:
                    cx2, cy2 = (x0+x1)/2, (y0+y1)/2
                    self.create_oval(cx2-9,cy2-9,cx2+9,cy2+9, fill=col, outline="")
                    self.create_text(cx2, cy2, text=str(count),
                                     font=("Segoe UI", 8, "bold"), fill="#000")
        # axis labels
        for i in range(5):
            self.create_text(PL+(i+.5)*cw, h-10, text=str(i+1),
                             font=("Segoe UI", 7), fill=T["t3"])
            self.create_text(12, PT+(4-i+.5)*ch, text=str(i+1),
                             font=("Segoe UI", 7), fill=T["t3"])
        self.create_text(PL+2.5*cw, h-2,  text="Вплив →",
                         font=("Segoe UI", 6), fill=T["t4"])


# ══════════════════════════════════════════════════════
#  ЛОГІН
# ══════════════════════════════════════════════════════

class LoginScreen(tk.Frame):
    def __init__(self, master):
        super().__init__(master, bg=T["bg"])
        self.pack(fill="both", expand=True)
        self._build()

    def _build(self):
        # BG pattern
        cv = tk.Canvas(self, bg=T["bg"], highlightthickness=0)
        cv.place(relx=0, rely=0, relwidth=1, relheight=1)
        self.after(20, lambda: self._bg_pattern(cv))

        # Center card
        wrap = tk.Frame(self, bg=T["card"], padx=0, pady=0)
        wrap.place(relx=.5, rely=.5, anchor="center", width=420)

        # Orange top stripe
        tk.Frame(wrap, bg=T["orange"], height=4).pack(fill="x")

        inner = tk.Frame(wrap, bg=T["card"], padx=48, pady=44)
        inner.pack(fill="x")

        # Logo
        logo_row = tk.Frame(inner, bg=T["card"])
        logo_row.pack(anchor="w", pady=(0, 28))
        icon_cv = tk.Canvas(logo_row, width=42, height=42,
                             bg=T["card"], highlightthickness=0)
        icon_cv.pack(side="left", padx=(0, 14))
        rounded_rect(icon_cv, 0, 0, 41, 41, 10,
                     fill=T["orange"]+"20", outline=T["orange"]+"60", width=1)
        icon_cv.create_text(21, 21, text="⚔", font=("Segoe UI", 18), fill=T["orange"])
        title_f = tk.Frame(logo_row, bg=T["card"])
        title_f.pack(side="left")
        tk.Label(title_f, text="Archer GRC",
                 font=("Segoe UI", 18, "bold"),
                 fg=T["t1"], bg=T["card"]).pack(anchor="w")
        tk.Label(title_f, text="Governance · Risk · Compliance",
                 font=("Segoe UI", 9), fg=T["t3"], bg=T["card"]).pack(anchor="w")

        # Separator line
        tk.Frame(inner, bg=T["b1"], height=1).pack(fill="x", pady=(0, 28))

        # Fields
        tk.Label(inner, text="Логін",
                 font=("Segoe UI", 8, "bold"), fg=T["t3"],
                 bg=T["card"], anchor="w").pack(fill="x")
        self.e_user = ModernEntry(inner)
        self.e_user.pack(fill="x", pady=(4, 16))

        tk.Label(inner, text="Пароль",
                 font=("Segoe UI", 8, "bold"), fg=T["t3"],
                 bg=T["card"], anchor="w").pack(fill="x")
        self.e_pass = ModernEntry(inner, show="●")
        self.e_pass.pack(fill="x", pady=(4, 28))

        login_cv = GlowButton(inner, "  Увійти →", self._login,
                              "primary", width=324, height=38)
        login_cv.pack(fill="x")

        tk.Label(inner, text="admin / admin123  •  auditor / pass123",
                 font=("Segoe UI", 8), fg=T["t4"],
                 bg=T["card"]).pack(pady=(20, 0))

        self.e_user.insert(0, "admin")
        self.e_user.entry.bind("<Return>", lambda _: self._login())
        self.e_pass.entry.bind("<Return>", lambda _: self._login())
        self.e_user.entry.focus()

    def _bg_pattern(self, cv):
        cv.delete("all")
        w = self.winfo_width()
        h = self.winfo_height()
        for i in range(-10, 30):
            o = i * 70
            cv.create_line(o, 0, o+h, h, fill=T["orange"]+"07", width=30)
        cv.create_rectangle(0, 0, w//3, h, fill=T["b1"]+"30", outline="")

    def _login(self):
        u = self.e_user.get().strip()
        p = self.e_pass.get().strip()
        db = get_db()
        row = db.execute(
            "SELECT * FROM users WHERE username=? AND password=? AND active=1", (u, p)
        ).fetchone()
        db.close()
        if row:
            self.master.login_success(dict(row))
        else:
            self.e_pass._border.configure(bg=T["red"])
            messagebox.showerror("Помилка входу", "Невірний логін або пароль")


# ══════════════════════════════════════════════════════
#  ЗАСТОСУНОК
# ══════════════════════════════════════════════════════

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Archer GRC Platform")
        self.geometry("1400x840")
        self.minsize(1100, 680)
        self.configure(bg=T["bg"])
        self.current_user = None
        apply_ttk_style(self)
        init_db()
        self._show_login()

    def _show_login(self):
        for w in self.winfo_children(): w.destroy()
        LoginScreen(self)

    def login_success(self, user):
        self.current_user = user
        for w in self.winfo_children(): w.destroy()
        MainLayout(self)

    def logout(self):
        self.current_user = None
        self._show_login()

    def log_action(self, action, module, rid=None, details=""):
        if not self.current_user: return
        db = get_db()
        db.execute("INSERT INTO audit_log(user_id,action,module,record_id,details) VALUES(?,?,?,?,?)",
                   (self.current_user["id"], action, module, rid, details))
        db.commit(); db.close()


# ══════════════════════════════════════════════════════
#  MAIN LAYOUT
# ══════════════════════════════════════════════════════

NAV = [
    (None, "НАВІГАЦІЯ"),
    ("🏠", "Дашборд",        "dashboard"),
    ("⚠",  "Ризики",         "risks"),
    ("🛡",  "Контролі",       "controls"),
    ("📋", "Політики",        "policies"),
    ("🏢", "Постачальники",   "suppliers"),
    ("🔍", "Аудит",           "audit"),
    ("🚨", "Інциденти",       "incidents"),
    ("⚖",  "Регулятори",      "regulations"),
    (None, "СИСТЕМА"),
    ("👥", "Користувачі",     "users"),
    ("📜", "Журнал дій",      "auditlog"),
]

class MainLayout(tk.Frame):
    def __init__(self, master):
        super().__init__(master, bg=T["bg"])
        self.pack(fill="both", expand=True)
        self._btns = {}
        self._active = None
        self._build()
        self._nav("dashboard")

    def _build(self):
        # ── SIDEBAR ──
        self.sb = tk.Frame(self, bg=T["surface"], width=220)
        self.sb.pack(side="left", fill="y")
        self.sb.pack_propagate(False)

        # top orange accent
        tk.Frame(self.sb, bg=T["orange"], height=3).pack(fill="x")

        # brand
        brand = tk.Frame(self.sb, bg=T["surface"], padx=20, pady=20)
        brand.pack(fill="x")
        br_cv = tk.Canvas(brand, width=32, height=32, bg=T["surface"],
                          highlightthickness=0)
        br_cv.pack(side="left", padx=(0,10))
        rounded_rect(br_cv, 0, 0, 31, 31, 7,
                     fill=T["orange"]+"22", outline=T["orange"]+"55", width=1)
        br_cv.create_text(16, 16, text="⚔", font=("Segoe UI", 14), fill=T["orange"])
        tf = tk.Frame(brand, bg=T["surface"])
        tf.pack(side="left")
        tk.Label(tf, text="Archer GRC",
                 font=("Segoe UI", 11, "bold"),
                 fg=T["t1"], bg=T["surface"]).pack(anchor="w")
        tk.Label(tf, text="v3.0",
                 font=("Segoe UI", 7), fg=T["t4"],
                 bg=T["surface"]).pack(anchor="w")

        # thin separator
        tk.Frame(self.sb, bg=T["b1"], height=1).pack(fill="x", padx=16)

        # nav list
        nav_f = tk.Frame(self.sb, bg=T["surface"])
        nav_f.pack(fill="both", expand=True, pady=10)

        for item in NAV:
            if item[0] is None:
                # section header
                tk.Label(nav_f, text=item[1],
                         font=("Segoe UI", 7, "bold"),
                         fg=T["t4"], bg=T["surface"],
                         anchor="w", padx=20).pack(fill="x", pady=(14, 4))
            else:
                icon, label, key = item
                self._make_nav_item(nav_f, icon, label, key)

        # bottom user panel
        tk.Frame(self.sb, bg=T["b1"], height=1).pack(fill="x", padx=16, side="bottom")
        user_panel = tk.Frame(self.sb, bg=T["surface"], padx=16, pady=14)
        user_panel.pack(fill="x", side="bottom")

        u = self.master.current_user
        av = tk.Canvas(user_panel, width=32, height=32,
                       bg=T["surface"], highlightthickness=0)
        av.pack(side="left", padx=(0, 10))
        av.create_oval(0, 0, 31, 31, fill=T["orange"]+"33", outline=T["orange"]+"66")
        av.create_text(16, 16, text=u["full_name"][0].upper(),
                       font=("Segoe UI", 11, "bold"), fill=T["orange"])

        info = tk.Frame(user_panel, bg=T["surface"])
        info.pack(side="left", fill="x", expand=True)
        name = u["full_name"]
        if len(name) > 18: name = name[:16] + "…"
        tk.Label(info, text=name,
                 font=("Segoe UI", 9, "bold"),
                 fg=T["t1"], bg=T["surface"]).pack(anchor="w")
        tk.Label(info, text=u["role"][:20],
                 font=("Segoe UI", 7), fg=T["t3"],
                 bg=T["surface"]).pack(anchor="w")

        logout_cv = tk.Canvas(user_panel, width=22, height=22,
                              bg=T["surface"], highlightthickness=0,
                              cursor="hand2")
        logout_cv.pack(side="right")
        logout_cv.create_text(11, 11, text="↩", font=("Segoe UI", 12),
                              fill=T["t3"])
        logout_cv.bind("<Button-1>", lambda _: self.master.logout())
        logout_cv.bind("<Enter>", lambda _: logout_cv.itemconfig(1, fill=T["orange"]))
        logout_cv.bind("<Leave>", lambda _: logout_cv.itemconfig(1, fill=T["t3"]))

        # ── MAIN ──
        main = tk.Frame(self, bg=T["bg"])
        main.pack(side="left", fill="both", expand=True)

        # header
        self.hdr = tk.Frame(main, bg=T["surface"], height=54)
        self.hdr.pack(fill="x")
        self.hdr.pack_propagate(False)
        tk.Frame(self.hdr, bg=T["b1"], height=1).pack(fill="x", side="bottom")

        self.h_title = tk.Label(self.hdr, text="",
                                font=("Segoe UI", 13, "bold"),
                                fg=T["t1"], bg=T["surface"])
        self.h_title.pack(side="left", padx=24, pady=14)

        # right header info
        hr = tk.Frame(self.hdr, bg=T["surface"])
        hr.pack(side="right", padx=20)
        dt = datetime.datetime.now()
        tk.Label(hr, text=dt.strftime("%d %b %Y"),
                 font=("Segoe UI", 9), fg=T["t3"],
                 bg=T["surface"]).pack(side="right", padx=(12, 0))

        # online dot + username
        usr_f = tk.Frame(hr, bg=T["surface"])
        usr_f.pack(side="right")
        tk.Canvas(usr_f, width=8, height=8, bg=T["surface"],
                  highlightthickness=0).pack(side="left", padx=(0,5), pady=3)
        dot_cv = usr_f.winfo_children()[0]
        dot_cv.create_oval(1,1,7,7, fill=T["green"], outline="")
        tk.Label(usr_f, text=u["username"],
                 font=("Segoe UI", 9, "bold"),
                 fg=T["t2"], bg=T["surface"]).pack(side="left")

        # content area
        self.content = tk.Frame(main, bg=T["bg"])
        self.content.pack(fill="both", expand=True)

    def _make_nav_item(self, parent, icon, label, key):
        # outer frame acts as the button
        item_f = tk.Frame(parent, bg=T["surface"], cursor="hand2", pady=0)
        item_f.pack(fill="x", padx=10, pady=1)

        # left indicator bar
        ind = tk.Frame(item_f, bg=T["surface"], width=3)
        ind.pack(side="left", fill="y")

        inner = tk.Frame(item_f, bg=T["surface"], padx=12, pady=9, cursor="hand2")
        inner.pack(side="left", fill="x", expand=True)

        icon_lbl = tk.Label(inner, text=icon,
                            font=("Segoe UI", 10),
                            fg=T["t3"], bg=T["surface"],
                            width=2, anchor="center")
        icon_lbl.pack(side="left", padx=(0, 10))

        text_lbl = tk.Label(inner, text=label,
                            font=("Segoe UI", 9),
                            fg=T["t3"], bg=T["surface"],
                            anchor="w")
        text_lbl.pack(side="left", fill="x", expand=True)

        self._btns[key] = (item_f, inner, ind, icon_lbl, text_lbl)

        def click(_=None, k=key):   self._nav(k)
        def enter(_=None, k=key):
            if self._active != k:
                for w in [item_f, inner]:
                    w.configure(bg=T["hover"])
                icon_lbl.configure(bg=T["hover"])
                text_lbl.configure(bg=T["hover"])
        def leave(_=None, k=key):
            if self._active != k:
                for w in [item_f, inner]:
                    w.configure(bg=T["surface"])
                icon_lbl.configure(bg=T["surface"])
                text_lbl.configure(bg=T["surface"])

        for w in [item_f, inner, icon_lbl, text_lbl]:
            w.bind("<Button-1>", click)
            w.bind("<Enter>",    enter)
            w.bind("<Leave>",    leave)

    def _nav(self, key):
        self._active = key
        label_map = {k: l for _, l, k in NAV if _ is not None}

        for k, (f, inner, ind, il, tl) in self._btns.items():
            if k == key:
                for w in [f, inner]: w.configure(bg=T["card"])
                il.configure(fg=T["orange"], bg=T["card"])
                tl.configure(fg=T["orange"], bg=T["card"],
                             font=("Segoe UI", 9, "bold"))
                ind.configure(bg=T["orange"])
            else:
                for w in [f, inner]: w.configure(bg=T["surface"])
                il.configure(fg=T["t3"], bg=T["surface"])
                tl.configure(fg=T["t3"], bg=T["surface"],
                             font=("Segoe UI", 9))
                ind.configure(bg=T["surface"])

        self.h_title.configure(text=label_map.get(key, ""))
        for w in self.content.winfo_children(): w.destroy()

        MODS = {
            "dashboard":   DashboardModule,
            "risks":       RisksModule,
            "controls":    ControlsModule,
            "policies":    PoliciesModule,
            "suppliers":   SuppliersModule,
            "audit":       AuditModule,
            "incidents":   IncidentsModule,
            "regulations": RegulationsModule,
            "users":       UsersModule,
            "auditlog":    AuditLogModule,
        }
        MODS.get(key, DashboardModule)(self.content, self.master)


# ══════════════════════════════════════════════════════
#  ДАШБОРД
# ══════════════════════════════════════════════════════

class DashboardModule(tk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent, bg=T["bg"])
        self.app = app
        self.pack(fill="both", expand=True)
        self._build()

    def _stats(self):
        db = get_db()
        s = {
            "open_risks":   db.execute("SELECT COUNT(*) FROM risks WHERE status='Відкритий'").fetchone()[0],
            "total_risks":  db.execute("SELECT COUNT(*) FROM risks").fetchone()[0],
            "controls":     db.execute("SELECT COUNT(*) FROM controls WHERE status='Активний'").fetchone()[0],
            "policies":     db.execute("SELECT COUNT(*) FROM policies WHERE status='Опублікована'").fetchone()[0],
            "incidents":    db.execute("SELECT COUNT(*) FROM incidents WHERE status NOT IN ('Вирішений','Скасовано')").fetchone()[0],
            "suppliers":    db.execute("SELECT COUNT(*) FROM suppliers WHERE status='Активний'").fetchone()[0],
            "compliant":    db.execute("SELECT COUNT(*) FROM regulations WHERE compliance_status='Відповідає'").fetchone()[0],
            "total_regs":   db.execute("SELECT COUNT(*) FROM regulations").fetchone()[0],
            "risks_raw":    [dict(r) for r in db.execute("SELECT likelihood,impact FROM risks WHERE status!='Закритий'").fetchall()],
            "top_risks":    [dict(r) for r in db.execute("""SELECT r.title,r.risk_score,r.status,u.full_name owner
                              FROM risks r LEFT JOIN users u ON r.owner_id=u.id
                              ORDER BY r.risk_score DESC LIMIT 5""").fetchall()],
            "regs":         [dict(r) for r in db.execute("SELECT title,compliance_status FROM regulations").fetchall()],
            "by_cat":       [dict(r) for r in db.execute("""SELECT category,COUNT(*) cnt FROM risks
                              WHERE status!='Закритий' GROUP BY category""").fetchall()],
        }
        db.close()
        return s

    def _build(self):
        s = self._stats()

        # scrollable canvas
        cv = tk.Canvas(self, bg=T["bg"], highlightthickness=0)
        vsb = ttk.Scrollbar(self, orient="vertical", command=cv.yview)
        cv.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        cv.pack(fill="both", expand=True)
        wrap = tk.Frame(cv, bg=T["bg"])
        win  = cv.create_window((0,0), window=wrap, anchor="nw")
        cv.bind("<Configure>", lambda e: cv.itemconfig(win, width=e.width))
        wrap.bind("<Configure>", lambda e: cv.configure(scrollregion=cv.bbox("all")))
        cv.bind_all("<MouseWheel>", lambda e: cv.yview_scroll(-1*(e.delta//120), "units"))

        P = 20  # outer padding
        p = tk.Frame(wrap, bg=T["bg"], padx=P, pady=P)
        p.pack(fill="both", expand=True)

        # ── ROW 1: KPI CARDS ──────────────────────────────────────
        kpi_f = tk.Frame(p, bg=T["bg"])
        kpi_f.pack(fill="x", pady=(0, 16))

        kpi_data = [
            ("Відкриті ризики",   str(s["open_risks"]),  f"/{s['total_risks']} загалом",   T["red"],    "⚠"),
            ("Активні контролі",  str(s["controls"]),    "контролів діють",                T["green"],  "🛡"),
            ("Опубл. політики",   str(s["policies"]),    "актуальних документів",          T["orange"], "📋"),
            ("Відкриті інциденти",str(s["incidents"]),   "потребують уваги",               T["yellow"], "🚨"),
            ("Постачальники",     str(s["suppliers"]),   "активних партнерів",             T["blue"],   "🏢"),
            ("Відповідність",     f"{s['compliant']}/{s['total_regs']}","регуляторів OK",  T["green"],  "⚖"),
        ]
        for i, (title, val, sub, color, icon) in enumerate(kpi_data):
            # card frame
            card = tk.Frame(kpi_f, bg=T["card"],
                            highlightthickness=1,
                            highlightbackground=T["b2"])
            card.grid(row=0, column=i, padx=(0, 12) if i<5 else (0,0), sticky="nsew")
            kpi_f.columnconfigure(i, weight=1)

            inner = tk.Frame(card, bg=T["card"], padx=16, pady=14)
            inner.pack(fill="both", expand=True)

            # top row: icon + value
            top = tk.Frame(inner, bg=T["card"])
            top.pack(fill="x")

            icon_cv2 = tk.Canvas(top, width=36, height=36,
                                 bg=T["card"], highlightthickness=0)
            icon_cv2.pack(side="right")
            rounded_rect(icon_cv2, 0, 0, 35, 35, 8,
                         fill=color+"18", outline=color+"33", width=1)
            icon_cv2.create_text(18, 18, text=icon,
                                 font=("Segoe UI", 14), fill=color)

            tk.Label(top, text=val,
                     font=("Segoe UI", 24, "bold"),
                     fg=color, bg=T["card"]).pack(side="left", anchor="sw")

            tk.Label(inner, text=title,
                     font=("Segoe UI", 8, "bold"),
                     fg=T["t2"], bg=T["card"]).pack(anchor="w", pady=(4,0))
            tk.Label(inner, text=sub,
                     font=("Segoe UI", 7),
                     fg=T["t4"], bg=T["card"]).pack(anchor="w")

            # bottom accent line
            tk.Frame(card, bg=color, height=2).pack(fill="x", side="bottom")

        # ── ROW 2: CHART + DONUT ──────────────────────────────────
        r2 = tk.Frame(p, bg=T["bg"])
        r2.pack(fill="x", pady=(0, 16))
        r2.columnconfigure(0, weight=3)
        r2.columnconfigure(1, weight=2)

        # Trend card
        tc = tk.Frame(r2, bg=T["card"],
                      highlightthickness=1, highlightbackground=T["b2"])
        tc.grid(row=0, column=0, padx=(0,12), sticky="nsew")
        tc_in = tk.Frame(tc, bg=T["card"], padx=18, pady=14)
        tc_in.pack(fill="both", expand=True)

        hdr_f = tk.Frame(tc_in, bg=T["card"])
        hdr_f.pack(fill="x", pady=(0,12))
        tk.Label(hdr_f, text="Динаміка ризиків",
                 font=("Segoe UI", 10, "bold"),
                 fg=T["t1"], bg=T["card"]).pack(side="left")
        val_f = tk.Frame(hdr_f, bg=T["card"])
        val_f.pack(side="right")
        tk.Label(val_f, text=str(s["open_risks"]),
                 font=("Segoe UI", 20, "bold"),
                 fg=T["orange"], bg=T["card"]).pack(side="left")
        tk.Label(val_f, text="  відкритих",
                 font=("Segoe UI", 8), fg=T["t3"],
                 bg=T["card"]).pack(side="left", pady=(8,0))

        base = s["total_risks"]
        trend = [max(1, base - 3 + random.randint(-1,2)) for _ in range(20)]
        trend[-1] = s["open_risks"]
        SparkLine(tc_in, trend, color=T["orange"], height=110).pack(fill="x")

        # period tabs
        pf = tk.Frame(tc_in, bg=T["card"])
        pf.pack(fill="x", pady=(8,0))
        for i, period in enumerate(["1Д","7Д","1М","3М","6М","Все"]):
            is_active = i == 0
            lbl_cv = tk.Canvas(pf, width=32, height=22,
                               bg=T["card"], highlightthickness=0,
                               cursor="hand2")
            lbl_cv.pack(side="left", padx=2)
            if is_active:
                rounded_rect(lbl_cv, 0, 0, 31, 21, 4,
                             fill=T["orange"]+"22", outline="")
                lbl_cv.create_text(16, 11, text=period,
                                   font=("Segoe UI", 7, "bold"), fill=T["orange"])
            else:
                lbl_cv.create_text(16, 11, text=period,
                                   font=("Segoe UI", 7), fill=T["t4"])

        # Donut card
        dc = tk.Frame(r2, bg=T["card"],
                      highlightthickness=1, highlightbackground=T["b2"])
        dc.grid(row=0, column=1, sticky="nsew")
        dc_in = tk.Frame(dc, bg=T["card"], padx=18, pady=14)
        dc_in.pack(fill="both", expand=True)
        tk.Label(dc_in, text="Категорії ризиків",
                 font=("Segoe UI", 10, "bold"),
                 fg=T["t1"], bg=T["card"]).pack(anchor="w", pady=(0,12))

        seg_cols = [T["orange"], T["red"], T["yellow"], T["blue"], T["green"], T["purple"]]
        segs = [(r["category"] or "Інше", r["cnt"], seg_cols[i % len(seg_cols)])
                for i, r in enumerate(s["by_cat"])] or [("Немає", 1, T["b3"])]

        drow = tk.Frame(dc_in, bg=T["card"])
        drow.pack(fill="x")
        DonutCanvas(drow, segs, size=120).pack(side="left")
        legend = tk.Frame(drow, bg=T["card"], padx=12)
        legend.pack(side="left", fill="y", pady=4)
        for label, val, color in segs:
            lf = tk.Frame(legend, bg=T["card"])
            lf.pack(fill="x", pady=2)
            dot = tk.Canvas(lf, width=8, height=8,
                            bg=T["card"], highlightthickness=0)
            dot.pack(side="left", padx=(0,6), pady=2)
            dot.create_oval(0,0,7,7, fill=color, outline="")
            tk.Label(lf, text=label[:16],
                     font=("Segoe UI", 8), fg=T["t2"],
                     bg=T["card"]).pack(side="left")
            tk.Label(lf, text=str(val),
                     font=("Segoe UI", 8, "bold"), fg=color,
                     bg=T["card"]).pack(side="right")

        # ── ROW 3: HEATMAP + TOP RISKS + COMPLIANCE ───────────────
        r3 = tk.Frame(p, bg=T["bg"])
        r3.pack(fill="x")
        r3.columnconfigure(0, weight=2)
        r3.columnconfigure(1, weight=2)
        r3.columnconfigure(2, weight=1)

        def section_card(parent, title, color=T["orange"]):
            card = tk.Frame(parent, bg=T["card"],
                            highlightthickness=1, highlightbackground=T["b2"])
            inn  = tk.Frame(card, bg=T["card"], padx=16, pady=14)
            inn.pack(fill="both", expand=True)
            hrow = tk.Frame(inn, bg=T["card"])
            hrow.pack(fill="x", pady=(0,12))
            tk.Frame(hrow, bg=color, width=3).pack(side="left", fill="y", padx=(0,10))
            tk.Label(hrow, text=title, font=("Segoe UI", 10, "bold"),
                     fg=T["t1"], bg=T["card"]).pack(side="left")
            return card, inn

        # Heatmap
        hm_card, hm_in = section_card(r3, "Теплова карта", T["red"])
        hm_card.grid(row=0, column=0, padx=(0,12), sticky="nsew")
        HeatmapCanvas(hm_in, s["risks_raw"], height=190).pack(fill="both", expand=True)

        # Top risks
        tr_card, tr_in = section_card(r3, "Топ-5 ризиків", T["yellow"])
        tr_card.grid(row=0, column=1, padx=(0,12), sticky="nsew")
        for risk in s["top_risks"]:
            score = risk.get("risk_score") or 0
            col   = risk_col(score)
            rf = tk.Frame(tr_in, bg=T["card2"],
                          highlightthickness=1, highlightbackground=T["b1"])
            rf.pack(fill="x", pady=3)
            rfi = tk.Frame(rf, bg=T["card2"], padx=10, pady=8)
            rfi.pack(fill="x")
            # score pill
            sc_cv = tk.Canvas(rfi, width=32, height=22,
                              bg=T["card2"], highlightthickness=0)
            sc_cv.pack(side="left", padx=(0,10))
            rounded_rect(sc_cv, 0, 0, 31, 21, 4,
                         fill=col+"33", outline=col+"66")
            sc_cv.create_text(16, 11, text=str(score),
                              font=("Segoe UI", 8, "bold"), fill=col)
            tk.Label(rfi, text=risk["title"][:30],
                     font=("Segoe UI", 9), fg=T["t1"],
                     bg=T["card2"]).pack(side="left")
            Pill(rfi, risk["status"]).pack(side="right")

        # Compliance
        cp_card, cp_in = section_card(r3, "Відповідність", T["green"])
        cp_card.grid(row=0, column=2, sticky="nsew")
        for reg in s["regs"]:
            col2, _ = STATUS_MAP.get(reg["compliance_status"], (T["t3"], ""))
            rf2 = tk.Frame(cp_in, bg=T["card2"], pady=7, padx=10)
            rf2.pack(fill="x", pady=2)
            dot2 = tk.Canvas(rf2, width=8, height=8,
                             bg=T["card2"], highlightthickness=0)
            dot2.pack(side="left", padx=(0,7), pady=1)
            dot2.create_oval(0,0,7,7, fill=col2, outline="")
            tk.Label(rf2, text=reg["title"][:22],
                     font=("Segoe UI", 8), fg=T["t2"],
                     bg=T["card2"], anchor="w").pack(side="left")


# ══════════════════════════════════════════════════════
#  БАЗОВИЙ CRUD
# ══════════════════════════════════════════════════════

class BaseCRUD(tk.Frame):
    TABLE   = ""
    COLUMNS = []

    def __init__(self, parent, app):
        super().__init__(parent, bg=T["bg"])
        self.app = app
        self.pack(fill="both", expand=True)
        self._build_toolbar()
        self._build_table()
        self.refresh()

    def _build_toolbar(self):
        bar = tk.Frame(self, bg=T["surface"],
                       highlightthickness=1, highlightbackground=T["b1"])
        bar.pack(fill="x")
        inner = tk.Frame(bar, bg=T["surface"], padx=16, pady=10)
        inner.pack(fill="x")

        GlowButton(inner, "+ Додати",    self._add,    "primary", 100, 32).pack(side="left", padx=(0,8))
        GlowButton(inner, "✎ Редагувати", self._edit,  "ghost",   110, 32).pack(side="left", padx=(0,8))
        GlowButton(inner, "✕ Видалити",  self._delete, "danger",  100, 32).pack(side="left")

        # search
        sf = tk.Frame(inner, bg=T["input"],
                      highlightthickness=1, highlightbackground=T["b2"])
        sf.pack(side="right")
        tk.Label(sf, text="⌕", font=("Segoe UI",11),
                 fg=T["t4"], bg=T["input"]).pack(side="left", padx=(8,2))
        self._sv = tk.StringVar()
        se = tk.Entry(sf, textvariable=self._sv,
                      bg=T["input"], fg=T["t1"],
                      insertbackground=T["orange"],
                      relief="flat", bd=0,
                      font=("Segoe UI", 9), width=22)
        se.pack(side="left", ipady=6, padx=(0,10))
        se.bind("<FocusIn>",  lambda _: sf.configure(highlightbackground=T["orange"]))
        se.bind("<FocusOut>", lambda _: sf.configure(highlightbackground=T["b2"]))
        self._sv.trace_add("write", lambda *_: self.refresh())

    def _build_table(self):
        wrap = tk.Frame(self, bg=T["bg"], padx=16, pady=12)
        wrap.pack(fill="both", expand=True)

        # Table header background
        head_f = tk.Frame(wrap, bg=T["card2"], pady=0)
        head_f.pack(fill="x")

        cols = [c[0] for c in self.COLUMNS]
        self.tree = ttk.Treeview(wrap, columns=cols, show="headings",
                                 selectmode="browse")
        vsb = ttk.Scrollbar(wrap, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(wrap, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        vsb.pack(side="right", fill="y")
        hsb.pack(side="bottom", fill="x")
        self.tree.pack(fill="both", expand=True)

        for key, label, width in self.COLUMNS:
            self.tree.heading(key, text=label, anchor="w")
            self.tree.column(key, width=width, minwidth=50, anchor="w")

        self.tree.tag_configure("even", background=T["card"])
        self.tree.tag_configure("odd",  background=T["card2"])
        self.tree.bind("<Double-1>", lambda _: self._edit())

    def get_rows(self, q=""): return []

    def refresh(self):
        q = self._sv.get().lower() if hasattr(self, "_sv") else ""
        self.tree.delete(*self.tree.get_children())
        for i, row in enumerate(self.get_rows(q)):
            vals = [row.get(c[0], "") or "" for c in self.COLUMNS]
            tag  = "even" if i % 2 == 0 else "odd"
            self.tree.insert("", "end", values=vals, iid=str(row["id"]), tags=(tag,))

    def sel_id(self):
        s = self.tree.selection()
        return int(s[0]) if s else None

    def _add(self):   self.open_form(None)
    def _edit(self):
        rid = self.sel_id()
        if not rid: messagebox.showwarning("Увага","Виберіть запис"); return
        self.open_form(rid)
    def _delete(self):
        rid = self.sel_id()
        if not rid: messagebox.showwarning("Увага","Виберіть запис"); return
        if messagebox.askyesno("Видалення","Видалити вибраний запис?"):
            db = get_db()
            db.execute(f"DELETE FROM {self.TABLE} WHERE id=?", (rid,))
            db.commit(); db.close()
            self.app.log_action("DELETE", self.TABLE, rid)
            self.refresh()
    def open_form(self, rid): pass


# ══════════════════════════════════════════════════════
#  UNIVERSAL FORM — сучасна модальна форма
# ══════════════════════════════════════════════════════

class UniversalForm(tk.Toplevel):
    def __init__(self, pm, app, rid, table, title, fields):
        super().__init__()
        self.pm, self.app = pm, app
        self.rid, self.table, self.fields = rid, table, fields
        self.title(f"{'Редагувати' if rid else 'Новий запис'} — {title}")
        self.geometry("560x580")
        self.configure(bg=T["bg"])
        self.resizable(False, True)

        db = get_db()
        self._users  = db.execute("SELECT id,full_name FROM users WHERE active=1").fetchall()
        db.close()
        self._unames = [u["full_name"] for u in self._users]
        self._uids   = [u["id"]        for u in self._users]
        self._wids   = {}

        tk.Frame(self, bg=T["orange"], height=3).pack(fill="x")
        self._build()
        if rid: self._populate()
        self.grab_set()

    def _build(self):
        cv = tk.Canvas(self, bg=T["bg"], highlightthickness=0)
        vsb = ttk.Scrollbar(self, orient="vertical", command=cv.yview)
        cv.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        cv.pack(fill="both", expand=True)
        f = tk.Frame(cv, bg=T["bg"], padx=28, pady=24)
        cv.create_window((0,0), window=f, anchor="nw")
        f.bind("<Configure>", lambda e: cv.configure(scrollregion=cv.bbox("all")))
        cv.bind("<MouseWheel>", lambda e: cv.yview_scroll(-1*(e.delta//120),"units"))

        for col, label, wtype, *opts in self.fields:
            tk.Label(f, text=label.upper(),
                     font=("Segoe UI", 7, "bold"),
                     fg=T["t3"], bg=T["bg"],
                     anchor="w").pack(fill="x", pady=(12, 3))

            if wtype == "entry":
                w = ModernEntry(f)
                w.pack(fill="x")
            elif wtype == "combo":
                wrap = tk.Frame(f, bg=T["b2"], pady=1)
                wrap.pack(fill="x")
                inner2 = tk.Frame(wrap, bg=T["input"], padx=4)
                inner2.pack(fill="x")
                w = ModernCombo(inner2, opts[0], width=55)
                w.pack(fill="x", ipady=4)
                w.bind("<FocusIn>",  lambda _, wr=wrap: wr.configure(bg=T["orange"]))
                w.bind("<FocusOut>", lambda _, wr=wrap: wr.configure(bg=T["b2"]))
                if opts[0]: w.set(opts[0][0])
            elif wtype == "text":
                w = ModernText(f, height=3)
                w.pack(fill="x")
            elif wtype == "user":
                wrap2 = tk.Frame(f, bg=T["b2"], pady=1)
                wrap2.pack(fill="x")
                inner3 = tk.Frame(wrap2, bg=T["input"], padx=4)
                inner3.pack(fill="x")
                w = ModernCombo(inner3, self._unames, width=55)
                w.pack(fill="x", ipady=4)
                w.bind("<FocusIn>",  lambda _, wr=wrap2: wr.configure(bg=T["orange"]))
                w.bind("<FocusOut>", lambda _, wr=wrap2: wr.configure(bg=T["b2"]))
            self._wids[col] = (wtype, w)

        tk.Frame(f, bg=T["b1"], height=1).pack(fill="x", pady=(24,16))
        bf = tk.Frame(f, bg=T["bg"])
        bf.pack(fill="x")
        GlowButton(bf, "💾  Зберегти", self._save, "primary",  130, 36).pack(side="left", padx=(0,8))
        GlowButton(bf, "✕  Скасувати", self.destroy, "ghost",  120, 36).pack(side="left")

    def _populate(self):
        db = get_db()
        row = db.execute(f"SELECT * FROM {self.table} WHERE id=?", (self.rid,)).fetchone()
        db.close()
        if not row: return
        row = dict(row)
        for col, _, wtype, *opts in self.fields:
            wt, w = self._wids[col]
            val = str(row.get(col, "") or "")
            if wt == "entry":   w.insert(0, val)
            elif wt in ("combo","user"):
                choices = opts[0] if wt == "combo" else self._unames
                if wt == "user":
                    uid = row.get(col)
                    if uid and uid in self._uids:
                        w.set(self._unames[self._uids.index(uid)])
                elif val in choices: w.set(val)
            elif wt == "text":  w.insert("1.0", val)

    def _save(self):
        first_col = self.fields[0][0]
        wt0, w0 = self._wids[first_col]
        v0 = w0.get() if wt0 != "text" else w0.get("1.0","end-1c")
        if not str(v0).strip():
            messagebox.showerror("Помилка", f"Поле '{self.fields[0][1]}' є обов'язковим")
            return
        data = {}
        for col, _, wtype, *opts in self.fields:
            wt, w = self._wids[col]
            if wt == "entry":   data[col] = w.get().strip() or None
            elif wt == "combo": data[col] = w.get() or None
            elif wt == "text":  data[col] = w.get("1.0","end-1c").strip() or None
            elif wt == "user":
                name = w.get()
                data[col] = self._uids[self._unames.index(name)] if name in self._unames else None
        db = get_db()
        if self.rid:
            sql = ", ".join(f"{k}=?" for k in data)
            db.execute(f"UPDATE {self.table} SET {sql} WHERE id=?", [*data.values(), self.rid])
            self.app.log_action("UPDATE", self.table, self.rid)
        else:
            keys = ",".join(data.keys())
            ph   = ",".join("?"*len(data))
            db.execute(f"INSERT INTO {self.table}({keys}) VALUES({ph})", list(data.values()))
            self.app.log_action("CREATE", self.table)
        db.commit(); db.close()
        self.pm.refresh()
        self.destroy()


# ══════════════════════════════════════════════════════
#  CRUD МОДУЛІ
# ══════════════════════════════════════════════════════

class RisksModule(BaseCRUD):
    TABLE   = "risks"
    COLUMNS = [("id","#",45),("title","Назва ризику",240),("category","Категорія",120),
               ("likelihood","Ймов.",60),("impact","Вплив",60),("risk_score","Оцінка",70),
               ("status","Статус",120),("owner","Власник",160),("review_date","Огляд",100)]
    def get_rows(self, q=""):
        db = get_db()
        rows = db.execute("""SELECT r.id,r.title,r.category,r.likelihood,r.impact,
                   r.risk_score,r.status,r.review_date,u.full_name owner
                   FROM risks r LEFT JOIN users u ON r.owner_id=u.id
                   ORDER BY r.risk_score DESC""").fetchall()
        db.close()
        return [dict(r) for r in rows if not q or q in r["title"].lower() or q in (r["category"] or "").lower()]
    def open_form(self, rid):
        UniversalForm(self, self.app, rid, "risks", "Ризик", [
            ("title","Назва *","entry"),
            ("category","Категорія","combo",["Кіберризик","Операційний","Фінансовий","Регуляторний","Репутаційний","Стратегічний"]),
            ("likelihood","Ймовірність (1–5)","combo",["1","2","3","4","5"]),
            ("impact","Вплив (1–5)","combo",["1","2","3","4","5"]),
            ("status","Статус","combo",["Відкритий","Пом'якшений","Прийнятий","Закритий"]),
            ("owner_id","Власник ризику","user"),
            ("department","Підрозділ","entry"),
            ("review_date","Дата огляду (РРРР-ММ-ДД)","entry"),
            ("mitigation_plan","План пом'якшення","text"),
            ("description","Опис","text"),
        ])

class ControlsModule(BaseCRUD):
    TABLE   = "controls"
    COLUMNS = [("id","#",45),("title","Назва контролю",240),("type","Тип",110),
               ("frequency","Частота",110),("status","Статус",100),
               ("effectiveness","Ефективність",160),("owner","Відповідальний",160),("last_tested","Тест",100)]
    def get_rows(self, q=""):
        db = get_db()
        rows = db.execute("""SELECT c.id,c.title,c.type,c.frequency,c.status,
                   c.effectiveness,c.last_tested,u.full_name owner
                   FROM controls c LEFT JOIN users u ON c.owner_id=u.id ORDER BY c.id""").fetchall()
        db.close()
        return [dict(r) for r in rows if not q or q in r["title"].lower()]
    def open_form(self, rid):
        UniversalForm(self, self.app, rid, "controls", "Контроль", [
            ("title","Назва *","entry"),
            ("type","Тип","combo",["Превентивний","Детективний","Відновлювальний","Компенсуючий"]),
            ("frequency","Частота","combo",["Постійно","Щодня","Щотижня","Щомісяця","Щоквартально","Щороку"]),
            ("status","Статус","combo",["Активний","Неактивний","Архівований"]),
            ("effectiveness","Ефективність","combo",["Ефективний","Частково ефективний","Неефективний","Не оцінено"]),
            ("owner_id","Відповідальний","user"),
            ("last_tested","Дата тестування","entry"),
            ("description","Опис","text"),
        ])

class PoliciesModule(BaseCRUD):
    TABLE   = "policies"
    COLUMNS = [("id","#",45),("title","Назва",240),("category","Категорія",130),
               ("version","Версія",70),("status","Статус",140),("owner","Власник",160),("review_date","Огляд",100)]
    def get_rows(self, q=""):
        db = get_db()
        rows = db.execute("""SELECT p.id,p.title,p.category,p.version,p.status,
                   p.review_date,u.full_name owner
                   FROM policies p LEFT JOIN users u ON p.owner_id=u.id ORDER BY p.id""").fetchall()
        db.close()
        return [dict(r) for r in rows if not q or q in r["title"].lower()]
    def open_form(self, rid):
        UniversalForm(self, self.app, rid, "policies", "Політика", [
            ("title","Назва *","entry"),
            ("category","Категорія","combo",["Інформаційна безпека","Комплаєнс","ІТ","HR","Фінанси"]),
            ("version","Версія","entry"),
            ("status","Статус","combo",["Чернетка","На затвердженні","Опублікована","Архівована"]),
            ("owner_id","Власник","user"),
            ("review_date","Дата огляду","entry"),
            ("description","Опис","text"),
            ("content","Зміст документу","text"),
        ])

class SuppliersModule(BaseCRUD):
    TABLE   = "suppliers"
    COLUMNS = [("id","#",45),("name","Назва постачальника",220),("category","Категорія",130),
               ("risk_level","Ризик",90),("status","Статус",100),("contact_name","Контакт",150),("contract_end","Договір до",100)]
    def get_rows(self, q=""):
        db = get_db()
        rows = db.execute("SELECT * FROM suppliers ORDER BY id").fetchall()
        db.close()
        return [dict(r) for r in rows if not q or q in r["name"].lower()]
    def open_form(self, rid):
        UniversalForm(self, self.app, rid, "suppliers", "Постачальник", [
            ("name","Назва *","entry"),
            ("category","Категорія","combo",["ІТ-послуги","Хмарна інфраструктура","Консалтинг","Охорона","Логістика"]),
            ("risk_level","Рівень ризику","combo",["Низький","Середній","Високий","Критичний"]),
            ("status","Статус","combo",["Активний","Неактивний","Призупинений"]),
            ("contact_name","Контактна особа","entry"),
            ("contact_email","Email","entry"),
            ("contract_start","Початок договору","entry"),
            ("contract_end","Кінець договору","entry"),
            ("notes","Примітки","text"),
        ])

class AuditModule(BaseCRUD):
    TABLE   = "audits"
    COLUMNS = [("id","#",45),("title","Назва аудиту",240),("type","Тип",110),
               ("department","Підрозділ",120),("start_date","Початок",100),
               ("end_date","Кінець",100),("status","Статус",120),("auditor","Аудитор",160)]
    def get_rows(self, q=""):
        db = get_db()
        rows = db.execute("""SELECT a.id,a.title,a.type,a.department,a.start_date,
                   a.end_date,a.status,u.full_name auditor
                   FROM audits a LEFT JOIN users u ON a.auditor_id=u.id ORDER BY a.id DESC""").fetchall()
        db.close()
        return [dict(r) for r in rows if not q or q in r["title"].lower()]
    def open_form(self, rid):
        UniversalForm(self, self.app, rid, "audits", "Аудит", [
            ("title","Назва *","entry"),
            ("type","Тип","combo",["Внутрішній","Зовнішній","Регуляторний","ІТ-аудит","Фінансовий"]),
            ("department","Підрозділ","entry"),
            ("auditor_id","Аудитор","user"),
            ("start_date","Дата початку","entry"),
            ("end_date","Дата кінця","entry"),
            ("status","Статус","combo",["Заплановано","В процесі","Завершено","Скасовано"]),
            ("scope","Область аудиту","text"),
            ("conclusion","Висновок","text"),
        ])

class IncidentsModule(BaseCRUD):
    TABLE   = "incidents"
    COLUMNS = [("id","#",45),("title","Назва",220),("category","Категорія",120),
               ("severity","Серйозність",100),("status","Статус",110),
               ("reporter","Репортер",150),("occurred_at","Дата",110)]
    def get_rows(self, q=""):
        db = get_db()
        rows = db.execute("""SELECT i.id,i.title,i.category,i.severity,i.status,
                   i.occurred_at,u.full_name reporter
                   FROM incidents i LEFT JOIN users u ON i.reporter_id=u.id ORDER BY i.id DESC""").fetchall()
        db.close()
        return [dict(r) for r in rows if not q or q in r["title"].lower()]
    def open_form(self, rid):
        UniversalForm(self, self.app, rid, "incidents", "Інцидент", [
            ("title","Назва *","entry"),
            ("category","Категорія","combo",["Кібербезпека","Операційний","ІТ-збій","Шахрайство","Витік даних"]),
            ("severity","Серйозність","combo",["Низька","Середня","Висока","Критична"]),
            ("status","Статус","combo",["Новий","В роботі","Вирішений","Скасовано"]),
            ("reporter_id","Репортер","user"),
            ("owner_id","Відповідальний","user"),
            ("occurred_at","Дата виникнення","entry"),
            ("root_cause","Першопричина","text"),
            ("corrective_action","Коригувальна дія","text"),
            ("description","Опис","text"),
        ])

class RegulationsModule(BaseCRUD):
    TABLE   = "regulations"
    COLUMNS = [("id","#",45),("title","Назва",240),("authority","Регулятор",100),
               ("category","Категорія",140),("compliance_status","Відповідність",160),
               ("effective_date","Чинний з",110),("responsible","Відповідальний",160)]
    def get_rows(self, q=""):
        db = get_db()
        rows = db.execute("""SELECT r.id,r.title,r.authority,r.category,
                   r.compliance_status,r.effective_date,u.full_name responsible
                   FROM regulations r LEFT JOIN users u ON r.responsible_id=u.id ORDER BY r.id""").fetchall()
        db.close()
        return [dict(r) for r in rows if not q or q in r["title"].lower()]
    def open_form(self, rid):
        UniversalForm(self, self.app, rid, "regulations", "Регуляторна вимога", [
            ("title","Назва *","entry"),
            ("authority","Регулятор","entry"),
            ("category","Категорія","combo",["Кіберзахист","Персональні дані","Інформаційна безпека","Платіжні системи"]),
            ("compliance_status","Статус відповідності","combo",["Відповідає","Частково відповідає","Не відповідає","В процесі","Не оцінено"]),
            ("effective_date","Набрання чинності","entry"),
            ("responsible_id","Відповідальний","user"),
            ("description","Опис","text"),
            ("notes","Примітки","text"),
        ])

class UsersModule(BaseCRUD):
    TABLE   = "users"
    COLUMNS = [("id","#",45),("username","Логін",130),("full_name","ПІБ",210),
               ("role","Роль",180),("department","Підрозділ",130),("email","Email",180),("active_str","Активний",80)]
    def get_rows(self, q=""):
        db = get_db()
        rows = db.execute("SELECT * FROM users ORDER BY id").fetchall()
        db.close()
        result = []
        for r in rows:
            d = dict(r); d["active_str"] = "✓" if d["active"] else "✗"
            if not q or q in d["full_name"].lower() or q in d["username"].lower():
                result.append(d)
        return result
    def open_form(self, rid): UserForm(self, self.app, rid)

class UserForm(tk.Toplevel):
    ROLES = ["Користувач (End User)","Власник ризику","Адміністратор додатка",
             "Системний адміністратор","Комплаєнс-офіцер","Аудитор"]
    def __init__(self, pm, app, rid=None):
        super().__init__()
        self.pm, self.app, self.rid = pm, app, rid
        self.title("Редагувати користувача" if rid else "Новий користувач")
        self.geometry("480x490")
        self.configure(bg=T["bg"])
        self.resizable(False, False)
        tk.Frame(self, bg=T["orange"], height=3).pack(fill="x")
        f = tk.Frame(self, bg=T["bg"], padx=28, pady=24)
        f.pack(fill="both", expand=True)

        def row(label, widget):
            tk.Label(f, text=label.upper(), font=("Segoe UI",7,"bold"),
                     fg=T["t3"], bg=T["bg"], anchor="w").pack(fill="x", pady=(12,3))
            widget.pack(fill="x")
            return widget

        self.f_user  = row("Логін *",   ModernEntry(f))
        self.f_pass  = row("Пароль *",  ModernEntry(f, show="●"))
        self.f_name  = row("ПІБ *",     ModernEntry(f))
        self.f_email = row("Email",     ModernEntry(f))
        self.f_dept  = row("Підрозділ", ModernEntry(f))

        tk.Label(f, text="РОЛЬ", font=("Segoe UI",7,"bold"),
                 fg=T["t3"], bg=T["bg"], anchor="w").pack(fill="x", pady=(12,3))
        wrap3 = tk.Frame(f, bg=T["b2"], pady=1)
        wrap3.pack(fill="x")
        inner4 = tk.Frame(wrap3, bg=T["input"], padx=4)
        inner4.pack(fill="x")
        self.f_role = ModernCombo(inner4, self.ROLES, width=55)
        self.f_role.pack(fill="x", ipady=4)
        self.f_role.set(self.ROLES[0])

        self.f_active = tk.BooleanVar(value=True)
        cb_f = tk.Frame(f, bg=T["bg"])
        cb_f.pack(fill="x", pady=(12,0))
        tk.Checkbutton(cb_f, variable=self.f_active,
                       bg=T["bg"], fg=T["t1"],
                       activebackground=T["bg"],
                       selectcolor=T["card"],
                       text="  Активний користувач",
                       font=("Segoe UI", 9)).pack(anchor="w")

        tk.Frame(f, bg=T["b1"], height=1).pack(fill="x", pady=(20,14))
        bf = tk.Frame(f, bg=T["bg"])
        bf.pack(fill="x")
        GlowButton(bf, "💾  Зберегти", self._save, "primary", 130, 36).pack(side="left", padx=(0,8))
        GlowButton(bf, "✕  Скасувати", self.destroy, "ghost",  120, 36).pack(side="left")

        if rid: self._populate()
        self.grab_set()

    def _populate(self):
        db = get_db()
        r = dict(db.execute("SELECT * FROM users WHERE id=?", (self.rid,)).fetchone())
        db.close()
        self.f_user.insert(0, r["username"]); self.f_name.insert(0, r["full_name"])
        self.f_email.insert(0, r.get("email") or ""); self.f_dept.insert(0, r.get("department") or "")
        if r["role"] in self.ROLES: self.f_role.set(r["role"])
        self.f_active.set(bool(r["active"]))

    def _save(self):
        u=self.f_user.get().strip(); p=self.f_pass.get().strip(); fn=self.f_name.get().strip()
        if not u or not fn: messagebox.showerror("Помилка","Логін та ПІБ обов'язкові"); return
        if not self.rid and not p: messagebox.showerror("Помилка","Пароль обов'язковий"); return
        db = get_db()
        try:
            if self.rid:
                upd = "username=?,full_name=?,email=?,department=?,role=?,active=?"
                vals=[u,fn,self.f_email.get().strip(),self.f_dept.get().strip(),self.f_role.get(),int(self.f_active.get())]
                if p: upd+=",password=?"; vals.append(p)
                vals.append(self.rid)
                db.execute(f"UPDATE users SET {upd} WHERE id=?", vals)
            else:
                db.execute("INSERT INTO users(username,password,full_name,email,department,role,active) VALUES(?,?,?,?,?,?,?)",
                           [u,p,fn,self.f_email.get().strip(),self.f_dept.get().strip(),self.f_role.get(),int(self.f_active.get())])
            db.commit()
        except sqlite3.IntegrityError:
            messagebox.showerror("Помилка","Такий логін вже існує"); db.close(); return
        db.close()
        self.app.log_action("USER_SAVE","users",self.rid,f"login={u}")
        self.pm.refresh(); self.destroy()


class AuditLogModule(BaseCRUD):
    TABLE   = "audit_log"
    COLUMNS = [("id","#",50),("created_at","Дата/Час",155),("user","Користувач",155),
               ("action","Дія",110),("module","Модуль",100),("record_id","ID",60),("details","Деталі",350)]
    def _build_toolbar(self):
        bar = tk.Frame(self, bg=T["surface"],
                       highlightthickness=1, highlightbackground=T["b1"])
        bar.pack(fill="x")
        inner = tk.Frame(bar, bg=T["surface"], padx=16, pady=10)
        inner.pack(fill="x")
        tk.Label(inner, text="Журнал аудиту — тільки читання",
                 font=("Segoe UI",10,"bold"), fg=T["t2"], bg=T["surface"]).pack(side="left")
        sf = tk.Frame(inner, bg=T["input"],
                      highlightthickness=1, highlightbackground=T["b2"])
        sf.pack(side="right")
        tk.Label(sf, text="⌕", font=("Segoe UI",11), fg=T["t4"],
                 bg=T["input"]).pack(side="left", padx=(8,2))
        self._sv = tk.StringVar()
        se = tk.Entry(sf, textvariable=self._sv, bg=T["input"], fg=T["t1"],
                      insertbackground=T["orange"], relief="flat", bd=0,
                      font=("Segoe UI",9), width=22)
        se.pack(side="left", ipady=6, padx=(0,10))
        self._sv.trace_add("write", lambda *_: self.refresh())
    def get_rows(self, q=""):
        db = get_db()
        rows = db.execute("""SELECT al.id,al.created_at,al.action,al.module,al.record_id,al.details,u.full_name user
                   FROM audit_log al LEFT JOIN users u ON al.user_id=u.id
                   ORDER BY al.id DESC LIMIT 400""").fetchall()
        db.close()
        return [dict(r) for r in rows if not q or q in (r["action"] or "").lower() or q in (r["module"] or "").lower()]
    def open_form(self, rid): pass
    def _add(self): pass
    def _edit(self): pass
    def _delete(self): pass


# ══════════════════════════════════════════════════════
#  ЗАПУСК
# ══════════════════════════════════════════════════════

if __name__ == "__main__":
    app = App()
    app.mainloop()
