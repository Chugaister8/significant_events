"""
Archer GRC Platform — Десктопна система управління ризиками та відповідністю
Версія: 2.0 | Квітень 2026 | Дизайн: Dark Fintech / Extej-style
"""

import tkinter as tk
from tkinter import ttk, messagebox
import sqlite3
import os
import math
import datetime
import random

# ─────────────────────────────────────────────────────────────────
# КОЛЬОРОВА ПАЛІТРА — Extej Dark Orange
# ─────────────────────────────────────────────────────────────────
C = {
    "bg":           "#0E0E0E",   # майже чорний фон
    "bg2":          "#141414",   # трохи світліший
    "panel":        "#181818",   # панелі/картки
    "panel2":       "#1E1E1E",   # вторинні панелі
    "border":       "#2A2A2A",   # межі
    "border2":      "#333333",
    "accent":       "#FF6B00",   # основний помаранчевий
    "accent2":      "#FF8C00",   # яскравіший помаранчевий
    "accent_dim":   "#FF6B0022", # напівпрозорий помаранчевий
    "accent_glow":  "#FF6B0044",
    "red":          "#FF4444",
    "green":        "#00D084",
    "yellow":       "#FFB800",
    "blue":         "#4A9EFF",
    "text":         "#F0F0F0",
    "text2":        "#A0A0A0",
    "text3":        "#606060",
    "sidebar_w":    210,
    "header_h":     56,
}

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "archer_grc.db")

# ─────────────────────────────────────────────────────────────────
# БАЗА ДАНИХ
# ─────────────────────────────────────────────────────────────────

def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn

def init_db():
    conn = get_db()
    c = conn.cursor()
    c.executescript("""
    CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT UNIQUE NOT NULL,
        password TEXT NOT NULL,
        full_name TEXT NOT NULL,
        role TEXT NOT NULL,
        email TEXT,
        department TEXT,
        active INTEGER DEFAULT 1,
        created_at TEXT DEFAULT (datetime('now'))
    );
    CREATE TABLE IF NOT EXISTS risks (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        title TEXT NOT NULL,
        description TEXT,
        category TEXT,
        likelihood INTEGER DEFAULT 3,
        impact INTEGER DEFAULT 3,
        risk_score INTEGER GENERATED ALWAYS AS (likelihood * impact) STORED,
        status TEXT DEFAULT 'Відкритий',
        owner_id INTEGER,
        department TEXT,
        mitigation_plan TEXT,
        review_date TEXT,
        created_at TEXT DEFAULT (datetime('now')),
        updated_at TEXT DEFAULT (datetime('now')),
        FOREIGN KEY(owner_id) REFERENCES users(id)
    );
    CREATE TABLE IF NOT EXISTS controls (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        title TEXT NOT NULL,
        description TEXT,
        type TEXT,
        frequency TEXT,
        owner_id INTEGER,
        status TEXT DEFAULT 'Активний',
        effectiveness TEXT DEFAULT 'Не оцінено',
        last_tested TEXT,
        created_at TEXT DEFAULT (datetime('now')),
        FOREIGN KEY(owner_id) REFERENCES users(id)
    );
    CREATE TABLE IF NOT EXISTS policies (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        title TEXT NOT NULL,
        description TEXT,
        category TEXT,
        version TEXT DEFAULT '1.0',
        status TEXT DEFAULT 'Чернетка',
        owner_id INTEGER,
        review_date TEXT,
        content TEXT,
        created_at TEXT DEFAULT (datetime('now')),
        FOREIGN KEY(owner_id) REFERENCES users(id)
    );
    CREATE TABLE IF NOT EXISTS suppliers (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL,
        category TEXT,
        contact_name TEXT,
        contact_email TEXT,
        risk_level TEXT DEFAULT 'Середній',
        status TEXT DEFAULT 'Активний',
        contract_start TEXT,
        contract_end TEXT,
        notes TEXT,
        created_at TEXT DEFAULT (datetime('now'))
    );
    CREATE TABLE IF NOT EXISTS audits (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        title TEXT NOT NULL,
        type TEXT,
        auditor_id INTEGER,
        department TEXT,
        start_date TEXT,
        end_date TEXT,
        status TEXT DEFAULT 'Заплановано',
        scope TEXT,
        conclusion TEXT,
        created_at TEXT DEFAULT (datetime('now')),
        FOREIGN KEY(auditor_id) REFERENCES users(id)
    );
    CREATE TABLE IF NOT EXISTS incidents (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        title TEXT NOT NULL,
        description TEXT,
        category TEXT,
        severity TEXT DEFAULT 'Середня',
        status TEXT DEFAULT 'Новий',
        reporter_id INTEGER,
        owner_id INTEGER,
        occurred_at TEXT,
        root_cause TEXT,
        corrective_action TEXT,
        created_at TEXT DEFAULT (datetime('now')),
        FOREIGN KEY(reporter_id) REFERENCES users(id),
        FOREIGN KEY(owner_id) REFERENCES users(id)
    );
    CREATE TABLE IF NOT EXISTS regulations (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        title TEXT NOT NULL,
        authority TEXT,
        category TEXT,
        description TEXT,
        effective_date TEXT,
        compliance_status TEXT DEFAULT 'Не оцінено',
        responsible_id INTEGER,
        notes TEXT,
        created_at TEXT DEFAULT (datetime('now')),
        FOREIGN KEY(responsible_id) REFERENCES users(id)
    );
    CREATE TABLE IF NOT EXISTS audit_log (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        user_id INTEGER,
        action TEXT,
        module TEXT,
        record_id INTEGER,
        details TEXT,
        created_at TEXT DEFAULT (datetime('now'))
    );
    """)

    c.execute("SELECT id FROM users WHERE username='admin'")
    if not c.fetchone():
        seed_users = [
            ('admin','admin123','Системний Адміністратор','Системний адміністратор','admin@company.ua','ІТ'),
            ('risk_officer','pass123','Іваненко Олексій','Власник ризику','risk@company.ua','Управління ризиками'),
            ('auditor','pass123','Петренко Марія','Аудитор','audit@company.ua','Аудит'),
            ('compliance','pass123','Коваль Дмитро','Комплаєнс-офіцер','compliance@company.ua','Комплаєнс'),
        ]
        for s in seed_users:
            c.execute("INSERT INTO users (username,password,full_name,role,email,department) VALUES (?,?,?,?,?,?)", s)

        risks_data = [
            ('Витік персональних даних','Несанкціонований доступ до ПДн клієнтів','Кіберризик',4,5,'Відкритий',2,'ІТ-безпека','Впровадження DLP','2026-06-30'),
            ('Збій ключових ІТ-систем','Відмова банківських систем у пікові години','Операційний',2,5,"Пом'якшений",2,'ІТ','Резервне копіювання, план DR','2026-09-30'),
            ('Недотримання вимог НБУ','Ризик штрафних санкцій за порушення нормативів','Регуляторний',2,4,'Відкритий',4,'Комплаєнс','Моніторинг нормативної бази','2026-07-31'),
            ('Шахрайство третіх осіб','Ризик шахрайських дій постачальників','Фінансовий',3,4,'Відкритий',2,'Безпека','Due diligence постачальників','2026-08-31'),
            ('Кібератака DDoS','Атака на відмову в обслуговуванні','Кіберризик',3,3,'Відкритий',2,'ІТ-безпека','Anti-DDoS, IPS/IDS','2026-05-31'),
            ('Фішинг атаки на персонал','Соціальна інженерія проти співробітників','Кіберризик',4,3,'Відкритий',2,'ІТ-безпека','Навчання та симуляції','2026-04-30'),
            ('Нелояльний персонал','Ризик зловживань з боку співробітників','Операційний',2,4,'Відкритий',2,'HR','Моніторинг активності','2026-12-31'),
        ]
        for r in risks_data:
            c.execute("INSERT INTO risks (title,description,category,likelihood,impact,status,owner_id,department,mitigation_plan,review_date) VALUES (?,?,?,?,?,?,?,?,?,?)", r)

        controls_data = [
            ('Антивірусний захист','Встановлений на всіх робочих станціях','Превентивний','Постійно',2,'Активний','Ефективний','2026-03-15'),
            ('Резервне копіювання','Щоденне backup критичних систем','Відновлювальний','Щодня',2,'Активний','Ефективний','2026-03-20'),
            ('Розмежування прав доступу','RBAC для всіх корпоративних систем','Превентивний','Постійно',4,'Активний','Частково ефективний','2026-02-28'),
            ('Журналювання подій','SIEM-моніторинг аномалій','Детективний','Постійно',2,'Активний','Ефективний','2026-03-01'),
            ('Навчання персоналу','Щоквартальні тренінги з кібербезпеки','Превентивний','Щоквартально',4,'Активний','Не оцінено','2026-01-15'),
        ]
        for ct in controls_data:
            c.execute("INSERT INTO controls (title,description,type,frequency,owner_id,status,effectiveness,last_tested) VALUES (?,?,?,?,?,?,?,?)", ct)

        policies_data = [
            ('Політика інформаційної безпеки','Основний документ з ІБ','Інформаційна безпека','2.1','Опублікована',4,'2026-12-31'),
            ('Політика управління паролями','Вимоги до паролів','Інформаційна безпека','1.3','Опублікована',4,'2026-06-30'),
            ('Антикорупційна політика','Запобігання корупційним ризикам','Комплаєнс','1.0','На затвердженні',4,'2027-01-31'),
            ('Політика BYOD','Використання особистих пристроїв','ІТ','1.1','Чернетка',2,'2026-09-30'),
        ]
        for p in policies_data:
            c.execute("INSERT INTO policies (title,description,category,version,status,owner_id,review_date) VALUES (?,?,?,?,?,?,?)", p)

        suppliers_data = [
            ('ТОВ "ТехноСофт"','ІТ-послуги','Сидоренко В.В.','v.s@technosoft.ua','Середній','Активний','2025-01-01','2026-12-31'),
            ('АТ "CloudServ"','Хмарна інфраструктура','Мороз О.П.','o.m@cloudserv.ua','Високий','Активний','2024-07-01','2026-06-30'),
            ('ФОП Бондаренко І.С.','Консалтинг','Бондаренко І.С.','i.b@gmail.com','Низький','Активний','2026-01-01','2026-12-31'),
        ]
        for s in suppliers_data:
            c.execute("INSERT INTO suppliers (name,category,contact_name,contact_email,risk_level,status,contract_start,contract_end) VALUES (?,?,?,?,?,?,?,?)", s)

        regulations_data = [
            ('Положення НБУ №95','НБУ','Кіберзахист','Відповідає',4,'2022-01-01'),
            ('Закон про захист ПДн','Верховна Рада','Персональні дані','Частково відповідає',4,'2010-06-01'),
            ('ISO/IEC 27001:2022','ISO','Інформаційна безпека','В процесі',4,'2022-10-25'),
            ('PCI DSS v4.0','PCI SSC','Платіжні системи','Відповідає',4,'2022-03-31'),
            ('GDPR','ЄС','Персональні дані','Не оцінено',4,'2018-05-25'),
        ]
        for r in regulations_data:
            c.execute("INSERT INTO regulations (title,authority,category,compliance_status,responsible_id,effective_date) VALUES (?,?,?,?,?,?)", r)

    conn.commit()
    conn.close()

# ─────────────────────────────────────────────────────────────────
# УТИЛІТИ UI
# ─────────────────────────────────────────────────────────────────

STATUS_COLOR = {
    "Відкритий": C["red"], "Пом'якшений": C["yellow"], "Закритий": C["green"],
    "Прийнятий": C["text3"], "Активний": C["green"], "Неактивний": C["text3"],
    "Ефективний": C["green"], "Частково ефективний": C["yellow"], "Неефективний": C["red"],
    "Не оцінено": C["text3"], "Опублікована": C["green"], "Чернетка": C["text3"],
    "На затвердженні": C["yellow"], "Архівована": C["text3"],
    "Новий": C["blue"], "В роботі": C["yellow"], "Вирішений": C["green"], "Скасовано": C["text3"],
    "Заплановано": C["blue"], "В процесі": C["yellow"], "Завершено": C["green"],
    "Відповідає": C["green"], "Частково відповідає": C["yellow"],
    "Не відповідає": C["red"], "Не оцінено": C["text3"],
    "Низький": C["green"], "Середній": C["yellow"], "Високий": C["red"], "Критичний": "#8B0000",
    "Відкрита": C["red"], "Закрита": C["green"],
}

def risk_color(score):
    if score <= 4:  return C["green"]
    if score <= 9:  return C["yellow"]
    if score <= 16: return C["red"]
    return "#8B0000"

def lbl(parent, text, size=9, bold=False, fg=None, bg=None, anchor="w", **kw):
    return tk.Label(parent, text=text,
                    font=("Consolas" if bold else "Segoe UI", size, "bold" if bold else "normal"),
                    fg=fg or C["text"], bg=bg or C["panel"],
                    anchor=anchor, **kw)

def entry(parent, width=36, show=None, **kw):
    e = tk.Entry(parent, width=width, bg=C["panel2"], fg=C["text"],
                 insertbackground=C["accent"], relief="flat",
                 font=("Segoe UI", 9), bd=0,
                 highlightthickness=1, highlightbackground=C["border2"],
                 highlightcolor=C["accent"], **kw)
    if show: e.configure(show=show)
    return e

def combo(parent, values, width=34, **kw):
    cb = ttk.Combobox(parent, values=values, width=width, state="readonly",
                      font=("Segoe UI", 9), **kw)
    return cb

def btn(parent, text, cmd=None, style="accent", width=None, **kw):
    configs = {
        "accent": (C["accent"],  "#000000", C["accent2"]),
        "ghost":  (C["panel2"],  C["text2"], C["border2"]),
        "danger": (C["red"],     "#ffffff",  "#FF6666"),
        "outline":(C["panel"],   C["accent"], C["accent"]),
    }
    bg_, fg_, hov = configs.get(style, configs["accent"])
    b = tk.Button(parent, text=text, bg=bg_, fg=fg_,
                  relief="flat", bd=0, font=("Segoe UI", 9, "bold"),
                  padx=14, pady=7, cursor="hand2",
                  activebackground=hov, activeforeground=fg_,
                  **kw)
    if cmd: b.configure(command=cmd)
    if width: b.configure(width=width)
    return b

def sep(parent, color=None, orient="h", **kw):
    if orient == "h":
        return tk.Frame(parent, bg=color or C["border"], height=1, **kw)
    return tk.Frame(parent, bg=color or C["border"], width=1, **kw)

def status_dot(parent, status_text, bg=None):
    bg = bg or C["panel"]
    color = STATUS_COLOR.get(status_text, C["text3"])
    f = tk.Frame(parent, bg=bg)
    tk.Canvas(f, width=8, height=8, bg=bg, highlightthickness=0).pack(side="left", padx=(0,5))
    tk.Label(f, text=status_text, font=("Segoe UI", 8), fg=color, bg=bg).pack(side="left")
    # draw dot
    c2 = f.winfo_children()[0]
    c2.create_oval(1, 1, 7, 7, fill=color, outline="")
    return f

def orange_badge(parent, text, bg=None):
    bg = bg or C["panel"]
    color = STATUS_COLOR.get(text, C["text3"])
    return tk.Label(parent, text=f" {text} ", bg=color + "33",
                    fg=color, font=("Segoe UI", 7, "bold"),
                    relief="flat", padx=4, pady=2)

# ─────────────────────────────────────────────────────────────────
# TTK СТИЛІ
# ─────────────────────────────────────────────────────────────────

def apply_styles(root):
    style = ttk.Style(root)
    style.theme_use("clam")
    style.configure("Treeview",
        background=C["panel2"],
        foreground=C["text"],
        fieldbackground=C["panel2"],
        rowheight=32,
        font=("Segoe UI", 9),
        borderwidth=0,
        relief="flat")
    style.configure("Treeview.Heading",
        background=C["panel"],
        foreground=C["text2"],
        font=("Segoe UI", 8, "bold"),
        relief="flat", borderwidth=0)
    style.map("Treeview",
        background=[("selected", C["accent"] + "33")],
        foreground=[("selected", C["accent"])])
    style.configure("TCombobox",
        fieldbackground=C["panel2"],
        background=C["panel2"],
        foreground=C["text"],
        arrowcolor=C["accent"],
        borderwidth=0, relief="flat")
    style.map("TCombobox",
        fieldbackground=[("readonly", C["panel2"])],
        foreground=[("readonly", C["text"])],
        selectbackground=[("readonly", C["panel2"])],
        selectforeground=[("readonly", C["text"])])
    style.configure("Vertical.TScrollbar",
        background=C["panel"],
        troughcolor=C["bg"],
        arrowcolor=C["text3"],
        relief="flat", borderwidth=0, width=6)
    style.configure("TNotebook",
        background=C["panel"],
        borderwidth=0)
    style.configure("TNotebook.Tab",
        background=C["panel"],
        foreground=C["text2"],
        padding=[16,7],
        font=("Segoe UI", 9))
    style.map("TNotebook.Tab",
        background=[("selected", C["panel2"])],
        foreground=[("selected", C["accent"])])

# ─────────────────────────────────────────────────────────────────
# CANVAS CHART — мінімалістичний лінійний
# ─────────────────────────────────────────────────────────────────

class MiniChart(tk.Canvas):
    def __init__(self, parent, data, color=None, fill=True, **kw):
        kw.setdefault("bg", C["panel2"])
        kw.setdefault("highlightthickness", 0)
        super().__init__(parent, **kw)
        self.data = data
        self.color = color or C["accent"]
        self.fill = fill
        self.bind("<Configure>", self._draw)

    def _draw(self, e=None):
        self.delete("all")
        w, h = self.winfo_width(), self.winfo_height()
        if w < 4 or h < 4 or not self.data: return
        pad = 6
        mn, mx = min(self.data), max(self.data)
        rng = mx - mn or 1
        n = len(self.data)
        xs = [pad + (w - 2*pad) * i / max(n-1, 1) for i in range(n)]
        ys = [h - pad - (h - 2*pad) * (v - mn) / rng for v in self.data]

        # Fill area
        if self.fill:
            pts = [pad, h] + [coord for xy in zip(xs, ys) for coord in xy] + [w - pad, h]
            grad_id = self.color
            self.create_polygon(pts, fill=self.color + "22", outline="")

        # Grid lines subtle
        for i in range(4):
            y = pad + (h - 2*pad) * i / 3
            self.create_line(pad, y, w-pad, y, fill=C["border"], width=1)

        # Line
        if n > 1:
            for i in range(n-1):
                self.create_line(xs[i], ys[i], xs[i+1], ys[i+1],
                                 fill=self.color, width=2, smooth=True)

        # Last point glow
        if n > 0:
            x0, y0 = xs[-1], ys[-1]
            self.create_oval(x0-5, y0-5, x0+5, y0+5,
                             fill=self.color+"44", outline="")
            self.create_oval(x0-3, y0-3, x0+3, y0+3,
                             fill=self.color, outline="")

# ─────────────────────────────────────────────────────────────────
# HEATMAP CANVAS
# ─────────────────────────────────────────────────────────────────

class HeatmapCanvas(tk.Canvas):
    def __init__(self, parent, risks, **kw):
        kw.setdefault("bg", C["panel2"])
        kw.setdefault("highlightthickness", 0)
        super().__init__(parent, **kw)
        self.risks = risks
        self.bind("<Configure>", self._draw)

    def _draw(self, e=None):
        self.delete("all")
        w, h = self.winfo_width(), self.winfo_height()
        if w < 10: return
        pad_l, pad_b = 32, 24
        cell_w = (w - pad_l - 8) / 5
        cell_h = (h - pad_b - 8) / 5

        for row_i, like in enumerate(range(5, 0, -1)):
            for col_j, imp in enumerate(range(1, 6)):
                score = like * imp
                count = sum(1 for r in self.risks
                            if r.get("likelihood") == like and r.get("impact") == imp)
                x0 = pad_l + col_j * cell_w + 2
                y0 = 4 + row_i * cell_h + 2
                x1 = x0 + cell_w - 4
                y1 = y0 + cell_h - 4

                col = risk_color(score)
                alpha_col = col + ("44" if count == 0 else "AA")
                # draw rect
                self.create_rectangle(x0, y0, x1, y1,
                                      fill=col + ("22" if count == 0 else "55"),
                                      outline=col + "44", width=1)
                if count > 0:
                    cx, cy = (x0+x1)/2, (y0+y1)/2
                    self.create_oval(cx-10, cy-10, cx+10, cy+10,
                                     fill=col, outline="")
                    self.create_text(cx, cy, text=str(count),
                                     fill="#000000", font=("Segoe UI", 9, "bold"))

        # Axis labels
        for i, v in enumerate(range(1, 6)):
            x = pad_l + (i + 0.5) * cell_w
            self.create_text(x, h - 10, text=str(v),
                             fill=C["text3"], font=("Segoe UI", 8))
        for i, v in enumerate(range(5, 0, -1)):
            y = 4 + (i + 0.5) * cell_h
            self.create_text(16, y, text=str(v),
                             fill=C["text3"], font=("Segoe UI", 8))
        self.create_text(pad_l + 2.5 * cell_w, h - 2, text="Вплив →",
                         fill=C["text3"], font=("Segoe UI", 7))
        self.create_text(8, h / 2, text="↑", fill=C["text3"], font=("Segoe UI", 8))

# ─────────────────────────────────────────────────────────────────
# DONUT CHART
# ─────────────────────────────────────────────────────────────────

class DonutChart(tk.Canvas):
    def __init__(self, parent, segments, **kw):
        """segments: list of (label, value, color)"""
        kw.setdefault("bg", C["panel2"])
        kw.setdefault("highlightthickness", 0)
        super().__init__(parent, **kw)
        self.segments = segments
        self.bind("<Configure>", self._draw)

    def _draw(self, e=None):
        self.delete("all")
        w, h = self.winfo_width(), self.winfo_height()
        if w < 10: return
        cx, cy = w // 2, h // 2
        r_out = min(cx, cy) - 6
        r_in  = r_out * 0.55
        total = sum(s[1] for s in self.segments) or 1
        angle = -90.0
        for label, val, color in self.segments:
            sweep = 360.0 * val / total
            self.create_arc(cx-r_out, cy-r_out, cx+r_out, cy+r_out,
                            start=angle, extent=sweep,
                            fill=color, outline=C["panel2"], width=2)
            angle += sweep
        # Inner circle
        self.create_oval(cx-r_in, cy-r_in, cx+r_in, cy+r_in,
                         fill=C["panel2"], outline="")
        # Center text
        total_val = sum(s[1] for s in self.segments)
        self.create_text(cx, cy - 6, text=str(total_val),
                         fill=C["text"], font=("Consolas", 14, "bold"))
        self.create_text(cx, cy + 10, text="всього",
                         fill=C["text3"], font=("Segoe UI", 7))

# ─────────────────────────────────────────────────────────────────
# ЛОГІН
# ─────────────────────────────────────────────────────────────────

class LoginScreen(tk.Frame):
    def __init__(self, master):
        super().__init__(master, bg=C["bg"])
        self.pack(fill="both", expand=True)
        self._build()

    def _build(self):
        # Diagonal accent lines on background
        bg_canvas = tk.Canvas(self, bg=C["bg"], highlightthickness=0)
        bg_canvas.place(relx=0, rely=0, relwidth=1, relheight=1)
        self.after(50, lambda: self._draw_bg(bg_canvas))

        center = tk.Frame(self, bg=C["panel"],
                          padx=48, pady=44)
        center.place(relx=0.5, rely=0.5, anchor="center")

        # Orange accent top bar
        tk.Frame(center, bg=C["accent"], height=3).pack(fill="x", pady=(0, 28))

        # Logo
        logo_f = tk.Frame(center, bg=C["panel"])
        logo_f.pack(pady=(0, 24))
        tk.Label(logo_f, text="⚔", font=("Segoe UI", 32),
                 fg=C["accent"], bg=C["panel"]).pack(side="left", padx=(0,12))
        title_f = tk.Frame(logo_f, bg=C["panel"])
        title_f.pack(side="left")
        tk.Label(title_f, text="ARCHER GRC",
                 font=("Consolas", 18, "bold"),
                 fg=C["text"], bg=C["panel"]).pack(anchor="w")
        tk.Label(title_f, text="Governance · Risk · Compliance",
                 font=("Segoe UI", 8), fg=C["text2"],
                 bg=C["panel"]).pack(anchor="w")

        sep(center).pack(fill="x", pady=(0,24))

        # Fields
        tk.Label(center, text="ЛОГІН", font=("Segoe UI", 8, "bold"),
                 fg=C["text2"], bg=C["panel"], anchor="w").pack(fill="x")
        self.e_user = entry(center, width=34)
        self.e_user.pack(fill="x", ipady=8, pady=(4,14))

        tk.Label(center, text="ПАРОЛЬ", font=("Segoe UI", 8, "bold"),
                 fg=C["text2"], bg=C["panel"], anchor="w").pack(fill="x")
        self.e_pass = entry(center, width=34, show="●")
        self.e_pass.pack(fill="x", ipady=8, pady=(4,24))

        login_btn = btn(center, "  УВІЙТИ →  ", self._login, "accent", width=34)
        login_btn.pack(fill="x", ipady=6)

        tk.Label(center, text="admin / admin123  •  auditor / pass123",
                 font=("Segoe UI", 8), fg=C["text3"],
                 bg=C["panel"]).pack(pady=(16, 0))

        self.e_user.insert(0, "admin")
        self.e_user.bind("<Return>", lambda e: self._login())
        self.e_pass.bind("<Return>", lambda e: self._login())
        self.e_user.focus()

    def _draw_bg(self, cv):
        cv.delete("all")
        w, h = self.winfo_width(), self.winfo_height()
        # diagonal lines
        for i in range(-5, 25):
            offset = i * 80
            cv.create_line(offset, 0, offset + h, h,
                           fill=C["accent"] + "08", width=40)
        # corner accent
        cv.create_polygon(0, 0, 180, 0, 0, 100,
                          fill=C["accent"] + "15", outline="")
        cv.create_polygon(w, h, w-180, h, w, h-100,
                          fill=C["accent"] + "15", outline="")

    def _login(self):
        u = self.e_user.get().strip()
        p = self.e_pass.get().strip()
        conn = get_db()
        row = conn.execute(
            "SELECT * FROM users WHERE username=? AND password=? AND active=1", (u, p)
        ).fetchone()
        conn.close()
        if row:
            self.master.login_success(dict(row))
        else:
            self.e_pass.configure(highlightbackground=C["red"], highlightcolor=C["red"])
            messagebox.showerror("Помилка", "Невірний логін або пароль")

# ─────────────────────────────────────────────────────────────────
# ГОЛОВНИЙ ДОДАТОК
# ─────────────────────────────────────────────────────────────────

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Archer GRC Platform")
        self.geometry("1380x820")
        self.minsize(1100, 680)
        self.configure(bg=C["bg"])
        self.current_user = None
        apply_styles(self)
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

    def log_action(self, action, module, record_id=None, details=""):
        if not self.current_user: return
        conn = get_db()
        conn.execute(
            "INSERT INTO audit_log (user_id, action, module, record_id, details) VALUES (?,?,?,?,?)",
            (self.current_user["id"], action, module, record_id, details))
        conn.commit()
        conn.close()

# ─────────────────────────────────────────────────────────────────
# ГОЛОВНИЙ LAYOUT
# ─────────────────────────────────────────────────────────────────

NAV_ITEMS = [
    ("PAGES", None),
    ("🏠  Дашборд",       "dashboard"),
    ("⚠   Ризики",        "risks"),
    ("🛡   Контролі",      "controls"),
    ("📋  Політики",       "policies"),
    ("🏢  Постачальники",  "suppliers"),
    ("🔍  Аудит",          "audit"),
    ("🚨  Інциденти",      "incidents"),
    ("⚖   Регулятори",     "regulations"),
    ("СИСТЕМА", None),
    ("👥  Користувачі",    "users"),
    ("📜  Журнал дій",     "auditlog"),
]

class MainLayout(tk.Frame):
    def __init__(self, master):
        super().__init__(master, bg=C["bg"])
        self.pack(fill="both", expand=True)
        self._nav_btns = {}
        self._active = tk.StringVar(value="dashboard")
        self._build()
        self._navigate("dashboard")

    def _build(self):
        # ── SIDEBAR ──
        self.sidebar = tk.Frame(self, bg=C["panel"], width=C["sidebar_w"])
        self.sidebar.pack(side="left", fill="y")
        self.sidebar.pack_propagate(False)

        # Logo top
        logo_f = tk.Frame(self.sidebar, bg=C["panel"], pady=0)
        logo_f.pack(fill="x")
        # orange top accent
        tk.Frame(logo_f, bg=C["accent"], height=3).pack(fill="x")
        inner_logo = tk.Frame(logo_f, bg=C["panel"], padx=18, pady=16)
        inner_logo.pack(fill="x")
        tk.Label(inner_logo, text="⚔ ARCHER",
                 font=("Consolas", 13, "bold"),
                 fg=C["text"], bg=C["panel"]).pack(anchor="w")
        tk.Label(inner_logo, text="GRC Platform v2.0",
                 font=("Segoe UI", 7), fg=C["text3"],
                 bg=C["panel"]).pack(anchor="w")

        sep(self.sidebar).pack(fill="x")

        # Nav
        nav_scroll = tk.Frame(self.sidebar, bg=C["panel"])
        nav_scroll.pack(fill="both", expand=True, pady=8)

        for label, key in NAV_ITEMS:
            if key is None:
                # Section header
                tk.Label(nav_scroll, text=label,
                         font=("Segoe UI", 7, "bold"),
                         fg=C["text3"], bg=C["panel"],
                         anchor="w", padx=18).pack(fill="x", pady=(12, 2))
            else:
                self._make_nav(nav_scroll, label, key)

        # Bottom: user info
        sep(self.sidebar).pack(fill="x", side="bottom")
        user_f = tk.Frame(self.sidebar, bg=C["panel"], padx=16, pady=12)
        user_f.pack(fill="x", side="bottom")
        u = self.master.current_user
        tk.Label(user_f, text=u["full_name"][:22],
                 font=("Segoe UI", 9, "bold"),
                 fg=C["text"], bg=C["panel"],
                 anchor="w").pack(fill="x")
        tk.Label(user_f, text=u["role"],
                 font=("Segoe UI", 7), fg=C["text3"],
                 bg=C["panel"], anchor="w").pack(fill="x")
        btn(user_f, "↩ Вийти", self.master.logout, "ghost").pack(fill="x", pady=(8,0))

        # ── MAIN AREA ──
        right = tk.Frame(self, bg=C["bg"])
        right.pack(side="left", fill="both", expand=True)

        # Header
        self.header = tk.Frame(right, bg=C["panel"], height=C["header_h"])
        self.header.pack(fill="x")
        self.header.pack_propagate(False)
        sep(self.header, color=C["border"]).pack(fill="x", side="bottom")

        # orange left accent on header
        tk.Frame(self.header, bg=C["accent"], width=3).pack(side="left", fill="y")

        self.h_title = tk.Label(self.header, text="",
                                font=("Consolas", 13, "bold"),
                                fg=C["text"], bg=C["panel"])
        self.h_title.pack(side="left", padx=20)

        # date/user badge right
        info_f = tk.Frame(self.header, bg=C["panel"])
        info_f.pack(side="right", padx=20)
        tk.Label(info_f, text=datetime.datetime.now().strftime("%d.%m.%Y"),
                 font=("Consolas", 9), fg=C["text2"],
                 bg=C["panel"]).pack(side="right", padx=(12,0))
        tk.Label(info_f,
                 text=f"● {self.master.current_user['username']}",
                 font=("Segoe UI", 9, "bold"),
                 fg=C["accent"], bg=C["panel"]).pack(side="right")

        # Content
        self.content = tk.Frame(right, bg=C["bg"])
        self.content.pack(fill="both", expand=True)

    def _make_nav(self, parent, label, key):
        f = tk.Frame(parent, bg=C["panel"], cursor="hand2")
        f.pack(fill="x", padx=10, pady=1)

        indicator = tk.Frame(f, bg=C["panel"], width=3)
        indicator.pack(side="left", fill="y")

        lbl_w = tk.Label(f, text=label,
                         font=("Segoe UI", 9),
                         fg=C["text2"], bg=C["panel"],
                         anchor="w", padx=12, pady=9,
                         cursor="hand2")
        lbl_w.pack(side="left", fill="x", expand=True)

        self._nav_btns[key] = (f, lbl_w, indicator)

        def click(k=key): self._navigate(k)
        def enter(e, k=key):
            if self._active.get() != k:
                lbl_w.configure(fg=C["text"], bg=C["panel2"])
                f.configure(bg=C["panel2"])
        def leave(e, k=key):
            if self._active.get() != k:
                lbl_w.configure(fg=C["text2"], bg=C["panel"])
                f.configure(bg=C["panel"])

        for w in [f, lbl_w]:
            w.bind("<Button-1>", lambda e, k=key: click(k))
            w.bind("<Enter>", enter)
            w.bind("<Leave>", leave)

    def _navigate(self, key):
        self._active.set(key)
        label_map = {k: l for l, k in NAV_ITEMS if k}

        for k, (f, lbl_w, ind) in self._nav_btns.items():
            if k == key:
                lbl_w.configure(fg=C["accent"], bg=C["panel2"],
                                 font=("Segoe UI", 9, "bold"))
                f.configure(bg=C["panel2"])
                ind.configure(bg=C["accent"])
            else:
                lbl_w.configure(fg=C["text2"], bg=C["panel"],
                                 font=("Segoe UI", 9))
                f.configure(bg=C["panel"])
                ind.configure(bg=C["panel"])

        self.h_title.configure(text=label_map.get(key, "").strip())

        for w in self.content.winfo_children():
            w.destroy()

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

# ─────────────────────────────────────────────────────────────────
# ДАШБОРД
# ─────────────────────────────────────────────────────────────────

class DashboardModule(tk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent, bg=C["bg"])
        self.app = app
        self.pack(fill="both", expand=True)
        self._build()

    def _stats(self):
        conn = get_db()
        s = {
            "open_risks":      conn.execute("SELECT COUNT(*) FROM risks WHERE status='Відкритий'").fetchone()[0],
            "total_risks":     conn.execute("SELECT COUNT(*) FROM risks").fetchone()[0],
            "controls":        conn.execute("SELECT COUNT(*) FROM controls WHERE status='Активний'").fetchone()[0],
            "policies":        conn.execute("SELECT COUNT(*) FROM policies WHERE status='Опублікована'").fetchone()[0],
            "incidents":       conn.execute("SELECT COUNT(*) FROM incidents WHERE status NOT IN ('Вирішений','Скасовано')").fetchone()[0],
            "suppliers":       conn.execute("SELECT COUNT(*) FROM suppliers WHERE status='Активний'").fetchone()[0],
            "compliant":       conn.execute("SELECT COUNT(*) FROM regulations WHERE compliance_status='Відповідає'").fetchone()[0],
            "total_regs":      conn.execute("SELECT COUNT(*) FROM regulations").fetchone()[0],
            "risks_data":      [dict(r) for r in conn.execute("SELECT likelihood, impact, risk_score FROM risks WHERE status!='Закритий'").fetchall()],
            "top_risks":       [dict(r) for r in conn.execute("""SELECT r.title, r.risk_score, r.status, u.full_name as owner
                                    FROM risks r LEFT JOIN users u ON r.owner_id=u.id
                                    ORDER BY r.risk_score DESC LIMIT 5""").fetchall()],
            "regs":            [dict(r) for r in conn.execute("SELECT title, compliance_status FROM regulations").fetchall()],
            "incidents_list":  [dict(r) for r in conn.execute("SELECT title, severity, status FROM incidents WHERE status NOT IN ('Вирішений','Скасовано') LIMIT 5").fetchall()],
        }
        conn.close()
        return s

    def _build(self):
        s = self._stats()

        # Scrollable
        cv = tk.Canvas(self, bg=C["bg"], highlightthickness=0)
        vsb = ttk.Scrollbar(self, orient="vertical", command=cv.yview)
        cv.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        cv.pack(fill="both", expand=True)
        inner = tk.Frame(cv, bg=C["bg"])
        win = cv.create_window((0,0), window=inner, anchor="nw")
        cv.bind("<Configure>", lambda e: cv.itemconfig(win, width=e.width))
        inner.bind("<Configure>", lambda e: cv.configure(scrollregion=cv.bbox("all")))
        cv.bind_all("<MouseWheel>", lambda e: cv.yview_scroll(-1*(e.delta//120), "units"))

        p = inner
        outer_pad = {"padx": 24, "pady": 0}

        # ── KPI ROW ──
        kpi_row = tk.Frame(p, bg=C["bg"])
        kpi_row.pack(fill="x", padx=24, pady=(20, 0))

        kpis = [
            ("ВІДКРИТІ РИЗИКИ",   str(s["open_risks"]),  f"з {s['total_risks']} загалом",     C["red"]),
            ("АКТИВНІ КОНТРОЛІ",  str(s["controls"]),    "контролів активно",                  C["green"]),
            ("ОПУБЛ. ПОЛІТИКИ",   str(s["policies"]),    "документів актуальні",               C["accent"]),
            ("ІНЦИДЕНТИ",         str(s["incidents"]),   "потребують уваги",                   C["yellow"]),
            ("ПОСТАЧАЛЬНИКИ",     str(s["suppliers"]),   "активних партнерів",                 C["blue"]),
            ("ВІДПОВІДНІСТЬ",     f"{s['compliant']}/{s['total_regs']}", "регуляторів виконано", C["green"]),
        ]
        for i, (title, val, sub, color) in enumerate(kpis):
            card = tk.Frame(kpi_row, bg=C["panel"], padx=18, pady=16)
            card.grid(row=0, column=i, padx=(0,12) if i < 5 else (0,0), sticky="nsew")
            kpi_row.columnconfigure(i, weight=1)

            # top accent line per card
            tk.Frame(card, bg=color, height=2).pack(fill="x", pady=(0, 12))
            tk.Label(card, text=val,
                     font=("Consolas", 26, "bold"),
                     fg=color, bg=C["panel"]).pack(anchor="w")
            tk.Label(card, text=title,
                     font=("Segoe UI", 7, "bold"),
                     fg=C["text2"], bg=C["panel"]).pack(anchor="w")
            tk.Label(card, text=sub,
                     font=("Segoe UI", 7),
                     fg=C["text3"], bg=C["panel"]).pack(anchor="w", pady=(2,0))

        # ── CHART ROW ──
        chart_row = tk.Frame(p, bg=C["bg"])
        chart_row.pack(fill="x", padx=24, pady=(16, 0))
        chart_row.columnconfigure(0, weight=3)
        chart_row.columnconfigure(1, weight=2)

        # Risk trend chart (simulated)
        trend_card = tk.Frame(chart_row, bg=C["panel"], padx=16, pady=14)
        trend_card.grid(row=0, column=0, padx=(0,12), sticky="nsew")
        tk.Frame(trend_card, bg=C["accent"], height=2).pack(fill="x", pady=(0,10))

        hdr_f = tk.Frame(trend_card, bg=C["panel"])
        hdr_f.pack(fill="x")
        tk.Label(hdr_f, text="Динаміка ризиків",
                 font=("Consolas", 10, "bold"),
                 fg=C["text"], bg=C["panel"]).pack(side="left")
        val_f = tk.Frame(hdr_f, bg=C["panel"])
        val_f.pack(side="right")
        tk.Label(val_f, text=f"{s['open_risks']}",
                 font=("Consolas", 18, "bold"),
                 fg=C["accent"], bg=C["panel"]).pack(side="left")
        tk.Label(val_f, text=" відкритих",
                 font=("Segoe UI", 8),
                 fg=C["text2"], bg=C["panel"]).pack(side="left", pady=(6,0))

        # Simulated trend data
        base = s["total_risks"]
        trend_data = [max(1, base - 2 + random.randint(-1, 2)) for _ in range(24)]
        trend_data[-1] = s["open_risks"]
        MiniChart(trend_card, trend_data, color=C["accent"],
                  height=130).pack(fill="x", pady=(10,0))

        # Period selector
        period_f = tk.Frame(trend_card, bg=C["panel"])
        period_f.pack(fill="x", pady=(8,0))
        for period in ["1Д","7Д","1М","3М","6М","ALL"]:
            cl = C["accent"] if period == "1Д" else C["text3"]
            tk.Label(period_f, text=period,
                     font=("Segoe UI", 7, "bold"),
                     fg=cl, bg=C["panel"],
                     padx=6, pady=3,
                     cursor="hand2").pack(side="left", padx=2)

        # Donut
        donut_card = tk.Frame(chart_row, bg=C["panel"], padx=16, pady=14)
        donut_card.grid(row=0, column=1, sticky="nsew")
        tk.Frame(donut_card, bg=C["blue"], height=2).pack(fill="x", pady=(0,10))
        tk.Label(donut_card, text="Розподіл ризиків",
                 font=("Consolas", 10, "bold"),
                 fg=C["text"], bg=C["panel"]).pack(anchor="w")

        conn = get_db()
        by_cat = conn.execute("""SELECT category, COUNT(*) as cnt FROM risks
                                  WHERE status!='Закритий' GROUP BY category""").fetchall()
        conn.close()
        seg_colors = [C["accent"], C["red"], C["yellow"], C["blue"], C["green"], "#AA44FF"]
        segments = [(r["category"] or "Інше", r["cnt"], seg_colors[i % len(seg_colors)])
                    for i, r in enumerate(by_cat)]
        if not segments:
            segments = [("Немає даних", 1, C["border"])]

        DonutChart(donut_card, segments, height=140).pack(fill="x", pady=(8,4))

        # Legend
        for label, val, color in segments:
            leg = tk.Frame(donut_card, bg=C["panel"])
            leg.pack(fill="x", pady=1)
            tk.Frame(leg, bg=color, width=8, height=8).pack(side="left", padx=(0,6))
            tk.Label(leg, text=f"{label}",
                     font=("Segoe UI", 7), fg=C["text2"],
                     bg=C["panel"]).pack(side="left")
            tk.Label(leg, text=str(val),
                     font=("Consolas", 7, "bold"), fg=color,
                     bg=C["panel"]).pack(side="right")

        # ── BOTTOM ROW ──
        bot_row = tk.Frame(p, bg=C["bg"])
        bot_row.pack(fill="x", padx=24, pady=(16, 24))
        bot_row.columnconfigure(0, weight=2)
        bot_row.columnconfigure(1, weight=1)
        bot_row.columnconfigure(2, weight=2)

        # Heatmap
        hmap_card = tk.Frame(bot_row, bg=C["panel"], padx=16, pady=14)
        hmap_card.grid(row=0, column=0, padx=(0,12), sticky="nsew")
        tk.Frame(hmap_card, bg=C["red"], height=2).pack(fill="x", pady=(0,10))
        tk.Label(hmap_card, text="Теплова карта ризиків",
                 font=("Consolas", 10, "bold"),
                 fg=C["text"], bg=C["panel"]).pack(anchor="w")
        HeatmapCanvas(hmap_card, s["risks_data"],
                      height=180).pack(fill="x", pady=(8,0), expand=True)

        # Compliance status
        comp_card = tk.Frame(bot_row, bg=C["panel"], padx=16, pady=14)
        comp_card.grid(row=0, column=1, padx=(0,12), sticky="nsew")
        tk.Frame(comp_card, bg=C["green"], height=2).pack(fill="x", pady=(0,10))
        tk.Label(comp_card, text="Відповідність",
                 font=("Consolas", 10, "bold"),
                 fg=C["text"], bg=C["panel"]).pack(anchor="w")
        for reg in s["regs"]:
            rf = tk.Frame(comp_card, bg=C["panel2"], pady=5, padx=8)
            rf.pack(fill="x", pady=2)
            color = STATUS_COLOR.get(reg["compliance_status"], C["text3"])
            tk.Label(rf, text="●", fg=color, bg=C["panel2"],
                     font=("Segoe UI", 8)).pack(side="left", padx=(0,5))
            tk.Label(rf, text=reg["title"][:20],
                     font=("Segoe UI", 8), fg=C["text"],
                     bg=C["panel2"]).pack(side="left")

        # Top risks
        tr_card = tk.Frame(bot_row, bg=C["panel"], padx=16, pady=14)
        tr_card.grid(row=0, column=2, sticky="nsew")
        tk.Frame(tr_card, bg=C["yellow"], height=2).pack(fill="x", pady=(0,10))
        tk.Label(tr_card, text="Топ ризики",
                 font=("Consolas", 10, "bold"),
                 fg=C["text"], bg=C["panel"]).pack(anchor="w")
        for risk in s["top_risks"]:
            rf = tk.Frame(tr_card, bg=C["panel2"], pady=6, padx=10)
            rf.pack(fill="x", pady=2)
            score = risk.get("risk_score") or 0
            col = risk_color(score)
            sc_f = tk.Frame(rf, bg=col, width=32, height=22)
            sc_f.pack(side="left", padx=(0,8))
            sc_f.pack_propagate(False)
            tk.Label(sc_f, text=str(score),
                     font=("Consolas", 9, "bold"),
                     fg="#000" if col == C["green"] else "#fff",
                     bg=col).place(relx=0.5, rely=0.5, anchor="center")
            tk.Label(rf, text=risk["title"][:28],
                     font=("Segoe UI", 8), fg=C["text"],
                     bg=C["panel2"]).pack(side="left")


# ─────────────────────────────────────────────────────────────────
# БАЗОВИЙ CRUD МОДУЛЬ
# ─────────────────────────────────────────────────────────────────

class BaseCRUD(tk.Frame):
    TABLE = ""
    COLUMNS = []  # (key, label, width)

    def __init__(self, parent, app):
        super().__init__(parent, bg=C["bg"])
        self.app = app
        self.pack(fill="both", expand=True)
        self._build_bar()
        self._build_table()
        self.refresh()

    def _build_bar(self):
        bar = tk.Frame(self, bg=C["panel"], padx=16, pady=10)
        bar.pack(fill="x")
        sep(bar, color=C["accent"], orient="v").pack(side="left", fill="y", padx=(0,14))

        btn(bar, "+  Додати", self._add, "accent").pack(side="left", padx=(0,8))
        btn(bar, "✎  Редагувати", self._edit, "ghost").pack(side="left", padx=(0,8))
        btn(bar, "✕  Видалити", self._delete, "danger").pack(side="left")

        # search right
        search_f = tk.Frame(bar, bg=C["panel2"],
                            highlightthickness=1,
                            highlightbackground=C["border2"],
                            highlightcolor=C["accent"])
        search_f.pack(side="right")
        tk.Label(search_f, text="⌕", font=("Segoe UI", 11),
                 fg=C["text3"], bg=C["panel2"]).pack(side="left", padx=(8,2))
        self._search_var = tk.StringVar()
        se = tk.Entry(search_f, textvariable=self._search_var,
                      bg=C["panel2"], fg=C["text"],
                      insertbackground=C["accent"],
                      relief="flat", bd=0,
                      font=("Segoe UI", 9), width=22)
        se.pack(side="left", ipady=6, padx=(0,8))
        self._search_var.trace_add("write", lambda *_: self.refresh())

        sep(self, color=C["border"]).pack(fill="x")

    def _build_table(self):
        f = tk.Frame(self, bg=C["bg"])
        f.pack(fill="both", expand=True, padx=16, pady=12)

        cols = [c[0] for c in self.COLUMNS]
        self.tree = ttk.Treeview(f, columns=cols, show="headings", selectmode="browse")
        vsb = ttk.Scrollbar(f, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        self.tree.pack(fill="both", expand=True)

        for key, label, width in self.COLUMNS:
            self.tree.heading(key, text=label, anchor="w")
            self.tree.column(key, width=width, minwidth=50, anchor="w")

        self.tree.tag_configure("even", background=C["panel"])
        self.tree.tag_configure("odd",  background=C["panel2"])
        self.tree.bind("<Double-1>", lambda e: self._edit())

    def get_rows(self, q=""):
        return []

    def refresh(self):
        q = self._search_var.get().lower() if hasattr(self, "_search_var") else ""
        self.tree.delete(*self.tree.get_children())
        for i, row in enumerate(self.get_rows(q)):
            vals = [row.get(c[0], "") or "" for c in self.COLUMNS]
            tag = "even" if i % 2 == 0 else "odd"
            self.tree.insert("", "end", values=vals,
                             iid=str(row.get("id")), tags=(tag,))

    def sel_id(self):
        s = self.tree.selection()
        return int(s[0]) if s else None

    def _add(self): self.open_form(None)

    def _edit(self):
        rid = self.sel_id()
        if not rid:
            messagebox.showwarning("Увага", "Виберіть запис")
            return
        self.open_form(rid)

    def _delete(self):
        rid = self.sel_id()
        if not rid:
            messagebox.showwarning("Увага", "Виберіть запис для видалення")
            return
        if messagebox.askyesno("Підтвердження", "Видалити вибраний запис?"):
            conn = get_db()
            conn.execute(f"DELETE FROM {self.TABLE} WHERE id=?", (rid,))
            conn.commit()
            conn.close()
            self.app.log_action("DELETE", self.TABLE, rid)
            self.refresh()

    def open_form(self, rid):
        pass

# ─────────────────────────────────────────────────────────────────
# UNIVERSAL FORM
# ─────────────────────────────────────────────────────────────────

class UniversalForm(tk.Toplevel):
    def __init__(self, pm, app, rid, table, title, fields):
        super().__init__()
        self.pm = pm
        self.app = app
        self.rid = rid
        self.table = table
        self.fields = fields
        self.title(f"{'Редагувати' if rid else 'Новий'}: {title}")
        self.geometry("600x560")
        self.configure(bg=C["bg"])
        self.resizable(False, True)

        conn = get_db()
        self._users = conn.execute("SELECT id, full_name FROM users WHERE active=1").fetchall()
        conn.close()
        self._unames = [u["full_name"] for u in self._users]
        self._uids   = [u["id"] for u in self._users]

        self._widgets = {}
        self._build()
        if rid:
            self._populate()
        self.grab_set()

    def _build(self):
        # Orange top bar
        tk.Frame(self, bg=C["accent"], height=3).pack(fill="x")

        cv = tk.Canvas(self, bg=C["bg"], highlightthickness=0)
        vsb = ttk.Scrollbar(self, orient="vertical", command=cv.yview)
        cv.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        cv.pack(fill="both", expand=True)
        f = tk.Frame(cv, bg=C["bg"], padx=28, pady=20)
        cv.create_window((0,0), window=f, anchor="nw")
        f.bind("<Configure>", lambda e: cv.configure(scrollregion=cv.bbox("all")))

        for col, label, wtype, *opts in self.fields:
            tk.Label(f, text=label.upper(),
                     font=("Segoe UI", 7, "bold"),
                     fg=C["text3"], bg=C["bg"],
                     anchor="w").pack(fill="x", pady=(10,2))

            if wtype == "entry":
                w = entry(f, width=60)
                w.pack(fill="x", ipady=7)
            elif wtype == "combo":
                w = combo(f, opts[0], width=57)
                w.pack(fill="x", ipady=4)
                if opts[0]: w.set(opts[0][0])
            elif wtype == "text":
                w = tk.Text(f, height=3, width=60,
                            bg=C["panel2"], fg=C["text"],
                            insertbackground=C["accent"],
                            relief="flat", font=("Segoe UI", 9),
                            highlightthickness=1,
                            highlightbackground=C["border2"],
                            highlightcolor=C["accent"],
                            wrap="word")
                w.pack(fill="x")
            elif wtype == "user":
                w = combo(f, self._unames, width=57)
                w.pack(fill="x", ipady=4)
            self._widgets[col] = (wtype, w)

        sep(f, color=C["border"]).pack(fill="x", pady=(20,16))
        bf = tk.Frame(f, bg=C["bg"])
        bf.pack(fill="x")
        btn(bf, "💾  Зберегти", self._save, "accent").pack(side="left", padx=(0,8), ipady=4)
        btn(bf, "✕  Скасувати", self.destroy, "ghost").pack(side="left", ipady=4)

    def _populate(self):
        conn = get_db()
        row = dict(conn.execute(f"SELECT * FROM {self.table} WHERE id=?", (self.rid,)).fetchone() or {})
        conn.close()
        for col, label, wtype, *opts in self.fields:
            wt, w = self._widgets[col]
            val = row.get(col, "") or ""
            if wt == "entry":
                w.insert(0, str(val))
            elif wt == "combo":
                choices = opts[0] if opts else []
                if str(val) in choices: w.set(str(val))
            elif wt == "text":
                w.insert("1.0", str(val))
            elif wt == "user":
                uid = row.get(col)
                if uid and uid in self._uids:
                    w.set(self._unames[self._uids.index(uid)])

    def _save(self):
        first_col = self.fields[0][0]
        wt0, w0 = self._widgets[first_col]
        v0 = w0.get().strip() if wt0 in ("entry","combo") else w0.get("1.0","end-1c").strip()
        if not v0:
            messagebox.showerror("Помилка", f"Поле '{self.fields[0][1]}' є обов'язковим")
            return
        data = {}
        for col, label, wtype, *opts in self.fields:
            wt, w = self._widgets[col]
            if wt == "entry":   data[col] = w.get().strip() or None
            elif wt == "combo": data[col] = w.get() or None
            elif wt == "text":  data[col] = w.get("1.0","end-1c").strip() or None
            elif wt == "user":
                name = w.get()
                data[col] = self._uids[self._unames.index(name)] if name in self._unames else None
        conn = get_db()
        if self.rid:
            sql = ", ".join(f"{k}=?" for k in data)
            conn.execute(f"UPDATE {self.table} SET {sql} WHERE id=?", [*data.values(), self.rid])
            self.app.log_action("UPDATE", self.table, self.rid)
        else:
            keys = ", ".join(data.keys())
            ph   = ", ".join("?" * len(data))
            conn.execute(f"INSERT INTO {self.table} ({keys}) VALUES ({ph})", list(data.values()))
            self.app.log_action("CREATE", self.table)
        conn.commit()
        conn.close()
        self.pm.refresh()
        self.destroy()

# ─────────────────────────────────────────────────────────────────
# МОДУЛІ
# ─────────────────────────────────────────────────────────────────

class RisksModule(BaseCRUD):
    TABLE = "risks"
    COLUMNS = [
        ("id","ID",50), ("title","Назва ризику",220), ("category","Категорія",120),
        ("likelihood","Ймов.",60), ("impact","Вплив",60), ("risk_score","Оцінка",70),
        ("status","Статус",110), ("owner","Власник",150), ("review_date","Огляд",100),
    ]
    def get_rows(self, q=""):
        conn = get_db()
        rows = conn.execute("""SELECT r.id, r.title, r.category, r.likelihood, r.impact,
                   r.risk_score, r.status, r.review_date, u.full_name as owner
                   FROM risks r LEFT JOIN users u ON r.owner_id=u.id
                   ORDER BY r.risk_score DESC""").fetchall()
        conn.close()
        return [dict(r) for r in rows if not q or q in r["title"].lower() or q in (r["category"] or "").lower()]

    def open_form(self, rid):
        UniversalForm(self, self.app, rid, "risks", "Ризик", [
            ("title","Назва *","entry"),
            ("category","Категорія","combo",["Кіберризик","Операційний","Фінансовий","Регуляторний","Репутаційний","Стратегічний"]),
            ("likelihood","Ймовірність (1-5)","combo",["1","2","3","4","5"]),
            ("impact","Вплив (1-5)","combo",["1","2","3","4","5"]),
            ("status","Статус","combo",["Відкритий","Пом'якшений","Прийнятий","Закритий"]),
            ("owner_id","Власник ризику","user"),
            ("department","Підрозділ","entry"),
            ("review_date","Дата огляду (РРРР-ММ-ДД)","entry"),
            ("mitigation_plan","План пом'якшення","text"),
            ("description","Опис","text"),
        ])

class ControlsModule(BaseCRUD):
    TABLE = "controls"
    COLUMNS = [
        ("id","ID",50), ("title","Назва контролю",220), ("type","Тип",110),
        ("frequency","Частота",110), ("status","Статус",100),
        ("effectiveness","Ефективність",150), ("owner","Відповідальний",150), ("last_tested","Тест",100),
    ]
    def get_rows(self, q=""):
        conn = get_db()
        rows = conn.execute("""SELECT c.id, c.title, c.type, c.frequency, c.status,
                   c.effectiveness, c.last_tested, u.full_name as owner
                   FROM controls c LEFT JOIN users u ON c.owner_id=u.id ORDER BY c.id""").fetchall()
        conn.close()
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
    TABLE = "policies"
    COLUMNS = [
        ("id","ID",50), ("title","Назва",220), ("category","Категорія",130),
        ("version","Версія",70), ("status","Статус",130),
        ("owner","Власник",150), ("review_date","Огляд",100),
    ]
    def get_rows(self, q=""):
        conn = get_db()
        rows = conn.execute("""SELECT p.id, p.title, p.category, p.version, p.status,
                   p.review_date, u.full_name as owner
                   FROM policies p LEFT JOIN users u ON p.owner_id=u.id ORDER BY p.id""").fetchall()
        conn.close()
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
    TABLE = "suppliers"
    COLUMNS = [
        ("id","ID",50), ("name","Назва",200), ("category","Категорія",130),
        ("risk_level","Ризик",80), ("status","Статус",100),
        ("contact_name","Контакт",150), ("contract_end","Договір до",100),
    ]
    def get_rows(self, q=""):
        conn = get_db()
        rows = conn.execute("SELECT * FROM suppliers ORDER BY id").fetchall()
        conn.close()
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
    TABLE = "audits"
    COLUMNS = [
        ("id","ID",50), ("title","Назва",220), ("type","Тип",110),
        ("department","Підрозділ",120), ("start_date","Початок",100),
        ("end_date","Кінець",100), ("status","Статус",110), ("auditor","Аудитор",150),
    ]
    def get_rows(self, q=""):
        conn = get_db()
        rows = conn.execute("""SELECT a.id, a.title, a.type, a.department, a.start_date,
                   a.end_date, a.status, u.full_name as auditor
                   FROM audits a LEFT JOIN users u ON a.auditor_id=u.id ORDER BY a.id DESC""").fetchall()
        conn.close()
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
    TABLE = "incidents"
    COLUMNS = [
        ("id","ID",50), ("title","Назва",200), ("category","Категорія",120),
        ("severity","Серйозність",100), ("status","Статус",110),
        ("reporter","Репортер",140), ("occurred_at","Дата виникнення",120),
    ]
    def get_rows(self, q=""):
        conn = get_db()
        rows = conn.execute("""SELECT i.id, i.title, i.category, i.severity, i.status,
                   i.occurred_at, u.full_name as reporter
                   FROM incidents i LEFT JOIN users u ON i.reporter_id=u.id ORDER BY i.id DESC""").fetchall()
        conn.close()
        return [dict(r) for r in rows if not q or q in r["title"].lower()]

    def open_form(self, rid):
        UniversalForm(self, self.app, rid, "incidents", "Інцидент", [
            ("title","Назва *","entry"),
            ("category","Категорія","combo",["Кібербезпека","Операційний","ІТ-збій","Шахрайство","Витік даних","Фізична безпека"]),
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
    TABLE = "regulations"
    COLUMNS = [
        ("id","ID",50), ("title","Назва",220), ("authority","Регулятор",100),
        ("category","Категорія",130), ("compliance_status","Відповідність",140),
        ("effective_date","Набрання чинності",120), ("responsible","Відповідальний",150),
    ]
    def get_rows(self, q=""):
        conn = get_db()
        rows = conn.execute("""SELECT r.id, r.title, r.authority, r.category,
                   r.compliance_status, r.effective_date, u.full_name as responsible
                   FROM regulations r LEFT JOIN users u ON r.responsible_id=u.id ORDER BY r.id""").fetchall()
        conn.close()
        return [dict(r) for r in rows if not q or q in r["title"].lower()]

    def open_form(self, rid):
        UniversalForm(self, self.app, rid, "regulations", "Регуляторна вимога", [
            ("title","Назва *","entry"),
            ("authority","Регулятор","entry"),
            ("category","Категорія","combo",["Кіберзахист","Персональні дані","Інформаційна безпека","Платіжні системи","Банківський нагляд"]),
            ("compliance_status","Статус відповідності","combo",["Відповідає","Частково відповідає","Не відповідає","В процесі","Не оцінено"]),
            ("effective_date","Набрання чинності","entry"),
            ("responsible_id","Відповідальний","user"),
            ("description","Опис","text"),
            ("notes","Примітки","text"),
        ])

# ─────────────────────────────────────────────────────────────────
# КОРИСТУВАЧІ
# ─────────────────────────────────────────────────────────────────

class UsersModule(BaseCRUD):
    TABLE = "users"
    COLUMNS = [
        ("id","ID",50), ("username","Логін",130), ("full_name","ПІБ",200),
        ("role","Роль",170), ("department","Підрозділ",130),
        ("email","Email",170), ("active_str","Активний",80),
    ]
    def get_rows(self, q=""):
        conn = get_db()
        rows = conn.execute("SELECT * FROM users ORDER BY id").fetchall()
        conn.close()
        result = []
        for r in rows:
            d = dict(r)
            d["active_str"] = "✓ Так" if d["active"] else "✗ Ні"
            if not q or q in d["full_name"].lower() or q in d["username"].lower():
                result.append(d)
        return result

    def open_form(self, rid):
        UserForm(self, self.app, rid)

class UserForm(tk.Toplevel):
    ROLES = ["Користувач (End User)","Власник ризику","Адміністратор додатка",
             "Системний адміністратор","Комплаєнс-офіцер","Аудитор"]

    def __init__(self, pm, app, rid=None):
        super().__init__()
        self.pm = pm
        self.app = app
        self.rid = rid
        self.title("Редагувати користувача" if rid else "Новий користувач")
        self.geometry("480x500")
        self.configure(bg=C["bg"])
        self.resizable(False, False)
        self._build()
        if rid: self._populate()
        self.grab_set()

    def _build(self):
        tk.Frame(self, bg=C["accent"], height=3).pack(fill="x")
        f = tk.Frame(self, bg=C["bg"], padx=28, pady=20)
        f.pack(fill="both", expand=True)

        def row(label, widget):
            tk.Label(f, text=label.upper(), font=("Segoe UI",7,"bold"),
                     fg=C["text3"], bg=C["bg"], anchor="w").pack(fill="x", pady=(10,2))
            widget.pack(fill="x", ipady=7)
            return widget

        self.f_user  = row("Логін *",    entry(f, width=50))
        self.f_pass  = row("Пароль *",   entry(f, width=50, show="●"))
        self.f_name  = row("ПІБ *",      entry(f, width=50))
        self.f_email = row("Email",      entry(f, width=50))
        self.f_dept  = row("Підрозділ",  entry(f, width=50))
        self.f_role  = row("Роль",       combo(f, self.ROLES, width=47))
        self.f_role.set(self.ROLES[0])

        tk.Label(f, text="АКТИВНИЙ", font=("Segoe UI",7,"bold"),
                 fg=C["text3"], bg=C["bg"], anchor="w").pack(fill="x", pady=(10,2))
        self.f_active = tk.BooleanVar(value=True)
        tk.Checkbutton(f, variable=self.f_active, bg=C["bg"],
                       fg=C["text"], activebackground=C["bg"],
                       selectcolor=C["panel2"],
                       text=" Активний користувач",
                       font=("Segoe UI",9)).pack(anchor="w")

        sep(f, color=C["border"]).pack(fill="x", pady=(16,12))
        bf = tk.Frame(f, bg=C["bg"])
        bf.pack(fill="x")
        btn(bf, "💾  Зберегти", self._save, "accent").pack(side="left", padx=(0,8), ipady=4)
        btn(bf, "✕  Скасувати", self.destroy, "ghost").pack(side="left", ipady=4)

    def _populate(self):
        conn = get_db()
        r = dict(conn.execute("SELECT * FROM users WHERE id=?", (self.rid,)).fetchone())
        conn.close()
        self.f_user.insert(0, r["username"])
        self.f_name.insert(0, r["full_name"])
        self.f_email.insert(0, r.get("email") or "")
        self.f_dept.insert(0, r.get("department") or "")
        if r["role"] in self.ROLES: self.f_role.set(r["role"])
        self.f_active.set(bool(r["active"]))

    def _save(self):
        u = self.f_user.get().strip()
        p = self.f_pass.get().strip()
        fn= self.f_name.get().strip()
        if not u or not fn:
            messagebox.showerror("Помилка","Логін та ПІБ обов'язкові"); return
        if not self.rid and not p:
            messagebox.showerror("Помилка","Пароль обов'язковий для нового користувача"); return
        conn = get_db()
        try:
            if self.rid:
                upd = "username=?, full_name=?, email=?, department=?, role=?, active=?"
                vals = [u, fn, self.f_email.get().strip(), self.f_dept.get().strip(),
                        self.f_role.get(), int(self.f_active.get())]
                if p: upd += ", password=?"; vals.append(p)
                vals.append(self.rid)
                conn.execute(f"UPDATE users SET {upd} WHERE id=?", vals)
            else:
                conn.execute("""INSERT INTO users (username,password,full_name,email,department,role,active)
                                VALUES (?,?,?,?,?,?,?)""",
                             [u,p,fn,self.f_email.get().strip(),self.f_dept.get().strip(),
                              self.f_role.get(), int(self.f_active.get())])
            conn.commit()
        except sqlite3.IntegrityError:
            messagebox.showerror("Помилка","Такий логін вже існує")
            conn.close(); return
        conn.close()
        self.app.log_action("USER_SAVE","users",self.rid,f"login={u}")
        self.pm.refresh()
        self.destroy()

# ─────────────────────────────────────────────────────────────────
# ЖУРНАЛ ДІЙ
# ─────────────────────────────────────────────────────────────────

class AuditLogModule(BaseCRUD):
    TABLE = "audit_log"
    COLUMNS = [
        ("id","ID",50), ("created_at","Дата/Час",150), ("user","Користувач",150),
        ("action","Дія",110), ("module","Модуль",100),
        ("record_id","Запис",70), ("details","Деталі",350),
    ]

    def _build_bar(self):
        bar = tk.Frame(self, bg=C["panel"], padx=16, pady=10)
        bar.pack(fill="x")
        sep(bar, color=C["accent"], orient="v").pack(side="left", fill="y", padx=(0,14))
        tk.Label(bar, text="ЖУРНАЛ АУДИТУ — тільки читання",
                 font=("Consolas", 10, "bold"),
                 fg=C["text"], bg=C["panel"]).pack(side="left")
        self._search_var = tk.StringVar()
        sf = tk.Frame(bar, bg=C["panel2"],
                      highlightthickness=1, highlightbackground=C["border2"])
        sf.pack(side="right")
        tk.Label(sf, text="⌕", fg=C["text3"], bg=C["panel2"],
                 font=("Segoe UI",11)).pack(side="left", padx=(8,2))
        se = tk.Entry(sf, textvariable=self._search_var,
                      bg=C["panel2"], fg=C["text"],
                      insertbackground=C["accent"],
                      relief="flat", bd=0, font=("Segoe UI",9), width=20)
        se.pack(side="left", ipady=6, padx=(0,8))
        self._search_var.trace_add("write", lambda *_: self.refresh())
        sep(self, color=C["border"]).pack(fill="x")

    def get_rows(self, q=""):
        conn = get_db()
        rows = conn.execute("""SELECT al.id, al.created_at, al.action, al.module,
                   al.record_id, al.details, u.full_name as user
                   FROM audit_log al LEFT JOIN users u ON al.user_id=u.id
                   ORDER BY al.id DESC LIMIT 400""").fetchall()
        conn.close()
        return [dict(r) for r in rows if not q
                or q in (r["action"] or "").lower()
                or q in (r["module"] or "").lower()]

    def open_form(self, rid): pass

# ─────────────────────────────────────────────────────────────────
# ЗАПУСК
# ─────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    app = App()
    app.mainloop()
