"""
NEXUS Business Platform — ядро (DB + моделі + теми)
"""
import sqlite3, os, random
from datetime import datetime, timedelta
from pathlib import Path

DB_PATH = Path.home() / ".nexus_data.db"

# ── СХЕМА ──────────────────────────────────────────────────────
SCHEMA = """
CREATE TABLE IF NOT EXISTS contacts (
    id      INTEGER PRIMARY KEY AUTOINCREMENT,
    name    TEXT NOT NULL,
    company TEXT DEFAULT '',
    email   TEXT DEFAULT '',
    phone   TEXT DEFAULT '',
    type    TEXT DEFAULT 'Клієнт',
    status  TEXT DEFAULT 'Активний',
    created TEXT DEFAULT (date('now'))
);
CREATE TABLE IF NOT EXISTS deals (
    id       INTEGER PRIMARY KEY AUTOINCREMENT,
    title    TEXT NOT NULL,
    client   TEXT DEFAULT '',
    amount   REAL DEFAULT 0,
    stage    TEXT DEFAULT 'Новий',
    manager  TEXT DEFAULT '',
    close_dt TEXT DEFAULT '',
    prob     INTEGER DEFAULT 50,
    created  TEXT DEFAULT (date('now'))
);
CREATE TABLE IF NOT EXISTS tasks (
    id       INTEGER PRIMARY KEY AUTOINCREMENT,
    title    TEXT NOT NULL,
    priority TEXT DEFAULT 'Середній',
    status   TEXT DEFAULT 'Відкрита',
    due      TEXT DEFAULT '',
    owner    TEXT DEFAULT '',
    module   TEXT DEFAULT 'CRM',
    created  TEXT DEFAULT (date('now'))
);
CREATE TABLE IF NOT EXISTS invoices (
    id       INTEGER PRIMARY KEY AUTOINCREMENT,
    number   TEXT NOT NULL,
    client   TEXT DEFAULT '',
    amount   REAL DEFAULT 0,
    vat      REAL DEFAULT 0,
    status   TEXT DEFAULT 'Очікує',
    due      TEXT DEFAULT '',
    created  TEXT DEFAULT (date('now'))
);
CREATE TABLE IF NOT EXISTS risks (
    id       INTEGER PRIMARY KEY AUTOINCREMENT,
    title    TEXT NOT NULL,
    category TEXT DEFAULT 'Операційний',
    impact   TEXT DEFAULT 'Середній',
    prob     TEXT DEFAULT 'Можливо',
    level    TEXT DEFAULT 'Середній',
    owner    TEXT DEFAULT '',
    status   TEXT DEFAULT 'Відкритий',
    created  TEXT DEFAULT (date('now'))
);
"""

SEED_CONTACTS = [
    ("Олег Петренко",   "Альфа Корп",    "o.petrenko@alpha.ua",  "+380501234567", "VIP",      "Активний"),
    ("Ірина Коваль",    "Бета ТОВ",      "i.koval@beta.ua",      "+380671234568", "Клієнт",   "Активний"),
    ("Денис Мельник",   "Гамма ФОП",     "d.melnyk@gamma.ua",    "+380631234569", "Партнер",  "Активний"),
    ("Тетяна Бойко",    "Дельта ЛТД",    "t.boyko@delta.ua",     "+380991234570", "Клієнт",   "Неактивний"),
    ("Андрій Шевченко", "Омега LLC",     "a.shev@omega.ua",      "+380731234571", "Лід",      "Новий"),
    ("Юлія Кравченко",  "Зета Груп",     "y.krav@zeta.ua",       "+380681234572", "Клієнт",   "Активний"),
    ("Максим Лисенко",  "Каппа Inc",     "m.lys@kappa.ua",       "+380501234573", "VIP",      "Активний"),
    ("Олена Павленко",  "Сигма Corp",    "o.pav@sigma.ua",       "+380671234574", "Клієнт",   "Активний"),
    ("Сергій Марченко", "Тета ТОВ",      "s.march@teta.ua",      "+380631234575", "Партнер",  "Активний"),
    ("Ніна Гончар",     "Йота ФОП",      "n.honch@iota.ua",      "+380991234576", "Лід",      "Новий"),
    ("Василь Кузьменко","Ламбда ЛТД",    "v.kuz@lambda.ua",      "+380731234577", "Клієнт",   "Активний"),
    ("Марина Ткаченко", "Мю Corp",       "m.tkach@mu.ua",        "+380681234578", "VIP",      "Активний"),
]
SEED_DEALS = [
    ("Впровадження ERP", "Альфа Корп",  480000, "Переговори", "О.Мельник",  "30.06.26", 70),
    ("Ліцензія CRM",     "Бета ТОВ",    120000, "Пропозиція", "І.Коваль",   "15.05.26", 55),
    ("Аудит ІТ",         "Гамма ФОП",    48000, "Закрито",    "Д.Петренко", "01.04.26", 100),
    ("Підтримка 2026",   "Дельта ЛТД",   96000, "Контакт",    "О.Мельник",  "30.09.26", 30),
    ("Модуль HR",        "Омега LLC",    210000, "Новий",      "І.Коваль",   "31.08.26", 20),
    ("Інтеграція API",   "Зета Груп",     72000, "Переговори", "Д.Петренко", "30.05.26", 65),
    ("Навчання команди", "Каппа Inc",     36000, "Закрито",    "О.Мельник",  "10.03.26", 100),
    ("Хмарна міграція",  "Сигма Corp",   360000, "Пропозиція", "І.Коваль",   "31.07.26", 45),
]
SEED_TASKS = [
    ("Зателефонувати Петренку",    "Висока",  "Відкрита",  "03.04.26", "О.Мельник",  "CRM"),
    ("Підготувати KP для Бета",    "Висока",  "В роботі",  "05.04.26", "І.Коваль",   "CRM"),
    ("Оновити реєстр ризиків",     "Середній","Відкрита",  "07.04.26", "Д.Петренко", "GRC"),
    ("Звірка залишків складу",     "Низька",  "Відкрита",  "10.04.26", "О.Мельник",  "ERP"),
    ("Презентація для Дельта ЛТД", "Висока",  "В роботі",  "04.04.26", "І.Коваль",   "CRM"),
    ("Сплатити рахунок #0042",     "Середній","Виконана",  "01.04.26", "Д.Петренко", "ERP"),
    ("Аудит доступів Q2",          "Висока",  "Відкрита",  "15.04.26", "О.Мельник",  "GRC"),
    ("Оновити договір Омега",      "Середній","Відкрита",  "20.04.26", "І.Коваль",   "CRM"),
]
SEED_INVOICES = [
    ("INV-0001","Альфа Корп",  48000,  9600,  "Оплачено",  "01.03.26"),
    ("INV-0002","Бета ТОВ",    12000,  2400,  "Оплачено",  "15.03.26"),
    ("INV-0003","Гамма ФОП",    6000,  1200,  "Очікує",    "30.04.26"),
    ("INV-0004","Дельта ЛТД",  96000, 19200,  "Прострочено","01.03.26"),
    ("INV-0005","Омега LLC",   210000, 42000,  "Очікує",    "15.05.26"),
    ("INV-0006","Зета Груп",    36000,  7200,  "Оплачено",  "20.03.26"),
    ("INV-0007","Каппа Inc",    18000,  3600,  "Оплачено",  "25.03.26"),
    ("INV-0008","Сигма Corp",  120000, 24000,  "Очікує",    "30.04.26"),
]
SEED_RISKS = [
    ("Витік персональних даних",   "ІТ",           "Критичний","Можливо",       "Критичний","О.Мельник","Відкритий"),
    ("Збій ERP під час закриття",  "Операційний",  "Високий",  "Малоймовірно",  "Середній", "І.Коваль", "Мітигований"),
    ("Штраф регулятора GDPR",      "Комплаєнс",    "Критичний","Можливо",       "Критичний","Д.Петренко","Відкритий"),
    ("Плинність ключових кадрів",  "HR",           "Середній", "Ймовірно",      "Середній", "О.Мельник","Відкритий"),
    ("Курсові ризики USD/UAH",     "Фінансовий",   "Середній", "Майже напевно", "Високий",  "І.Коваль", "Відкритий"),
    ("Атака ransomware",           "ІТ",           "Критичний","Малоймовірно",  "Середній", "Д.Петренко","Мітигований"),
    ("Зрив постачання Q3",         "Операційний",  "Середній", "Можливо",       "Середній", "О.Мельник","Відкритий"),
    ("Судовий позов від Дельта",   "Юридичний",    "Високий",  "Малоймовірно",  "Середній", "І.Коваль", "Відкритий"),
]


class DB:
    _conn: sqlite3.Connection | None = None

    @classmethod
    def connect(cls):
        cls._conn = sqlite3.connect(str(DB_PATH), check_same_thread=False)
        cls._conn.row_factory = sqlite3.Row
        cls._conn.executescript(SCHEMA)
        cls._conn.commit()
        cls._seed()

    @classmethod
    def _seed(cls):
        if cls.one("SELECT COUNT(*) as c FROM contacts")["c"] == 0:
            cls._conn.executemany(
                "INSERT INTO contacts(name,company,email,phone,type,status) VALUES(?,?,?,?,?,?)",
                SEED_CONTACTS)
        if cls.one("SELECT COUNT(*) as c FROM deals")["c"] == 0:
            cls._conn.executemany(
                "INSERT INTO deals(title,client,amount,stage,manager,close_dt,prob) VALUES(?,?,?,?,?,?,?)",
                SEED_DEALS)
        if cls.one("SELECT COUNT(*) as c FROM tasks")["c"] == 0:
            cls._conn.executemany(
                "INSERT INTO tasks(title,priority,status,due,owner,module) VALUES(?,?,?,?,?,?)",
                SEED_TASKS)
        if cls.one("SELECT COUNT(*) as c FROM invoices")["c"] == 0:
            cls._conn.executemany(
                "INSERT INTO invoices(number,client,amount,vat,status,due) VALUES(?,?,?,?,?,?)",
                SEED_INVOICES)
        if cls.one("SELECT COUNT(*) as c FROM risks")["c"] == 0:
            cls._conn.executemany(
                "INSERT INTO risks(title,category,impact,prob,level,owner,status) VALUES(?,?,?,?,?,?,?)",
                SEED_RISKS)
        cls._conn.commit()

    @classmethod
    def all(cls, sql: str, params=()) -> list:
        return cls._conn.execute(sql, params).fetchall()

    @classmethod
    def one(cls, sql: str, params=()) -> sqlite3.Row | None:
        return cls._conn.execute(sql, params).fetchone()

    @classmethod
    def run(cls, sql: str, params=()) -> int:
        cur = cls._conn.execute(sql, params)
        cls._conn.commit()
        return cur.lastrowid


# ── ТЕМИ ───────────────────────────────────────────────────────
DARK = {
    "bg0":"#07090F","bg1":"#0D1117","bg2":"#111927",
    "bg3":"#172032","bg4":"#1D2A40","bgh":"#243350",
    "b0":"#1A2335","b1":"#243045","b2":"#2E3D55",
    "t0":"#EEF4FF","t1":"#8899BB","t2":"#4A5F80","t3":"#2A3850",
    "a":"#00BFFF","ad":"#003D55","ag":"#00BFFF28",
    "a2":"#7C6EFF","a2d":"#1A1640",
    "g":"#00CC88","gd":"#00261A",
    "r":"#FF4466","rd":"#2E0A14",
    "w":"#FFAA00","wd":"#2A1A00",
    "b":"#3388FF","bd":"#091830",
    "sep":"#1A2335","inv":"#07090F",
    "sb":"#0D1117","sbth":"#1A2335",
    "inp":"#0D1117","inp_b":"#243045","inp_bf":"#00BFFF",
    "btn_p":"#00BFFF","btn_pt":"#07090F",
    "btn_g":"#172032","btn_gt":"#8899BB","btn_gb":"#243045",
    "btn_d":"#2E0A14","btn_dt":"#FF4466","btn_db":"#FF4466",
    "modal":"#111927","overlay":"#00000099",
}
LIGHT = {
    "bg0":"#E8EDF5","bg1":"#F4F7FC","bg2":"#FFFFFF",
    "bg3":"#EEF2FA","bg4":"#E4EAF5","bgh":"#DDE4F0",
    "b0":"#D8E0EE","b1":"#C8D4E8","b2":"#B0C0DC",
    "t0":"#0A1428","t1":"#2A3F60","t2":"#6070A0","t3":"#A0B0CC",
    "a":"#0077CC","ad":"#E0F0FF","ag":"#0077CC22",
    "a2":"#5544BB","a2d":"#EEECff",
    "g":"#007755","gd":"#E6F8F2",
    "r":"#CC2244","rd":"#FCEEF2",
    "w":"#AA6600","wd":"#FFF6E6",
    "b":"#1155BB","bd":"#EEF2FF",
    "sep":"#D8E0EE","inv":"#FFFFFF",
    "sb":"#F4F7FC","sbth":"#C8D4E8",
    "inp":"#FFFFFF","inp_b":"#C8D4E8","inp_bf":"#0077CC",
    "btn_p":"#0077CC","btn_pt":"#FFFFFF",
    "btn_g":"#EEF2FA","btn_gt":"#2A3F60","btn_gb":"#C8D4E8",
    "btn_d":"#FCEEF2","btn_dt":"#CC2244","btn_db":"#CC2244",
    "modal":"#FFFFFF","overlay":"#00000044",
}

FONTS = {
    "h1":("Segoe UI",18,"bold"),  "h2":("Segoe UI",14,"bold"),
    "h3":("Segoe UI",11,"bold"),  "h4":("Segoe UI",10,"bold"),
    "body":("Segoe UI",10,"normal"),"sm":("Segoe UI",9,"normal"),
    "xs":("Segoe UI",8,"normal"), "cap":("Segoe UI",8,"bold"),
    "bold":("Segoe UI",10,"bold"),"mono":("Consolas",9,"normal"),
    "num":("Segoe UI",22,"bold"), "numsm":("Segoe UI",15,"bold"),
    "logo":("Segoe UI",13,"bold"),
}
