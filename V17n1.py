"""
╔══════════════════════════════════════════════════════╗
║  NEXUS  Business Platform  v3.0                      ║
║  CRM · ERP · GRC  |  Python 3.10+  |  Tkinter        ║
║  Запуск:  python nexus_app.py                        ║
╚══════════════════════════════════════════════════════╝
"""
import tkinter as tk
from tkinter import ttk, messagebox
import platform, sys
from datetime import datetime

# ── імпорт ядра ────────────────────────────────────────
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
from nexus_core import DB, DARK, LIGHT, FONTS


# ══════════════════════════════════════════════════════
#  ГЛОБАЛЬНИЙ СТАН ТЕМИ
# ══════════════════════════════════════════════════════

class Theme:
    _name = "dark"
    c     = DARK.copy()
    _cbs  = []

    @classmethod
    def toggle(cls):
        cls._name = "light" if cls._name == "dark" else "dark"
        cls.c = (LIGHT if cls._name == "light" else DARK).copy()
        for cb in list(cls._cbs):
            try: cb()
            except Exception: pass

    @classmethod
    def sub(cls, cb):   cls._cbs.append(cb)
    @classmethod
    def unsub(cls, cb):
        if cb in cls._cbs: cls._cbs.remove(cb)

T = Theme
F = FONTS


# ══════════════════════════════════════════════════════
#  TOAST
# ══════════════════════════════════════════════════════

class Toast:
    _root = None
    _q    = []
    _busy = False

    @classmethod
    def init(cls, root): cls._root = root

    @classmethod
    def show(cls, msg, kind="info"):
        cls._q.append((msg, kind))
        if not cls._busy: cls._fire()

    @classmethod
    def _fire(cls):
        if not cls._q or not cls._root:
            cls._busy = False; return
        cls._busy = True
        msg, kind = cls._q.pop(0)
        colors = {
            "ok":  (T.c["g"],  T.c["gd"]),
            "err": (T.c["r"],  T.c["rd"]),
            "warn":(T.c["w"],  T.c["wd"]),
            "info":(T.c["a"],  T.c["ad"]),
        }
        fg, bg = colors.get(kind, colors["info"])
        icons  = {"ok":"✔","err":"✖","warn":"⚠","info":"ℹ"}

        w = tk.Toplevel(cls._root)
        w.overrideredirect(True)
        w.attributes("-topmost", True)
        w.configure(bg=fg)
        fr = tk.Frame(w, bg=bg, padx=12, pady=8)
        fr.pack(fill="both", expand=True, padx=1, pady=1)
        tk.Label(fr, text=icons.get(kind,"ℹ"),
                 bg=bg, fg=fg, font=("Segoe UI",12,"bold")).pack(side="left", padx=(0,8))
        tk.Label(fr, text=msg, bg=bg, fg=T.c["t0"],
                 font=F["body"], wraplength=260).pack(side="left")
        cls._root.update_idletasks()
        sw = cls._root.winfo_screenwidth()
        sh = cls._root.winfo_screenheight()
        w.update_idletasks()
        ww, wh = w.winfo_reqwidth(), w.winfo_reqheight()
        w.geometry(f"{ww}x{wh}+{sw-ww-18}+{sh-wh-54}")
        w.after(3000, lambda: cls._dismiss(w))

    @classmethod
    def _dismiss(cls, w):
        try: w.destroy()
        except: pass
        cls._root.after(120, cls._fire)


# ══════════════════════════════════════════════════════
#  БАЗОВІ ВІДЖЕТИ  (без Canvas, без after=, без side= в Frame)
# ══════════════════════════════════════════════════════

def sep_h(parent):
    f = tk.Frame(parent, bg=T.c["sep"], height=1)
    f.pack(fill="x")
    return f

def sep_v(parent):
    f = tk.Frame(parent, bg=T.c["sep"], width=1)
    f.pack(side="left", fill="y")
    return f

def lbl(parent, text, style="body", color=None, bg=None, **kw):
    return tk.Label(parent, text=text,
                    font=F.get(style, F["body"]),
                    fg=color or T.c["t1"],
                    bg=bg or parent.cget("bg"), **kw)

def card_frame(parent, padx=16, pady=14):
    outer = tk.Frame(parent, bg=T.c["b0"])
    inner = tk.Frame(outer,  bg=T.c["bg2"], padx=padx, pady=pady)
    inner.pack(fill="both", expand=True, padx=1, pady=1)
    return outer, inner


# ══════════════════════════════════════════════════════
#  КНОПКА
# ══════════════════════════════════════════════════════

class Btn(tk.Frame):
    def __init__(self, parent, text, style="ghost",
                 icon="", command=None, **kw):
        super().__init__(parent, bg=parent.cget("bg"))
        styles = {
            "primary": ("btn_p", "btn_pt", "btn_p"),
            "ghost":   ("btn_g", "btn_gt", "b2"),
            "danger":  ("btn_d", "btn_dt", "btn_db"),
            "outline": ("bg2",   "t1",     "b1"),
        }
        bk, fk, brk = styles.get(style, styles["ghost"])
        self._bg  = T.c[bk]
        self._fg  = T.c[fk]
        self._br  = T.c[brk]
        self._cmd = command
        ltext = (icon + "  " + text).strip() if icon else text

        self._inner = tk.Frame(self,
            bg=self._bg,
            highlightbackground=self._br,
            highlightthickness=1)
        self._inner.pack(fill="both", expand=True)

        self._lbl = tk.Label(self._inner,
            text=ltext,
            bg=self._bg, fg=self._fg,
            font=F["bold"] if style == "primary" else F["body"],
            padx=12, pady=4, cursor="hand2")
        self._lbl.pack()

        for w in (self._inner, self._lbl):
            w.bind("<Button-1>",          self._click)
            w.bind("<Enter>",             self._enter)
            w.bind("<Leave>",             self._leave)

    def _click(self, _=None):
        if self._cmd: self._cmd()

    def _enter(self, _=None):
        self._inner.configure(bg=T.c["bgh"])
        self._lbl.configure(bg=T.c["bgh"])
        self._inner.configure(highlightbackground=T.c["a"])

    def _leave(self, _=None):
        self._inner.configure(bg=self._bg)
        self._lbl.configure(bg=self._bg)
        self._inner.configure(highlightbackground=self._br)


# ══════════════════════════════════════════════════════
#  ПОЛЕ ВВОДУ
# ══════════════════════════════════════════════════════

class Field(tk.Frame):
    def __init__(self, parent, label="", ph="", required=False, **kw):
        super().__init__(parent, bg=parent.cget("bg"))
        if label:
            hdr = tk.Frame(self, bg=self.cget("bg"))
            hdr.pack(fill="x", pady=(0,3))
            tk.Label(hdr, text=label, bg=self.cget("bg"),
                     fg=T.c["t2"], font=F["cap"]).pack(side="left")
            if required:
                tk.Label(hdr, text=" *", bg=self.cget("bg"),
                         fg=T.c["r"], font=F["cap"]).pack(side="left")

        self._border = tk.Frame(self, bg=T.c["inp_b"])
        self._border.pack(fill="x")
        self._inner  = tk.Frame(self._border, bg=T.c["inp"])
        self._inner.pack(fill="x", padx=1, pady=1)

        self._var = tk.StringVar()
        self._ph  = ph
        self._ph_on = False

        self._e = tk.Entry(self._inner,
            textvariable=self._var,
            bg=T.c["inp"], fg=T.c["t0"],
            insertbackground=T.c["a"],
            relief="flat", bd=5, font=F["body"], **kw)
        self._e.pack(fill="x")

        if ph:
            self._e.insert(0, ph)
            self._e.configure(fg=T.c["t2"])
            self._ph_on = True

        self._e.bind("<FocusIn>",  self._fin)
        self._e.bind("<FocusOut>", self._fout)

    def _fin(self, _=None):
        self._border.configure(bg=T.c["inp_bf"])
        if self._ph_on:
            self._e.delete(0, "end")
            self._e.configure(fg=T.c["t0"])
            self._ph_on = False

    def _fout(self, _=None):
        self._border.configure(bg=T.c["inp_b"])
        if not self._e.get() and self._ph:
            self._e.insert(0, self._ph)
            self._e.configure(fg=T.c["t2"])
            self._ph_on = True

    def get(self):
        return "" if self._ph_on else self._var.get()

    def set(self, v):
        self._ph_on = False
        self._e.delete(0, "end")
        self._e.insert(0, str(v))
        self._e.configure(fg=T.c["t0"])

    def focus(self): self._e.focus_set()

    def bind_return(self, fn):
        self._e.bind("<Return>", lambda _: fn())


# ══════════════════════════════════════════════════════
#  КОМБОБОКС
# ══════════════════════════════════════════════════════

class Pick(tk.Frame):
    def __init__(self, parent, label="", options=None, **kw):
        super().__init__(parent, bg=parent.cget("bg"))
        self._opts = options or []
        if label:
            tk.Label(self, text=label, bg=self.cget("bg"),
                     fg=T.c["t2"], font=F["cap"]).pack(anchor="w", pady=(0,3))
        s = ttk.Style()
        s.theme_use("default")
        s.configure("N.TCombobox",
            fieldbackground=T.c["inp"],
            background=T.c["bg3"],
            foreground=T.c["t0"],
            arrowcolor=T.c["t1"],
            borderwidth=0, relief="flat", padding=(6,4))
        s.map("N.TCombobox",
            fieldbackground=[("readonly", T.c["inp"])],
            selectbackground=[("readonly", T.c["ad"])],
            selectforeground=[("readonly", T.c["t0"])])
        self._var = tk.StringVar()
        self._cb  = ttk.Combobox(self,
            textvariable=self._var,
            values=self._opts,
            state="readonly",
            font=F["body"],
            style="N.TCombobox", **kw)
        self._cb.pack(fill="x")
        if self._opts: self._cb.current(0)

    def get(self): return self._var.get()
    def set(self, v): self._var.set(v)


# ══════════════════════════════════════════════════════
#  ТАБЛИЦЯ
# ══════════════════════════════════════════════════════

class Table(tk.Frame):
    def __init__(self, parent, cols, rows=None,
                 on_dbl=None, ctx=None):
        """
        cols = [{"id":"x","label":"X","w":100,"anchor":"w"}, ...]
        ctx  = [{"label":"Дія","cmd": fn(values)}, "|", ...]
        """
        super().__init__(parent, bg=T.c["bg2"])
        self._cols  = cols
        self._dbl   = on_dbl
        self._ctx   = ctx or []
        self._sort_col = None
        self._sort_asc = True
        self._apply_style()
        self._build()
        if rows: self.load(rows)

    def _apply_style(self):
        s = ttk.Style()
        s.theme_use("default")
        s.configure("N.Treeview",
            background=T.c["bg2"],
            foreground=T.c["t1"],
            fieldbackground=T.c["bg2"],
            rowheight=30, font=F["body"],
            borderwidth=0, relief="flat")
        s.configure("N.Treeview.Heading",
            background=T.c["bg3"],
            foreground=T.c["t2"],
            font=F["cap"], relief="flat",
            padding=(8,5))
        s.map("N.Treeview",
            background=[("selected", T.c["ad"])],
            foreground=[("selected", T.c["a"])])
        s.map("N.Treeview.Heading",
            background=[("active", T.c["bg4"])])
        # scrollbar
        s.configure("N.Vertical.TScrollbar",
            background=T.c["sb"],
            troughcolor=T.c["sb"],
            arrowcolor=T.c["t2"],
            borderwidth=0, relief="flat")
        s.map("N.Vertical.TScrollbar",
            background=[("active", T.c["sbth"])])

    def _build(self):
        ids = [c["id"] for c in self._cols]
        vsb = ttk.Scrollbar(self, orient="vertical",
                            style="N.Vertical.TScrollbar")
        self._tv = ttk.Treeview(self, columns=ids,
            show="headings", style="N.Treeview",
            yscrollcommand=vsb.set, selectmode="extended")
        vsb.configure(command=self._tv.yview)
        vsb.pack(side="right", fill="y")
        self._tv.pack(fill="both", expand=True)

        for c in self._cols:
            self._tv.heading(c["id"],
                text=c.get("label", c["id"]),
                anchor=c.get("anchor","w"),
                command=lambda cid=c["id"]: self._sort(cid))
            self._tv.column(c["id"],
                width=c.get("w",120),
                minwidth=c.get("mw",50),
                anchor=c.get("anchor","w"),
                stretch=c.get("str",True))

        self._cnt = tk.Label(self, text="",
            bg=T.c["bg3"], fg=T.c["t2"], font=F["xs"], padx=10, pady=3)
        self._cnt.pack(fill="x")

        if self._dbl:
            self._tv.bind("<Double-1>",
                lambda _: self._dbl(self._tv.item(
                    self._tv.focus(), "values")))
        if self._ctx:
            self._tv.bind("<Button-3>", self._rmb)
        self._tv.bind("<<TreeviewSelect>>", lambda _: self._upd_cnt())

    def load(self, rows):
        for i in self._tv.get_children(): self._tv.delete(i)
        for n, row in enumerate(rows):
            tag = "alt" if n % 2 else "norm"
            self._tv.insert("", "end", values=row, tags=(tag,))
        self._tv.tag_configure("norm", background=T.c["bg2"])
        self._tv.tag_configure("alt",  background=T.c["bg1"])
        self._upd_cnt()

    def _upd_cnt(self):
        tot = len(self._tv.get_children())
        sel = len(self._tv.selection())
        self._cnt.configure(
            text=f"Вибрано: {sel}  /  Всього: {tot}" if sel
                 else f"Всього: {tot} записів")

    def _sort(self, cid):
        if self._sort_col == cid:
            self._sort_asc = not self._sort_asc
        else:
            self._sort_col = cid
            self._sort_asc = True
        items = [(self._tv.set(i, cid), i)
                 for i in self._tv.get_children()]
        items.sort(key=lambda x: x[0].lower()
                   if isinstance(x[0], str) else x[0],
                   reverse=not self._sort_asc)
        for n, (_, iid) in enumerate(items):
            self._tv.move(iid, "", n)
            self._tv.item(iid, tags=("alt" if n%2 else "norm",))
        arr = " ▲" if self._sort_asc else " ▼"
        for c in self._cols:
            suffix = arr if c["id"] == cid else ""
            self._tv.heading(c["id"],
                text=c.get("label", c["id"]) + suffix)

    def _rmb(self, e):
        iid = self._tv.identify_row(e.y)
        if not iid: return
        if iid not in self._tv.selection():
            self._tv.selection_set(iid)
        vals = self._tv.item(iid, "values")
        m = tk.Menu(self, tearoff=0,
            bg=T.c["bg3"], fg=T.c["t0"],
            activebackground=T.c["bg4"],
            activeforeground=T.c["a"],
            relief="flat", bd=0, font=F["body"])
        for item in self._ctx:
            if item == "|":
                m.add_separator()
            else:
                m.add_command(label=item["label"],
                    command=lambda fn=item["cmd"],v=vals: fn(v))
        m.tk_popup(e.x_root, e.y_root)

    def selection(self):
        return [self._tv.item(i,"values")
                for i in self._tv.selection()]


# ══════════════════════════════════════════════════════
#  ДІАЛОГ (базовий)
# ══════════════════════════════════════════════════════

class Dlg(tk.Toplevel):
    def __init__(self, parent, title, w=500, h=380):
        super().__init__(parent)
        self.title(title)
        self.resizable(False, False)
        self.transient(parent)
        self.configure(bg=T.c["modal"])
        self.grab_set()
        # центрування
        self.update_idletasks()
        px = parent.winfo_rootx() + (parent.winfo_width()  - w) // 2
        py = parent.winfo_rooty() + (parent.winfo_height() - h) // 2
        self.geometry(f"{w}x{h}+{px}+{py}")
        self.bind("<Escape>", lambda _: self.destroy())
        self._chrome(title)

    def _chrome(self, title):
        tk.Frame(self, bg=T.c["a"], height=2).pack(fill="x")
        hdr = tk.Frame(self, bg=T.c["modal"], padx=18, pady=12)
        hdr.pack(fill="x")
        tk.Label(hdr, text=title, bg=T.c["modal"],
                 fg=T.c["t0"], font=F["h3"]).pack(side="left")
        tk.Button(hdr, text="✕",
                  bg=T.c["modal"], fg=T.c["t2"],
                  relief="flat", bd=0, cursor="hand2",
                  font=("Segoe UI",11),
                  activebackground=T.c["bg3"],
                  activeforeground=T.c["t0"],
                  command=self.destroy).pack(side="right")
        tk.Frame(self, bg=T.c["sep"], height=1).pack(fill="x")
        self._body = tk.Frame(self, bg=T.c["modal"], padx=18, pady=14)
        self._body.pack(fill="both", expand=True)
        self.body(self._body)
        tk.Frame(self, bg=T.c["sep"], height=1).pack(fill="x")
        foot = tk.Frame(self, bg=T.c["modal"], padx=18, pady=10)
        foot.pack(fill="x")
        Btn(foot, "Підтвердити", "primary", command=self._ok).pack(side="right", padx=(6,0))
        Btn(foot, "Скасувати",   "outline", command=self.destroy).pack(side="right")

    def body(self, m): pass

    def _ok(self):
        if self.validate():
            self.save()
            self.destroy()

    def validate(self): return True
    def save(self): pass


# ══════════════════════════════════════════════════════
#  ФОРМА КОНТАКТУ
# ══════════════════════════════════════════════════════

class ContactDlg(Dlg):
    def __init__(self, parent, row=None):
        self._row = row
        super().__init__(parent,
            "Редагувати контакт" if row else "Новий контакт",
            520, 400)

    def body(self, m):
        r1 = tk.Frame(m, bg=m.cget("bg"))
        r1.pack(fill="x", pady=(0,10))
        self._name = Field(r1, "Імʼя та прізвище *", required=True)
        self._name.pack(fill="x")

        r2 = tk.Frame(m, bg=m.cget("bg"))
        r2.pack(fill="x", pady=(0,10))
        self._comp = Field(r2, "Компанія")
        self._comp.pack(side="left", fill="x", expand=True, padx=(0,8))
        self._type = Pick(r2, "Тип",
            ["Клієнт","Партнер","Лід","VIP"])
        self._type.pack(side="left", fill="x", expand=True)

        r3 = tk.Frame(m, bg=m.cget("bg"))
        r3.pack(fill="x", pady=(0,10))
        self._email = Field(r3, "Email")
        self._email.pack(side="left", fill="x", expand=True, padx=(0,8))
        self._phone = Field(r3, "Телефон")
        self._phone.pack(side="left", fill="x", expand=True)

        self._status = Pick(m, "Статус",
            ["Активний","Неактивний","Новий"])
        self._status.pack(fill="x")

        if self._row:
            self._name.set(self._row[1])
            self._comp.set(self._row[2])
            self._email.set(self._row[3])
            self._phone.set(self._row[4])
            self._type.set(self._row[5])
            self._status.set(self._row[6])

    def validate(self):
        if not self._name.get():
            Toast.show("Введіть імʼя контакту","warn")
            return False
        return True

    def save(self):
        if self._row:
            DB.run("""UPDATE contacts SET name=?,company=?,email=?,
                      phone=?,type=?,status=? WHERE id=?""",
                   (self._name.get(), self._comp.get(),
                    self._email.get(), self._phone.get(),
                    self._type.get(), self._status.get(),
                    self._row[0]))
            Toast.show("Контакт оновлено","ok")
        else:
            DB.run("""INSERT INTO contacts
                      (name,company,email,phone,type,status)
                      VALUES(?,?,?,?,?,?)""",
                   (self._name.get(), self._comp.get(),
                    self._email.get(), self._phone.get(),
                    self._type.get(), self._status.get()))
            Toast.show("Контакт створено","ok")


# ══════════════════════════════════════════════════════
#  ФОРМА УГОДИ
# ══════════════════════════════════════════════════════

class DealDlg(Dlg):
    def __init__(self, parent, row=None):
        self._row = row
        super().__init__(parent,
            "Редагувати угоду" if row else "Нова угода",
            520, 420)

    def body(self, m):
        self._title = Field(m, "Назва угоди *", required=True)
        self._title.pack(fill="x", pady=(0,10))

        r2 = tk.Frame(m, bg=m.cget("bg"))
        r2.pack(fill="x", pady=(0,10))
        self._client = Field(r2, "Клієнт")
        self._client.pack(side="left", fill="x", expand=True, padx=(0,8))
        self._amount = Field(r2, "Сума (₴)")
        self._amount.pack(side="left", fill="x", expand=True)

        r3 = tk.Frame(m, bg=m.cget("bg"))
        r3.pack(fill="x", pady=(0,10))
        self._stage = Pick(r3, "Стадія",
            ["Новий","Контакт","Пропозиція","Переговори","Закрито"])
        self._stage.pack(side="left", fill="x", expand=True, padx=(0,8))
        self._mgr = Field(r3, "Менеджер")
        self._mgr.pack(side="left", fill="x", expand=True)

        r4 = tk.Frame(m, bg=m.cget("bg"))
        r4.pack(fill="x")
        self._close = Field(r4, "Дата закриття")
        self._close.pack(side="left", fill="x", expand=True, padx=(0,8))
        self._prob = Field(r4, "Ймовірність %")
        self._prob.pack(side="left", fill="x", expand=True)

        if self._row:
            self._title.set(self._row[1])
            self._client.set(self._row[2])
            self._amount.set(str(self._row[3]))
            self._stage.set(self._row[4])
            self._mgr.set(self._row[5])
            self._close.set(self._row[6])
            self._prob.set(str(self._row[7]))

    def validate(self):
        if not self._title.get():
            Toast.show("Введіть назву угоди","warn"); return False
        return True

    def save(self):
        try: amt = float(self._amount.get() or 0)
        except: amt = 0
        try: prob = int(self._prob.get() or 50)
        except: prob = 50

        if self._row:
            DB.run("""UPDATE deals SET title=?,client=?,amount=?,
                      stage=?,manager=?,close_dt=?,prob=? WHERE id=?""",
                   (self._title.get(), self._client.get(), amt,
                    self._stage.get(), self._mgr.get(),
                    self._close.get(), prob, self._row[0]))
            Toast.show("Угоду оновлено","ok")
        else:
            DB.run("""INSERT INTO deals
                      (title,client,amount,stage,manager,close_dt,prob)
                      VALUES(?,?,?,?,?,?,?)""",
                   (self._title.get(), self._client.get(), amt,
                    self._stage.get(), self._mgr.get(),
                    self._close.get(), prob))
            Toast.show("Угоду створено","ok")


# ══════════════════════════════════════════════════════
#  ФОРМА ЗАВДАННЯ
# ══════════════════════════════════════════════════════

class TaskDlg(Dlg):
    def __init__(self, parent, row=None):
        self._row = row
        super().__init__(parent,
            "Редагувати завдання" if row else "Нове завдання",
            500, 360)

    def body(self, m):
        self._title = Field(m, "Завдання *", required=True)
        self._title.pack(fill="x", pady=(0,10))
        r2 = tk.Frame(m, bg=m.cget("bg"))
        r2.pack(fill="x", pady=(0,10))
        self._prio = Pick(r2, "Пріоритет",
            ["Низька","Середній","Висока","Критична"])
        self._prio.pack(side="left", fill="x", expand=True, padx=(0,8))
        self._stat = Pick(r2, "Статус",
            ["Відкрита","В роботі","Виконана","Відкладена"])
        self._stat.pack(side="left", fill="x", expand=True)
        r3 = tk.Frame(m, bg=m.cget("bg"))
        r3.pack(fill="x")
        self._due   = Field(r3, "Термін")
        self._due.pack(side="left", fill="x", expand=True, padx=(0,8))
        self._owner = Field(r3, "Відповідальний")
        self._owner.pack(side="left", fill="x", expand=True)
        if self._row:
            self._title.set(self._row[1])
            self._prio.set(self._row[2])
            self._stat.set(self._row[3])
            self._due.set(self._row[4])
            self._owner.set(self._row[5])

    def validate(self):
        if not self._title.get():
            Toast.show("Введіть назву завдання","warn"); return False
        return True

    def save(self):
        if self._row:
            DB.run("""UPDATE tasks SET title=?,priority=?,status=?,
                      due=?,owner=? WHERE id=?""",
                   (self._title.get(), self._prio.get(),
                    self._stat.get(), self._due.get(),
                    self._owner.get(), self._row[0]))
            Toast.show("Завдання оновлено","ok")
        else:
            DB.run("""INSERT INTO tasks(title,priority,status,due,owner)
                      VALUES(?,?,?,?,?)""",
                   (self._title.get(), self._prio.get(),
                    self._stat.get(), self._due.get(),
                    self._owner.get()))
            Toast.show("Завдання створено","ok")


# ══════════════════════════════════════════════════════
#  ФОРМА РИЗИКУ
# ══════════════════════════════════════════════════════

class RiskDlg(Dlg):
    def __init__(self, parent, row=None):
        self._row = row
        super().__init__(parent,
            "Редагувати ризик" if row else "Новий ризик",
            520, 420)

    def body(self, m):
        self._title = Field(m, "Назва ризику *", required=True)
        self._title.pack(fill="x", pady=(0,10))
        r2 = tk.Frame(m, bg=m.cget("bg"))
        r2.pack(fill="x", pady=(0,10))
        self._cat = Pick(r2, "Категорія",
            ["Операційний","Фінансовий","ІТ","Комплаєнс","HR","Юридичний","Репутаційний"])
        self._cat.pack(side="left", fill="x", expand=True, padx=(0,8))
        self._lvl = Pick(r2, "Рівень",
            ["Низький","Середній","Високий","Критичний"])
        self._lvl.pack(side="left", fill="x", expand=True)
        r3 = tk.Frame(m, bg=m.cget("bg"))
        r3.pack(fill="x", pady=(0,10))
        self._imp = Pick(r3, "Вплив",
            ["Незначний","Малий","Середній","Великий","Критичний"])
        self._imp.pack(side="left", fill="x", expand=True, padx=(0,8))
        self._prb = Pick(r3, "Ймовірність",
            ["Майже неможливо","Малоймовірно","Можливо","Ймовірно","Майже напевно"])
        self._prb.pack(side="left", fill="x", expand=True)
        r4 = tk.Frame(m, bg=m.cget("bg"))
        r4.pack(fill="x")
        self._own = Field(r4, "Власник")
        self._own.pack(side="left", fill="x", expand=True, padx=(0,8))
        self._stat = Pick(r4, "Статус",
            ["Відкритий","Мітигований","Прийнятий","Закритий"])
        self._stat.pack(side="left", fill="x", expand=True)
        if self._row:
            self._title.set(self._row[1])
            self._cat.set(self._row[2])
            self._imp.set(self._row[3])
            self._prb.set(self._row[4])
            self._lvl.set(self._row[5])
            self._own.set(self._row[6])
            self._stat.set(self._row[7])

    def validate(self):
        if not self._title.get():
            Toast.show("Введіть назву ризику","warn"); return False
        return True

    def save(self):
        if self._row:
            DB.run("""UPDATE risks SET title=?,category=?,impact=?,
                      prob=?,level=?,owner=?,status=? WHERE id=?""",
                   (self._title.get(), self._cat.get(),
                    self._imp.get(), self._prb.get(),
                    self._lvl.get(), self._own.get(),
                    self._stat.get(), self._row[0]))
            Toast.show("Ризик оновлено","ok")
        else:
            DB.run("""INSERT INTO risks
                      (title,category,impact,prob,level,owner,status)
                      VALUES(?,?,?,?,?,?,?)""",
                   (self._title.get(), self._cat.get(),
                    self._imp.get(), self._prb.get(),
                    self._lvl.get(), self._own.get(),
                    self._stat.get()))
            Toast.show("Ризик створено","ok")


# ══════════════════════════════════════════════════════
#  МОДУЛЬ — BASE
# ══════════════════════════════════════════════════════

class Module(tk.Frame):
    ID    = "base"
    TITLE = "Модуль"
    ICON  = "◈"
    GROUP = ""

    def __init__(self, parent):
        super().__init__(parent, bg=T.c["bg1"])
        self._ready = False

    def activate(self):
        if not self._ready:
            self.build()
            self._ready = True
        self.refresh()

    def build(self):
        lbl(self, "Порожній модуль", "h3",
            T.c["t2"]).pack(expand=True)

    def refresh(self): pass

    # ── helpers ─────────────────────────────────────
    def page_header(self, title, sub="", actions=None):
        hdr = tk.Frame(self, bg=T.c["bg1"])
        hdr.pack(fill="x")
        inner = tk.Frame(hdr, bg=T.c["bg1"], padx=22, pady=14)
        inner.pack(fill="x")
        left = tk.Frame(inner, bg=T.c["bg1"])
        left.pack(side="left", fill="x", expand=True)
        if sub:
            lbl(left, sub, "xs", T.c["t3"]).pack(anchor="w")
        lbl(left, title, "h2", T.c["t0"]).pack(anchor="w")
        if actions:
            right = tk.Frame(inner, bg=T.c["bg1"])
            right.pack(side="right")
            for a in reversed(actions):
                Btn(right, a["label"], a.get("style","ghost"),
                    icon=a.get("icon",""),
                    command=a.get("cmd")
                    ).pack(side="right", padx=(4,0), ipady=2)
        tk.Frame(self, bg=T.c["sep"], height=1).pack(fill="x")

    def scrollable(self):
        cnv = tk.Canvas(self, bg=T.c["bg1"],
                        highlightthickness=0, bd=0)
        vsb = ttk.Scrollbar(self, orient="vertical",
                            style="N.Vertical.TScrollbar",
                            command=cnv.yview)
        cnv.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        cnv.pack(fill="both", expand=True)
        inner = tk.Frame(cnv, bg=T.c["bg1"])
        win   = cnv.create_window((0,0), window=inner, anchor="nw")

        def _cfg(e):
            cnv.configure(scrollregion=cnv.bbox("all"))
            cnv.itemconfigure(win, width=cnv.winfo_width())
        inner.bind("<Configure>", _cfg)
        cnv.bind("<MouseWheel>",
            lambda e: cnv.yview_scroll(int(-e.delta/60),"units"))
        return inner

    def kpi_card(self, parent, title, value, delta="",
                 pos=True, accent=None):
        acc = accent or T.c["a"]
        f = tk.Frame(parent, bg=T.c["bg2"],
                     highlightthickness=1,
                     highlightbackground=T.c["b0"])
        tk.Frame(f, bg=acc, height=2).pack(fill="x")
        body = tk.Frame(f, bg=T.c["bg2"], padx=14, pady=12)
        body.pack(fill="both", expand=True)
        lbl(body, title, "cap", T.c["t2"]).pack(anchor="w")
        lbl(body, value, "num", T.c["t0"]).pack(anchor="w", pady=(6,0))
        if delta:
            dcol = T.c["g"] if pos else T.c["r"]
            darr = "▲" if pos else "▼"
            lbl(body, f"{darr} {delta}", "sm", dcol).pack(anchor="w", pady=(2,0))
        return f


# ══════════════════════════════════════════════════════
#  DASHBOARD
# ══════════════════════════════════════════════════════

class Dashboard(Module):
    ID = "dashboard"; TITLE = "Dashboard"; ICON = "⊞"

    def build(self):
        self.page_header("Панель керування",
            "NEXUS › Dashboard",
            actions=[
                {"label":"Оновити","icon":"↻","style":"ghost",
                 "cmd": self.refresh},
                {"label":"Новий запис","icon":"+","style":"primary",
                 "cmd": lambda: TaskDlg(self.winfo_toplevel())},
            ])
        self._sc = self.scrollable()
        self._build_kpi()
        self._build_pipeline()
        self._build_tables()

    def _build_kpi(self):
        f = tk.Frame(self._sc, bg=T.c["bg1"], padx=22, pady=16)
        f.pack(fill="x")
        lbl(f, "КЛЮЧОВІ ПОКАЗНИКИ", "cap", T.c["t3"]).pack(anchor="w", pady=(0,10))
        row = tk.Frame(f, bg=T.c["bg1"])
        row.pack(fill="x")
        row.columnconfigure((0,1,2,3), weight=1, uniform="k")
        kpis = [
            ("КЛІЄНТИ",        None,         "+18%",  True,  T.c["a"]),
            ("УГОДИ (всього)", None,         "+11%",  True,  T.c["g"]),
            ("ВІДКР.ЗАДАЧІ",   None,         "",      True,  T.c["w"]),
            ("РИЗИКИ",         None,         "+3 нові",False, T.c["r"]),
        ]
        self._kpi_cards = []
        for col,(title,_,delta,pos,acc) in enumerate(kpis):
            c = self.kpi_card(row, title, "…", delta, pos, acc)
            c.grid(row=0, column=col,
                   padx=(0, 10 if col < 3 else 0), sticky="nsew")
            self._kpi_cards.append(c)

    def _build_pipeline(self):
        f = tk.Frame(self._sc, bg=T.c["bg1"], padx=22)
        f.pack(fill="x", pady=(0,16))
        lbl(f, "ВОРОНКА ПРОДАЖІВ", "cap", T.c["t3"]).pack(anchor="w", pady=(0,8))
        self._pipe_frame = tk.Frame(f, bg=T.c["bg1"])
        self._pipe_frame.pack(fill="x")

    def _build_tables(self):
        f = tk.Frame(self._sc, bg=T.c["bg1"], padx=22, pady=(0,24))
        f.pack(fill="both", expand=True)
        f.columnconfigure(0, weight=3)
        f.columnconfigure(1, weight=2)

        # остання угоди
        left, li = card_frame(f)
        left.grid(row=0, column=0, sticky="nsew", padx=(0,10))
        lbl(li, "Останні угоди", "h4", T.c["t0"]).pack(anchor="w", pady=(0,8))
        cols_d = [
            {"id":"title",  "label":"НАЗВА",   "w":180},
            {"id":"client", "label":"КЛІЄНТ",  "w":120},
            {"id":"amount", "label":"СУМА",    "w":90,"anchor":"e"},
            {"id":"stage",  "label":"СТАДІЯ",  "w":100},
        ]
        self._dash_deals = Table(li, cols_d,
            on_dbl=lambda v: Toast.show(f"Угода: {v[1]}","info"))
        self._dash_deals.pack(fill="both", expand=True)

        # завдання
        right, ri = card_frame(f)
        right.grid(row=0, column=1, sticky="nsew")
        lbl(ri, "Мої завдання", "h4", T.c["t0"]).pack(anchor="w", pady=(0,8))
        cols_t = [
            {"id":"title", "label":"ЗАВДАННЯ", "w":160},
            {"id":"prio",  "label":"ПРІОР.",   "w":80},
            {"id":"due",   "label":"ТЕРМІН",   "w":70},
        ]
        self._dash_tasks = Table(ri, cols_t,
            on_dbl=lambda v: Toast.show(f"Задача: {v[0]}","info"))
        self._dash_tasks.pack(fill="both", expand=True)

    def refresh(self):
        # KPI з БД
        cnt_contacts = DB.one("SELECT COUNT(*) c FROM contacts")["c"]
        cnt_deals    = DB.one("SELECT COUNT(*) c FROM deals")["c"]
        cnt_tasks    = DB.one("SELECT COUNT(*) c FROM tasks WHERE status!='Виконана'")["c"]
        cnt_risks    = DB.one("SELECT COUNT(*) c FROM risks WHERE status='Відкритий'")["c"]

        vals = [str(cnt_contacts), str(cnt_deals),
                str(cnt_tasks),    str(cnt_risks)]
        for card, val in zip(self._kpi_cards, vals):
            # оновлюємо числовий label всередині картки
            body = card.winfo_children()[1]  # body frame (другий після stripe)
            for w in body.winfo_children():
                if hasattr(w,"cget") and w.cget("font") == str(F["num"]):
                    w.configure(text=val)
                    break

        # pipeline
        for w in self._pipe_frame.winfo_children():
            w.destroy()
        stages = ["Новий","Контакт","Пропозиція","Переговори","Закрито"]
        colors = [T.c["a"],T.c["a2"],T.c["g"],T.c["w"],T.c["r"]]
        total  = max(DB.one("SELECT COUNT(*) c FROM deals")["c"], 1)
        for i,(stage,col) in enumerate(zip(stages,colors)):
            cnt = DB.one(
                "SELECT COUNT(*) c FROM deals WHERE stage=?", (stage,))["c"]
            r = tk.Frame(self._pipe_frame, bg=T.c["bg1"])
            r.pack(fill="x", pady=2)
            lbl(r, stage, "sm", T.c["t1"], width=12, anchor="w").pack(side="left")
            # прогрес-бар через Frame
            bar_bg = tk.Frame(r, bg=T.c["bg3"], height=16)
            bar_bg.pack(side="left", fill="x", expand=True, padx=(6,6))
            bar_bg.update_idletasks()
            bw = bar_bg.winfo_width() or 200
            pct = cnt / total
            bar_fill = tk.Frame(bar_bg, bg=col, height=16,
                                width=max(4, int(bw * pct)))
            bar_fill.place(x=0, y=0, relheight=1)
            lbl(r, str(cnt), "sm", T.c["t2"], width=4).pack(side="left")

        # deals table
        rows = DB.all("SELECT title,client,amount,stage FROM deals ORDER BY id DESC LIMIT 8")
        self._dash_deals.load(
            [(r["title"], r["client"],
              f"₴{r['amount']:,.0f}".replace(",","\u202f"),
              r["stage"]) for r in rows])

        # tasks table
        trows = DB.all("""SELECT title,priority,due FROM tasks
                          WHERE status!='Виконана' ORDER BY id LIMIT 8""")
        self._dash_tasks.load(
            [(r["title"], r["priority"], r["due"]) for r in trows])


# ══════════════════════════════════════════════════════
#  CONTACTS
# ══════════════════════════════════════════════════════

class Contacts(Module):
    ID = "contacts"; TITLE = "Контакти"; ICON = "◉"; GROUP = "CRM"

    def build(self):
        self.page_header("Контакти", "NEXUS › CRM › Контакти",
            actions=[
                {"label":"+ Контакт","style":"primary","icon":"+",
                 "cmd": lambda: self._open(None)},
            ])
        tb = tk.Frame(self, bg=T.c["bg1"], padx=22, pady=8)
        tb.pack(fill="x")
        self._sf = Field(tb, ph="⌕  Пошук за іменем…")
        self._sf.pack(side="left")
        self._sf.bind_return(self.refresh)
        Btn(tb, "Знайти", "ghost",
            command=self.refresh).pack(side="left", padx=(6,0), ipady=2)

        wrap = tk.Frame(self, bg=T.c["bg1"], padx=22, pady=(0,16))
        wrap.pack(fill="both", expand=True)
        cols = [
            {"id":"id",     "label":"ID",        "w":44, "str":False},
            {"id":"name",   "label":"КОНТАКТ",   "w":170},
            {"id":"company","label":"КОМПАНІЯ",  "w":140},
            {"id":"email",  "label":"EMAIL",      "w":170},
            {"id":"phone",  "label":"ТЕЛЕФОН",    "w":130},
            {"id":"type",   "label":"ТИП",        "w":80},
            {"id":"status", "label":"СТАТУС",     "w":90},
        ]
        self._tbl = Table(wrap, cols,
            on_dbl=lambda v: self._open(v),
            ctx=[
                {"label":"✏  Редагувати","cmd": self._open},
                {"label":"✉  Написати",
                 "cmd": lambda v: Toast.show(f"Email → {v[3]}","info")},
                "|",
                {"label":"✖  Видалити",  "cmd": self._delete},
            ])
        self._tbl.pack(fill="both", expand=True)

    def refresh(self):
        q = self._sf.get() if hasattr(self,"_sf") else ""
        if q:
            rows = DB.all(
                "SELECT id,name,company,email,phone,type,status "
                "FROM contacts WHERE name LIKE ? ORDER BY name",
                (f"%{q}%",))
        else:
            rows = DB.all(
                "SELECT id,name,company,email,phone,type,status "
                "FROM contacts ORDER BY name")
        self._tbl.load([tuple(r) for r in rows])

    def _open(self, v):
        def _after():
            self.refresh()
        if v:
            row = DB.one("SELECT * FROM contacts WHERE id=?", (v[0],))
            d = ContactDlg(self.winfo_toplevel(), tuple(row))
        else:
            d = ContactDlg(self.winfo_toplevel())
        d.bind("<Destroy>", lambda _: self.after(100, self.refresh))

    def _delete(self, v):
        if messagebox.askyesno("Видалити",
                f"Видалити контакт «{v[1]}»?"):
            DB.run("DELETE FROM contacts WHERE id=?", (v[0],))
            Toast.show("Контакт видалено","ok")
            self.refresh()


# ══════════════════════════════════════════════════════
#  DEALS
# ══════════════════════════════════════════════════════

class Deals(Module):
    ID = "deals"; TITLE = "Угоди"; ICON = "◇"; GROUP = "CRM"

    def build(self):
        self.page_header("Угоди", "NEXUS › CRM › Угоди",
            actions=[
                {"label":"+ Угода","style":"primary","icon":"+",
                 "cmd": lambda: self._open(None)},
            ])
        # stage tabs
        self._stage_var = tk.StringVar(value="Всі")
        tb = tk.Frame(self, bg=T.c["bg1"], padx=22, pady=6)
        tb.pack(fill="x")
        for stage in ["Всі","Новий","Контакт","Пропозиція","Переговори","Закрито"]:
            b = tk.Button(tb, text=stage,
                bg=T.c["bg3"], fg=T.c["t1"],
                relief="flat", bd=0, cursor="hand2",
                font=F["sm"], padx=10, pady=4,
                command=lambda s=stage: (self._stage_var.set(s), self.refresh()))
            b.pack(side="left", padx=(0,4))

        wrap = tk.Frame(self, bg=T.c["bg1"], padx=22, pady=(0,16))
        wrap.pack(fill="both", expand=True)
        cols = [
            {"id":"id",     "label":"ID",         "w":50,"str":False},
            {"id":"title",  "label":"НАЗВА",       "w":200},
            {"id":"client", "label":"КЛІЄНТ",      "w":140},
            {"id":"amount", "label":"СУМА",        "w":100,"anchor":"e"},
            {"id":"stage",  "label":"СТАДІЯ",      "w":110},
            {"id":"manager","label":"МЕНЕДЖЕР",    "w":110},
            {"id":"close",  "label":"ЗАКРИТТЯ",    "w":90},
            {"id":"prob",   "label":"ЙМОВ.%",      "w":70,"anchor":"e"},
        ]
        self._tbl = Table(wrap, cols,
            on_dbl=lambda v: self._open(v),
            ctx=[
                {"label":"✏  Редагувати", "cmd": self._open},
                "|",
                {"label":"✖  Видалити",   "cmd": self._delete},
            ])
        self._tbl.pack(fill="both", expand=True)

    def refresh(self):
        stage = self._stage_var.get() if hasattr(self,"_stage_var") else "Всі"
        if stage == "Всі":
            rows = DB.all(
                "SELECT id,title,client,amount,stage,manager,close_dt,prob "
                "FROM deals ORDER BY id DESC")
        else:
            rows = DB.all(
                "SELECT id,title,client,amount,stage,manager,close_dt,prob "
                "FROM deals WHERE stage=? ORDER BY id DESC", (stage,))
        self._tbl.load([
            (r["id"], r["title"], r["client"],
             f"₴{r['amount']:,.0f}".replace(",","\u202f"),
             r["stage"], r["manager"], r["close_dt"],
             f"{r['prob']}%") for r in rows])

    def _open(self, v):
        if v:
            row = DB.one("SELECT * FROM deals WHERE id=?", (v[0],))
            d = DealDlg(self.winfo_toplevel(), tuple(row))
        else:
            d = DealDlg(self.winfo_toplevel())
        d.bind("<Destroy>", lambda _: self.after(100, self.refresh))

    def _delete(self, v):
        if messagebox.askyesno("Видалити", f"Видалити угоду «{v[1]}»?"):
            DB.run("DELETE FROM deals WHERE id=?", (v[0],))
            Toast.show("Угоду видалено","ok")
            self.refresh()


# ══════════════════════════════════════════════════════
#  TASKS
# ══════════════════════════════════════════════════════

class Tasks(Module):
    ID = "tasks"; TITLE = "Завдання"; ICON = "✔"; GROUP = "CRM"

    def build(self):
        self.page_header("Завдання", "NEXUS › CRM › Завдання",
            actions=[
                {"label":"+ Завдання","style":"primary","icon":"+",
                 "cmd": lambda: self._open(None)},
            ])
        tb = tk.Frame(self, bg=T.c["bg1"], padx=22, pady=6)
        tb.pack(fill="x")
        self._st_var = tk.StringVar(value="Всі")
        for st in ["Всі","Відкрита","В роботі","Виконана","Відкладена"]:
            tk.Button(tb, text=st,
                bg=T.c["bg3"], fg=T.c["t1"],
                relief="flat", bd=0, cursor="hand2",
                font=F["sm"], padx=10, pady=4,
                command=lambda s=st: (self._st_var.set(s), self.refresh())
            ).pack(side="left", padx=(0,4))

        wrap = tk.Frame(self, bg=T.c["bg1"], padx=22, pady=(0,16))
        wrap.pack(fill="both", expand=True)
        cols = [
            {"id":"id",     "label":"ID",          "w":44,"str":False},
            {"id":"title",  "label":"ЗАВДАННЯ",    "w":240},
            {"id":"prio",   "label":"ПРІОРИТЕТ",   "w":90},
            {"id":"status", "label":"СТАТУС",      "w":100},
            {"id":"due",    "label":"ТЕРМІН",      "w":80},
            {"id":"owner",  "label":"ВІДПОВІДАЛЬНИЙ","w":130},
        ]
        self._tbl = Table(wrap, cols,
            on_dbl=lambda v: self._open(v),
            ctx=[
                {"label":"✏  Редагувати", "cmd": self._open},
                {"label":"✔  Виконано",
                 "cmd": lambda v: (
                     DB.run("UPDATE tasks SET status='Виконана' WHERE id=?", (v[0],)),
                     Toast.show("Позначено як виконане","ok"),
                     self.refresh())},
                "|",
                {"label":"✖  Видалити",   "cmd": self._delete},
            ])
        self._tbl.pack(fill="both", expand=True)

    def refresh(self):
        st = self._st_var.get() if hasattr(self,"_st_var") else "Всі"
        if st == "Всі":
            rows = DB.all(
                "SELECT id,title,priority,status,due,owner "
                "FROM tasks ORDER BY id DESC")
        else:
            rows = DB.all(
                "SELECT id,title,priority,status,due,owner "
                "FROM tasks WHERE status=? ORDER BY id DESC", (st,))
        self._tbl.load([tuple(r) for r in rows])

    def _open(self, v):
        if v:
            row = DB.one("SELECT * FROM tasks WHERE id=?", (v[0],))
            d = TaskDlg(self.winfo_toplevel(), tuple(row))
        else:
            d = TaskDlg(self.winfo_toplevel())
        d.bind("<Destroy>", lambda _: self.after(100, self.refresh))

    def _delete(self, v):
        if messagebox.askyesno("Видалити", f"Видалити «{v[1]}»?"):
            DB.run("DELETE FROM tasks WHERE id=?", (v[0],))
            Toast.show("Завдання видалено","ok")
            self.refresh()


# ══════════════════════════════════════════════════════
#  FINANCE
# ══════════════════════════════════════════════════════

class Finance(Module):
    ID = "finance"; TITLE = "Фінанси"; ICON = "₴"; GROUP = "ERP"

    def build(self):
        self.page_header("Фінанси", "NEXUS › ERP › Фінанси",
            actions=[
                {"label":"+ Рахунок","style":"primary","icon":"+",
                 "cmd": lambda: Toast.show("Форма рахунку","info")},
            ])
        sc = self.scrollable()
        kf = tk.Frame(sc, bg=T.c["bg1"], padx=22, pady=14)
        kf.pack(fill="x")
        kf.columnconfigure((0,1,2), weight=1, uniform="k")
        self._kf = kf
        for col,(title,acc) in enumerate([
            ("ДОХІД (₴)",      T.c["g"]),
            ("ВИТРАТИ (₴)",    T.c["w"]),
            ("ПРИБУТОК (₴)",   T.c["a"]),
        ]):
            c = self.kpi_card(kf, title, "…", accent=acc)
            c.grid(row=0,column=col,padx=(0,10 if col<2 else 0),sticky="nsew")
        self._kpis = kf.winfo_children()

        wrap = tk.Frame(sc, bg=T.c["bg1"], padx=22, pady=(0,24))
        wrap.pack(fill="both", expand=True)
        cols = [
            {"id":"id",     "label":"№",           "w":90,"str":False},
            {"id":"client", "label":"КОНТРАГЕНТ",  "w":160},
            {"id":"amount", "label":"СУМА",        "w":100,"anchor":"e"},
            {"id":"vat",    "label":"ПДВ",         "w":80,"anchor":"e"},
            {"id":"total",  "label":"РАЗОМ",       "w":100,"anchor":"e"},
            {"id":"status", "label":"СТАТУС",      "w":110},
            {"id":"due",    "label":"ТЕРМІН",      "w":80},
        ]
        self._tbl = Table(wrap, cols,
            on_dbl=lambda v: Toast.show(f"Рахунок {v[0]}","info"),
            ctx=[
                {"label":"⬇  PDF",
                 "cmd": lambda v: Toast.show(f"PDF {v[0]}","ok")},
                {"label":"✔  Позначити оплаченим",
                 "cmd": lambda v: Toast.show(f"Оплачено {v[0]}","ok")},
            ])
        self._tbl.pack(fill="both", expand=True)

    def refresh(self):
        rows = DB.all(
            "SELECT number,client,amount,vat,status,due FROM invoices ORDER BY id DESC")
        data = []
        for r in rows:
            total = r["amount"] + r["vat"]
            data.append((r["number"], r["client"],
                f"₴{r['amount']:,.0f}".replace(",","\u202f"),
                f"₴{r['vat']:,.0f}".replace(",","\u202f"),
                f"₴{total:,.0f}".replace(",","\u202f"),
                r["status"], r["due"]))
        self._tbl.load(data)

        paid  = DB.one("SELECT SUM(amount+vat) s FROM invoices WHERE status='Оплачено'")["s"] or 0
        exp   = paid * 0.35
        prof  = paid - exp

        def fmt(n): return f"₴{n:,.0f}".replace(",","\u202f")
        vals  = [fmt(paid), fmt(exp), fmt(prof)]
        for card, val in zip(self._kf.winfo_children(), vals):
            body = card.winfo_children()
            if len(body) < 2: continue
            body2 = body[1]
            for w in body2.winfo_children():
                try:
                    if w.cget("font") == str(F["num"]):
                        w.configure(text=val)
                        break
                except: pass


# ══════════════════════════════════════════════════════
#  RISKS
# ══════════════════════════════════════════════════════

class Risks(Module):
    ID = "risks"; TITLE = "Ризики"; ICON = "△"; GROUP = "GRC"

    LEVEL_COLORS = {
        "Критичний": "r","Високий":"w","Середній":"b","Низький":"g",
    }

    def build(self):
        self.page_header("Реєстр ризиків", "NEXUS › GRC › Ризики",
            actions=[
                {"label":"+ Ризик","style":"primary","icon":"+",
                 "cmd": lambda: self._open(None)},
            ])
        sc = self.scrollable()
        cf = tk.Frame(sc, bg=T.c["bg1"], padx=22, pady=14)
        cf.pack(fill="x")
        cf.columnconfigure((0,1,2,3), weight=1, uniform="k")
        self._cf = cf
        for col,(title,acc) in enumerate([
            ("ВСЬОГО",     T.c["t0"]),
            ("КРИТИЧНИХ",  T.c["r"]),
            ("ВІДКРИТИХ",  T.c["w"]),
            ("ЗАКРИТИХ",   T.c["g"]),
        ]):
            c = self.kpi_card(cf, title, "…", accent=acc)
            c.grid(row=0,column=col,padx=(0,10 if col<3 else 0),sticky="nsew")

        wrap = tk.Frame(sc, bg=T.c["bg1"], padx=22, pady=(0,24))
        wrap.pack(fill="both", expand=True)
        cols = [
            {"id":"id",    "label":"ID",         "w":50,"str":False},
            {"id":"title", "label":"РИЗИК",       "w":230},
            {"id":"cat",   "label":"КАТЕГОРІЯ",   "w":120},
            {"id":"impact","label":"ВПЛИВ",       "w":90},
            {"id":"prob",  "label":"ЙМОВІРНІСТЬ", "w":130},
            {"id":"level", "label":"РІВЕНЬ",      "w":90},
            {"id":"owner", "label":"ВЛАСНИК",     "w":110},
            {"id":"status","label":"СТАТУС",      "w":90},
        ]
        self._tbl = Table(wrap, cols,
            on_dbl=lambda v: self._open(v),
            ctx=[
                {"label":"✏  Редагувати",   "cmd": self._open},
                {"label":"✔  Закрити ризик",
                 "cmd": lambda v: (
                     DB.run("UPDATE risks SET status='Закритий' WHERE id=?",(v[0],)),
                     Toast.show("Ризик закрито","ok"),
                     self.refresh())},
                "|",
                {"label":"✖  Видалити",     "cmd": self._delete},
            ])
        self._tbl.pack(fill="both", expand=True)

    def refresh(self):
        rows = DB.all(
            "SELECT id,title,category,impact,prob,level,owner,status "
            "FROM risks ORDER BY id")
        self._tbl.load([tuple(r) for r in rows])

        def cnt(q, *p): return DB.one(q, p)["c"]
        total  = cnt("SELECT COUNT(*) c FROM risks")
        crit   = cnt("SELECT COUNT(*) c FROM risks WHERE level=?","Критичний")
        opened = cnt("SELECT COUNT(*) c FROM risks WHERE status='Відкритий'")
        closed = cnt("SELECT COUNT(*) c FROM risks WHERE status='Закритий'")
        vals   = [str(total), str(crit), str(opened), str(closed)]
        for card, val in zip(self._cf.winfo_children(), vals):
            body = card.winfo_children()
            if len(body) < 2: continue
            for w in body[1].winfo_children():
                try:
                    if w.cget("font") == str(F["num"]):
                        w.configure(text=val); break
                except: pass

    def _open(self, v):
        if v:
            row = DB.one("SELECT * FROM risks WHERE id=?", (v[0],))
            d = RiskDlg(self.winfo_toplevel(), tuple(row))
        else:
            d = RiskDlg(self.winfo_toplevel())
        d.bind("<Destroy>", lambda _: self.after(100, self.refresh))

    def _delete(self, v):
        if messagebox.askyesno("Видалити", f"Видалити «{v[1]}»?"):
            DB.run("DELETE FROM risks WHERE id=?", (v[0],))
            Toast.show("Ризик видалено","ok")
            self.refresh()


# ══════════════════════════════════════════════════════
#  SETTINGS
# ══════════════════════════════════════════════════════

class Settings(Module):
    ID = "settings"; TITLE = "Налаштування"; ICON = "⚙"

    def build(self):
        self.page_header("Налаштування", "NEXUS › Налаштування")
        sc = self.scrollable()
        pad = tk.Frame(sc, bg=T.c["bg1"], padx=32, pady=20)
        pad.pack(fill="both", expand=True)

        self._section(pad, "ЗОВНІШНІЙ ВИГЛЯД")
        r = tk.Frame(pad, bg=T.c["bg1"])
        r.pack(fill="x", pady=(8,0))
        lbl(r,"Тема інтерфейсу","bold",T.c["t1"],
            width=24, anchor="w").pack(side="left")
        self._tv = tk.StringVar(value=T._name)
        for val, text in [("dark","◑  Темна"),("light","☀  Світла")]:
            tk.Radiobutton(r, text=text,
                variable=self._tv, value=val,
                bg=T.c["bg1"], fg=T.c["t0"],
                selectcolor=T.c["bg3"],
                activebackground=T.c["bg1"],
                font=F["body"],
                command=lambda v=val: T.apply(v) or T.toggle() or T.toggle()
            ).pack(side="left", padx=10)

        self._section(pad, "ЗАГАЛЬНІ")
        for k,v in [
            ("Організація",     "Моя компанія"),
            ("Версія",          "3.0.0"),
            ("База даних",      str(DB._conn)),
            ("Мова",            "Українська"),
            ("Часовий пояс",    "Europe/Kyiv"),
        ]:
            self._row(pad, k, v)

        self._section(pad, "СИСТЕМА")
        for k,v in [
            ("Python",   sys.version.split()[0]),
            ("Tkinter",  str(tk.TkVersion)),
            ("ОС",       platform.system()+" "+platform.release()),
            ("Процесор", platform.machine()),
        ]:
            self._row(pad, k, v, mono=True)

        tk.Frame(pad, bg=T.c["sep"], height=1).pack(fill="x", pady=20)
        br = tk.Frame(pad, bg=T.c["bg1"])
        br.pack(anchor="w")
        Btn(br,"Зберегти","primary",
            command=lambda: Toast.show("Налаштування збережено","ok")
            ).pack(side="left", ipady=3)
        Btn(br,"Скинути БД","danger",
            command=self._reset_db
            ).pack(side="left", padx=8, ipady=3)

    def _section(self, p, title):
        tk.Frame(p, bg=T.c["sep"], height=1).pack(fill="x", pady=(16,6))
        lbl(p, title, "cap", T.c["t3"]).pack(anchor="w")

    def _row(self, p, k, v, mono=False):
        r = tk.Frame(p, bg=T.c["bg1"])
        r.pack(fill="x", pady=5)
        lbl(r, k, "body", T.c["t2"], width=22, anchor="w").pack(side="left")
        lbl(r, str(v), "mono" if mono else "bold", T.c["t1"]).pack(side="left")

    def _reset_db(self):
        if messagebox.askyesno("Скинути базу даних",
                "Видалити всі дані та заново заповнити тестовими?"):
            import os
            try: os.remove(str(DB._conn))
            except: pass
            DB.connect()
            Toast.show("БД скинута і заповнена","ok")

    def refresh(self):
        # оновити значення теми
        if hasattr(self, "_tv"):
            self._tv.set(T._name)


# ══════════════════════════════════════════════════════
#  REGISTRY + NAV
# ══════════════════════════════════════════════════════

MODULES = {
    "dashboard": Dashboard,
    "contacts":  Contacts,
    "deals":     Deals,
    "tasks":     Tasks,
    "finance":   Finance,
    "risks":     Risks,
    "settings":  Settings,
}

NAV = [
    {"id":"dashboard","label":"Dashboard",    "icon":"⊞"},
    {"group":"CRM","icon":"◈","items":[
        {"id":"contacts","label":"Контакти",  "icon":"◉"},
        {"id":"deals",   "label":"Угоди",     "icon":"◇"},
        {"id":"tasks",   "label":"Завдання",  "icon":"✔"},
    ]},
    {"group":"ERP","icon":"◫","items":[
        {"id":"finance", "label":"Фінанси",   "icon":"₴"},
    ]},
    {"group":"GRC","icon":"◬","items":[
        {"id":"risks",   "label":"Ризики",    "icon":"△"},
    ]},
    {"id":"settings","label":"Налаштування","icon":"⚙"},
]


# ══════════════════════════════════════════════════════
#  TOP BAR
# ══════════════════════════════════════════════════════

class TopBar(tk.Frame):
    def __init__(self, parent, on_burger, on_search):
        super().__init__(parent,
            bg=T.c["bg1"], height=48)
        self.pack_propagate(False)

        # left
        left = tk.Frame(self, bg=T.c["bg1"])
        left.pack(side="left", fill="y", padx=(6,0))
        tk.Button(left, text="☰",
            bg=T.c["bg1"], fg=T.c["t2"],
            relief="flat", bd=0, cursor="hand2",
            font=("Segoe UI",15),
            activebackground=T.c["bg3"],
            activeforeground=T.c["t0"],
            command=on_burger).pack(side="left", padx=(4,8))

        logo = tk.Frame(left, bg=T.c["a"], width=24, height=24)
        logo.pack_propagate(False)
        logo.pack(side="left")
        tk.Label(logo, text="N", bg=T.c["a"],
                 fg=T.c["inv"], font=("Segoe UI",11,"bold")
                 ).place(relx=.5,rely=.5,anchor="center")
        lbl(left, " NEXUS", "logo", T.c["t0"]).pack(side="left")

        # center: search
        mid = tk.Frame(self, bg=T.c["bg1"])
        mid.pack(side="left", fill="both", expand=True, padx=16)
        sw = tk.Frame(mid, bg=T.c["bg3"],
                      highlightbackground=T.c["b0"],
                      highlightthickness=1)
        sw.pack(side="left", fill="y", pady=10)
        lbl(sw, "⌕", "body", T.c["t2"],
            bg=T.c["bg3"]).pack(side="left", padx=(8,0))
        self._sv = tk.StringVar()
        self._se = tk.Entry(sw,
            textvariable=self._sv,
            bg=T.c["bg3"], fg=T.c["t0"],
            insertbackground=T.c["a"],
            relief="flat", bd=0, width=30, font=F["body"])
        self._se.pack(side="left", pady=2, padx=(2,8))
        self._se.insert(0,"Пошук…")
        self._se.configure(fg=T.c["t2"])
        self._se.bind("<FocusIn>",  self._si)
        self._se.bind("<FocusOut>", self._so)
        self._se.bind("<Return>",   lambda _: on_search(self._sv.get()))

        # right
        right = tk.Frame(self, bg=T.c["bg1"])
        right.pack(side="right", fill="y", padx=(0,10))
        tk.Button(right, text="◑",
            bg=T.c["bg1"], fg=T.c["t2"],
            relief="flat", bd=0, cursor="hand2",
            font=("Segoe UI",13),
            activebackground=T.c["bg3"],
            activeforeground=T.c["t0"],
            command=T.toggle).pack(side="left", padx=4)

        tk.Frame(right, bg=T.c["sep"], width=1).pack(
            side="left", fill="y", pady=10, padx=6)

        av = tk.Frame(right, bg=T.c["a2"], width=28, height=28)
        av.pack_propagate(False)
        av.pack(side="left")
        tk.Label(av, text="ОМ", bg=T.c["a2"],
                 fg="#FFF", font=("Segoe UI",8,"bold")
                 ).place(relx=.5,rely=.5,anchor="center")
        lbl(right, "  Адмін", "bold", T.c["t0"]).pack(side="left")

        # bottom sep — пакується ззовні
        self._sep = tk.Frame(parent, bg=T.c["sep"], height=1)

    def _si(self, _=None):
        if self._se.get() == "Пошук…":
            self._se.delete(0,"end")
            self._se.configure(fg=T.c["t0"])

    def _so(self, _=None):
        if not self._se.get():
            self._se.insert(0,"Пошук…")
            self._se.configure(fg=T.c["t2"])

    def focus_search(self):
        self._se.focus_set()
        self._si()


# ══════════════════════════════════════════════════════
#  SIDE BAR
# ══════════════════════════════════════════════════════

class SideBar(tk.Frame):
    W  = 208
    WC = 44

    def __init__(self, parent, on_nav):
        super().__init__(parent, bg=T.c["bg1"], width=self.W)
        self.pack_propagate(False)
        self._on_nav  = on_nav
        self._active  = "dashboard"
        self._open    = {"CRM"}
        self._items   = {}
        self._collapsed = False
        self._draw()

    def _draw(self):
        for w in self.winfo_children(): w.destroy()
        self._items.clear()
        tk.Frame(self, bg=T.c["bg1"], height=8).pack(fill="x")
        for entry in NAV:
            if "group" in entry:
                self._group(entry)
            else:
                self._item(entry)
        # bottom
        tk.Frame(self, bg=T.c["sep"], height=1).pack(
            fill="x", side="bottom")
        tk.Label(self, text=f"v3.0.0",
                 bg=T.c["bg1"], fg=T.c["t3"],
                 font=F["xs"]).pack(side="bottom", pady=6)

    def _group(self, entry):
        grp   = entry["group"]
        icon  = entry["icon"]
        items = entry["items"]
        is_open = grp in self._open

        hdr = tk.Frame(self, bg=T.c["bg1"], cursor="hand2")
        hdr.pack(fill="x", pady=(6,0))
        tk.Label(hdr, text=icon,
                 bg=T.c["bg1"], fg=T.c["t3"],
                 font=("Segoe UI",10), padx=12).pack(side="left")
        tk.Label(hdr, text=grp,
                 bg=T.c["bg1"], fg=T.c["t3"],
                 font=F["cap"]).pack(side="left", fill="x", expand=True)
        av = tk.StringVar(value="▾" if is_open else "›")
        al = tk.Label(hdr, textvariable=av,
                      bg=T.c["bg1"], fg=T.c["t3"],
                      font=F["sm"], padx=10)
        al.pack(side="right")

        kf = tk.Frame(self, bg=T.c["bg1"])
        if is_open: kf.pack(fill="x")

        for it in items:
            self._item(it, kf, depth=1)

        def toggle(_=None):
            if grp in self._open:
                self._open.discard(grp); kf.pack_forget(); av.set("›")
            else:
                self._open.add(grp); kf.pack(fill="x"); av.set("▾")

        for w in (hdr, al):
            w.bind("<Button-1>", toggle)
            w.bind("<Enter>", lambda _,h=hdr: self._hov(h, True))
            w.bind("<Leave>", lambda _,h=hdr: self._hov(h, False))

    def _item(self, entry, parent=None, depth=0):
        parent = parent or self
        iid   = entry["id"]
        icon  = entry["icon"]
        label = entry["label"]
        act   = (iid == self._active)
        bg    = T.c["bg4"] if act else T.c["bg1"]
        fg    = T.c["t0"]  if act else T.c["t1"]

        row = tk.Frame(parent, bg=bg, cursor="hand2")
        row.pack(fill="x", pady=(1,0))
        ind = tk.Frame(row, bg=T.c["a"] if act else bg, width=3)
        ind.pack(side="left", fill="y")
        if depth:
            tk.Frame(row, bg=bg, width=depth*14).pack(side="left")
        tk.Label(row, text=icon, bg=bg,
                 fg=T.c["a"] if act else T.c["t3"],
                 font=("Segoe UI",11), padx=8, pady=8).pack(side="left")
        ll = tk.Label(row, text=label, bg=bg, fg=fg,
                      font=F["bold"] if act else F["body"], anchor="w")
        ll.pack(side="left", fill="x", expand=True, pady=8)
        self._items[iid] = (row, ind, ll)

        def click(_=None, _id=iid):
            self._activate(_id)
            self._on_nav(_id)

        for w in (row, ll):
            w.bind("<Button-1>", click)
            w.bind("<Enter>", lambda _,r=row,_id=iid: (
                r.configure(bg=T.c["bgh"]) if _id!=self._active else None,
                [c.configure(bg=T.c["bgh"]) for c in r.winfo_children()
                 if _id!=self._active]))
            w.bind("<Leave>", lambda _,r=row,_id=iid,_bg=bg: (
                r.configure(bg=_bg) if _id!=self._active else None,
                [c.configure(bg=_bg) for c in r.winfo_children()
                 if _id!=self._active]))

    def _hov(self, h, on):
        bg = T.c["bg3"] if on else T.c["bg1"]
        h.configure(bg=bg)
        for c in h.winfo_children():
            try: c.configure(bg=bg)
            except: pass

    def _activate(self, iid):
        old = self._items.get(self._active)
        if old:
            r,i,l = old
            r.configure(bg=T.c["bg1"])
            i.configure(bg=T.c["bg1"])
            l.configure(bg=T.c["bg1"], fg=T.c["t1"], font=F["body"])
            for c in r.winfo_children():
                try: c.configure(bg=T.c["bg1"])
                except: pass
        self._active = iid
        new = self._items.get(iid)
        if new:
            r,i,l = new
            r.configure(bg=T.c["bg4"])
            i.configure(bg=T.c["a"])
            l.configure(bg=T.c["bg4"], fg=T.c["t0"], font=F["bold"])
            for c in r.winfo_children():
                try: c.configure(bg=T.c["bg4"])
                except: pass
            l.configure(bg=T.c["bg4"])

    def toggle_collapse(self):
        self._collapsed = not self._collapsed
        self.configure(width=self.WC if self._collapsed else self.W)


# ══════════════════════════════════════════════════════
#  STATUS BAR
# ══════════════════════════════════════════════════════

class StatusBar(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent, bg=T.c["bg1"], height=24)
        self.pack_propagate(False)
        # sep пакується ззовні
        self._sep = tk.Frame(parent, bg=T.c["sep"], height=1)

        self._sv = tk.StringVar(value="Готово")
        self._tv = tk.StringVar()

        tk.Label(self, textvariable=self._sv,
                 bg=T.c["bg1"], fg=T.c["t2"],
                 font=F["xs"], padx=12).pack(side="left", fill="y")
        tk.Label(self, text="●  З'єднання активне",
                 bg=T.c["bg1"], fg=T.c["g"],
                 font=F["xs"]).pack(side="left")
        tk.Label(self, textvariable=self._tv,
                 bg=T.c["bg1"], fg=T.c["t3"],
                 font=F["mono"], padx=12).pack(side="right", fill="y")
        self._tick()

    def msg(self, text):  self._sv.set(text)
    def _tick(self):
        self._tv.set(datetime.now().strftime("%H:%M:%S   %d.%m.%Y"))
        self.after(1000, self._tick)


# ══════════════════════════════════════════════════════
#  WORKSPACE
# ══════════════════════════════════════════════════════

class Workspace(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent, bg=T.c["bg1"])
        self._cache   = {}
        self._current = None

    def show(self, mid):
        if self._current == mid: return
        if self._current in self._cache:
            self._cache[self._current].pack_forget()
        if mid not in self._cache:
            cls  = MODULES.get(mid, Module)
            inst = cls(self)
            self._cache[mid] = inst
        self._cache[mid].pack(fill="both", expand=True)
        self._cache[mid].activate()
        self._current = mid


# ══════════════════════════════════════════════════════
#  APPLICATION
# ══════════════════════════════════════════════════════

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        # HiDPI
        if platform.system() == "Windows":
            try:
                from ctypes import windll
                windll.shcore.SetProcessDpiAwareness(1)
            except: pass

        self.title("NEXUS  —  Business Platform  v3.0")
        self.geometry("1300x800")
        self.minsize(960, 620)
        self.configure(bg=T.c["bg0"])

        DB.connect()
        Toast.init(self)
        T.sub(self._rebuild)

        self._build()
        self._hotkeys()
        self.after(700, lambda: Toast.show("Ласкаво просимо до NEXUS!","ok"))

    # ── build ─────────────────────────────────────────
    def _build(self):
        self.configure(bg=T.c["bg0"])

        self._topbar = TopBar(self,
            on_burger=self._burger,
            on_search=self._search)
        self._topbar.pack(fill="x", side="top")
        self._topbar._sep.pack(fill="x", side="top")

        self._main = tk.Frame(self, bg=T.c["bg0"])
        self._main.pack(fill="both", expand=True, side="top")

        self._sidebar = SideBar(self._main, on_nav=self._nav)
        self._sidebar.pack(side="left", fill="y")
        tk.Frame(self._main, bg=T.c["sep"], width=1).pack(
            side="left", fill="y")

        self._ws = Workspace(self._main)
        self._ws.pack(side="left", fill="both", expand=True)

        self._sb = StatusBar(self)
        self._sb._sep.pack(fill="x", side="bottom")
        self._sb.pack(fill="x", side="bottom")

        self._nav("dashboard")

    # ── rebuild on theme change ────────────────────────
    def _rebuild(self):
        for w in self.winfo_children():
            w.destroy()
        self._build()

    # ── actions ────────────────────────────────────────
    def _nav(self, mid):
        self._ws.show(mid)
        self._sb.msg(f"Відкрито: {mid}")

    def _burger(self):
        self._sidebar.toggle_collapse()

    def _search(self, q):
        if q and q != "Пошук…":
            Toast.show(f"Пошук: «{q}»","info")

    # ── hotkeys ────────────────────────────────────────
    def _hotkeys(self):
        mids = list(MODULES.keys())
        for i, mid in enumerate(mids, 1):
            self.bind(f"<Control-Key-{i}>",
                      lambda _, m=mid: self._nav(m))
        self.bind("<Control-t>", lambda _: T.toggle())
        self.bind("<Control-T>", lambda _: T.toggle())
        self.bind("<Control-f>", lambda _: self._topbar.focus_search())
        self.bind("<Control-F>", lambda _: self._topbar.focus_search())
        self.bind("<F5>",        lambda _: self._refresh())
        self.bind("<Control-q>", lambda _: self.quit())

    def _refresh(self):
        cur = self._ws._current
        if cur and cur in self._ws._cache:
            self._ws._cache[cur].refresh()
            Toast.show("Оновлено","ok")


# ══════════════════════════════════════════════════════
#  ENTRY POINT
# ══════════════════════════════════════════════════════

if __name__ == "__main__":
    App().mainloop()
