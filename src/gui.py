"""
gui.py — графічний інтерфейс генератора тестових завдань Part 147
Запуск: python src/gui.py
"""

import os
import sys
import sqlite3
import platform
import threading
import subprocess
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

_IS_MAC = platform.system() == "Darwin"

_SRC = os.path.dirname(os.path.abspath(__file__))
_ROOT = os.path.dirname(_SRC)
sys.path.insert(0, _SRC)

import database as db
import generator as gen
import exporter_docx as ex_docx
import parser as prs
import quota_loader as ql

# ─── Палітра кольорів ─────────────────────────────────────────────────────────
C_SIDEBAR   = "#1E3A5F"
C_SIDEBAR_H = "#2D5480"
C_ACCENT    = "#2E86AB"
C_BG        = "#F0F4F8"
C_CARD      = "#FFFFFF"
C_TEXT      = "#2C3E50"
C_MUTED     = "#7F8C8D"
C_BORDER    = "#DDE1E7"
C_SUCCESS   = "#27AE60"
C_WARNING   = "#E67E22"
C_DANGER    = "#E74C3C"
C_ROW_ALT   = "#F7FAFC"
C_STEP_DONE = "#D4EDDA"

FONT        = ("Helvetica", 12)
FONT_SM     = ("Helvetica", 11)
FONT_BOLD   = ("Helvetica", 12, "bold")
FONT_TITLE  = ("Helvetica", 18, "bold")
FONT_MONO   = ("Courier", 11)


# ─── Допоміжні компоненти ─────────────────────────────────────────────────────

def card(parent, title="", **kw):
    """Картка з рамкою і заголовком."""
    outer = tk.Frame(parent, bg=C_BORDER, padx=1, pady=1)
    inner = tk.Frame(outer, bg=C_CARD, **kw)
    inner.pack(fill="both", expand=True)
    if title:
        tk.Label(inner, text=title, bg=C_CARD, fg=C_SIDEBAR,
                 font=FONT_BOLD).pack(anchor="w", padx=16, pady=(12, 4))
    return outer, inner


def accent_btn(parent, text, command, width=None):
    kw = {"width": width} if width else {}
    return tk.Button(parent, text=text, command=command,
                     bg=C_ACCENT, fg="white", activebackground=C_SIDEBAR,
                     activeforeground="white", relief="flat", bd=0,
                     font=FONT_BOLD, padx=14, pady=8, cursor="hand2", **kw)


def ghost_btn(parent, text, command, width=None):
    kw = {"width": width} if width else {}
    return tk.Button(parent, text=text, command=command,
                     bg=C_BORDER, fg=C_TEXT, activebackground="#C8D0DA",
                     relief="flat", bd=0, font=FONT_SM,
                     padx=10, pady=6, cursor="hand2", **kw)


# ─── CounterWidget ────────────────────────────────────────────────────────────

class CounterWidget(tk.Frame):
    """Компонент «−  значення  +» з прив'язаною IntVar."""

    def __init__(self, parent, var: tk.IntVar,
                 min_val: int = 0, max_val: int = 999,
                 btn_size: int = 28, font_size: int = 13,
                 bg=C_CARD, **kw):
        super().__init__(parent, bg=bg, **kw)
        self._var = var
        self._min = min_val
        self._max = max_val

        # Кнопка «−»
        self._btn_minus = tk.Label(
            self, text="−", bg="#E8EDF2", fg=C_TEXT,
            font=("Helvetica", font_size, "bold"),
            width=2, cursor="hand2",
            relief="flat",
        )
        self._btn_minus.pack(side="left")

        # Значення
        self._lbl = tk.Label(
            self, textvariable=var,
            bg=bg, fg=C_TEXT,
            font=("Helvetica", font_size, "bold"),
            width=4, anchor="center",
        )
        self._lbl.pack(side="left", padx=4)

        # Кнопка «+»
        self._btn_plus = tk.Label(
            self, text="+", bg=C_ACCENT, fg="white",
            font=("Helvetica", font_size, "bold"),
            width=2, cursor="hand2",
            relief="flat",
        )
        self._btn_plus.pack(side="left")

        self._btn_minus.bind("<Button-1>", self._dec)
        self._btn_plus.bind("<Button-1>",  self._inc)
        self._var.trace_add("write", self._update_states)
        self._update_states()

    def _dec(self, _=None):
        v = self._var.get()
        if v > self._min:
            self._var.set(v - 1)

    def _inc(self, _=None):
        v = self._var.get()
        if v < self._max:
            self._var.set(v + 1)

    def _update_states(self, *_):
        try:
            v = self._var.get()
        except tk.TclError:
            return
        self._btn_minus.configure(
            bg="#C8D0DA" if v <= self._min else "#E8EDF2",
            fg=C_MUTED if v <= self._min else C_TEXT,
        )
        self._btn_plus.configure(
            bg=C_MUTED if v >= self._max else C_ACCENT,
        )

    def set_max(self, max_val: int):
        self._max = max_val
        self._update_states()


# ─── Головне вікно ────────────────────────────────────────────────────────────

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Part 147 — Генератор тестових завдань")
        self.geometry("1150x720")
        self.minsize(960, 620)
        self.configure(bg=C_BG)

        self.db_path    = tk.StringVar(value=db.DB_PATH)
        self.output_dir = tk.StringVar(value=os.path.join(_ROOT, "output"))

        self._build()
        self._nav("generate")

    def _build(self):
        # Бічна панель
        self.sidebar = tk.Frame(self, bg=C_SIDEBAR, width=210)
        self.sidebar.pack(side="left", fill="y")
        self.sidebar.pack_propagate(False)

        tk.Label(self.sidebar, text="✈", bg=C_SIDEBAR, fg="white",
                 font=("Helvetica", 28)).pack(pady=(28, 2))
        tk.Label(self.sidebar, text="PART 147", bg=C_SIDEBAR, fg="white",
                 font=("Helvetica", 15, "bold")).pack()
        tk.Label(self.sidebar, text="Генератор тестів", bg=C_SIDEBAR,
                 fg="#8DAFC8", font=("Helvetica", 10)).pack(pady=(0, 16))

        tk.Frame(self.sidebar, bg="#2D4F70", height=1).pack(fill="x", padx=20, pady=4)

        self._nav_btns = {}
        self._active_nav = None
        nav = [
            ("generate", "✈   Генерація"),
            ("modules",  "📦  Модулі"),
            ("stats",    "📊  Статистика"),
            ("settings", "⚙️   Налаштування"),
        ]
        for key, label in nav:
            row = tk.Frame(self.sidebar, bg=C_SIDEBAR, cursor="hand2")
            row.pack(fill="x", pady=1)

            accent_bar = tk.Frame(row, bg=C_SIDEBAR, width=4)
            accent_bar.pack(side="left", fill="y")
            accent_bar.pack_propagate(False)

            lbl = tk.Label(row, text=label, anchor="w",
                           bg=C_SIDEBAR, fg="#8DAFC8",
                           font=("Helvetica", 12), padx=14, pady=9)
            lbl.pack(side="left", fill="x", expand=True)

            self._nav_btns[key] = {"row": row, "bar": accent_bar, "lbl": lbl}

            def _click(e, k=key): self._nav(k)
            def _enter(e, k=key):
                if k != self._active_nav:
                    self._nav_btns[k]["row"].configure(bg="#243552")
                    self._nav_btns[k]["lbl"].configure(bg="#243552", fg="white")
            def _leave(e, k=key):
                if k != self._active_nav:
                    self._nav_btns[k]["row"].configure(bg=C_SIDEBAR)
                    self._nav_btns[k]["lbl"].configure(bg=C_SIDEBAR, fg="#8DAFC8")

            for w in (row, accent_bar, lbl):
                w.bind("<Button-1>", _click)
                w.bind("<Enter>", _enter)
                w.bind("<Leave>", _leave)

        tk.Label(self.sidebar, text="v1.0  •  EASA Part 147",
                 bg=C_SIDEBAR, fg="#5A7A9A",
                 font=("Helvetica", 9)).pack(side="bottom", pady=14)

        # Основний контент
        content_wrap = tk.Frame(self, bg=C_BG)
        content_wrap.pack(side="left", fill="both", expand=True)

        self.statusbar = tk.Frame(content_wrap, bg="#E2E8F0", height=28)
        self.statusbar.pack(side="bottom", fill="x")
        self.statusbar.pack_propagate(False)
        self._status_lbl = tk.Label(self.statusbar, text="Готово",
                                    bg="#E2E8F0", fg=C_MUTED,
                                    font=("Helvetica", 10), anchor="w")
        self._status_lbl.pack(side="left", padx=14)

        self.container = tk.Frame(content_wrap, bg=C_BG)
        self.container.pack(fill="both", expand=True)

        self._frames = {
            "generate": GenerateFrame(self.container, self),
            "modules":  ModulesFrame(self.container, self),
            "stats":    StatsFrame(self.container, self),
            "settings": SettingsFrame(self.container, self),
        }
        for f in self._frames.values():
            f.place(relx=0, rely=0, relwidth=1, relheight=1)

    def _nav(self, key):
        self._active_nav = key
        for k, w in self._nav_btns.items():
            if k == key:
                w["row"].configure(bg=C_SIDEBAR_H)
                w["bar"].configure(bg=C_ACCENT)
                w["lbl"].configure(bg=C_SIDEBAR_H, fg="white")
            else:
                w["row"].configure(bg=C_SIDEBAR)
                w["bar"].configure(bg=C_SIDEBAR)
                w["lbl"].configure(bg=C_SIDEBAR, fg="#8DAFC8")
        self._frames[key].tkraise()
        if hasattr(self._frames[key], "on_show"):
            self._frames[key].on_show()

    def set_status(self, text, color=C_MUTED):
        self._status_lbl.configure(text=text, fg=color)
        self.update_idletasks()


# ─── Вкладка: Генерація (5-крокова форма-wizard) ─────────────────────────────

class GenerateFrame(tk.Frame):
    STEPS = 4   # 0..4

    def __init__(self, parent, app):
        super().__init__(parent, bg=C_BG)
        self.app = app
        self._step = 0

        # Змінні вибору
        self._cat_var       = tk.StringVar(value="B1.1")
        self._variants_var  = tk.IntVar(value=1)
        self._quota_vars    = {}   # section_id -> tk.IntVar
        self._avail         = {}   # section_id -> int
        self._section_rows  = {}   # section_id -> (row_frame, counter_widget)
        self._module_expanded = {} # module_id  -> bool

        self._build()

    # ── Побудова загального каркасу ───────────────────────────────────────────

    def _build(self):
        # Верхній індикатор кроків
        self._indicator = tk.Frame(self, bg=C_BG)
        self._indicator.pack(fill="x", padx=32, pady=(18, 0))
        self._step_lbls = []
        steps = ["Категорія", "Квоти", "Варіанти", "Папка", "Генерація"]
        for i, s in enumerate(steps):
            col = tk.Frame(self._indicator, bg=C_BG)
            col.pack(side="left", expand=True)

            circle = tk.Label(col, text=str(i + 1), width=2,
                              bg=C_BORDER, fg=C_MUTED,
                              font=("Helvetica", 10, "bold"),
                              relief="flat")
            circle.pack()
            lbl = tk.Label(col, text=s, bg=C_BG, fg=C_MUTED,
                           font=("Helvetica", 9))
            lbl.pack()
            self._step_lbls.append((circle, lbl))

            if i < len(steps) - 1:
                tk.Frame(self._indicator, bg=C_BORDER,
                         height=2, width=40).pack(side="left", expand=True,
                                                   pady=8)

        # Контейнер кроків
        self._pages_wrap = tk.Frame(self, bg=C_BG)
        self._pages_wrap.pack(fill="both", expand=True, padx=0, pady=0)

        self._pages = {}
        builders = [
            self._build_step0,
            self._build_step1,
            self._build_step2,
            self._build_step3,
            self._build_step4,
        ]
        for i, builder in enumerate(builders):
            page = tk.Frame(self._pages_wrap, bg=C_BG)
            page.place(relx=0, rely=0, relwidth=1, relheight=1)
            self._pages[i] = page
            builder(page)

        self._go_to(0)

    # ── Навігація між кроками ────────────────────────────────────────────────

    def _go_to(self, step: int):
        self._step = step
        self._pages[step].tkraise()
        self._update_indicator()

    def _update_indicator(self):
        colors = {
            "done":    (C_SUCCESS,  "white",  C_SUCCESS),
            "active":  (C_ACCENT,   "white",  C_ACCENT),
            "pending": (C_BORDER,   C_MUTED,  C_MUTED),
        }
        for i, (circle, lbl) in enumerate(self._step_lbls):
            if i < self._step:
                state = "done"
                circle.configure(text="✓")
            elif i == self._step:
                state = "active"
                circle.configure(text=str(i + 1))
            else:
                state = "pending"
                circle.configure(text=str(i + 1))
            bg, fg, lbfg = colors[state]
            circle.configure(bg=bg, fg=fg)
            lbl.configure(fg=lbfg)

    # ── Крок 0: Вибір категорії ───────────────────────────────────────────────

    def _build_step0(self, page):
        # Центрований контент
        wrap = tk.Frame(page, bg=C_BG)
        wrap.place(relx=0.5, rely=0.5, anchor="center")

        tk.Label(wrap, text="✈", bg=C_BG, fg=C_ACCENT,
                 font=("Helvetica", 48)).pack(pady=(0, 8))
        tk.Label(wrap, text="Оберіть категорію здобувача",
                 bg=C_BG, fg=C_SIDEBAR, font=("Helvetica", 16, "bold")).pack()
        tk.Label(wrap, text="EASA Part 147 — B1.1 / B1.3 / B2",
                 bg=C_BG, fg=C_MUTED, font=("Helvetica", 11)).pack(pady=(2, 24))

        cards_row = tk.Frame(wrap, bg=C_BG)
        cards_row.pack()

        cat_info = {
            "B1.1": ("Авіаційний механік\nповітряне судно\nз ГТД", "✈"),
            "B1.3": ("Авіаційний механік\nповітряне судно\nз ПД", "🔧"),
            "B2":   ("Авіаційний механік\nавіоніка", "📡"),
        }

        for cat, (desc, icon) in cat_info.items():
            c = tk.Frame(cards_row, bg=C_CARD,
                         highlightthickness=2,
                         highlightbackground=C_BORDER,
                         cursor="hand2", padx=20, pady=18)
            c.pack(side="left", padx=10)

            tk.Label(c, text=icon, bg=C_CARD,
                     font=("Helvetica", 26)).pack()
            tk.Label(c, text=cat, bg=C_CARD, fg=C_SIDEBAR,
                     font=("Helvetica", 15, "bold")).pack(pady=(4, 2))
            tk.Label(c, text=desc, bg=C_CARD, fg=C_MUTED,
                     font=("Helvetica", 9), justify="center").pack()

            def _select(e=None, cv=cat, frame=c):
                self._cat_var.set(cv)
                for child in cards_row.winfo_children():
                    child.configure(highlightbackground=C_BORDER,
                                    bg=C_CARD)
                    for w in child.winfo_children():
                        w.configure(bg=C_CARD)
                frame.configure(highlightbackground=C_ACCENT, bg="#EBF5FB")
                for w in frame.winfo_children():
                    w.configure(bg="#EBF5FB")

            for w in [c] + c.winfo_children():
                w.bind("<Button-1>", _select)

        self._cat_cards_row = cards_row

        # Вибираємо першу картку за замовчуванням при показі
        self._step0_initialized = False

        accent_btn(wrap, "Далі  →", self._from_step0_to_1).pack(pady=(28, 0))

    def _from_step0_to_1(self):
        cat = self._cat_var.get()
        dp = self.app.db_path.get()

        if not os.path.exists(dp):
            messagebox.showwarning("Увага",
                "База даних не знайдена.\nПерейдіть у Налаштування → Ініціалізувати БД.")
            return

        try:
            cat_id = db.get_category_id(cat, dp)
        except Exception:
            messagebox.showwarning("Увага",
                "БД не ініціалізована. Перейдіть у Налаштування → Ініціалізувати БД.")
            return

        if not cat_id:
            messagebox.showwarning("Увага",
                f"Категорія {cat} ще не завантажена.\nСпочатку завантажте квоти на вкладці «Модулі».")
            return

        self._current_cat_id = cat_id
        self._fill_quota_list()
        self._go_to(1)

    # ── Крок 1: Квоти за розділами ────────────────────────────────────────────

    def _build_step1(self, page):
        hdr = tk.Frame(page, bg=C_BG)
        hdr.pack(fill="x", padx=28, pady=(14, 0))

        ghost_btn(hdr, "← Назад", lambda: self._go_to(0)).pack(side="left")
        tk.Label(hdr, text="Квоти питань за розділами",
                 bg=C_BG, fg=C_SIDEBAR, font=FONT_TITLE).pack(side="left", padx=16)

        # Легенда
        tk.Label(page,
                 text="Натисніть на модуль, щоб розгорнути розділи. "
                      "«В БД» — доступних питань. Задайте квоту для кожного розділу.",
                 bg=C_BG, fg=C_MUTED, font=("Helvetica", 10),
                 justify="left").pack(anchor="w", padx=28, pady=(4, 6))

        # Прокручуваний список
        wrap = tk.Frame(page, bg=C_BG)
        wrap.pack(fill="both", expand=True, padx=28, pady=(0, 8))
        wrap.rowconfigure(0, weight=1)
        wrap.columnconfigure(0, weight=1)

        self._q1_canvas = tk.Canvas(wrap, bg=C_BG, highlightthickness=0)
        vsb = ttk.Scrollbar(wrap, orient="vertical", command=self._q1_canvas.yview)
        self._q1_canvas.configure(yscrollcommand=vsb.set)
        self._q1_canvas.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")

        self._q1_inner = tk.Frame(self._q1_canvas, bg=C_BG)
        self._q1_win = self._q1_canvas.create_window(
            (0, 0), window=self._q1_inner, anchor="nw")

        def _on_configure(e):
            self._q1_canvas.configure(scrollregion=self._q1_canvas.bbox("all"))
        def _on_canvas_resize(e):
            self._q1_canvas.itemconfig(self._q1_win, width=e.width)
        def _on_mw(e):
            delta = -e.delta if _IS_MAC else -(e.delta // 120)
            self._q1_canvas.yview_scroll(delta, "units")

        self._q1_inner.bind("<Configure>", _on_configure)
        self._q1_canvas.bind("<Configure>", _on_canvas_resize)
        self._q1_canvas.bind("<MouseWheel>", _on_mw)
        self._q1_inner.bind("<MouseWheel>", _on_mw)

        # Нижня панель
        bot = tk.Frame(page, bg=C_BG)
        bot.pack(fill="x", padx=28, pady=(0, 12))

        self._step1_total_lbl = tk.Label(
            bot, text="", bg=C_BG, fg=C_MUTED, font=FONT_SM)
        self._step1_total_lbl.pack(side="left")

        accent_btn(bot, "Далі  →", self._from_step1_to_2).pack(side="right")

    def _fill_quota_list(self):
        inner = self._q1_inner
        for w in inner.winfo_children():
            w.destroy()
        self._quota_vars.clear()
        self._avail.clear()
        self._section_rows.clear()
        self._module_expanded.clear()

        dp = self.app.db_path.get()
        cat_id = self._current_cat_id
        cat_code = self._cat_var.get()

        try:
            sections = db.get_sections_by_category(cat_id, dp)
        except Exception as e:
            tk.Label(inner, text=f"Помилка завантаження: {e}",
                     bg=C_BG, fg=C_DANGER, font=FONT_SM).pack(pady=20)
            return

        if not sections:
            tk.Label(inner,
                     text=f"Немає питань для категорії {cat_code}.\n"
                          "Спочатку імпортуйте модулі на вкладці «Модулі».",
                     bg=C_BG, fg=C_WARNING, font=FONT_SM,
                     justify="left").pack(pady=20, padx=4)
            return

        # Групуємо по модулях
        modules = {}
        for row in sections:
            mid = row["module_id"]
            if mid not in modules:
                modules[mid] = {
                    "code": row["module_code"],
                    "name": row["module_name"],
                    "sections": [],
                }
            modules[mid]["sections"].append(row)

        def _bind_mw(widget):
            widget.bind("<MouseWheel>",
                lambda e, c=self._q1_canvas: c.yview_scroll(
                    -e.delta if _IS_MAC else -(e.delta // 120), "units"))

        for mid, mdata in modules.items():
            # Заголовок модуля (кнопка розгортання)
            mod_frame = tk.Frame(inner, bg=C_BG)
            mod_frame.pack(fill="x", pady=(4, 0))

            hdr = tk.Frame(mod_frame, bg=C_SIDEBAR_H, cursor="hand2")
            hdr.pack(fill="x")
            _bind_mw(hdr)

            arrow_lbl = tk.Label(hdr, text="▶", bg=C_SIDEBAR_H, fg="white",
                                  font=("Helvetica", 10), padx=8, pady=6)
            arrow_lbl.pack(side="left")
            tk.Label(hdr,
                     text=f"  М{mdata['code']}  {mdata['name']}",
                     bg=C_SIDEBAR_H, fg="white",
                     font=("Helvetica", 11, "bold"),
                     anchor="w").pack(side="left", fill="x", expand=True,
                                      pady=6)

            avail_sum = sum(s["available"] for s in mdata["sections"])
            tk.Label(hdr, text=f"{avail_sum} пит.",
                     bg=C_SIDEBAR_H, fg="#A8C8E8",
                     font=("Helvetica", 10), padx=12).pack(side="right")
            _bind_mw(hdr)

            # Тіло модуля (розділи)
            body = tk.Frame(mod_frame, bg=C_CARD)
            self._module_expanded[mid] = False

            def _toggle(e=None, b=body, a=arrow_lbl, m=mid):
                if self._module_expanded[m]:
                    b.pack_forget()
                    a.configure(text="▶")
                    self._module_expanded[m] = False
                else:
                    b.pack(fill="x")
                    a.configure(text="▼")
                    self._module_expanded[m] = True
                self._q1_inner.update_idletasks()
                self._q1_canvas.configure(
                    scrollregion=self._q1_canvas.bbox("all"))

            for w in hdr.winfo_children():
                w.bind("<Button-1>", _toggle)
            hdr.bind("<Button-1>", _toggle)

            # Заголовок таблиці розділів
            col_hdr = tk.Frame(body, bg="#EEF2F7")
            col_hdr.pack(fill="x")
            for txt, w, side in [
                ("Розділ", 10, "left"),
                ("Назва", 0, "left"),
                ("В БД", 6, "right"),
                ("Квота", 9, "right"),
            ]:
                tk.Label(col_hdr, text=txt, width=w if w else None,
                         bg="#EEF2F7", fg=C_SIDEBAR,
                         font=("Helvetica", 10, "bold"),
                         anchor="w" if side == "left" else "e",
                         padx=4, pady=4).pack(
                    side="left" if side == "left" else "right")

            for i, sec in enumerate(mdata["sections"]):
                sec_id = sec["id"]
                avail = sec["available"]
                saved = min(sec["saved_quota"], avail)

                self._avail[sec_id] = avail
                var = tk.IntVar(value=saved)
                self._quota_vars[sec_id] = var

                bg = C_CARD if i % 2 == 0 else C_ROW_ALT
                row = tk.Frame(body, bg=bg)
                row.pack(fill="x")
                _bind_mw(row)

                tk.Label(row, text=sec["section_code"], width=10,
                         bg=bg, fg=C_MUTED, font=FONT_SM,
                         anchor="w", padx=4, pady=4).pack(side="left")

                name_txt = sec["section_name"][:32]
                tk.Label(row, text=name_txt,
                         bg=bg, fg=C_TEXT, font=FONT_SM,
                         anchor="w").pack(side="left", fill="x", expand=True,
                                          padx=4)

                avail_color = C_SUCCESS if avail > 0 else C_DANGER
                tk.Label(row, text=str(avail), width=5,
                         bg=bg, fg=avail_color,
                         font=("Helvetica", 11, "bold"),
                         anchor="e").pack(side="right", padx=(0, 8))

                cw = CounterWidget(row, var,
                                   min_val=0, max_val=avail,
                                   font_size=11, bg=bg)
                cw.pack(side="right", padx=(0, 8), pady=2)
                _bind_mw(cw)

                self._section_rows[sec_id] = (row, cw)

                var.trace_add("write", lambda *_, v=var, a=avail, lbl=avail_color:
                              self._update_quota_total())

            self._update_quota_total()

    def _update_quota_total(self):
        total = sum(v.get() for v in self._quota_vars.values())
        if hasattr(self, "_step1_total_lbl"):
            self._step1_total_lbl.configure(
                text=f"Всього питань у варіанті: {total}")

    def _from_step1_to_2(self):
        total = sum(v.get() for v in self._quota_vars.values())
        if total == 0:
            messagebox.showwarning("Увага",
                "Задайте квоту хоча б для одного розділу.")
            return
        # Зберігаємо квоти у БД
        dp = self.app.db_path.get()
        cat_id = self._current_cat_id
        for sec_id, var in self._quota_vars.items():
            db.insert_quota(sec_id, cat_id, var.get(), dp)
        self._go_to(2)

    # ── Крок 2: Кількість варіантів ───────────────────────────────────────────

    def _build_step2(self, page):
        wrap = tk.Frame(page, bg=C_BG)
        wrap.place(relx=0.5, rely=0.5, anchor="center")

        tk.Label(wrap, text="Скільки варіантів згенерувати?",
                 bg=C_BG, fg=C_SIDEBAR, font=("Helvetica", 16, "bold")).pack(pady=(0, 6))
        tk.Label(wrap, text="Кожен варіант — унікальна вибірка питань",
                 bg=C_BG, fg=C_MUTED, font=("Helvetica", 11)).pack(pady=(0, 28))

        cw = CounterWidget(wrap, self._variants_var,
                           min_val=1, max_val=50,
                           btn_size=20, font_size=22,
                           bg=C_BG)
        cw.pack(pady=8)

        tk.Label(wrap, text="варіант(ів)",
                 bg=C_BG, fg=C_MUTED, font=("Helvetica", 11)).pack(pady=(4, 32))

        btn_row = tk.Frame(wrap, bg=C_BG)
        btn_row.pack()
        ghost_btn(btn_row, "← Назад", lambda: self._go_to(1)).pack(side="left", padx=8)
        accent_btn(btn_row, "Далі  →", lambda: self._go_to(3)).pack(side="left", padx=8)

    # ── Крок 3: Вибір папки ───────────────────────────────────────────────────

    def _build_step3(self, page):
        wrap = tk.Frame(page, bg=C_BG)
        wrap.place(relx=0.5, rely=0.5, anchor="center")

        tk.Label(wrap, text="📁", bg=C_BG,
                 font=("Helvetica", 40)).pack(pady=(0, 8))
        tk.Label(wrap, text="Оберіть папку для збереження",
                 bg=C_BG, fg=C_SIDEBAR, font=("Helvetica", 16, "bold")).pack()
        tk.Label(wrap, text="Тести та ключі відповідей будуть збережені в цю папку",
                 bg=C_BG, fg=C_MUTED, font=("Helvetica", 11)).pack(pady=(2, 20))

        dir_row = tk.Frame(wrap, bg=C_BG)
        dir_row.pack(fill="x", pady=(0, 6))

        dir_entry = tk.Entry(dir_row, textvariable=self.app.output_dir,
                             font=FONT_SM, bg="white", fg=C_TEXT,
                             relief="flat", bd=1,
                             highlightthickness=1,
                             highlightbackground=C_BORDER,
                             width=46)
        dir_entry.pack(side="left", padx=(0, 6))
        ghost_btn(dir_row, "📂", self._browse_output).pack(side="left")

        self._step3_info = tk.Label(wrap, text="",
                                    bg=C_BG, fg=C_MUTED,
                                    font=("Helvetica", 10))
        self._step3_info.pack(pady=(4, 28))

        btn_row = tk.Frame(wrap, bg=C_BG)
        btn_row.pack()
        ghost_btn(btn_row, "← Назад", lambda: self._go_to(2)).pack(side="left", padx=8)
        accent_btn(btn_row, "🚀  Генерувати", self._start_generate).pack(side="left", padx=8)

    def _browse_output(self):
        d = filedialog.askdirectory(
            title="Оберіть папку для збереження",
            initialdir=self.app.output_dir.get())
        if d:
            self.app.output_dir.set(d)

    # ── Крок 4: Генерація та результат ───────────────────────────────────────

    def _build_step4(self, page):
        self._step4_page = page

        wrap = tk.Frame(page, bg=C_BG)
        wrap.place(relx=0.5, rely=0.5, anchor="center")
        self._step4_wrap = wrap

        self._step4_icon = tk.Label(wrap, text="⚙️", bg=C_BG,
                                    font=("Helvetica", 44))
        self._step4_icon.pack(pady=(0, 10))

        self._step4_title = tk.Label(wrap, text="Генерація…",
                                     bg=C_BG, fg=C_SIDEBAR,
                                     font=("Helvetica", 16, "bold"))
        self._step4_title.pack()

        self._step4_sub = tk.Label(wrap, text="",
                                   bg=C_BG, fg=C_MUTED,
                                   font=("Helvetica", 11))
        self._step4_sub.pack(pady=(4, 16))

        style = ttk.Style()
        style.configure("Wiz.Horizontal.TProgressbar",
                        troughcolor=C_BORDER, background=C_ACCENT,
                        bordercolor=C_BORDER, lightcolor=C_ACCENT,
                        darkcolor=C_ACCENT, thickness=10)
        self._step4_pb_var = tk.DoubleVar(value=0)
        self._step4_pb = ttk.Progressbar(wrap, variable=self._step4_pb_var,
                                         maximum=100, length=340,
                                         style="Wiz.Horizontal.TProgressbar")
        self._step4_pb.pack(fill="x", pady=(0, 20))

        self._step4_btns = tk.Frame(wrap, bg=C_BG)
        self._step4_btns.pack()
        # Кнопки з'являться після завершення

    def _start_generate(self):
        out = self.app.output_dir.get().strip()
        if not out:
            messagebox.showwarning("Увага", "Оберіть папку для збереження.")
            return
        os.makedirs(out, exist_ok=True)

        self._go_to(4)
        self._step4_icon.configure(text="⚙️")
        self._step4_title.configure(text="Генерація…", fg=C_SIDEBAR)
        self._step4_sub.configure(text="Зачекайте, будь ласка")
        self._step4_pb_var.set(0)
        for w in self._step4_btns.winfo_children():
            w.destroy()

        n = self._variants_var.get()
        cat_code = self._cat_var.get()
        dp = self.app.db_path.get()

        threading.Thread(
            target=self._generate_worker,
            args=(cat_code, n, out, dp),
            daemon=True,
        ).start()

    def _generate_worker(self, cat_code, n_variants, out_dir, dp):
        ok, errors = [], []

        for i in range(n_variants):
            label = f"Варіант №{i + 1}"
            self.after(0, lambda lbl=label, idx=i: (
                self._step4_sub.configure(
                    text=f"Формую {idx + 1}/{n_variants}: {lbl}"),
                self._step4_pb_var.set(idx / n_variants * 100),
                self.app.set_status(f"Генерація {idx + 1}/{n_variants}"),
            ))
            try:
                variant = gen.generate_test(cat_code, label, db_path=dp)
                ex_docx.export_student_docx(variant, out_dir)
                ex_docx.export_answer_key_docx(variant, out_dir)
                ok.append(label)
            except Exception as e:
                errors.append(f"{label}: {e}")

        def _finish():
            self._step4_pb_var.set(100)
            if not errors:
                self._step4_icon.configure(text="✅")
                self._step4_title.configure(text="Готово!", fg=C_SUCCESS)
                self._step4_sub.configure(
                    text=f"Згенеровано {len(ok)} варіант(ів) у папці:\n{out_dir}")
                self.app.set_status(f"Готово: {len(ok)} варіант(ів)", C_SUCCESS)
            else:
                self._step4_icon.configure(text="⚠️")
                self._step4_title.configure(text="Завершено з помилками",
                                             fg=C_WARNING)
                err_txt = "\n".join(errors[:3])
                if len(errors) > 3:
                    err_txt += f"\n… ще {len(errors)-3} помилок"
                self._step4_sub.configure(
                    text=f"✓ {len(ok)} варіант(ів)  ✗ {len(errors)} помилок\n{err_txt}")
                self.app.set_status(f"Готово: {len(ok)} ok, {len(errors)} помилок",
                                    C_WARNING)

            for w in self._step4_btns.winfo_children():
                w.destroy()

            def _open():
                subprocess.Popen(["open", out_dir])

            accent_btn(self._step4_btns, "📂  Відкрити папку", _open).pack(
                side="left", padx=8)
            ghost_btn(self._step4_btns, "↩  Створити ще",
                      lambda: self._go_to(0)).pack(side="left", padx=8)

        self.after(0, _finish)

    def on_show(self):
        pass   # нічого не робимо при перемиканні вкладки


# ─── Вкладка: Модулі ─────────────────────────────────────────────────────────

class ModulesFrame(tk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent, bg=C_BG)
        self.app = app
        self._build()

    def _build(self):
        tk.Label(self, text="Модулі та імпорт даних",
                 bg=C_BG, fg=C_SIDEBAR, font=FONT_TITLE).pack(
            anchor="w", padx=24, pady=(20, 12))

        body = tk.Frame(self, bg=C_BG)
        body.pack(fill="both", expand=True, padx=24, pady=4)
        body.columnconfigure(0, weight=3)
        body.columnconfigure(1, weight=2)
        body.rowconfigure(0, weight=1)

        # ── Ліво: таблиця завантажених модулів ───────────────────────────────
        left_outer, left_inner = card(body, "Завантажені модулі")
        left_outer.grid(row=0, column=0, sticky="nsew", padx=(0, 8), pady=(0, 16))

        style = ttk.Style()
        style.configure("Mod.Treeview",
                        font=FONT_SM, rowheight=24,
                        background=C_CARD, fieldbackground=C_CARD,
                        foreground=C_TEXT, borderwidth=0)
        style.configure("Mod.Treeview.Heading",
                        font=("Helvetica", 10, "bold"),
                        background="#EEF2F7", foreground=C_SIDEBAR,
                        relief="flat")
        style.map("Mod.Treeview",
                  background=[("selected", "#D6EAF8")],
                  foreground=[("selected", C_TEXT)])

        tree_wrap = tk.Frame(left_inner, bg=C_CARD)
        tree_wrap.pack(fill="both", expand=True, padx=12, pady=(0, 12))
        tree_wrap.rowconfigure(0, weight=1)
        tree_wrap.columnconfigure(0, weight=1)

        cols = ("mod", "name", "b11", "b13", "b2")
        self._mod_tree = ttk.Treeview(tree_wrap, columns=cols,
                                       show="headings",
                                       style="Mod.Treeview")
        for col, lbl, w, anc in [
            ("mod",  "Код",    55,  "center"),
            ("name", "Назва",  200, "w"),
            ("b11",  "B1.1",  65,  "center"),
            ("b13",  "B1.3",  65,  "center"),
            ("b2",   "B2",    65,  "center"),
        ]:
            self._mod_tree.heading(col, text=lbl)
            self._mod_tree.column(col, width=w, anchor=anc,
                                   stretch=(col == "name"))
        self._mod_tree.tag_configure("even", background=C_CARD)
        self._mod_tree.tag_configure("odd",  background=C_ROW_ALT)

        vsb = ttk.Scrollbar(tree_wrap, orient="vertical",
                             command=self._mod_tree.yview)
        self._mod_tree.configure(yscrollcommand=vsb.set)
        vsb.grid(row=0, column=1, sticky="ns")
        self._mod_tree.grid(row=0, column=0, sticky="nsew")

        refresh_row = tk.Frame(left_inner, bg=C_CARD)
        refresh_row.pack(fill="x", padx=12, pady=(0, 8))
        ghost_btn(refresh_row, "🔄  Оновити", self._refresh_modules).pack(side="left")

        # ── Право: форма імпорту + журнал ────────────────────────────────────
        right = tk.Frame(body, bg=C_BG)
        right.grid(row=0, column=1, sticky="nsew", pady=(0, 16))
        right.rowconfigure(2, weight=1)

        # Таблиця квот
        outer1, inner1 = card(right, "1.  Таблиця квот")
        outer1.grid(row=0, column=0, sticky="ew", pady=(0, 8))

        self.quota_path = tk.StringVar()
        self._file_row(inner1, self.quota_path, "Таблиця по категоріям.docx", "*.docx")
        accent_btn(inner1, "Завантажити квоти",
                   self._import_quotas).pack(fill="x", padx=16, pady=(6, 14))

        # Файл модуля
        outer2, inner2 = card(right, "2.  Файл модуля (питання)")
        outer2.grid(row=1, column=0, sticky="ew", pady=(0, 8))

        self.module_path = tk.StringVar()
        self._file_row(inner2, self.module_path, "Файл модуля (.docx)", "*.docx")

        meta = tk.Frame(inner2, bg=C_CARD)
        meta.pack(fill="x", padx=16, pady=(8, 0))

        tk.Label(meta, text="Код:", bg=C_CARD,
                 fg=C_TEXT, font=FONT_SM).grid(row=0, column=0, sticky="w")
        self.mod_code = tk.Entry(meta, width=6, font=FONT_SM,
                                 bg="#F7FAFC", relief="flat",
                                 highlightthickness=1,
                                 highlightbackground=C_BORDER)
        self.mod_code.insert(0, "1")
        self.mod_code.grid(row=0, column=1, padx=(4, 14), pady=3, sticky="w")

        tk.Label(meta, text="Назва:", bg=C_CARD,
                 fg=C_TEXT, font=FONT_SM).grid(row=0, column=2, sticky="w")
        self.mod_name = tk.Entry(meta, font=FONT_SM, bg="#F7FAFC",
                                 relief="flat", highlightthickness=1,
                                 highlightbackground=C_BORDER)
        self.mod_name.insert(0, "Математика")
        self.mod_name.grid(row=0, column=3, padx=4, pady=3, sticky="ew")
        meta.columnconfigure(3, weight=1)

        accent_btn(inner2, "Імпортувати модуль",
                   self._import_module).pack(fill="x", padx=16, pady=(10, 14))

        # Журнал
        outer3, inner3 = card(right, "Журнал")
        outer3.grid(row=2, column=0, sticky="nsew", pady=(0, 0))
        right.columnconfigure(0, weight=1)

        log_wrap = tk.Frame(inner3, bg=C_CARD)
        log_wrap.pack(fill="both", expand=True, padx=12, pady=(0, 8))
        log_wrap.rowconfigure(0, weight=1)
        log_wrap.columnconfigure(0, weight=1)

        self.log = tk.Text(log_wrap, font=FONT_MONO, bg="#F7FAFC", fg=C_TEXT,
                           relief="flat", bd=0, state="disabled",
                           highlightthickness=1, highlightbackground=C_BORDER,
                           wrap="word", height=8)
        vsb2 = ttk.Scrollbar(log_wrap, command=self.log.yview)
        self.log.configure(yscrollcommand=vsb2.set)
        self.log.grid(row=0, column=0, sticky="nsew")
        vsb2.grid(row=0, column=1, sticky="ns")

        self.log.tag_configure("ok",   foreground=C_SUCCESS)
        self.log.tag_configure("err",  foreground=C_DANGER)
        self.log.tag_configure("info", foreground=C_ACCENT)
        self.log.tag_configure("warn", foreground=C_WARNING)
        self._log("Готово до роботи.", "info")

    def _file_row(self, parent, var, placeholder, pattern):
        row = tk.Frame(parent, bg=C_CARD)
        row.pack(fill="x", padx=16, pady=(4, 0))
        tk.Entry(row, textvariable=var, font=FONT_SM,
                 bg="#F7FAFC", fg=C_TEXT, relief="flat",
                 highlightthickness=1,
                 highlightbackground=C_BORDER).pack(
            side="left", fill="x", expand=True)
        ghost_btn(row, "📂",
                  lambda: self._browse(var, placeholder, pattern)
                  ).pack(side="left", padx=(6, 0))

    def _browse(self, var, label, pattern):
        f = filedialog.askopenfilename(
            title=f"Оберіть: {label}",
            filetypes=[(label, pattern), ("Всі файли", "*.*")],
        )
        if f:
            var.set(f)

    def _log(self, msg, tag=""):
        self.log.configure(state="normal")
        self.log.insert("end", msg + "\n", tag)
        self.log.see("end")
        self.log.configure(state="disabled")

    def _refresh_modules(self):
        for item in self._mod_tree.get_children():
            self._mod_tree.delete(item)

        dp = self.app.db_path.get()
        if not os.path.exists(dp):
            return
        try:
            conn = sqlite3.connect(dp)
            conn.row_factory = sqlite3.Row
            rows = conn.execute("""
                SELECT m.code AS mc, m.name AS mn,
                    SUM(CASE WHEN cat.code='B1.1' THEN 1 ELSE 0 END) AS b11,
                    SUM(CASE WHEN cat.code='B1.3' THEN 1 ELSE 0 END) AS b13,
                    SUM(CASE WHEN cat.code='B2'   THEN 1 ELSE 0 END) AS b2
                FROM questions q
                JOIN sections  s    ON s.id   = q.section_id
                JOIN modules   m    ON m.id   = s.module_id
                JOIN categories cat ON cat.id = q.category_id
                GROUP BY m.id
                ORDER BY CAST(m.code AS INTEGER)
            """).fetchall()
            conn.close()
            for i, r in enumerate(rows):
                tag = "even" if i % 2 == 0 else "odd"
                self._mod_tree.insert("", "end", tags=(tag,), values=(
                    f"M{r['mc']}", r["mn"],
                    r["b11"] or "—",
                    r["b13"] or "—",
                    r["b2"]  or "—",
                ))
        except Exception:
            pass

    def on_show(self):
        self._refresh_modules()

    def _import_quotas(self):
        path = self.quota_path.get().strip()
        if not path or not os.path.exists(path):
            messagebox.showwarning("Увага", "Оберіть файл таблиці квот!")
            return
        dp = self.app.db_path.get()
        self._log(f"→ Завантаження квот: {os.path.basename(path)}", "info")

        def worker():
            try:
                db.init_db(dp)
                ql.load_quotas_from_docx(path, dp)
                s = db.get_db_stats(dp)
                self._log(
                    f"✓ Квоти завантажено. Розділів: {s['sections']}, квот: {s['quotas']}",
                    "ok")
                self.app.set_status("Квоти завантажено", C_SUCCESS)
                self.after(0, self._refresh_modules)
            except Exception as e:
                self._log(f"✗ Помилка: {e}", "err")
                self.app.set_status(f"Помилка: {e}", C_DANGER)

        threading.Thread(target=worker, daemon=True).start()

    def _import_module(self):
        path = self.module_path.get().strip()
        code = self.mod_code.get().strip()
        name = self.mod_name.get().strip()
        if not path or not os.path.exists(path):
            messagebox.showwarning("Увага", "Оберіть файл модуля!")
            return
        if not code or not name:
            messagebox.showwarning("Увага", "Введіть код і назву модуля!")
            return

        dp = self.app.db_path.get()
        self._log(f"→ Імпорт M{code} — {name}: {os.path.basename(path)}", "info")

        def worker():
            try:
                result = prs.parse_module_file(path, code, name, dp)
                total = sum(result.values())
                self._log(f"✓ Імпортовано {total} питань: {result}", "ok")
                self.app.set_status(f"Імпорт завершено: {total} питань", C_SUCCESS)
                self.after(0, self._refresh_modules)
            except Exception as e:
                self._log(f"✗ Помилка: {e}", "err")
                self.app.set_status(f"Помилка: {e}", C_DANGER)

        threading.Thread(target=worker, daemon=True).start()


# ─── Вкладка: Статистика ─────────────────────────────────────────────────────

class StatsFrame(tk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent, bg=C_BG)
        self.app = app
        self._build()

    def _build(self):
        hdr = tk.Frame(self, bg=C_BG)
        hdr.pack(fill="x", padx=24, pady=(20, 12))
        tk.Label(hdr, text="Статистика бази даних",
                 bg=C_BG, fg=C_SIDEBAR, font=FONT_TITLE).pack(side="left")
        ghost_btn(hdr, "🔄  Оновити", self.on_show).pack(side="right")

        self._cards_row = tk.Frame(self, bg=C_BG)
        self._cards_row.pack(fill="x", padx=24, pady=(0, 12))

        outer, inner = card(self, "Кількість питань по модулях та категоріях")
        outer.pack(fill="both", expand=True, padx=24, pady=(0, 16))

        style = ttk.Style()
        style.configure("Stats.Treeview",
                        font=FONT_SM, rowheight=26,
                        background=C_CARD, fieldbackground=C_CARD,
                        foreground=C_TEXT, borderwidth=0)
        style.configure("Stats.Treeview.Heading",
                        font=("Helvetica", 11, "bold"),
                        background="#EEF2F7", foreground=C_SIDEBAR,
                        relief="flat")
        style.map("Stats.Treeview",
                  background=[("selected", "#D6EAF8")],
                  foreground=[("selected", C_TEXT)])

        tree_wrap = tk.Frame(inner, bg=C_CARD)
        tree_wrap.pack(fill="both", expand=True, padx=16, pady=(4, 12))
        tree_wrap.rowconfigure(0, weight=1)
        tree_wrap.columnconfigure(0, weight=1)

        cols = ("module", "b11", "b13", "b2", "total")
        self.tree = ttk.Treeview(tree_wrap, columns=cols,
                                  show="headings", style="Stats.Treeview")
        for col, lbl, w, anc in [
            ("module", "Модуль",  280, "w"),
            ("b11",    "B1.1",   100, "center"),
            ("b13",    "B1.3",   100, "center"),
            ("b2",     "B2",     100, "center"),
            ("total",  "Всього", 100, "center"),
        ]:
            self.tree.heading(col, text=lbl)
            self.tree.column(col, width=w, anchor=anc, stretch=(col == "module"))

        vsb = ttk.Scrollbar(tree_wrap, orient="vertical",
                             command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        vsb.grid(row=0, column=1, sticky="ns")
        self.tree.grid(row=0, column=0, sticky="nsew")

        self.tree.tag_configure("even", background=C_CARD)
        self.tree.tag_configure("odd",  background=C_ROW_ALT)

    def on_show(self):
        for w in self._cards_row.winfo_children():
            w.destroy()

        dp = self.app.db_path.get()
        if not os.path.exists(dp):
            tk.Label(self._cards_row, text="База даних не знайдена",
                     bg=C_BG, fg=C_DANGER, font=FONT).pack(pady=10)
            return

        try:
            stats = db.get_db_stats(dp)
        except Exception:
            tk.Label(self._cards_row, text="БД не ініціалізована",
                     bg=C_BG, fg=C_DANGER, font=FONT).pack(pady=10)
            return

        try:
            conn = sqlite3.connect(dp)
            mods_with_q = conn.execute(
                "SELECT COUNT(DISTINCT m.id) FROM modules m "
                "JOIN sections s ON s.module_id=m.id "
                "JOIN questions q ON q.section_id=s.id"
            ).fetchone()[0]
            conn.close()
        except Exception:
            mods_with_q = 0

        for icon, value, label, color in [
            ("📚", stats["questions"],     "Питань",                C_ACCENT),
            ("📂", mods_with_q,            "Модулів\n(з питаннями)", C_SUCCESS),
            ("📋", stats["modules"],       "Модулів\n(у квотах)",   C_SIDEBAR),
            ("🗂",  stats["sections"],     "Розділів",              C_SIDEBAR),
            ("🔢",  stats["quotas"],       "Активних квот",         C_ACCENT),
        ]:
            c = tk.Frame(self._cards_row, bg=C_CARD,
                         relief="flat", bd=0,
                         highlightthickness=1,
                         highlightbackground=C_BORDER)
            c.pack(side="left", fill="x", expand=True, padx=(0, 8))
            tk.Label(c, text=icon, bg=C_CARD,
                     font=("Helvetica", 22)).pack(pady=(12, 2))
            tk.Label(c, text=str(value), bg=C_CARD, fg=color,
                     font=("Helvetica", 20, "bold")).pack()
            tk.Label(c, text=label, bg=C_CARD, fg=C_MUTED,
                     font=("Helvetica", 10), justify="center").pack(pady=(0, 12))

        for item in self.tree.get_children():
            self.tree.delete(item)

        try:
            conn = sqlite3.connect(dp)
            conn.row_factory = sqlite3.Row
            rows = conn.execute("""
                SELECT m.code AS mc, m.name AS mn,
                    SUM(CASE WHEN cat.code='B1.1' THEN 1 ELSE 0 END) AS b11,
                    SUM(CASE WHEN cat.code='B1.3' THEN 1 ELSE 0 END) AS b13,
                    SUM(CASE WHEN cat.code='B2'   THEN 1 ELSE 0 END) AS b2,
                    COUNT(*) AS total
                FROM questions q
                JOIN sections  s   ON s.id   = q.section_id
                JOIN modules   m   ON m.id   = s.module_id
                JOIN categories cat ON cat.id = q.category_id
                GROUP BY m.id
                ORDER BY CAST(m.code AS INTEGER)
            """).fetchall()
            conn.close()

            for i, r in enumerate(rows):
                tag = "even" if i % 2 == 0 else "odd"
                self.tree.insert("", "end", tags=(tag,), values=(
                    f"M{r['mc']}   {r['mn']}",
                    r["b11"] or "—",
                    r["b13"] or "—",
                    r["b2"]  or "—",
                    r["total"],
                ))
        except Exception:
            pass


# ─── Вкладка: Налаштування ───────────────────────────────────────────────────

class SettingsFrame(tk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent, bg=C_BG)
        self.app = app
        self._build()

    def _build(self):
        tk.Label(self, text="Налаштування",
                 bg=C_BG, fg=C_SIDEBAR, font=FONT_TITLE).pack(
            anchor="w", padx=24, pady=(20, 12))

        wrap = tk.Frame(self, bg=C_BG)
        wrap.pack(fill="x", padx=24)

        outer1, inner1 = card(wrap, "База даних (SQLite)")
        outer1.pack(fill="x", pady=(0, 12))

        db_row = tk.Frame(inner1, bg=C_CARD)
        db_row.pack(fill="x", padx=16, pady=(4, 0))
        tk.Entry(db_row, textvariable=self.app.db_path, font=FONT_SM,
                 bg="#F7FAFC", relief="flat", highlightthickness=1,
                 highlightbackground=C_BORDER).pack(
            side="left", fill="x", expand=True)
        ghost_btn(db_row, "📂", self._browse_db).pack(side="left", padx=(6, 0))

        btn_row = tk.Frame(inner1, bg=C_CARD)
        btn_row.pack(fill="x", padx=16, pady=(8, 14))

        self._db_status = tk.Label(btn_row, text="", bg=C_CARD,
                                    fg=C_MUTED, font=FONT_SM)

        def init_and_report():
            try:
                db.init_db(self.app.db_path.get())
                self._db_status.configure(text="✓ БД ініціалізована", fg=C_SUCCESS)
            except Exception as e:
                self._db_status.configure(text=f"✗ {e}", fg=C_DANGER)

        accent_btn(btn_row, "Ініціалізувати БД",
                   init_and_report).pack(side="left")
        self._db_status.pack(side="left", padx=12)

        outer2, inner2 = card(wrap, "Папка для збереження тестів")
        outer2.pack(fill="x", pady=(0, 12))

        dir_row = tk.Frame(inner2, bg=C_CARD)
        dir_row.pack(fill="x", padx=16, pady=(4, 14))
        tk.Entry(dir_row, textvariable=self.app.output_dir, font=FONT_SM,
                 bg="#F7FAFC", relief="flat", highlightthickness=1,
                 highlightbackground=C_BORDER).pack(
            side="left", fill="x", expand=True)
        ghost_btn(dir_row, "📂",
                  lambda: self._browse_dir(self.app.output_dir)
                  ).pack(side="left", padx=(6, 0))

        outer3, inner3 = card(wrap, "Про програму")
        outer3.pack(fill="x")
        tk.Label(
            inner3,
            text=(
                "Система автоматизованого формування тестових завдань\n"
                "для навчального центру технічного обслуговування повітряних суден\n"
                "відповідно до вимог EASA Part 147.\n\n"
                "Підтримувані категорії:  B1.1  ·  B1.3  ·  B2\n"
                "Формати виводу:  DOCX"
            ),
            bg=C_CARD, fg=C_TEXT, font=FONT_SM, justify="left",
        ).pack(anchor="w", padx=16, pady=(4, 16))

    def _browse_db(self):
        f = filedialog.askopenfilename(
            title="Оберіть файл бази даних",
            filetypes=[("SQLite DB", "*.db"), ("Всі файли", "*.*")],
        )
        if f:
            self.app.db_path.set(f)

    def _browse_dir(self, var):
        d = filedialog.askdirectory(title="Оберіть папку")
        if d:
            var.set(d)


# ─── Запуск ───────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    app = App()
    app.mainloop()
