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
import seed_db as seed

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


def accent_btn(parent, text, command, width=None, big=False):
    """
    Кнопка-акцент (Frame+Label) — коректно відображає колір на macOS.
    big=True — збільшений варіант для головних CTA (кнопка «Далі», «Генерувати»).
    """
    px = 28 if big else 18
    py = 13 if big else 9
    fs = 14 if big else 12

    outer = tk.Frame(parent, bg=C_ACCENT, cursor="hand2")
    if width:
        outer.configure(width=width)

    lbl = tk.Label(outer, text=text, bg=C_ACCENT, fg="white",
                   font=("Helvetica", fs, "bold"),
                   padx=px, pady=py, cursor="hand2")
    lbl.pack()

    C_HOVER = "#1A6A8A"
    def _enter(e): outer.configure(bg=C_HOVER);  lbl.configure(bg=C_HOVER)
    def _leave(e): outer.configure(bg=C_ACCENT); lbl.configure(bg=C_ACCENT)
    def _click(e): command()

    for w in (outer, lbl):
        w.bind("<Enter>",    _enter)
        w.bind("<Leave>",    _leave)
        w.bind("<Button-1>", _click)

    return outer


def ghost_btn(parent, text, command, width=None):
    """Вторинна кнопка (Frame+Label)."""
    outer = tk.Frame(parent, bg=C_BORDER, cursor="hand2")
    if width:
        outer.configure(width=width)

    lbl = tk.Label(outer, text=text, bg=C_BORDER, fg=C_TEXT,
                   font=FONT_SM, padx=12, pady=8, cursor="hand2")
    lbl.pack()

    C_HOVER = "#C8D0DA"
    def _enter(e): outer.configure(bg=C_HOVER); lbl.configure(bg=C_HOVER)
    def _leave(e): outer.configure(bg=C_BORDER); lbl.configure(bg=C_BORDER)
    def _click(e): command()

    for w in (outer, lbl):
        w.bind("<Enter>",    _enter)
        w.bind("<Leave>",    _leave)
        w.bind("<Button-1>", _click)

    return outer


# ─── CounterWidget ────────────────────────────────────────────────────────────

class CounterWidget(tk.Frame):
    """
    Компонент «−  [поле вводу]  +» з прив'язаною IntVar.
    Підтримує: кнопки +/−, пряме введення з клавіатури, Enter/Tab для підтвердження.
    """

    def __init__(self, parent, var: tk.IntVar,
                 min_val: int = 0, max_val: int = 999,
                 btn_size: int = 28, font_size: int = 13,
                 bg=C_CARD, **kw):
        super().__init__(parent, bg=bg, **kw)
        self._var = var
        self._min = min_val
        self._max = max_val
        self._fs  = font_size

        # Кнопка «−»
        self._btn_minus = tk.Label(
            self, text="−", bg="#E8EDF2", fg=C_TEXT,
            font=("Helvetica", font_size, "bold"),
            width=2, cursor="hand2", relief="flat",
        )
        self._btn_minus.pack(side="left")

        # Поле введення (замість Label — Entry)
        self._entry_var = tk.StringVar(value=str(var.get()))
        self._entry = tk.Entry(
            self,
            textvariable=self._entry_var,
            font=("Helvetica", font_size, "bold"),
            width=4,
            justify="center",
            bg=bg,
            fg=C_TEXT,
            relief="flat",
            highlightthickness=1,
            highlightbackground=C_BORDER,
            highlightcolor=C_ACCENT,
            insertbackground=C_TEXT,
        )
        self._entry.pack(side="left", padx=4)

        # Кнопка «+»
        self._btn_plus = tk.Label(
            self, text="+", bg=C_ACCENT, fg="white",
            font=("Helvetica", font_size, "bold"),
            width=2, cursor="hand2", relief="flat",
        )
        self._btn_plus.pack(side="left")

        # Прив'язки кнопок
        self._btn_minus.bind("<Button-1>", self._dec)
        self._btn_plus.bind("<Button-1>",  self._inc)

        # Клавіатурне введення: підтверджуємо по Enter / Tab / втраті фокуса
        self._entry.bind("<Return>",    self._commit)
        self._entry.bind("<Tab>",       self._commit)
        self._entry.bind("<FocusOut>",  self._commit)
        # Стрілки вгору/вниз у полі
        self._entry.bind("<Up>",        lambda e: self._inc())
        self._entry.bind("<Down>",      lambda e: self._dec())
        # Дозволяємо лише цифри
        vcmd = (self.register(self._validate_char), "%S")
        self._entry.configure(validate="key", validatecommand=vcmd)

        # Синхронізація IntVar → поле (якщо змінили var ззовні)
        self._var.trace_add("write", self._sync_from_var)
        self._update_states()

    # ── Валідація та commit ────────────────────────────────────────────────

    @staticmethod
    def _validate_char(char: str) -> bool:
        """Дозволяє вводити тільки цифри."""
        return char.isdigit()

    def _commit(self, _=None):
        """Читає текст поля, затискає в [min, max] і оновлює IntVar."""
        raw = self._entry_var.get().strip()
        try:
            val = int(raw)
        except ValueError:
            val = self._min
        val = max(self._min, min(self._max, val))
        # Оновлюємо без рекурсії: тимчасово знімаємо trace
        self._var.set(val)
        self._entry_var.set(str(val))
        self._update_states()

    def _sync_from_var(self, *_):
        """Синхронізує поле при зміні IntVar (напр. з +/−)."""
        try:
            self._entry_var.set(str(self._var.get()))
        except tk.TclError:
            pass
        self._update_states()

    # ── +/− ───────────────────────────────────────────────────────────────

    def _dec(self, _=None):
        v = self._var.get()
        if v > self._min:
            self._var.set(v - 1)

    def _inc(self, _=None):
        v = self._var.get()
        if v < self._max:
            self._var.set(v + 1)

    # ── Стан кнопок ───────────────────────────────────────────────────────

    def _update_states(self, *_):
        try:
            v = self._var.get()
        except tk.TclError:
            return
        self._btn_minus.configure(
            bg="#C8D0DA" if v <= self._min else "#E8EDF2",
            fg=C_MUTED   if v <= self._min else C_TEXT,
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

        # Авто-ініціалізація БД при першому запуску
        self._auto_seed()

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
            ("settings", "📖  Інструкція"),
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
        # Скидаємо глобальний скрол — on_show кожної вкладки встановить свій
        try:
            self.unbind_all("<MouseWheel>")
        except Exception:
            pass
        self._frames[key].tkraise()
        if hasattr(self._frames[key], "on_show"):
            self._frames[key].on_show()

    def set_status(self, text, color=C_MUTED):
        self._status_lbl.configure(text=text, fg=color)
        self.update_idletasks()

    def _auto_seed(self):
        """Автоматично ініціалізує БД нормами при першому запуску."""
        dp = self.db_path.get()
        if seed.needs_seeding(dp):
            ok = seed.auto_seed_if_needed(dp)
            if not ok:
                # Файл таблиці норм не знайдено — повідомимо користувача пізніше
                # через вкладку Модулі (не блокуємо запуск)
                pass


# ─── Вкладка: Генерація (5-крокова форма-wizard) ─────────────────────────────

class GenerateFrame(tk.Frame):
    STEPS = 4   # 0..4

    def __init__(self, parent, app):
        super().__init__(parent, bg=C_BG)
        self.app = app
        self._step = 0

        # Змінні вибору
        self._cat_var       = tk.StringVar(value="TB1.1")
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
        steps = ["Категорія", "Норми", "Варіанти", "Папка", "Генерація"]
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
        tk.Label(wrap, text="EASA Part 147 — TB1.1 / TB1.3 / TB2",
                 bg=C_BG, fg=C_MUTED, font=("Helvetica", 11)).pack(pady=(2, 24))

        cards_row = tk.Frame(wrap, bg=C_BG)
        cards_row.pack()

        cat_info = {
            "TB1.1": ("Авіаційний механік\nгазотурбінний літак", "⚙️"),
            "TB1.3": ("Авіаційний механік\nгазотурбінний вертольот", "🔧"),
            "TB2":   ("Авіаційний механік\nавіоніка", "📡"),
        }

        for cat, (desc, icon) in cat_info.items():
            c = tk.Frame(cards_row, bg=C_CARD,
                         highlightthickness=3,
                         highlightbackground=C_BORDER,
                         cursor="hand2",
                         padx=32, pady=28)
            c.pack(side="left", padx=14)

            tk.Label(c, text=icon, bg=C_CARD,
                     font=("Helvetica", 40)).pack(pady=(0, 4))
            tk.Label(c, text=cat, bg=C_CARD, fg=C_SIDEBAR,
                     font=("Helvetica", 18, "bold")).pack(pady=(0, 6))
            tk.Label(c, text=desc, bg=C_CARD, fg=C_MUTED,
                     font=("Helvetica", 11), justify="center").pack()

            def _select(e=None, cv=cat, frame=c):
                self._cat_var.set(cv)
                for child in cards_row.winfo_children():
                    child.configure(highlightbackground=C_BORDER, bg=C_CARD)
                    for w in child.winfo_children():
                        w.configure(bg=C_CARD)
                frame.configure(highlightbackground=C_ACCENT, bg="#EBF5FB")
                for w in frame.winfo_children():
                    w.configure(bg="#EBF5FB")

            for w in [c] + c.winfo_children():
                w.bind("<Button-1>", _select)

        self._cat_cards_row = cards_row

        self._step0_initialized = False

        accent_btn(wrap, "Далі  →", self._from_step0_to_1, big=True).pack(pady=(36, 0))

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
                f"Категорія {cat} ще не завантажена.\nСпочатку завантажте норми на вкладці «Модулі».")
            return

        self._current_cat_id = cat_id
        self._fill_quota_list()
        self._go_to(1)

    # ── Крок 1: Норми за розділами ────────────────────────────────────────────

    def _build_step1(self, page):
        hdr = tk.Frame(page, bg=C_BG)
        hdr.pack(fill="x", padx=28, pady=(14, 0))

        ghost_btn(hdr, "← Назад", lambda: self._go_to(0)).pack(side="left")
        tk.Label(hdr, text="Норми питань за розділами",
                 bg=C_BG, fg=C_SIDEBAR, font=FONT_TITLE).pack(side="left", padx=16)

        # Легенда
        tk.Label(page,
                 text="Натисніть на модуль, щоб розгорнути розділи. "
                      "«В БД» — доступних питань. Задайте норму для кожного розділу.",
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
        self._q1_scroll_fn = _on_mw

        # Нижня панель
        bot = tk.Frame(page, bg=C_BG)
        bot.pack(fill="x", padx=28, pady=(0, 12))

        self._step1_total_lbl = tk.Label(
            bot, text="", bg=C_BG, fg=C_MUTED, font=FONT_SM)
        self._step1_total_lbl.pack(side="left")

        accent_btn(bot, "Далі  →", self._from_step1_to_2, big=True).pack(side="right")

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
            # Використовуємо нову функцію — лише розділи з quota>0 для категорії
            sections = db.get_sections_for_category_with_quotas(cat_id, dp)
        except Exception as e:
            tk.Label(inner, text=f"Помилка завантаження: {e}",
                     bg=C_BG, fg=C_DANGER, font=FONT_SM).pack(pady=20)
            return

        if not sections:
            tk.Label(inner,
                     text=f"Норми для категорії {cat_code} не завантажені.\n"
                          "Спочатку завантажте таблицю норм на вкладці «Модулі».",
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
                ("Норма", 5, "right"),
                ("В БД", 6, "right"),
                ("Задати", 9, "right"),
            ]:
                tk.Label(col_hdr, text=txt, width=w if w else None,
                         bg="#EEF2F7", fg=C_SIDEBAR,
                         font=("Helvetica", 10, "bold"),
                         anchor="w" if side == "left" else "e",
                         padx=4, pady=4).pack(
                    side="left" if side == "left" else "right")

            for i, sec in enumerate(mdata["sections"]):
                sec_id = sec["id"]
                avail  = sec["available"]   # питань у БД (може бути 0)
                quota  = sec["quota"]        # норма з таблиці (завжди > 0)

                self._avail[sec_id] = avail
                # Початкове значення лічильника = норма з таблиці (але не більше доступних)
                init_val = min(quota, avail) if avail > 0 else quota
                var = tk.IntVar(value=init_val)
                self._quota_vars[sec_id] = var

                bg = C_CARD if i % 2 == 0 else C_ROW_ALT
                row = tk.Frame(body, bg=bg)
                row.pack(fill="x")
                _bind_mw(row)

                tk.Label(row, text=sec["section_code"], width=10,
                         bg=bg, fg=C_MUTED, font=FONT_SM,
                         anchor="w", padx=4, pady=4).pack(side="left")

                tk.Label(row, text=sec["section_name"],
                         bg=bg, fg=C_TEXT, font=FONT_SM,
                         anchor="w", wraplength=480, justify="left").pack(
                    side="left", fill="x", expand=True, padx=4)

                # Стовпець "Норма" (з таблиці)
                tk.Label(row, text=str(quota), width=5,
                         bg=bg, fg=C_MUTED,
                         font=("Helvetica", 10),
                         anchor="e").pack(side="right", padx=(0, 4))

                # Стовпець "В БД" — зелений якщо є, червоний якщо немає
                avail_color = C_SUCCESS if avail > 0 else C_DANGER
                tk.Label(row, text=str(avail), width=5,
                         bg=bg, fg=avail_color,
                         font=("Helvetica", 11, "bold"),
                         anchor="e").pack(side="right", padx=(0, 8))

                cw = CounterWidget(row, var,
                                   min_val=0, max_val=max(avail, quota),
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
                "Задайте норму хоча б для одного розділу.")
            return
        # Зберігаємо норми у БД
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
        accent_btn(btn_row, "Далі  →", lambda: self._go_to(3), big=True).pack(side="left", padx=8)

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
        accent_btn(btn_row, "🚀  Генерувати", self._start_generate, big=True).pack(side="left", padx=8)

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

            accent_btn(self._step4_btns, "📂  Відкрити папку", _open,
                       big=True).pack(side="left", padx=8)
            ghost_btn(self._step4_btns, "↩  Створити ще",
                      lambda: self._go_to(0)).pack(side="left", padx=8)

        self.after(0, _finish)

    def on_show(self):
        # Активуємо скрол для вкладки генерації
        if hasattr(self, "_q1_scroll_fn"):
            self.bind_all("<MouseWheel>", self._q1_scroll_fn)


# ─── Вкладка: Модулі ─────────────────────────────────────────────────────────

class ModulesFrame(tk.Frame):
    """
    Вкладка «Модулі».

    Зверху — горизонтальна сітка плиток модулів (4 колонки):
      🟢 зелена рамка + кількість питань  — модуль завантажено
      ⚪ сіра рамка                        — модуль не завантажено
    Клік на плитку — виділяє модуль і відкриває панель завантаження знизу.

    Знизу — два блоки: таблиця норм | завантаження обраного модуля.
    """

    _DEFAULT_MODULES = [
        ("1",  "Математика"),
        ("2",  "Фізика"),
        ("3",  "Основи електротехніки"),
        ("4",  "Основи електроніки"),
        ("5",  "Цифрові техніки / Системи приладів"),
        ("6",  "Матеріали та обладнання"),
        ("7",  "Практика технічного обслуговування"),
        ("8",  "Основи аеродинаміки"),
        ("9",  "Людський фактор"),
        ("10", "Авіаційне законодавство"),
        ("11A", "Аеродинаміка, конструкції та системи ПС з ГТД"),
        ("11B", "Аеродинаміка, конструкції та системи ПС з ПД"),
        ("12", "Аеродинаміка, конструкції та системи вертольотів"),
        ("13", "Конструкція та системи ПС"),
        ("14", "Двигун"),
        ("15", "Газова турбіна"),
        ("16", "Поршневий двигун"),
        ("17", "Гвинт"),
    ]
    _GRID_COLS = 4  # плиток у рядку

    def __init__(self, parent, app):
        super().__init__(parent, bg=C_BG)
        self.app = app
        self._tile_info: dict[str, dict] = {}   # code -> {frame, ...}
        self._selected_code: str | None = None
        self._build()

    # ── Побудова каркасу ──────────────────────────────────────────────────────

    def _build(self):
        tk.Label(self, text="Модулі та імпорт даних",
                 bg=C_BG, fg=C_SIDEBAR, font=FONT_TITLE).pack(
            anchor="w", padx=24, pady=(20, 10))

        # Легенда
        leg = tk.Frame(self, bg=C_BG)
        leg.pack(anchor="w", padx=24, pady=(0, 8))
        for color, text in [(C_SUCCESS, "завантажено"), ("#BDC3C7", "не завантажено")]:
            tk.Label(leg, text="●", bg=C_BG, fg=color,
                     font=("Helvetica", 14)).pack(side="left")
            tk.Label(leg, text=text + "    ", bg=C_BG, fg=C_MUTED,
                     font=("Helvetica", 10)).pack(side="left")

        # ── Прокручувана сітка плиток ─────────────────────────────────────────
        grid_outer = tk.Frame(self, bg=C_BG)
        grid_outer.pack(fill="x", padx=24, pady=(0, 12))

        grid_canvas = tk.Canvas(grid_outer, bg=C_BG, highlightthickness=0)
        grid_vsb = ttk.Scrollbar(grid_outer, orient="vertical",
                                  command=grid_canvas.yview)
        grid_canvas.configure(yscrollcommand=grid_vsb.set, height=230)
        grid_canvas.pack(side="left", fill="x", expand=True)
        grid_vsb.pack(side="right", fill="y")

        self._grid_frame = tk.Frame(grid_canvas, bg=C_BG)
        self._grid_win = grid_canvas.create_window(
            (0, 0), window=self._grid_frame, anchor="nw")

        self._grid_frame.bind("<Configure>",
            lambda e: grid_canvas.configure(
                scrollregion=grid_canvas.bbox("all")))
        grid_canvas.bind("<Configure>",
            lambda e: grid_canvas.itemconfig(self._grid_win, width=e.width))

        def _mw(e): grid_canvas.yview_scroll(
            -e.delta if _IS_MAC else -(e.delta // 120), "units")
        grid_canvas.bind("<MouseWheel>", _mw)
        self._grid_frame.bind("<MouseWheel>", _mw)
        self._grid_canvas = grid_canvas
        self._grid_scroll_fn = _mw

        # ── Нижня панель: норми | завантаження ───────────────────────────────
        bot = tk.Frame(self, bg=C_BG)
        bot.pack(fill="x", padx=24, pady=(0, 16))
        bot.columnconfigure(0, weight=1)
        bot.columnconfigure(1, weight=1)

        # Блок норм
        outer1, inner1 = card(bot, "Таблиця норм  (завантажується один раз)")
        outer1.grid(row=0, column=0, sticky="nsew", padx=(0, 8))

        self.quota_path = tk.StringVar()
        self._file_row(inner1, self.quota_path, "Таблиця по категоріям.docx", "*.docx")

        self._norm_status_lbl = tk.Label(inner1, text="",
                                          bg=C_CARD, fg=C_MUTED,
                                          font=("Helvetica", 10))
        self._norm_status_lbl.pack(anchor="w", padx=16, pady=(4, 0))
        accent_btn(inner1, "Завантажити норми",
                   self._import_quotas).pack(fill="x", padx=16, pady=(8, 14))

        # Блок завантаження модуля
        outer2, inner2 = card(bot, "Завантаження модуля")
        outer2.grid(row=0, column=1, sticky="nsew")

        self._sel_lbl = tk.Label(inner2,
                                  text="↑ Оберіть модуль зі списку вище",
                                  bg=C_CARD, fg=C_MUTED,
                                  font=("Helvetica", 11, "italic"))
        self._sel_lbl.pack(anchor="w", padx=16, pady=(10, 4))

        self.module_path = tk.StringVar()
        self._file_row(inner2, self.module_path, "Файл модуля (.docx)", "*.docx")

        # Статус імпорту
        self._import_status_lbl = tk.Label(inner2, text="",
                                            bg=C_CARD, fg=C_MUTED,
                                            font=("Helvetica", 10))
        self._import_status_lbl.pack(anchor="w", padx=16, pady=(4, 0))

        accent_btn(inner2, "Імпортувати модуль",
                   self._import_module).pack(fill="x", padx=16, pady=(8, 14))

    # ── Сітка плиток ─────────────────────────────────────────────────────────

    def _build_module_grid(self):
        """Відмальовує плитки модулів у сітці."""
        frame = self._grid_frame
        for w in frame.winfo_children():
            w.destroy()
        self._tile_info.clear()

        dp = self.app.db_path.get()
        q_counts: dict[str, int] = {}
        module_names: dict[str, str] = {}

        if os.path.exists(dp):
            try:
                conn = sqlite3.connect(dp)
                conn.row_factory = sqlite3.Row
                rows = conn.execute("""
                    SELECT m.code, m.name, COUNT(q.id) AS cnt
                    FROM modules m
                    LEFT JOIN sections s ON s.module_id = m.id
                    LEFT JOIN questions q ON q.section_id = s.id
                    GROUP BY m.id
                    ORDER BY CAST(m.code AS INTEGER)
                """).fetchall()
                conn.close()
                for r in rows:
                    q_counts[r["code"]] = r["cnt"]
                    module_names[r["code"]] = r["name"]
            except Exception:
                pass

        display = (
            [(c, module_names[c]) for c in sorted(
                module_names.keys(),
                key=lambda x: int(x) if x.isdigit() else 99)]
            if module_names else self._DEFAULT_MODULES
        )

        # Налаштовуємо колонки сітки
        for col in range(self._GRID_COLS):
            frame.columnconfigure(col, weight=1, uniform="tile")

        for idx, (code, name) in enumerate(display):
            cnt    = q_counts.get(code, 0)
            loaded = cnt > 0
            row_i  = idx // self._GRID_COLS
            col_i  = idx % self._GRID_COLS

            brd_color = C_SUCCESS if loaded else "#BDC3C7"
            tile = tk.Frame(frame,
                            bg=C_CARD,
                            highlightthickness=2,
                            highlightbackground=brd_color,
                            cursor="hand2",
                            padx=10, pady=8)
            tile.grid(row=row_i, column=col_i,
                      sticky="nsew", padx=4, pady=4)

            # Рядок: індикатор + код
            top = tk.Frame(tile, bg=C_CARD)
            top.pack(fill="x")
            tk.Label(top, text="●", bg=C_CARD,
                     fg=brd_color,
                     font=("Helvetica", 11)).pack(side="left")
            tk.Label(top, text=f"M{code}", bg=C_CARD,
                     fg=C_SIDEBAR,
                     font=("Helvetica", 11, "bold")).pack(side="left", padx=4)
            if loaded:
                tk.Label(top, text=f"{cnt} пит.", bg=C_CARD,
                         fg=C_SUCCESS,
                         font=("Helvetica", 9)).pack(side="right")

            # Назва (обрізаємо якщо довга)
            short_name = name if len(name) <= 28 else name[:26] + "…"
            tk.Label(tile, text=short_name, bg=C_CARD,
                     fg=C_TEXT if loaded else C_MUTED,
                     font=("Helvetica", 9),
                     justify="left", anchor="w",
                     wraplength=160).pack(anchor="w", pady=(4, 0))

            self._tile_info[code] = {
                "tile": tile, "loaded": loaded,
                "name": name, "code": code,
                "brd_base": brd_color,
            }

            def _sel(e=None, c=code, n=name): self._on_select(c, n)
            def _mw(e):
                self._grid_canvas.yview_scroll(
                    -e.delta if _IS_MAC else -(e.delta // 120), "units")

            for w in [tile] + tile.winfo_children() + \
                     [gc for t in tile.winfo_children()
                      for gc in (t.winfo_children() if hasattr(t, "winfo_children") else [])]:
                try:
                    w.bind("<Button-1>", _sel)
                    w.bind("<MouseWheel>", _mw)
                except Exception:
                    pass

        # Відновлюємо виділення
        if self._selected_code and self._selected_code in self._tile_info:
            self._highlight_tile(self._selected_code)

    def _on_select(self, code: str, name: str):
        self._selected_code = code
        self._highlight_tile(code)
        self._sel_lbl.configure(
            text=f"Вибрано: M{code} — {name}", fg=C_SIDEBAR)
        self.module_path.set("")
        self._import_status_lbl.configure(text="")

    def _highlight_tile(self, code: str):
        C_SEL = "#1A6A8A"
        for c, info in self._tile_info.items():
            if c == code:
                info["tile"].configure(highlightbackground=C_SEL, bg="#EBF5FB")
                for ch in info["tile"].winfo_children():
                    try:
                        ch.configure(bg="#EBF5FB")
                        for gc in ch.winfo_children():
                            try: gc.configure(bg="#EBF5FB")
                            except Exception: pass
                    except Exception:
                        pass
            else:
                brd = info["brd_base"]
                info["tile"].configure(highlightbackground=brd, bg=C_CARD)
                for ch in info["tile"].winfo_children():
                    try:
                        ch.configure(bg=C_CARD)
                        for gc in ch.winfo_children():
                            try: gc.configure(bg=C_CARD)
                            except Exception: pass
                    except Exception:
                        pass

    # ── Допоміжні методи ─────────────────────────────────────────────────────

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

    def _update_norm_status(self):
        dp = self.app.db_path.get()
        if not os.path.exists(dp):
            self._norm_status_lbl.configure(text="БД не знайдена", fg=C_DANGER)
            return
        try:
            stats = db.get_db_stats(dp)
            if stats["quotas"] > 0:
                self._norm_status_lbl.configure(
                    text=f"✓ {stats['sections']} розділів, {stats['quotas']} норм",
                    fg=C_SUCCESS)
            else:
                self._norm_status_lbl.configure(
                    text="Норми не завантажені", fg=C_WARNING)
        except Exception:
            self._norm_status_lbl.configure(text="Помилка читання БД", fg=C_DANGER)

    def on_show(self):
        self._build_module_grid()
        self._update_norm_status()
        if hasattr(self, "_grid_scroll_fn"):
            self.bind_all("<MouseWheel>", self._grid_scroll_fn)

    # ── Операції ─────────────────────────────────────────────────────────────

    def _import_quotas(self):
        path = self.quota_path.get().strip()
        if not path or not os.path.exists(path):
            messagebox.showwarning("Увага", "Оберіть файл таблиці норм!")
            return
        dp = self.app.db_path.get()
        self.app.set_status("Завантажуємо норми…", C_ACCENT)

        def worker():
            try:
                db.init_db(dp)
                ql.load_quotas_from_docx(path, dp)
                s = db.get_db_stats(dp)
                self.app.set_status(
                    f"✓ Норми завантажено: {s['sections']} розділів", C_SUCCESS)
                self.after(0, self._build_module_grid)
                self.after(0, self._update_norm_status)
            except Exception as e:
                self.app.set_status(f"Помилка: {e}", C_DANGER)

        threading.Thread(target=worker, daemon=True).start()

    def _import_module(self):
        if not self._selected_code:
            messagebox.showwarning("Увага", "Спочатку оберіть модуль зі списку!")
            return
        path = self.module_path.get().strip()
        if not path or not os.path.exists(path):
            messagebox.showwarning("Увага", "Оберіть файл модуля (.docx)!")
            return

        code = self._selected_code
        name = self._tile_info[code]["name"] if code in self._tile_info else code
        dp   = self.app.db_path.get()

        self._import_status_lbl.configure(text="Імпортуємо…", fg=C_ACCENT)
        self.app.set_status(f"Імпорт M{code}…", C_ACCENT)

        def worker():
            try:
                result = prs.parse_module_file(path, code, name, dp)
                total  = sum(result.values())
                self.after(0, lambda: self._import_status_lbl.configure(
                    text=f"✓ Імпортовано {total} питань", fg=C_SUCCESS))
                self.app.set_status(f"Імпорт завершено: {total} питань", C_SUCCESS)
                self.after(0, self._build_module_grid)
            except Exception as e:
                self.after(0, lambda: self._import_status_lbl.configure(
                    text=f"✗ {e}", fg=C_DANGER))
                self.app.set_status(f"Помилка: {e}", C_DANGER)

        threading.Thread(target=worker, daemon=True).start()


# ─── Вкладка: Статистика ─────────────────────────────────────────────────────

class StatsFrame(tk.Frame):
    """
    Вкладка «Статистика».

    Верхній блок — три плитки готовності по категоріях (B1.1 / B1.3 / B2):
    відсоток завантажених питань відносно сумарної норми + колірний статус.

    Нижній блок — картки по кожному модулю:
    назва + горизонтальна смуга прогресу + бейджики категорій.
    """

    # Кольори рівнів готовності
    _READY_COLORS = [
        (80, C_SUCCESS,  "Готовий"),
        (40, C_WARNING,  "Частково"),
        (0,  C_DANGER,   "Не готовий"),
    ]

    def __init__(self, parent, app):
        super().__init__(parent, bg=C_BG)
        self.app = app
        self._scroll_fn = None
        self._build()

    # ── Побудова каркасу ──────────────────────────────────────────────────────

    def _build(self):
        hdr = tk.Frame(self, bg=C_BG)
        hdr.pack(fill="x", padx=24, pady=(20, 10))
        tk.Label(hdr, text="Статистика бази даних",
                 bg=C_BG, fg=C_SIDEBAR, font=FONT_TITLE).pack(side="left")
        ghost_btn(hdr, "🔄  Оновити", self.on_show).pack(side="right")

        # Плитки готовності категорій (заповнюються в on_show)
        self._cat_tiles = tk.Frame(self, bg=C_BG)
        self._cat_tiles.pack(fill="x", padx=24, pady=(0, 14))

        # Прокручувана область з картками модулів
        wrap = tk.Frame(self, bg=C_BG)
        wrap.pack(fill="both", expand=True, padx=24, pady=(0, 16))
        wrap.rowconfigure(0, weight=1)
        wrap.columnconfigure(0, weight=1)

        self._canvas = tk.Canvas(wrap, bg=C_BG, highlightthickness=0)
        vsb = ttk.Scrollbar(wrap, orient="vertical", command=self._canvas.yview)
        self._canvas.configure(yscrollcommand=vsb.set)
        self._canvas.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")

        self._mod_frame = tk.Frame(self._canvas, bg=C_BG)
        self._win_id = self._canvas.create_window(
            (0, 0), window=self._mod_frame, anchor="nw")

        self._mod_frame.bind("<Configure>",
            lambda e: self._canvas.configure(
                scrollregion=self._canvas.bbox("all")))
        self._canvas.bind("<Configure>",
            lambda e: self._canvas.itemconfig(self._win_id, width=e.width))

        def _mw(e):
            self._canvas.yview_scroll(
                -e.delta if _IS_MAC else -(e.delta // 120), "units")
        self._canvas.bind("<MouseWheel>", _mw)
        self._mod_frame.bind("<MouseWheel>", _mw)
        self._scroll_fn = _mw

    # ── Оновлення даних ───────────────────────────────────────────────────────

    def on_show(self):
        # Активуємо глобальний скрол на весь час поки вкладка відкрита
        if self._scroll_fn:
            self.bind_all("<MouseWheel>", self._scroll_fn)

        dp = self.app.db_path.get()

        # Очищення
        for w in self._cat_tiles.winfo_children():
            w.destroy()
        for w in self._mod_frame.winfo_children():
            w.destroy()

        if not os.path.exists(dp):
            tk.Label(self._cat_tiles, text="База даних не знайдена",
                     bg=C_BG, fg=C_DANGER, font=FONT).pack(pady=10)
            return

        try:
            conn = sqlite3.connect(dp)
            conn.row_factory = sqlite3.Row

            # ── Дані для плиток категорій ─────────────────────────────────
            cat_data = {}
            for cat_code in ("TB1.1", "TB1.3", "TB2"):
                row = conn.execute("""
                    SELECT
                        COALESCE(SUM(qt.count), 0) AS quota_total,
                        COALESCE(SUM(
                            MIN(qt.count,
                                (SELECT COUNT(*) FROM questions q2
                                 WHERE q2.section_id = s.id AND q2.category_id = c.id))
                        ), 0) AS q_available
                    FROM quotas qt
                    JOIN sections s  ON s.id = qt.section_id
                    JOIN categories c ON c.id = qt.category_id AND c.code = ?
                    WHERE qt.count > 0
                """, (cat_code,)).fetchone()
                quota = row["quota_total"] or 0
                avail = row["q_available"] or 0
                pct   = int(avail / quota * 100) if quota > 0 else 0
                cat_data[cat_code] = {"quota": quota, "avail": avail, "pct": pct}

            # ── Дані для карток модулів ───────────────────────────────────
            mod_rows = conn.execute("""
                SELECT m.id, m.code, m.name,
                    (SELECT COUNT(*) FROM questions q2 JOIN sections s2 ON q2.section_id=s2.id
                     JOIN categories c2 ON q2.category_id=c2.id
                     WHERE s2.module_id=m.id AND c2.code='TB1.1') AS b11,
                    (SELECT COUNT(*) FROM questions q2 JOIN sections s2 ON q2.section_id=s2.id
                     JOIN categories c2 ON q2.category_id=c2.id
                     WHERE s2.module_id=m.id AND c2.code='TB1.3') AS b13,
                    (SELECT COUNT(*) FROM questions q2 JOIN sections s2 ON q2.section_id=s2.id
                     JOIN categories c2 ON q2.category_id=c2.id
                     WHERE s2.module_id=m.id AND c2.code='TB2') AS b2,
                    (SELECT COUNT(*) FROM questions q2 JOIN sections s2 ON q2.section_id=s2.id
                     WHERE s2.module_id=m.id) AS total,
                    (SELECT COALESCE(SUM(qt2.count),0)
                     FROM quotas qt2 JOIN sections s2 ON s2.id=qt2.section_id
                     WHERE s2.module_id=m.id) AS quota_sum
                FROM modules m
                ORDER BY CAST(m.code AS INTEGER)
            """).fetchall()
            conn.close()
        except Exception as e:
            tk.Label(self._cat_tiles, text=f"Помилка читання БД: {e}",
                     bg=C_BG, fg=C_DANGER, font=FONT_SM).pack(pady=10)
            return

        # ── Відмалювати плитки категорій ─────────────────────────────────────
        self._draw_cat_tiles(cat_data)

        # ── Відмалювати картки модулів ────────────────────────────────────────
        self._draw_module_cards(mod_rows)

    def _ready_color(self, pct: int) -> tuple[str, str]:
        """Повертає (колір, текст) для відсотка готовності."""
        for threshold, color, label in self._READY_COLORS:
            if pct >= threshold:
                return color, label
        return C_DANGER, "Не готовий"

    def _draw_cat_tiles(self, cat_data: dict):
        """Три великі плитки TB1.1 / TB1.3 / TB2."""
        CAT_ICONS = {"TB1.1": "⚙️", "TB1.3": "🔧", "TB2": "📡"}
        CAT_NAMES = {
            "TB1.1": "Авіамеханік (ГТД)",
            "TB1.3": "Авіамеханік (ПД)",
            "TB2":   "Авіоніка",
        }

        for cat_code, data in cat_data.items():
            pct   = data["pct"]
            color, status = self._ready_color(pct)
            avail = data["avail"]
            quota = data["quota"]

            tile = tk.Frame(self._cat_tiles, bg=C_CARD,
                            highlightthickness=2,
                            highlightbackground=color)
            tile.pack(side="left", fill="x", expand=True, padx=(0, 10))

            # Верхній рядок: іконка + код + статус
            top = tk.Frame(tile, bg=C_CARD)
            top.pack(fill="x", padx=16, pady=(14, 4))

            tk.Label(top, text=CAT_ICONS[cat_code], bg=C_CARD,
                     font=("Helvetica", 22)).pack(side="left")
            tk.Label(top, text=cat_code, bg=C_CARD, fg=C_SIDEBAR,
                     font=("Helvetica", 18, "bold")).pack(side="left", padx=8)
            status_lbl = tk.Label(top, text=status, bg=color, fg="white",
                                   font=("Helvetica", 10, "bold"),
                                   padx=8, pady=3)
            status_lbl.pack(side="right")

            # Назва
            tk.Label(tile, text=CAT_NAMES[cat_code],
                     bg=C_CARD, fg=C_MUTED,
                     font=("Helvetica", 10)).pack(anchor="w", padx=16)

            # Прогрес-бар
            pb_bg = tk.Frame(tile, bg=C_BORDER, height=8)
            pb_bg.pack(fill="x", padx=16, pady=(10, 4))
            pb_bg.update_idletasks()

            def _draw_bar(frame=pb_bg, p=pct, c=color):
                w = frame.winfo_width()
                fill_w = max(4, int(w * p / 100))
                tk.Frame(frame, bg=c, height=8, width=fill_w).place(x=0, y=0)
            pb_bg.bind("<Configure>", lambda e, f=pb_bg, p=pct, c=color:
                       tk.Frame(f, bg=c, height=8,
                                width=max(4, int(e.width * p / 100))).place(x=0, y=0))

            # Підпис
            tk.Label(tile, text=f"{avail} з {quota} питань  ({pct}%)",
                     bg=C_CARD, fg=color,
                     font=("Helvetica", 11, "bold")).pack(
                anchor="w", padx=16, pady=(0, 14))

    def _draw_module_cards(self, mod_rows):
        """Картки-прогресбари по кожному модулю."""
        # Заголовок секції
        tk.Label(self._mod_frame,
                 text="Завантаження по модулях",
                 bg=C_BG, fg=C_SIDEBAR,
                 font=("Helvetica", 13, "bold")).pack(
            anchor="w", pady=(0, 8))

        for r in mod_rows:
            total = r["total"] or 0
            quota = r["quota_sum"] or 0
            pct   = int(total / quota * 100) if quota > 0 else 0
            color, _ = self._ready_color(pct)

            b11 = r["b11"] or 0
            b13 = r["b13"] or 0
            b2  = r["b2"]  or 0

            # Картка
            c_outer = tk.Frame(self._mod_frame, bg=C_BORDER, pady=1, padx=1)
            c_outer.pack(fill="x", pady=(0, 6))
            c_inner = tk.Frame(c_outer, bg=C_CARD, padx=14, pady=10)
            c_inner.pack(fill="both")

            # Рядок 1: код + назва + кількість питань
            row1 = tk.Frame(c_inner, bg=C_CARD)
            row1.pack(fill="x")

            tk.Label(row1, text=f"M{r['code']}", bg=C_CARD, fg=C_SIDEBAR,
                     font=("Helvetica", 11, "bold"), width=5,
                     anchor="w").pack(side="left")
            tk.Label(row1, text=r["name"], bg=C_CARD, fg=C_TEXT,
                     font=FONT_SM, anchor="w").pack(
                side="left", fill="x", expand=True)

            # Бейджики категорій (лише де є питання)
            badge_row = tk.Frame(row1, bg=C_CARD)
            badge_row.pack(side="right")
            for cat_code, cnt in [("TB1.1", b11), ("TB1.3", b13), ("TB2", b2)]:
                if cnt > 0:
                    tk.Label(badge_row, text=f"{cat_code}: {cnt}",
                             bg=C_ACCENT, fg="white",
                             font=("Helvetica", 9, "bold"),
                             padx=6, pady=2).pack(side="left", padx=2)

            # Прогрес-бар
            pb_bg = tk.Frame(c_inner, bg=C_BORDER, height=6)
            pb_bg.pack(fill="x", pady=(8, 4))
            pb_bg.bind("<Configure>",
                       lambda e, c=color, p=pct:
                       tk.Frame(e.widget, bg=c, height=6,
                                width=max(0, int(e.width * p / 100))).place(x=0, y=0))

            # Підпис прогресу
            row2 = tk.Frame(c_inner, bg=C_CARD)
            row2.pack(fill="x")
            if quota > 0:
                tk.Label(row2,
                         text=f"{total} питань завантажено  /  норма {quota}  ({pct}%)",
                         bg=C_CARD, fg=color,
                         font=("Helvetica", 10)).pack(side="left")
            else:
                tk.Label(row2,
                         text=f"{total} питань  (норма не задана)",
                         bg=C_CARD, fg=C_MUTED,
                         font=("Helvetica", 10)).pack(side="left")

            # Скрол для карток забезпечується через bind_all у on_show


# ─── Вкладка: Інструкція користувача ─────────────────────────────────────────

class SettingsFrame(tk.Frame):
    """
    Вкладка «Інструкція користувача» — красиво відформатований посібник
    по роботі з програмою, розбитий на кольорові блоки з іконками.
    """

    _SECTIONS = [
        {
            "icon": "📦",
            "title": "Крок 1 — Підготовка бази даних",
            "color": "#EBF5FB",
            "border": "#2E86AB",
            "steps": [
                ("1.1", "Перейдіть у вкладку «Модулі»"),
                ("1.2", "У блоці «Таблиця норм» оберіть файл «Таблиця по категоріям.docx» "
                        "та натисніть «Завантажити норми». Це потрібно зробити лише один раз — "
                        "норми зберігаються у базі даних і залишаються після перезапуску."),
                ("1.3", "Для кожного модуля, що використовується у вашій категорії, "
                        "оберіть модуль зі списку (клік на рядок), потім у блоці "
                        "«Завантаження модуля» вкажіть файл .docx та натисніть "
                        "«Імпортувати модуль»."),
                ("1.4", "Зелений індикатор ● біля модуля означає, що питання успішно "
                        "завантажено до бази даних."),
            ],
        },
        {
            "icon": "✈",
            "title": "Крок 2 — Генерація тестів",
            "color": "#EAFAF1",
            "border": "#27AE60",
            "steps": [
                ("2.1", "Перейдіть у вкладку «Генерація» та оберіть категорію здобувача: "
                        "TB1.1 (ГТД), TB1.3 (ПД) або TB2 (авіоніка)."),
                ("2.2", "На кроці «Норми» відображаються лише ті модулі та розділи, "
                        "що передбачені для обраної категорії. Для кожного розділу "
                        "стовпець «Норма» показує рекомендовану кількість питань "
                        "із таблиці, «В БД» — скільки питань доступно, "
                        "«Задати» — поле для введення потрібної кількості."),
                ("2.3", "На кроці «Варіанти» за допомогою лічильника вкажіть скільки "
                        "варіантів тесту потрібно згенерувати."),
                ("2.4", "На кроці «Папка» оберіть директорію для збереження файлів."),
                ("2.5", "Натисніть «Генерувати» — програма автоматично створить файли "
                        "тестів та листів відповідей для кожного варіанту."),
            ],
        },
        {
            "icon": "📄",
            "title": "Вихідні файли",
            "color": "#FEF9E7",
            "border": "#D4AC0D",
            "steps": [
                ("●", "Файл здобувача — тестові завдання з бланком відповідей "
                        "(Варіант №1, №2, …)."),
                ("●", "Файл викладача — лист відповідей: лише номер питання та правильна відповідь."),
                ("●", "Всі файли зберігаються у вказану вами папку у форматі .docx."),
            ],
        },
        {
            "icon": "📊",
            "title": "Статистика",
            "color": "#F4ECF7",
            "border": "#7D3C98",
            "steps": [
                ("●", "Вкладка «Статистика» показує кількість завантажених питань "
                        "по кожному модулю та категорії."),
                ("●", "Використовуйте її для перевірки повноти завантаження модулів "
                        "перед генерацією тестів."),
            ],
        },
        {
            "icon": "💡",
            "title": "Корисні поради",
            "color": "#FDFEFE",
            "border": "#BDC3C7",
            "steps": [
                ("●", "База даних зберігається у файлі questions.db у папці data/. "
                        "При перенесенні програми на інший комп'ютер скопіюйте разом "
                        "із нею файли data/questions.db та data/Таблиця по категоріям.docx."),
                ("●", "Якщо потрібно оновити питання модуля — просто завантажте новий "
                        "файл для того самого модуля. Старі питання замінюються новими."),
                ("●", "Категорії TB1.1 і TB1.3 використовують різні набори модулів — "
                        "на кроці «Норми» показуються тільки ті, що відповідають "
                        "обраній категорії."),
            ],
        },
    ]

    def __init__(self, parent, app):
        super().__init__(parent, bg=C_BG)
        self.app = app
        self._build()

    def _build(self):
        # Заголовок
        hdr = tk.Frame(self, bg=C_BG)
        hdr.pack(fill="x", padx=24, pady=(20, 4))
        tk.Label(hdr, text="Інструкція користувача",
                 bg=C_BG, fg=C_SIDEBAR, font=FONT_TITLE).pack(side="left")

        tk.Label(self,
                 text="Система автоматизованого формування тестів EASA Part 147  ·  "
                      "Підтримувані категорії: TB1.1  TB1.3  TB2",
                 bg=C_BG, fg=C_MUTED, font=("Helvetica", 11)).pack(
            anchor="w", padx=24, pady=(0, 12))

        # Прокручуваний контент
        outer = tk.Frame(self, bg=C_BG)
        outer.pack(fill="both", expand=True, padx=24, pady=(0, 16))
        outer.rowconfigure(0, weight=1)
        outer.columnconfigure(0, weight=1)

        canvas = tk.Canvas(outer, bg=C_BG, highlightthickness=0)
        vsb = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)
        canvas.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")

        content = tk.Frame(canvas, bg=C_BG)
        win_id = canvas.create_window((0, 0), window=content, anchor="nw")

        def _on_conf(e):
            canvas.configure(scrollregion=canvas.bbox("all"))
        def _on_canvas_conf(e):
            canvas.itemconfig(win_id, width=e.width)

        def _scroll(e):
            canvas.yview_scroll(
                -e.delta if _IS_MAC else -(e.delta // 120), "units")

        content.bind("<Configure>", _on_conf)
        canvas.bind("<Configure>", _on_canvas_conf)
        canvas.bind("<MouseWheel>", _scroll)
        content.bind("<MouseWheel>", _scroll)
        # Зберігаємо функцію скролу для передачі в _build_section
        self._instr_scroll = _scroll
        self._settings_canvas = canvas

        # Блоки інструкції
        for sec in self._SECTIONS:
            self._build_section(content, sec)

        # Нижній колонтитул
        tk.Label(content,
                 text="Розроблено: Деревецька Поліна Михайлівна  ·  "
                      "Криворізький фаховий коледж КАІ  ·  2026",
                 bg=C_BG, fg=C_MUTED,
                 font=("Helvetica", 10, "italic")).pack(
            anchor="center", pady=(16, 8))

    def on_show(self):
        if hasattr(self, "_instr_scroll"):
            self.bind_all("<MouseWheel>", self._instr_scroll)

    def _build_section(self, parent, sec: dict):
        """Малює один кольоровий блок інструкції."""
        bg  = sec["color"]
        brd = sec["border"]
        scr = getattr(self, "_instr_scroll", None)

        def _bind_scroll(w):
            if scr:
                w.bind("<MouseWheel>", scr)

        # Зовнішній фрейм з кольоровою лівою смугою
        outer = tk.Frame(parent, bg=brd, padx=3, pady=0)
        outer.pack(fill="x", pady=(0, 10))
        _bind_scroll(outer)

        inner = tk.Frame(outer, bg=bg)
        inner.pack(fill="both", expand=True)
        _bind_scroll(inner)

        # Заголовок блоку
        hdr = tk.Frame(inner, bg=bg)
        hdr.pack(fill="x", padx=14, pady=(12, 6))
        _bind_scroll(hdr)

        icon_lbl = tk.Label(hdr, text=sec["icon"], bg=bg,
                             font=("Helvetica", 20))
        icon_lbl.pack(side="left", padx=(0, 8))
        _bind_scroll(icon_lbl)

        title_lbl = tk.Label(hdr, text=sec["title"], bg=bg, fg=C_SIDEBAR,
                              font=("Helvetica", 13, "bold"))
        title_lbl.pack(side="left")
        _bind_scroll(title_lbl)

        # Кроки
        for num, text in sec["steps"]:
            row = tk.Frame(inner, bg=bg)
            row.pack(fill="x", padx=14, pady=(0, 6))
            _bind_scroll(row)

            num_lbl = tk.Label(row, text=num, bg=bg, fg=brd,
                               font=("Helvetica", 11, "bold"),
                               width=4, anchor="nw")
            num_lbl.pack(side="left", pady=2)
            _bind_scroll(num_lbl)

            txt_lbl = tk.Label(row, text=text, bg=bg, fg=C_TEXT,
                               font=("Helvetica", 11),
                               justify="left", wraplength=700,
                               anchor="nw")
            txt_lbl.pack(side="left", fill="x", expand=True, pady=2)
            _bind_scroll(txt_lbl)

        spacer = tk.Frame(inner, bg=bg, height=6)
        spacer.pack()
        _bind_scroll(spacer)


# ─── Запуск ───────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    app = App()
    app.mainloop()
