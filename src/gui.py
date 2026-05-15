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

FONT        = ("Helvetica", 12)
FONT_SM     = ("Helvetica", 11)
FONT_BOLD   = ("Helvetica", 12, "bold")
FONT_TITLE  = ("Helvetica", 18, "bold")
FONT_MONO   = ("Courier", 11)


# ─── Головне вікно ────────────────────────────────────────────────────────────

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Part 147 — Генератор тестових завдань")
        self.geometry("1150x720")
        self.minsize(960, 620)
        self.configure(bg=C_BG)

        self.db_path   = tk.StringVar(value=db.DB_PATH)
        self.output_dir = tk.StringVar(
            value=os.path.join(_ROOT, "output"))

        self._build()
        self._nav("generate")

    # ── Побудова каркасу ──────────────────────────────────────────────────────

    def _build(self):
        # Бічна панель
        self.sidebar = tk.Frame(self, bg=C_SIDEBAR, width=210)
        self.sidebar.pack(side="left", fill="y")
        self.sidebar.pack_propagate(False)

        # Логотип
        tk.Label(self.sidebar, text="✈", bg=C_SIDEBAR, fg="white",
                 font=("Helvetica", 28)).pack(pady=(28, 2))
        tk.Label(self.sidebar, text="PART 147", bg=C_SIDEBAR, fg="white",
                 font=("Helvetica", 15, "bold")).pack()
        tk.Label(self.sidebar, text="Генератор тестів", bg=C_SIDEBAR,
                 fg="#8DAFC8", font=("Helvetica", 10)).pack(pady=(0, 16))

        # Роздільник
        tk.Frame(self.sidebar, bg="#2D4F70", height=1).pack(fill="x",
                                                              padx=20, pady=4)

        # Навігаційні кнопки
        self._nav_btns = {}
        self._active_nav = None
        nav = [
            ("generate", "🎲  Генерація"),
            ("import",   "📥  Імпорт даних"),
            ("stats",    "📊  Статистика"),
            ("settings", "⚙️   Налаштування"),
        ]
        for key, label in nav:
            row = tk.Frame(self.sidebar, bg=C_SIDEBAR, cursor="hand2")
            row.pack(fill="x", pady=1)

            accent_bar = tk.Frame(row, bg=C_SIDEBAR, width=4)
            accent_bar.pack(side="left", fill="y")
            accent_bar.pack_propagate(False)

            lbl = tk.Label(
                row, text=label, anchor="w",
                bg=C_SIDEBAR, fg="#8DAFC8",
                font=("Helvetica", 12), padx=14, pady=9,
            )
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

        # Версія внизу бічної панелі
        tk.Label(self.sidebar, text="v1.0  •  EASA Part 147",
                 bg=C_SIDEBAR, fg="#5A7A9A",
                 font=("Helvetica", 9)).pack(side="bottom", pady=14)

        # Основний контент
        content_wrap = tk.Frame(self, bg=C_BG)
        content_wrap.pack(side="left", fill="both", expand=True)

        # Статус-рядок
        self.statusbar = tk.Frame(content_wrap, bg="#E2E8F0", height=28)
        self.statusbar.pack(side="bottom", fill="x")
        self.statusbar.pack_propagate(False)
        self._status_lbl = tk.Label(self.statusbar, text="Готово",
                                    bg="#E2E8F0", fg=C_MUTED, font=("Helvetica", 10),
                                    anchor="w")
        self._status_lbl.pack(side="left", padx=14)

        # Контейнер для фреймів
        self.container = tk.Frame(content_wrap, bg=C_BG)
        self.container.pack(fill="both", expand=True)

        # Фрейми екранів
        self._frames = {
            "generate": GenerateFrame(self.container, self),
            "import":   ImportFrame(self.container, self),
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


# ─── Допоміжні компоненти ─────────────────────────────────────────────────────

def card(parent, title="", **kw):
    """Картка з тінню і заголовком."""
    outer = tk.Frame(parent, bg=C_BORDER, padx=1, pady=1)
    inner = tk.Frame(outer, bg=C_CARD, **kw)
    inner.pack(fill="both", expand=True)
    if title:
        tk.Label(inner, text=title, bg=C_CARD, fg=C_SIDEBAR,
                 font=FONT_BOLD).pack(anchor="w", padx=16, pady=(12, 4))
    return outer, inner


def section_label(parent, text):
    tk.Label(parent, text=text, bg=C_BG, fg=C_MUTED,
             font=("Helvetica", 10, "bold")).pack(anchor="w", pady=(8, 2))


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


# ─── Вкладка: Генерація ───────────────────────────────────────────────────────

class GenerateFrame(tk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent, bg=C_BG)
        self.app = app
        self._quota_vars = {}   # (section_id, cat_code) -> tk.IntVar
        self._avail      = {}   # (section_id, cat_code) -> int
        self._build()

    def _build(self):
        # Заголовок
        hdr = tk.Frame(self, bg=C_BG)
        hdr.pack(fill="x", padx=24, pady=(20, 10))
        tk.Label(hdr, text="Генерація тестових завдань",
                 bg=C_BG, fg=C_SIDEBAR, font=FONT_TITLE).pack(side="left")

        # Два стовпці
        body = tk.Frame(self, bg=C_BG)
        body.pack(fill="both", expand=True, padx=24, pady=4)
        body.columnconfigure(0, weight=5)
        body.columnconfigure(1, weight=2, minsize=200)
        body.rowconfigure(0, weight=1)

        self._build_quota(body)
        self._build_options(body)

    # ── Таблиця квот ─────────────────────────────────────────────────────────

    def _build_quota(self, parent):
        outer, inner = card(parent, "Квоти питань за розділами")
        outer.grid(row=0, column=0, sticky="nsew", padx=(0, 8), pady=(0, 16))

        # Вкладки категорій
        nb = ttk.Notebook(inner)
        nb.pack(fill="both", expand=True, padx=12, pady=(0, 12))
        self._nb = nb

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("TNotebook",        background=C_CARD, borderwidth=0, tabmargins=0)
        style.configure("TNotebook.Tab",    font=FONT_SM, padding=[16, 7],
                        background=C_BORDER, foreground=C_TEXT)
        style.map("TNotebook.Tab",
                  background=[("selected", C_ACCENT), ("!selected", C_BORDER),
                               ("active",  C_SIDEBAR_H)],
                  foreground=[("selected", "white"),  ("!selected", C_TEXT),
                               ("active",  "white")])

        self._tab_frames = {}
        for cat in ("B1.1", "B1.3", "B2"):
            tab = tk.Frame(nb, bg=C_CARD)
            nb.add(tab, text=f"  {cat}  ")
            self._tab_frames[cat] = tab
            self._build_quota_tab(tab, cat)

    def _build_quota_tab(self, parent, cat_code):
        # Заголовок таблиці
        # padx правий = 17px (ширина scrollbar) щоб заголовки вирівнялись з рядками
        hdr = tk.Frame(parent, bg="#EEF2F7")
        hdr.pack(fill="x", padx=(4, 21), pady=(6, 0))
        # width у символах — має збігатись з шириною Label у рядках даних
        for txt, w, anchor, expand in [
            ("Модуль",       6,  "center", False),
            ("Розділ",       9,  "center", False),
            ("Назва розділу", 0, "w",      True),
            ("В БД ✦",       7,  "center", False),
            ("Квота",        5,  "center", False),
        ]:
            tk.Label(hdr, text=txt, width=w if w else None,
                     bg="#EEF2F7", fg=C_SIDEBAR, font=("Helvetica", 10, "bold"),
                     anchor=anchor).pack(
                side="left",
                fill="x" if expand else None,
                expand=expand,
                padx=4, pady=5,
            )

        # Легенда
        tk.Label(parent, text="✦ В БД — кількість питань цього розділу в базі. Червоний = менше ніж потрібна квота.",
                 bg=C_CARD, fg=C_MUTED, font=("Helvetica", 9),
                 anchor="w").pack(fill="x", padx=8, pady=(0, 2))

        # Прокручуваний список
        wrap = tk.Frame(parent, bg=C_CARD)
        wrap.pack(fill="both", expand=True, padx=4, pady=2)
        wrap.rowconfigure(0, weight=1)
        wrap.columnconfigure(0, weight=1)

        canvas = tk.Canvas(wrap, bg=C_CARD, highlightthickness=0)
        vsb = ttk.Scrollbar(wrap, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)
        canvas.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")

        scroll_frame = tk.Frame(canvas, bg=C_CARD)
        window_id = canvas.create_window((0, 0), window=scroll_frame, anchor="nw")

        def on_configure(e):
            canvas.configure(scrollregion=canvas.bbox("all"))
        def on_canvas_resize(e):
            canvas.itemconfig(window_id, width=e.width)

        def _on_mousewheel(e):
            delta = -e.delta if _IS_MAC else -(e.delta // 120)
            canvas.yview_scroll(delta, "units")

        scroll_frame.bind("<Configure>", on_configure)
        canvas.bind("<Configure>", on_canvas_resize)
        canvas.bind("<MouseWheel>", _on_mousewheel)
        # Також прив'язуємо до scroll_frame і всіх дочірніх віджетів
        scroll_frame.bind("<MouseWheel>", _on_mousewheel)

        setattr(self, f"_rows_{cat_code}", scroll_frame)
        setattr(self, f"_canvas_{cat_code}", canvas)

    def _fill_quota_rows(self, cat_code):
        frame = getattr(self, f"_rows_{cat_code}")
        for w in frame.winfo_children():
            w.destroy()

        dp = self.app.db_path.get()
        if not os.path.exists(dp):
            tk.Label(frame, text="База даних не знайдена. Перейдіть у Налаштування → Ініціалізувати БД.",
                     bg=C_CARD, fg=C_WARNING, font=FONT_SM, wraplength=380, justify="left").pack(pady=20, padx=12)
            return

        try:
            cat_id = db.get_category_id(cat_code, dp)
        except Exception:
            tk.Label(frame, text="БД не ініціалізована. Перейдіть у Налаштування → Ініціалізувати БД.",
                     bg=C_CARD, fg=C_WARNING, font=FONT_SM, wraplength=380, justify="left").pack(pady=20, padx=12)
            return

        if not cat_id:
            tk.Label(frame, text=f"Категорія {cat_code} відсутня в БД. Завантажте таблицю квот.",
                     bg=C_CARD, fg=C_WARNING, font=FONT_SM, wraplength=380, justify="left").pack(pady=20, padx=12)
            return

        try:
            quotas = db.get_quotas_for_category(cat_id, dp)
        except Exception:
            quotas = []

        if not quotas:
            tk.Label(frame, text="Квоти не знайдено. Завантажте таблицю квот на вкладці «Імпорт даних».",
                     bg=C_CARD, fg=C_WARNING, font=FONT_SM, wraplength=380, justify="left").pack(pady=20, padx=12)
            return

        # Отримуємо canvas для прив'язки скролу до рядків
        _canvas = getattr(self, f"_canvas_{cat_code}", None)

        def _bind_scroll(widget):
            if _canvas:
                widget.bind("<MouseWheel>",
                    lambda e, c=_canvas: c.yview_scroll(
                        -e.delta if _IS_MAC else -(e.delta // 120), "units"))

        for i, q in enumerate(quotas):
            bg = C_CARD if i % 2 == 0 else C_ROW_ALT
            row = tk.Frame(frame, bg=bg)
            row.pack(fill="x", pady=0)
            _bind_scroll(row)

            avail = db.count_questions(q["section_id"], cat_id, dp)
            self._avail[(q["section_id"], cat_code)] = avail

            for lbl_widget in [
                tk.Label(row, text=f"M{q['module_code']}", width=6,
                         bg=bg, fg=C_MUTED, font=FONT_SM, anchor="center"),
            ]:
                lbl_widget.pack(side="left", padx=4, pady=4)
                _bind_scroll(lbl_widget)

            sec_lbl = tk.Label(row, text=q["section_code"], width=9,
                               bg=bg, fg=C_TEXT, font=FONT_SM, anchor="center")
            sec_lbl.pack(side="left", padx=2)
            _bind_scroll(sec_lbl)

            # Назва (займає вільний простір)
            name_lbl = tk.Label(row, text=q["section_name"][:34],
                                bg=bg, fg=C_TEXT, font=FONT_SM, anchor="w")
            name_lbl.pack(side="left", fill="x", expand=True, padx=4)
            _bind_scroll(name_lbl)

            # В БД — колір залежить від достатності
            avail_color = C_SUCCESS if avail >= q["count"] else C_DANGER
            avail_lbl = tk.Label(row, text=str(avail), width=7,
                                 bg=bg, fg=avail_color, font=("Helvetica", 11, "bold"),
                                 anchor="center")
            avail_lbl.pack(side="left", padx=2)
            _bind_scroll(avail_lbl)

            # Спінбокс квоти
            key = (q["section_id"], cat_code)
            # Якщо збережена квота перевищує доступну — скидаємо до avail
            saved_quota = q["count"] if q["count"] <= avail else avail
            if key not in self._quota_vars:
                self._quota_vars[key] = tk.IntVar(value=saved_quota)

            spin = ttk.Spinbox(row, from_=0, to=avail,
                               textvariable=self._quota_vars[key],
                               width=5, font=FONT_SM)
            spin.pack(side="left", padx=(4, 8), pady=3)
            _bind_scroll(spin)

            # Підсвічуємо червоним одразу при введенні більше ніж є в БД
            def _make_validator(var, sp, mx):
                def _check(*_):
                    try:
                        val = var.get()
                        sp.configure(foreground=C_DANGER if val > mx else C_TEXT)
                    except tk.TclError:
                        sp.configure(foreground=C_DANGER)
                return _check

            self._quota_vars[key].trace_add(
                "write", _make_validator(self._quota_vars[key], spin, avail)
            )

            # При виході з поля — затискаємо значення в [0, avail]
            def _clamp(e, var=self._quota_vars[key], mx=avail):
                try:
                    val = var.get()
                    if val > mx:
                        var.set(mx)
                    elif val < 0:
                        var.set(0)
                except tk.TclError:
                    var.set(0)

            spin.bind("<FocusOut>", _clamp)

    # ── Права панель: здобувачі + налаштування + кнопка ─────────────────────

    def _build_options(self, parent):
        right = tk.Frame(parent, bg=C_BG)
        right.grid(row=0, column=1, sticky="nsew", pady=(0, 16))
        right.rowconfigure(4, weight=1)

        # Здобувачі
        outer, inner = card(right, "Здобувачі")
        outer.pack(fill="x", pady=(0, 8))

        self.students_text = tk.Text(inner, height=7, font=FONT_SM,
                                     bg="#F7FAFC", fg=C_TEXT,
                                     relief="flat", bd=0,
                                     padx=8, pady=6,
                                     highlightthickness=1,
                                     highlightbackground=C_BORDER)
        self.students_text.pack(fill="x", padx=16, pady=(0, 4))
        self.students_text.insert("1.0", "Іваненко Іван Іванович")

        tk.Label(inner, text="Один рядок = один здобувач",
                 bg=C_CARD, fg=C_MUTED, font=("Helvetica", 10)).pack(
            anchor="w", padx=16, pady=(0, 10))

        # Категорія
        outer2, inner2 = card(right, "Категорія")
        outer2.pack(fill="x", pady=(0, 8))

        self.cat_var = tk.StringVar(value="B1.1")
        cat_row = tk.Frame(inner2, bg=C_CARD)
        cat_row.pack(fill="x", padx=16, pady=(0, 12))
        for cat in ("B1.1", "B1.3", "B2"):
            rb = tk.Radiobutton(cat_row, text=cat, variable=self.cat_var, value=cat,
                                bg=C_CARD, fg=C_TEXT, selectcolor=C_CARD,
                                activebackground=C_CARD,
                                font=FONT_SM, cursor="hand2",
                                command=self._on_cat_change)
            rb.pack(side="left", padx=(0, 16))

        # Папка збереження
        outer4, inner4 = card(right, "Папка для збереження")
        outer4.pack(fill="x", pady=(0, 8))

        dir_row = tk.Frame(inner4, bg=C_CARD)
        dir_row.pack(fill="x", padx=16, pady=(0, 12))
        tk.Entry(dir_row, textvariable=self.app.output_dir,
                 font=FONT_SM, bg="#F7FAFC", fg=C_TEXT,
                 relief="flat", bd=1, highlightthickness=1,
                 highlightbackground=C_BORDER).pack(side="left", fill="x", expand=True)
        ghost_btn(dir_row, "📁", self._browse_output).pack(side="left", padx=(6, 0))

        # Прогрес + кнопка генерації
        outer5, inner5 = card(right)
        outer5.pack(fill="x", pady=(0, 4))

        self._prog_var = tk.DoubleVar(value=0)
        self._prog_lbl = tk.Label(inner5, text="", bg=C_CARD, fg=C_MUTED, font=FONT_SM)
        self._prog_lbl.pack(anchor="w", padx=16, pady=(10, 2))

        _pb_style = ttk.Style()
        _pb_style.configure("Accent.Horizontal.TProgressbar",
                            troughcolor=C_BORDER, background=C_ACCENT,
                            bordercolor=C_BORDER, lightcolor=C_ACCENT,
                            darkcolor=C_ACCENT, thickness=10)
        self._progressbar = ttk.Progressbar(inner5, variable=self._prog_var,
                                             maximum=100, length=200,
                                             style="Accent.Horizontal.TProgressbar")
        self._progressbar.pack(fill="x", padx=16, pady=(0, 8))

        self._gen_btn = accent_btn(inner5, "🎲   Генерувати тести",
                                   self._start_generate)
        self._gen_btn.pack(fill="x", padx=16, pady=(0, 14))

    def _on_cat_change(self):
        """Синхронізуємо активну вкладку таблиці квот з вибраною категорією."""
        cat = self.cat_var.get()
        cats = ("B1.1", "B1.3", "B2")
        if cat in cats:
            self._nb.select(cats.index(cat))

    def _browse_output(self):
        d = filedialog.askdirectory(title="Оберіть папку для збереження",
                                    initialdir=self.app.output_dir.get())
        if d:
            self.app.output_dir.set(d)

    def on_show(self):
        for cat in ("B1.1", "B1.3", "B2"):
            self._fill_quota_rows(cat)

    # ── Генерація у фоновому потоці ───────────────────────────────────────────

    def _start_generate(self):
        names = [n.strip()
                 for n in self.students_text.get("1.0", "end").splitlines()
                 if n.strip()]
        if not names:
            messagebox.showwarning("Увага", "Введіть хоча б одне ім'я здобувача.")
            return

        cat_code = self.cat_var.get()
        dp  = self.app.db_path.get()
        cat_id = db.get_category_id(cat_code, dp)
        if not cat_id:
            messagebox.showerror("Помилка",
                                 f"Категорія {cat_code} відсутня в БД.\n"
                                 "Спочатку завантажте квоти та модуль.")
            return

        # Записуємо змінені квоти назад у БД
        for (sec_id, cat), var in self._quota_vars.items():
            if cat == cat_code:
                db.insert_quota(sec_id, cat_id, var.get(), dp)

        out = self.app.output_dir.get()
        os.makedirs(out, exist_ok=True)

        self._gen_btn.configure(state="disabled", bg=C_MUTED)
        self._prog_var.set(0)

        threading.Thread(
            target=self._generate_worker,
            args=(cat_code, names, out, dp),
            daemon=True,
        ).start()

    def _generate_worker(self, cat_code, names, out_dir, dp):
        total = len(names)
        ok, errors = [], []

        for i, name in enumerate(names):
            self._prog_lbl.configure(text=f"Генерую {i+1}/{total}: {name}…")
            self._prog_var.set(i / total * 100)
            self.app.set_status(f"Генерація {i+1}/{total}: {name}")
            try:
                variant = gen.generate_test(cat_code, name, db_path=dp)
                ex_docx.export_student_docx(variant, out_dir)
                ex_docx.export_answer_key_docx(variant, out_dir)
                ok.append(name)
            except Exception as e:
                errors.append(f"{name}: {e}")

        self._prog_var.set(100)

        msg = f"✓ Згенеровано варіантів: {len(ok)}"
        if errors:
            msg += f"\n\n⚠ Помилки ({len(errors)}):\n" + "\n".join(errors[:5])
        color = C_SUCCESS if not errors else C_WARNING
        lbl_text = f"Готово: {len(ok)} варіант(ів)" + (f", {len(errors)} помилок" if errors else "")

        # Усі tkinter-виклики — тільки з головного потоку
        def _finish():
            self._gen_btn.configure(state="normal", bg=C_ACCENT)
            self._prog_lbl.configure(text=lbl_text)
            self.app.set_status(f"Готово: {len(ok)} варіант(ів)", color)
            if ok and messagebox.askyesno("Готово", msg + "\n\nВідкрити папку з файлами?"):
                subprocess.Popen(["open", out_dir])
            elif errors:
                messagebox.showerror("Помилки генерації", msg)

        self.after(0, _finish)


# ─── Вкладка: Імпорт ─────────────────────────────────────────────────────────

class ImportFrame(tk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent, bg=C_BG)
        self.app = app
        self._build()

    def _build(self):
        tk.Label(self, text="Імпорт даних",
                 bg=C_BG, fg=C_SIDEBAR, font=FONT_TITLE).pack(
            anchor="w", padx=24, pady=(20, 12))

        body = tk.Frame(self, bg=C_BG)
        body.pack(fill="both", expand=True, padx=24, pady=4)
        body.columnconfigure(0, weight=1)
        body.columnconfigure(1, weight=1)
        body.rowconfigure(1, weight=1)

        # ── Картка 1: Таблиця квот ────────────────────────────────────────────
        outer1, inner1 = card(body, "1.  Таблиця квот")
        outer1.grid(row=0, column=0, sticky="ew", padx=(0, 8), pady=(0, 10))

        self.quota_path = tk.StringVar()
        self._file_row(inner1, self.quota_path,
                       "Таблиця по категоріям.docx", "*.docx")
        accent_btn(inner1, "Завантажити квоти",
                   self._import_quotas).pack(fill="x", padx=16, pady=(6, 14))

        # ── Картка 2: Файл модуля ─────────────────────────────────────────────
        outer2, inner2 = card(body, "2.  Файл модуля (питання)")
        outer2.grid(row=0, column=1, sticky="ew", padx=(8, 0), pady=(0, 10))

        self.module_path = tk.StringVar()
        self._file_row(inner2, self.module_path, "Файл модуля (.docx)", "*.docx")

        meta = tk.Frame(inner2, bg=C_CARD)
        meta.pack(fill="x", padx=16, pady=(8, 0))

        tk.Label(meta, text="Код модуля:", bg=C_CARD,
                 fg=C_TEXT, font=FONT_SM).grid(row=0, column=0, sticky="w")
        self.mod_code = tk.Entry(meta, width=8, font=FONT_SM,
                                 bg="#F7FAFC", relief="flat",
                                 highlightthickness=1, highlightbackground=C_BORDER)
        self.mod_code.insert(0, "1")
        self.mod_code.grid(row=0, column=1, padx=(6, 20), pady=3, sticky="w")

        tk.Label(meta, text="Назва модуля:", bg=C_CARD,
                 fg=C_TEXT, font=FONT_SM).grid(row=0, column=2, sticky="w")
        self.mod_name = tk.Entry(meta, font=FONT_SM, bg="#F7FAFC",
                                 relief="flat", highlightthickness=1,
                                 highlightbackground=C_BORDER)
        self.mod_name.insert(0, "Математика")
        self.mod_name.grid(row=0, column=3, padx=6, pady=3, sticky="ew")
        meta.columnconfigure(3, weight=1)

        accent_btn(inner2, "Імпортувати модуль",
                   self._import_module).pack(fill="x", padx=16, pady=(10, 14))

        # ── Журнал ───────────────────────────────────────────────────────────
        outer3, inner3 = card(body, "Журнал операцій")
        outer3.grid(row=1, column=0, columnspan=2, sticky="nsew", pady=(0, 16))

        log_wrap = tk.Frame(inner3, bg=C_CARD)
        log_wrap.pack(fill="both", expand=True, padx=16, pady=(0, 12))
        log_wrap.rowconfigure(0, weight=1)
        log_wrap.columnconfigure(0, weight=1)

        self.log = tk.Text(log_wrap, font=FONT_MONO, bg="#F7FAFC", fg=C_TEXT,
                           relief="flat", bd=0, state="disabled",
                           highlightthickness=1, highlightbackground=C_BORDER,
                           wrap="word")
        vsb = ttk.Scrollbar(log_wrap, command=self.log.yview)
        self.log.configure(yscrollcommand=vsb.set)
        self.log.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")

        # Теги для кольорових рядків
        self.log.tag_configure("ok",   foreground=C_SUCCESS)
        self.log.tag_configure("err",  foreground=C_DANGER)
        self.log.tag_configure("info", foreground=C_ACCENT)
        self.log.tag_configure("warn", foreground=C_WARNING)

        self._log("Готово до роботи.", "info")

    def _file_row(self, parent, var, placeholder, pattern):
        row = tk.Frame(parent, bg=C_CARD)
        row.pack(fill="x", padx=16, pady=(4, 0))
        e = tk.Entry(row, textvariable=var, font=FONT_SM,
                     bg="#F7FAFC", fg=C_TEXT, relief="flat",
                     highlightthickness=1, highlightbackground=C_BORDER)
        e.insert(0, "")
        e.pack(side="left", fill="x", expand=True)
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
                self._log(f"✓ Квоти завантажено. "
                          f"Розділів: {s['sections']}, квот: {s['quotas']}", "ok")
                self.app.set_status("Квоти завантажено", C_SUCCESS)
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

        # Рядок карток-лічильників
        self._cards_row = tk.Frame(self, bg=C_BG)
        self._cards_row.pack(fill="x", padx=24, pady=(0, 12))

        # Таблиця деталей
        outer, inner = card(self, "Кількість питань по модулях та категоріях")
        outer.pack(fill="both", expand=True, padx=24, pady=(0, 16))

        # ttk.Treeview
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

        # Чергування кольорів рядків
        self.tree.tag_configure("even", background=C_CARD)
        self.tree.tag_configure("odd",  background=C_ROW_ALT)

    def on_show(self):
        # Очищаємо картки
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

        # Рахуємо модулі з питаннями окремо
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

        # Картки-лічильники
        for icon, value, label, color in [
            ("📚", stats["questions"],     "Питань",              C_ACCENT),
            ("📂", mods_with_q,            "Модулів\n(з питаннями)", C_SUCCESS),
            ("📋", stats["modules"],       "Модулів\n(у квотах)",  C_SIDEBAR),
            ("🗂",  stats["sections"],     "Розділів",             C_SIDEBAR),
            ("🔢",  stats["quotas"],       "Активних квот",        C_ACCENT),
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

        # Таблиця
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

        # ── База даних ────────────────────────────────────────────────────────
        outer1, inner1 = card(wrap, "База даних (SQLite)")
        outer1.pack(fill="x", pady=(0, 12))

        db_row = tk.Frame(inner1, bg=C_CARD)
        db_row.pack(fill="x", padx=16, pady=(4, 0))
        tk.Entry(db_row, textvariable=self.app.db_path, font=FONT_SM,
                 bg="#F7FAFC", relief="flat", highlightthickness=1,
                 highlightbackground=C_BORDER).pack(side="left", fill="x", expand=True)
        ghost_btn(db_row, "📂", self._browse_db).pack(side="left", padx=(6, 0))

        btn_row = tk.Frame(inner1, bg=C_CARD)
        btn_row.pack(fill="x", padx=16, pady=(8, 14))

        self._db_status = tk.Label(btn_row, text="", bg=C_CARD,
                                    fg=C_MUTED, font=FONT_SM)

        def init_and_report():
            try:
                db.init_db(self.app.db_path.get())
                self._db_status.configure(
                    text="✓ БД ініціалізована", fg=C_SUCCESS)
            except Exception as e:
                self._db_status.configure(text=f"✗ {e}", fg=C_DANGER)

        accent_btn(btn_row, "Ініціалізувати БД",
                   init_and_report).pack(side="left")
        self._db_status.pack(side="left", padx=12)

        # ── Папка виводу ─────────────────────────────────────────────────────
        outer2, inner2 = card(wrap, "Папка для збереження тестів")
        outer2.pack(fill="x", pady=(0, 12))

        dir_row = tk.Frame(inner2, bg=C_CARD)
        dir_row.pack(fill="x", padx=16, pady=(4, 14))
        tk.Entry(dir_row, textvariable=self.app.output_dir, font=FONT_SM,
                 bg="#F7FAFC", relief="flat", highlightthickness=1,
                 highlightbackground=C_BORDER).pack(side="left", fill="x", expand=True)
        ghost_btn(dir_row, "📂",
                  lambda: self._browse_dir(self.app.output_dir)
                  ).pack(side="left", padx=(6, 0))

        # ── Про програму ─────────────────────────────────────────────────────
        outer3, inner3 = card(wrap, "Про програму")
        outer3.pack(fill="x")
        tk.Label(
            inner3,
            text=(
                "Система автоматизованого формування тестових завдань\n"
                "для навчального центру технічного обслуговування повітряних суден\n"
                "відповідно до вимог EASA Part 147.\n\n"
                "Підтримувані категорії:  B1.1  ·  B1.3  ·  B2\n"
                "Формати виводу:  DOCX  ·  PDF"
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
