
import os
import sys
import sqlite3
import platform
import threading
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox

import customtkinter as ctk

_IS_MAC = platform.system() == "Darwin"

_SRC  = os.path.dirname(os.path.abspath(__file__))
_ROOT = os.path.dirname(_SRC)
sys.path.insert(0, _SRC)

import database as db
import generator as gen
import exporter_docx as ex_docx
import parser as prs
import quota_loader as ql
import seed_db as seed

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

C_SIDEBAR   = "#0F172A"
C_SIDEBAR_H = "#1E293B"

C_ACCENT    = ("#2563EB", "#3B82F6")
C_ACCENT_H  = ("#1D4ED8", "#2563EB")

C_BG        = ("#F8FAFC", "#1E293B")
C_CARD      = ("#FFFFFF", "#0F172A")
C_TEXT      = ("#1E293B", "#E2E8F0")
C_MUTED     = ("#64748B", "#94A3B8")
C_BORDER    = ("#E2E8F0", "#334155")
C_ROW_ALT   = ("#F1F5F9", "#172033")
C_SUCCESS   = ("#16A34A", "#22C55E")
C_WARNING   = ("#D97706", "#FBBF24")
C_DANGER    = ("#DC2626", "#F87171")

def _pick(color_tuple):
    if isinstance(color_tuple, str):
        return color_tuple
    mode = ctk.get_appearance_mode().lower()
    return color_tuple[1] if mode == "dark" else color_tuple[0]

def _bind_mousewheel(scrollable_frame: "ctk.CTkScrollableFrame"):
    canvas = scrollable_frame._parent_canvas

    def _on_wheel(event):
        if _IS_MAC:
            canvas.yview_scroll(-1 * event.delta, "units")
        else:
            canvas.yview_scroll(-1 * (event.delta // 120), "units")

    def _bind_widget(widget):
        widget.unbind("<MouseWheel>")
        widget.bind("<MouseWheel>", _on_wheel)
        if _IS_MAC:
            widget.unbind("<Button-4>")
            widget.unbind("<Button-5>")
            widget.bind("<Button-4>", lambda e: canvas.yview_scroll(-1, "units"))
            widget.bind("<Button-5>", lambda e: canvas.yview_scroll(1, "units"))
        for child in widget.winfo_children():
            _bind_widget(child)

    canvas.unbind("<MouseWheel>")
    canvas.bind("<MouseWheel>", _on_wheel)

    _pending = [False]

    def _on_configure(e):
        if not _pending[0]:
            _pending[0] = True
            scrollable_frame.after_idle(lambda: (_bind_widget(scrollable_frame), _pending.__setitem__(0, False)))

    _bind_widget(scrollable_frame)
    scrollable_frame.bind("<Configure>", _on_configure, add="+")

FONT       = ("Helvetica", 12)
FONT_SM    = ("Helvetica", 11)
FONT_BOLD  = ("Helvetica", 12, "bold")
FONT_TITLE = ("Helvetica", 18, "bold")

import json
from datetime import datetime

_HISTORY_PATH = os.path.join(_ROOT, "data", "history.json")

def _load_history() -> list[dict]:
    try:
        with open(_HISTORY_PATH, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return []

def _save_history_entry(category: str, variants: int, out_dir: str):
    history = _load_history()
    history.insert(0, {
        "date":     datetime.now().strftime("%d.%m.%Y %H:%M"),
        "category": category,
        "variants": variants,
        "dir":      out_dir,
    })
    history = history[:100]
    os.makedirs(os.path.dirname(_HISTORY_PATH), exist_ok=True)
    try:
        with open(_HISTORY_PATH, "w", encoding="utf-8") as f:
            json.dump(history, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

def accent_btn(parent, text, command, width=None, big=False):
    kw = dict(
        text=text, command=command,
        fg_color=C_ACCENT, hover_color=C_ACCENT_H,
        text_color="white", corner_radius=8,
        font=("Helvetica", 14, "bold"),
        height=46 if big else 36,
    )
    if width:
        kw["width"] = width
    return ctk.CTkButton(parent, **kw)

def ghost_btn(parent, text, command, width=None):
    kw = dict(
        text=text, command=command,
        fg_color=C_BORDER, hover_color=C_ROW_ALT,
        text_color=C_TEXT, corner_radius=8,
        font=FONT_SM, height=36,
    )
    if width:
        kw["width"] = width
    return ctk.CTkButton(parent, **kw)

def card_frame(parent, title="", **kw):
    frame = ctk.CTkFrame(parent, fg_color=C_CARD,
                         border_width=1, border_color=C_BORDER,
                         corner_radius=10, **kw)
    if title:
        ctk.CTkLabel(frame, text=title,
                     text_color=C_TEXT,
                     font=FONT_BOLD,
                     fg_color="transparent").pack(anchor="w", padx=16, pady=(12, 4))
    return frame

class CounterWidget(ctk.CTkFrame):

    def __init__(self, parent, var: tk.IntVar,
                 min_val: int = 0, max_val: int = 999,
                 font_size: int = 13, **kw):
        super().__init__(parent, fg_color="transparent", **kw)
        self._var = var
        self._min = min_val
        self._max = max_val

        self._btn_minus = ctk.CTkButton(
            self, text="−", width=36, height=36,
            fg_color=C_BORDER, hover_color=C_ROW_ALT,
            text_color=C_TEXT,
            font=("Helvetica", 12, "bold"),
            corner_radius=8,
            command=self._dec,
        )
        self._btn_minus.pack(side="left")

        self._entry_var = tk.StringVar(value=str(var.get()))
        self._entry = ctk.CTkEntry(
            self,
            textvariable=self._entry_var,
            font=("Helvetica", 12, "bold"),
            width=64, height=36,
            justify="center",
            corner_radius=8,
            border_width=1,
            border_color=C_BORDER,
        )
        self._entry.pack(side="left", padx=4)

        self._btn_plus = ctk.CTkButton(
            self, text="+", width=36, height=36,
            fg_color=C_ACCENT, hover_color=C_ACCENT_H,
            text_color="white",
            font=("Helvetica", 12, "bold"),
            corner_radius=8,
            command=self._inc,
        )
        self._btn_plus.pack(side="left")

        self._entry.bind("<Return>",   self._commit)
        self._entry.bind("<Tab>",      self._commit)
        self._entry.bind("<FocusOut>", self._commit)
        self._entry.bind("<Up>",       lambda e: self._inc())
        self._entry.bind("<Down>",     lambda e: self._dec())

        vcmd = (self.register(lambda s: s.isdigit()), "%S")
        self._entry.configure(validate="key", validatecommand=vcmd)

        self._var.trace_add("write", self._sync_from_var)
        self._update_states()

    def _commit(self, _=None):
        raw = self._entry_var.get().strip()
        try:
            val = int(raw)
        except ValueError:
            val = self._min
        val = max(self._min, min(self._max, val))
        self._var.set(val)
        self._entry_var.set(str(val))
        self._update_states()

    def _sync_from_var(self, *_):
        try:
            self._entry_var.set(str(self._var.get()))
        except tk.TclError:
            pass
        self._update_states()

    def _dec(self):
        self._commit()
        v = self._var.get()
        if v > self._min:
            self._var.set(v - 1)

    def _inc(self):
        self._commit()
        v = self._var.get()
        if v < self._max:
            self._var.set(v + 1)

    def _update_states(self, *_):
        try:
            v = self._var.get()
        except tk.TclError:
            return
        self._btn_minus.configure(
            state="disabled" if v <= self._min else "normal")
        self._btn_plus.configure(
            state="disabled" if v >= self._max else "normal")

    def set_max(self, max_val: int):
        self._max = max_val
        self._update_states()

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Part 147 — Генератор тестових завдань")
        self.geometry("1150x720")
        self.minsize(960, 620)

        self.db_path    = tk.StringVar(value=db.DB_PATH)
        self.output_dir = tk.StringVar(value=os.path.join(_ROOT, "output"))

        self._auto_seed()
        self._build()
        self._nav("generate")

    def _build(self):
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.sidebar = ctk.CTkFrame(self, fg_color=C_SIDEBAR,
                                    width=220, corner_radius=0)
        self.sidebar.grid(row=0, column=0, sticky="nsew")
        self.sidebar.grid_rowconfigure(7, weight=1)
        self.sidebar.grid_propagate(False)

        logo = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        logo.grid(row=0, column=0, padx=16, pady=(24, 0), sticky="ew")

        ctk.CTkLabel(logo, text="✈", font=("Helvetica", 22),
                     fg_color="transparent", text_color="white").pack(side="left")
        name_col = ctk.CTkFrame(logo, fg_color="transparent")
        name_col.pack(side="left", padx=(8, 0))
        ctk.CTkLabel(name_col, text="PART 147",
                     font=("Helvetica", 13, "bold"),
                     fg_color="transparent", text_color="white",
                     anchor="w").pack(anchor="w")
        ctk.CTkLabel(name_col, text="Генератор тестів",
                     font=("Helvetica", 9),
                     fg_color="transparent", text_color="#64748B",
                     anchor="w").pack(anchor="w")

        ctk.CTkFrame(self.sidebar, fg_color="#1E293B",
                     height=1, corner_radius=0).grid(
            row=1, column=0, padx=16, pady=(18, 4), sticky="ew")

        self._nav_btns = {}
        self._active_nav = None

        nav_frame = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        nav_frame.grid(row=2, column=0, sticky="ew", padx=8)

        def _section_lbl(parent, text):
            ctk.CTkLabel(parent, text=text,
                         font=("Helvetica", 9, "bold"),
                         fg_color="transparent",
                         text_color="#475569",
                         anchor="w").pack(fill="x", padx=12, pady=(10, 2))

        _nav_groups = [
            ("ГОЛОВНЕ", [
                ("generate", "✈", "Генерація"),
                ("modules",  "📦", "Модулі"),
                ("stats",    "📋", "Журнал"),
            ]),
            ("ДОВІДКА", [
                ("settings", "📖", "Інструкція"),
            ]),
        ]

        for section_title, items in _nav_groups:
            _section_lbl(nav_frame, section_title)
            for key, icon, label in items:
                btn = ctk.CTkButton(
                    nav_frame,
                    text=f"  {icon}  {label}",
                    anchor="w",
                    height=40, corner_radius=8,
                    fg_color="transparent",
                    hover_color=C_SIDEBAR_H,
                    text_color="#94A3B8",
                    font=("Helvetica", 12),
                    command=lambda k=key: self._nav(k),
                )
                btn.pack(fill="x", pady=1)
                self._nav_btns[key] = btn

        ctk.CTkFrame(self.sidebar, fg_color="#1E293B",
                     height=1, corner_radius=0).grid(
            row=3, column=0, padx=16, pady=(16, 6), sticky="ew")

        cat_sec = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        cat_sec.grid(row=4, column=0, padx=16, sticky="ew")

        ctk.CTkLabel(cat_sec, text="КАТЕГОРІЇ",
                     font=("Helvetica", 9, "bold"),
                     fg_color="transparent", text_color="#475569",
                     anchor="w").pack(anchor="w")

        badges = ctk.CTkFrame(cat_sec, fg_color="transparent")
        badges.pack(fill="x", pady=(6, 0))
        for txt, col in (("B1.1", "#1D4ED8"), ("B1.3", "#15803D"), ("B2", "#7C3AED")):
            ctk.CTkLabel(badges, text=txt,
                         fg_color=col, text_color="white",
                         font=("Helvetica", 9, "bold"),
                         corner_radius=6,
                         width=40, height=22).pack(side="left", padx=(0, 4))

        ctk.CTkFrame(self.sidebar, fg_color="transparent").grid(
            row=7, column=0, sticky="nsew")

        self._theme_btn = ctk.CTkButton(
            self.sidebar,
            text="🌙  Темна тема",
            anchor="w",
            height=36, corner_radius=8,
            fg_color=C_SIDEBAR_H,
            hover_color="#2D3F55",
            text_color="#94A3B8",
            font=("Helvetica", 10),
            command=self._toggle_theme,
        )
        self._theme_btn.grid(row=8, column=0, padx=16, pady=(0, 6), sticky="ew")

        ctk.CTkLabel(self.sidebar, text="v1.0  •  EASA Part 147",
                     fg_color="transparent", text_color="#334155",
                     font=("Helvetica", 9),
                     anchor="w").grid(row=9, column=0, padx=16,
                                      pady=(0, 14), sticky="ew")

        self._content_wrap = ctk.CTkFrame(self, fg_color=C_BG, corner_radius=0)
        self._content_wrap.grid(row=0, column=1, sticky="nsew")
        self._content_wrap.grid_rowconfigure(0, weight=1)
        self._content_wrap.grid_columnconfigure(0, weight=1)

        self.container = ctk.CTkFrame(self._content_wrap, fg_color=C_BG,
                                      corner_radius=0)
        self.container.grid(row=0, column=0, sticky="nsew")
        self.container.grid_rowconfigure(0, weight=1)
        self.container.grid_columnconfigure(0, weight=1)

        self.statusbar = ctk.CTkFrame(self._content_wrap, fg_color=C_BORDER,
                                      height=28, corner_radius=0)
        self.statusbar.grid(row=1, column=0, sticky="ew")
        self.statusbar.grid_propagate(False)
        self._status_lbl = ctk.CTkLabel(
            self.statusbar, text="Готово",
            text_color=C_MUTED, fg_color="transparent",
            font=("Helvetica", 10), anchor="w")
        self._status_lbl.pack(side="left", padx=14)

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
        for k, btn in self._nav_btns.items():
            if k == key:
                btn.configure(
                    fg_color="#1E3A5F",
                    text_color="white",
                    font=("Helvetica", 12, "bold"),
                )
            else:
                btn.configure(
                    fg_color="transparent",
                    text_color="#94A3B8",
                    font=("Helvetica", 12),
                )
        self._frames[key].tkraise()
        if hasattr(self._frames[key], "on_show"):
            self._frames[key].on_show()

    def set_status(self, text, color=None):
        self._status_lbl.configure(
            text=text,
            text_color=color if color else _pick(C_MUTED),
        )
        self.update_idletasks()

    def _toggle_theme(self):
        current = ctk.get_appearance_mode().lower()
        new = "dark" if current == "light" else "light"
        ctk.set_appearance_mode(new)
        self._theme_btn.configure(
            text="☀  Світла тема" if new == "dark" else "🌙  Темна тема"
        )
        self._rebuild_content()

    def _rebuild_content(self):
        active = self._active_nav
        for f in self._frames.values():
            f.destroy()
        self._frames = {
            "generate": GenerateFrame(self.container, self),
            "modules":  ModulesFrame(self.container, self),
            "stats":    StatsFrame(self.container, self),
            "settings": SettingsFrame(self.container, self),
        }
        for f in self._frames.values():
            f.place(relx=0, rely=0, relwidth=1, relheight=1)
        if active:
            self._nav(active)

    def _auto_seed(self):
        dp = self.db_path.get()
        if seed.needs_seeding(dp):
            seed.auto_seed_if_needed(dp)

class GenerateFrame(ctk.CTkFrame):
    def __init__(self, parent, app):
        super().__init__(parent, fg_color=C_BG, corner_radius=0)
        self.app = app
        self._step = 0

        self._cat_var      = tk.StringVar(value="B1.1")
        self._variants_var = tk.IntVar(value=1)
        self._quota_vars   = {}
        self._avail        = {}
        self._section_rows = {}
        self._module_expanded = {}

        self._build()

    def _build(self):

        self._indicator = ctk.CTkFrame(self, fg_color="transparent")
        self._indicator.pack(fill="x", padx=32, pady=(18, 0))
        self._step_lbls = []

        steps = ["Категорія", "Норми", "Варіанти", "Папка", "Генерація"]
        for i, s in enumerate(steps):
            col = ctk.CTkFrame(self._indicator, fg_color="transparent")
            col.pack(side="left", expand=True)

            circle = ctk.CTkLabel(
                col, text=str(i + 1), width=28, height=28,
                fg_color=C_BORDER, text_color=C_MUTED,
                font=("Helvetica", 10, "bold"),
                corner_radius=14)
            circle.pack()
            lbl = ctk.CTkLabel(col, text=s, fg_color="transparent",
                               text_color=C_MUTED,
                               font=("Helvetica", 9))
            lbl.pack()
            self._step_lbls.append((circle, lbl))

            if i < len(steps) - 1:
                ctk.CTkFrame(self._indicator, fg_color=C_BORDER,
                             height=2, width=40,
                             corner_radius=1).pack(side="left", expand=True, pady=8)

        self._pages_wrap = ctk.CTkFrame(self, fg_color="transparent",
                                        corner_radius=0)
        self._pages_wrap.pack(fill="both", expand=True)

        self._pages = {}
        builders = [self._build_step0, self._build_step1,
                    self._build_step2, self._build_step3, self._build_step4]
        for i, builder in enumerate(builders):
            page = ctk.CTkFrame(self._pages_wrap, fg_color=C_BG, corner_radius=0)
            page.place(relx=0, rely=0, relwidth=1, relheight=1)
            self._pages[i] = page
            builder(page)

        self._go_to(0)

    def _go_to(self, step: int):
        self._step = step
        self._pages[step].tkraise()
        self._update_indicator()

    def _update_indicator(self):
        for i, (circle, lbl) in enumerate(self._step_lbls):
            if i < self._step:
                circle.configure(fg_color=C_SUCCESS, text_color="white", text="✓")
                lbl.configure(text_color=C_SUCCESS)
            elif i == self._step:
                circle.configure(fg_color=C_ACCENT, text_color="white",
                                 text=str(i + 1))
                lbl.configure(text_color=C_ACCENT)
            else:
                circle.configure(fg_color=C_BORDER, text_color=C_MUTED,
                                 text=str(i + 1))
                lbl.configure(text_color=C_MUTED)

    def _build_step0(self, page):
        wrap = ctk.CTkFrame(page, fg_color="transparent")
        wrap.place(relx=0.5, rely=0.5, anchor="center")

        ctk.CTkLabel(wrap, text="✈",
                     text_color=C_ACCENT, fg_color="transparent",
                     font=("Helvetica", 48)).pack(pady=(0, 8))
        ctk.CTkLabel(wrap, text="Оберіть категорію здобувача",
                     text_color=C_TEXT, fg_color="transparent",
                     font=("Helvetica", 16, "bold")).pack()
        ctk.CTkLabel(wrap, text="EASA Part 147 — B1.1 / B1.3 / B2",
                     text_color=C_MUTED, fg_color="transparent",
                     font=("Helvetica", 11)).pack(pady=(2, 24))

        cards_row = ctk.CTkFrame(wrap, fg_color="transparent")
        cards_row.pack()

        cat_info = {
            "B1.1": ("Авіаційний механік\nгазотурбінний літак", "⚙️"),
            "B1.3": ("Авіаційний механік\nгазотурбінний вертольот", "🔧"),
            "B2":   ("Авіаційний механік\nавіоніка", "📡"),
        }

        self._cat_card_widgets = {}
        for cat, (desc, icon) in cat_info.items():
            c = ctk.CTkFrame(
                cards_row,
                fg_color=C_CARD,
                border_width=2,
                border_color=C_BORDER,
                corner_radius=12,
                cursor="hand2",
            )
            c.pack(side="left", padx=14, ipadx=24, ipady=20)

            ctk.CTkLabel(c, text=icon, fg_color="transparent",
                         font=("Helvetica", 40)).pack(pady=(0, 4))
            ctk.CTkLabel(c, text=cat, fg_color="transparent",
                         text_color=C_TEXT,
                         font=("Helvetica", 18, "bold")).pack(pady=(0, 6))
            ctk.CTkLabel(c, text=desc, fg_color="transparent",
                         text_color=C_MUTED,
                         font=("Helvetica", 11),
                         justify="center").pack()

            self._cat_card_widgets[cat] = c

            def _select(e=None, cv=cat):
                self._cat_var.set(cv)
                for k, w in self._cat_card_widgets.items():
                    if k == cv:
                        w.configure(border_color=_pick(C_ACCENT),
                                    fg_color=("#EBF5FB", "#1E3A5F"))
                    else:
                        w.configure(border_color=_pick(C_BORDER),
                                    fg_color=C_CARD)

            c.bind("<Button-1>", _select)
            for child in c.winfo_children():
                child.bind("<Button-1>", _select)

        accent_btn(wrap, "Далі  →", self._from_step0_to_1,
                   big=True).pack(pady=(36, 0))

    def _from_step0_to_1(self):
        cat = self._cat_var.get()
        dp  = self.app.db_path.get()

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
                f"Категорія {cat} ще не завантажена.\n"
                "Спочатку завантажте норми на вкладці «Модулі».")
            return

        self._current_cat_id = cat_id
        self._fill_quota_list()
        self._go_to(1)

    def _build_step1(self, page):
        hdr = ctk.CTkFrame(page, fg_color="transparent")
        hdr.pack(fill="x", padx=28, pady=(14, 0))
        ghost_btn(hdr, "← Назад", lambda: self._go_to(0)).pack(side="left")
        ctk.CTkLabel(hdr, text="Норми питань за розділами",
                     text_color=C_TEXT, fg_color="transparent",
                     font=FONT_TITLE).pack(side="left", padx=16)

        ctk.CTkLabel(page,
                     text="Натисніть на модуль, щоб розгорнути розділи. "
                          "«В БД» — доступних питань. Задайте норму для кожного розділу.",
                     text_color=C_MUTED, fg_color="transparent",
                     font=("Helvetica", 10),
                     justify="left").pack(anchor="w", padx=28, pady=(4, 6))

        self._q1_scroll = ctk.CTkScrollableFrame(
            page, fg_color=C_BG, corner_radius=0,
            scrollbar_button_color=C_BORDER,
        )
        self._q1_scroll.pack(fill="both", expand=True, padx=28, pady=(0, 4))
        _bind_mousewheel(self._q1_scroll)

        bot = ctk.CTkFrame(page, fg_color="transparent")
        bot.pack(fill="x", padx=28, pady=(0, 12))

        self._step1_total_lbl = ctk.CTkLabel(
            bot, text="", fg_color="transparent",
            text_color=C_MUTED, font=FONT_SM)
        self._step1_total_lbl.pack(side="left")

        accent_btn(bot, "Далі  →", self._from_step1_to_2,
                   big=True).pack(side="right")

    def _fill_quota_list(self):
        inner = self._q1_scroll
        for w in inner.winfo_children():
            w.destroy()
        self._quota_vars.clear()
        self._avail.clear()
        self._section_rows.clear()
        self._module_expanded.clear()

        dp       = self.app.db_path.get()
        cat_id   = self._current_cat_id
        cat_code = self._cat_var.get()

        try:
            sections = db.get_sections_for_category_with_quotas(cat_id, dp)
        except Exception as e:
            ctk.CTkLabel(inner, text=f"Помилка завантаження: {e}",
                         text_color=C_DANGER, fg_color="transparent",
                         font=FONT_SM).pack(pady=20)
            return

        if not sections:
            ctk.CTkLabel(inner,
                         text=f"Норми для категорії {cat_code} не завантажені.\n"
                              "Спочатку завантажте таблицю норм на вкладці «Модулі».",
                         text_color=C_WARNING, fg_color="transparent",
                         font=FONT_SM, justify="left").pack(pady=20, padx=4)
            return

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

        for mid, mdata in modules.items():
            mod_frame = ctk.CTkFrame(inner, fg_color="transparent",
                                     corner_radius=0)
            mod_frame.pack(fill="x", pady=(4, 0))

            hdr = ctk.CTkFrame(mod_frame, fg_color=C_SIDEBAR_H,
                               corner_radius=8, cursor="hand2")
            hdr.pack(fill="x")

            arrow_lbl = ctk.CTkLabel(hdr, text="▶",
                                     fg_color="transparent",
                                     text_color="white",
                                     font=("Helvetica", 10),
                                     width=24)
            arrow_lbl.pack(side="left", padx=(8, 0), pady=6)

            ctk.CTkLabel(hdr,
                         text=f"  М{mdata['code']}  {mdata['name']}",
                         fg_color="transparent",
                         text_color="white",
                         font=("Helvetica", 11, "bold"),
                         anchor="w").pack(side="left", fill="x",
                                          expand=True, pady=6)

            avail_sum = sum(s["available"] for s in mdata["sections"])
            ctk.CTkLabel(hdr, text=f"{avail_sum} пит.",
                         fg_color="transparent",
                         text_color="#A8C8E8",
                         font=("Helvetica", 10),
                         width=60).pack(side="right", padx=12)

            body = ctk.CTkFrame(mod_frame, fg_color=C_CARD,
                                corner_radius=0, border_width=1,
                                border_color=C_BORDER)
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
                    self._q1_scroll.event_generate("<Configure>")

            for w in hdr.winfo_children():
                w.bind("<Button-1>", _toggle)
            hdr.bind("<Button-1>", _toggle)

            col_hdr = ctk.CTkFrame(body, fg_color=("#EEF2F7", "#1A2540"),
                                   corner_radius=0)
            col_hdr.pack(fill="x")
            for txt, ww, anchor in [
                ("Розділ", 100, "w"), ("Назва", 0, "w"),
                ("Норма", 60, "e"), ("В БД", 60, "e"), ("Задати", 90, "e"),
            ]:
                kw = dict(text=txt, fg_color="transparent",
                          text_color=C_TEXT,
                          font=("Helvetica", 10, "bold"),
                          anchor=anchor)
                if ww:
                    kw["width"] = ww
                lbl = ctk.CTkLabel(col_hdr, **kw)
                lbl.pack(
                    side="left" if anchor == "w" else "right",
                    padx=4, pady=4,
                    fill="x" if (ww == 0) else None,
                    expand=(ww == 0),
                )

            for i, sec in enumerate(mdata["sections"]):
                sec_id = sec["id"]
                avail  = sec["available"]
                quota  = sec["quota"]

                self._avail[sec_id] = avail
                init_val = min(quota, avail) if avail > 0 else quota
                var = tk.IntVar(value=init_val)
                self._quota_vars[sec_id] = var

                row_bg = C_CARD if i % 2 == 0 else C_ROW_ALT
                row = ctk.CTkFrame(body, fg_color=row_bg, corner_radius=0)
                row.pack(fill="x")

                ctk.CTkLabel(row, text=sec["section_code"],
                             width=100, fg_color="transparent",
                             text_color=C_MUTED, font=FONT_SM,
                             anchor="w").pack(side="left", padx=4, pady=4)

                ctk.CTkLabel(row, text=sec["section_name"],
                             fg_color="transparent",
                             text_color=C_TEXT, font=FONT_SM,
                             anchor="w", wraplength=440,
                             justify="left").pack(
                    side="left", fill="x", expand=True, padx=4)

                ctk.CTkLabel(row, text=str(quota),
                             width=60, fg_color="transparent",
                             text_color=C_MUTED,
                             font=("Helvetica", 10),
                             anchor="e").pack(side="right", padx=(0, 4))

                a_color = C_SUCCESS if avail > 0 else C_DANGER
                ctk.CTkLabel(row, text=str(avail),
                             width=60, fg_color="transparent",
                             text_color=a_color,
                             font=("Helvetica", 11, "bold"),
                             anchor="e").pack(side="right", padx=(0, 8))

                cw = CounterWidget(row, var,
                                   min_val=0, max_val=max(avail, quota),
                                   font_size=11)
                cw.pack(side="right", padx=(0, 8), pady=2)
                self._section_rows[sec_id] = (row, cw)

                var.trace_add("write", lambda *_: self._update_quota_total())

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
        dp     = self.app.db_path.get()
        cat_id = self._current_cat_id
        for sec_id, var in self._quota_vars.items():
            db.insert_quota(sec_id, cat_id, var.get(), dp)
        self._go_to(2)

    def _build_step2(self, page):
        wrap = ctk.CTkFrame(page, fg_color="transparent")
        wrap.place(relx=0.5, rely=0.5, anchor="center")

        ctk.CTkLabel(wrap, text="Скільки варіантів згенерувати?",
                     fg_color="transparent", text_color=C_TEXT,
                     font=("Helvetica", 16, "bold")).pack(pady=(0, 6))
        ctk.CTkLabel(wrap, text="Кожен варіант — унікальна вибірка питань",
                     fg_color="transparent", text_color=C_MUTED,
                     font=("Helvetica", 11)).pack(pady=(0, 28))

        CounterWidget(wrap, self._variants_var,
                      min_val=1, max_val=50, font_size=22).pack(pady=8)

        ctk.CTkLabel(wrap, text="варіант(ів)",
                     fg_color="transparent", text_color=C_MUTED,
                     font=("Helvetica", 11)).pack(pady=(4, 32))

        btn_row = ctk.CTkFrame(wrap, fg_color="transparent")
        btn_row.pack()
        ghost_btn(btn_row, "← Назад", lambda: self._go_to(1)).pack(
            side="left", padx=8)
        accent_btn(btn_row, "Далі  →", lambda: self._go_to(3),
                   big=True).pack(side="left", padx=8)

    def _build_step3(self, page):
        wrap = ctk.CTkFrame(page, fg_color="transparent")
        wrap.place(relx=0.5, rely=0.5, anchor="center")

        ctk.CTkLabel(wrap, text="📁", fg_color="transparent",
                     font=("Helvetica", 40)).pack(pady=(0, 8))
        ctk.CTkLabel(wrap, text="Оберіть папку для збереження",
                     fg_color="transparent", text_color=C_TEXT,
                     font=("Helvetica", 16, "bold")).pack()
        ctk.CTkLabel(wrap, text="Тести та ключі відповідей будуть збережені в цю папку",
                     fg_color="transparent", text_color=C_MUTED,
                     font=("Helvetica", 11)).pack(pady=(2, 20))

        dir_row = ctk.CTkFrame(wrap, fg_color="transparent")
        dir_row.pack(fill="x", pady=(0, 6))

        dir_entry = ctk.CTkEntry(
            dir_row, textvariable=self.app.output_dir,
            font=FONT_SM, width=380, height=36,
            corner_radius=8, border_width=1, border_color=C_BORDER,
        )
        dir_entry.pack(side="left", padx=(0, 6))
        ghost_btn(dir_row, "📂", self._browse_output).pack(side="left")

        self._step3_info = ctk.CTkLabel(wrap, text="",
                                         fg_color="transparent",
                                         text_color=C_MUTED,
                                         font=("Helvetica", 10))
        self._step3_info.pack(pady=(4, 28))

        btn_row = ctk.CTkFrame(wrap, fg_color="transparent")
        btn_row.pack()
        ghost_btn(btn_row, "← Назад", lambda: self._go_to(2)).pack(
            side="left", padx=8)
        accent_btn(btn_row, "🚀  Генерувати", self._start_generate,
                   big=True).pack(side="left", padx=8)

    def _browse_output(self):
        d = filedialog.askdirectory(
            title="Оберіть папку для збереження",
            initialdir=self.app.output_dir.get())
        if d:
            self.app.output_dir.set(d)

    def _build_step4(self, page):
        wrap = ctk.CTkFrame(page, fg_color="transparent")
        wrap.place(relx=0.5, rely=0.5, anchor="center")
        self._step4_wrap = wrap

        self._step4_icon = ctk.CTkLabel(wrap, text="⚙️",
                                         fg_color="transparent",
                                         font=("Helvetica", 44))
        self._step4_icon.pack(pady=(0, 10))

        self._step4_title = ctk.CTkLabel(wrap, text="Генерація…",
                                          fg_color="transparent",
                                          text_color=C_TEXT,
                                          font=("Helvetica", 16, "bold"))
        self._step4_title.pack()

        self._step4_sub = ctk.CTkLabel(wrap, text="",
                                        fg_color="transparent",
                                        text_color=C_MUTED,
                                        font=("Helvetica", 11))
        self._step4_sub.pack(pady=(4, 16))

        self._step4_pb = ctk.CTkProgressBar(wrap, width=340,
                                             progress_color=C_ACCENT,
                                             fg_color=C_BORDER,
                                             height=10, corner_radius=5)
        self._step4_pb.set(0)
        self._step4_pb.pack(fill="x", pady=(0, 20))

        self._step4_btns = ctk.CTkFrame(wrap, fg_color="transparent")
        self._step4_btns.pack()

    def _start_generate(self):
        out = self.app.output_dir.get().strip()
        if not out:
            messagebox.showwarning("Увага", "Оберіть папку для збереження.")
            return
        os.makedirs(out, exist_ok=True)

        self._go_to(4)
        self._step4_icon.configure(text="⚙️")
        self._step4_title.configure(text="Генерація…", text_color=_pick(C_TEXT))
        self._step4_sub.configure(text="Зачекайте, будь ласка")
        self._step4_pb.set(0)
        for w in self._step4_btns.winfo_children():
            w.destroy()

        n        = self._variants_var.get()
        cat_code = self._cat_var.get()
        dp       = self.app.db_path.get()

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
                self._step4_pb.set((idx + 1) / n_variants),
                self.app.set_status(f"Генерація {idx + 1}/{n_variants}"),
            ))
            try:
                variant  = gen.generate_test(cat_code, label, db_path=dp)
                used_ids = {q.id for q in variant.questions}
                desc_qs  = gen.generate_descriptive_questions(
                    cat_code, exclude_ids=used_ids, db_path=dp)
                ex_docx.export_student_docx(variant, out_dir, descriptive=desc_qs)
                ex_docx.export_answer_key_docx(variant, out_dir, descriptive=desc_qs)
                ok.append(label)
            except Exception as e:
                errors.append(f"{label}: {e}")

        def _finish():
            self._step4_pb.set(1)
            if ok:
                _save_history_entry(cat_code, len(ok), out_dir)
            if not errors:
                self._step4_icon.configure(text="✅")
                self._step4_title.configure(text="Готово!",
                                             text_color=_pick(C_SUCCESS))
                self._step4_sub.configure(
                    text=f"Згенеровано {len(ok)} варіант(ів) у папці:\n{out_dir}")
                self.app.set_status(f"Готово: {len(ok)} варіант(ів)",
                                    _pick(C_SUCCESS))
            else:
                self._step4_icon.configure(text="⚠️")
                self._step4_title.configure(text="Завершено з помилками",
                                             text_color=_pick(C_WARNING))
                err_txt = "\n".join(errors[:3])
                if len(errors) > 3:
                    err_txt += f"\n… ще {len(errors)-3} помилок"
                self._step4_sub.configure(
                    text=f"✓ {len(ok)} варіант(ів)  ✗ {len(errors)} помилок\n{err_txt}")
                self.app.set_status(
                    f"Готово: {len(ok)} ok, {len(errors)} помилок",
                    _pick(C_WARNING))

            for w in self._step4_btns.winfo_children():
                w.destroy()

            accent_btn(self._step4_btns, "📂  Відкрити папку",
                       lambda: subprocess.Popen(["open", out_dir]),
                       big=True).pack(side="left", padx=8)
            ghost_btn(self._step4_btns, "↩  Створити ще",
                      lambda: self._go_to(0)).pack(side="left", padx=8)

        self.after(0, _finish)

    def on_show(self):
        pass

class ModulesFrame(ctk.CTkFrame):
    _DEFAULT_MODULES = [
        ("1",  "Математика"), ("2",  "Фізика"),
        ("3",  "Основи електротехніки"), ("4",  "Основи електроніки"),
        ("5",  "Цифрові техніки / Системи приладів"),
        ("6",  "Матеріали та обладнання"),
        ("7",  "Практика технічного обслуговування"),
        ("8",  "Основи аеродинаміки"), ("9",  "Людський фактор"),
        ("10", "Авіаційне законодавство"),
        ("11A", "Аеродинаміка, конструкції та системи ПС з ГТД"),
        ("11B", "Аеродинаміка, конструкції та системи ПС з ПД"),
        ("12", "Аеродинаміка, конструкції та системи вертольотів"),
        ("13", "Конструкція та системи ПС"), ("14", "Двигун"),
        ("15", "Газова турбіна"), ("16", "Поршневий двигун"), ("17", "Гвинт"),
    ]
    _CAT_COLORS = {"B1.1": "#1D4ED8", "B1.3": "#15803D", "B2": "#7C3AED"}

    def __init__(self, parent, app):
        super().__init__(parent, fg_color=C_BG, corner_radius=0)
        self.app = app
        self._selected_code: str | None = None
        self._module_rows: dict[str, dict] = {}
        self._build()

    def _build(self):

        hdr = ctk.CTkFrame(self, fg_color="transparent")
        hdr.pack(fill="x", padx=24, pady=(18, 0))
        ctk.CTkLabel(hdr, text="Модулі та імпорт даних",
                     fg_color="transparent", text_color=C_TEXT,
                     font=FONT_TITLE).pack(side="left")
        self._readiness_row = ctk.CTkFrame(self, fg_color="transparent")
        self._readiness_row.pack(fill="x", padx=24, pady=(12, 0))

        ctk.CTkFrame(self, fg_color=C_BORDER, height=1,
                     corner_radius=0).pack(fill="x", padx=24, pady=(12, 0))

        body = ctk.CTkFrame(self, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=24, pady=(12, 16))
        body.columnconfigure(0, weight=3)
        body.columnconfigure(1, weight=2)
        body.rowconfigure(0, weight=1)

        self._mod_list = ctk.CTkScrollableFrame(
            body, fg_color=C_BG, corner_radius=0,
            scrollbar_button_color=C_BORDER,
        )
        self._mod_list.pack(side="left", fill="both", expand=True,
                            padx=(0, 12))
        _bind_mousewheel(self._mod_list)

        right = ctk.CTkFrame(body, fg_color="transparent")
        right.pack(side="left", fill="y")

        c1 = card_frame(right, "Таблиця норм")
        c1.pack(fill="x", pady=(0, 10))
        self.quota_path = tk.StringVar()
        self._file_row(c1, self.quota_path, "Таблиця по категоріям.docx", "*.docx")
        self._norm_status_lbl = ctk.CTkLabel(
            c1, text="", fg_color="transparent",
            text_color=C_MUTED, font=("Helvetica", 10))
        self._norm_status_lbl.pack(anchor="w", padx=16, pady=(4, 0))
        accent_btn(c1, "Завантажити норми",
                   self._import_quotas).pack(fill="x", padx=16, pady=(8, 14))

        c2 = card_frame(right, "Імпорт модуля")
        c2.pack(fill="x")
        self._sel_lbl = ctk.CTkLabel(
            c2, text="← Оберіть модуль зі списку",
            fg_color="transparent", text_color=C_MUTED,
            font=("Helvetica", 11, "italic"))
        self._sel_lbl.pack(anchor="w", padx=16, pady=(10, 4))
        self.module_path = tk.StringVar()
        self._file_row(c2, self.module_path, "Файл модуля (.docx)", "*.docx")
        self._import_status_lbl = ctk.CTkLabel(
            c2, text="", fg_color="transparent",
            text_color=C_MUTED, font=("Helvetica", 10))
        self._import_status_lbl.pack(anchor="w", padx=16, pady=(4, 0))
        accent_btn(c2, "Імпортувати модуль",
                   self._import_module).pack(fill="x", padx=16, pady=(8, 14))

    def _build_readiness(self, conn):
        for w in self._readiness_row.winfo_children():
            w.destroy()

        CAT_INFO = {
            "B1.1": ("⚙️", "Авіамеханік ГТД",  "#1D4ED8"),
            "B1.3": ("🔧", "Авіамеханік ПД",   "#15803D"),
            "B2":   ("📡", "Авіоніка",          "#7C3AED"),
        }
        for cat_code, (icon, name, color) in CAT_INFO.items():
            try:
                row = conn.execute("""
                    SELECT COALESCE(SUM(qt.count),0) AS quota_total,
                           COALESCE(SUM(MIN(qt.count,
                               (SELECT COUNT(*) FROM questions q2
                                WHERE q2.section_id=s.id AND q2.category_id=c.id)
                           )),0) AS q_available
                    FROM quotas qt
                    JOIN sections s   ON s.id=qt.section_id
                    JOIN categories c ON c.id=qt.category_id AND c.code=?
                    WHERE qt.count>0
                """, (cat_code,)).fetchone()
                quota = row["quota_total"] or 0
                avail = row["q_available"] or 0
                pct   = int(avail / quota * 100) if quota > 0 else 0
            except Exception:
                quota, avail, pct = 0, 0, 0

            tile = ctk.CTkFrame(self._readiness_row,
                                fg_color=C_CARD,
                                border_width=1, border_color=color,
                                corner_radius=10)
            tile.pack(side="left", fill="x", expand=True, padx=(0, 10))

            top = ctk.CTkFrame(tile, fg_color="transparent")
            top.pack(fill="x", padx=12, pady=(10, 4))
            ctk.CTkLabel(top, text=icon, fg_color="transparent",
                         font=("Helvetica", 18)).pack(side="left")
            ctk.CTkLabel(top, text=f"  {cat_code}",
                         fg_color="transparent", text_color=C_TEXT,
                         font=("Helvetica", 14, "bold")).pack(side="left")
            ctk.CTkLabel(top, text=f"{pct}%",
                         fg_color="transparent", text_color=color,
                         font=("Helvetica", 13, "bold")).pack(side="right")

            ctk.CTkLabel(tile, text=name,
                         fg_color="transparent", text_color=C_MUTED,
                         font=("Helvetica", 9)).pack(anchor="w", padx=12)

            pb = ctk.CTkProgressBar(tile, height=6, corner_radius=3,
                                    fg_color=C_BORDER,
                                    progress_color=color)
            pb.set(pct / 100)
            pb.pack(fill="x", padx=12, pady=(6, 4))

            ctk.CTkLabel(tile, text=f"{avail} / {quota} питань",
                         fg_color="transparent", text_color=C_MUTED,
                         font=("Helvetica", 9)).pack(
                anchor="w", padx=12, pady=(0, 10))

    def _build_module_list(self, conn):
        for w in self._mod_list.winfo_children():
            w.destroy()
        self._module_rows.clear()

        try:
            rows = conn.execute("""
                SELECT m.code, m.name,
                    COUNT(DISTINCT q.id) AS total,
                    (SELECT COALESCE(SUM(qt2.count),0)
                     FROM quotas qt2 JOIN sections s2 ON s2.id=qt2.section_id
                     WHERE s2.module_id=m.id) AS quota_sum,
                    (SELECT COUNT(*) FROM questions q2
                     JOIN sections s2 ON q2.section_id=s2.id
                     JOIN categories c2 ON q2.category_id=c2.id
                     WHERE s2.module_id=m.id AND c2.code='B1.1') AS b11,
                    (SELECT COUNT(*) FROM questions q2
                     JOIN sections s2 ON q2.section_id=s2.id
                     JOIN categories c2 ON q2.category_id=c2.id
                     WHERE s2.module_id=m.id AND c2.code='B1.3') AS b13,
                    (SELECT COUNT(*) FROM questions q2
                     JOIN sections s2 ON q2.section_id=s2.id
                     JOIN categories c2 ON q2.category_id=c2.id
                     WHERE s2.module_id=m.id AND c2.code='B2') AS b2
                FROM modules m
                LEFT JOIN sections s ON s.module_id=m.id
                LEFT JOIN questions q ON q.section_id=s.id
                GROUP BY m.id ORDER BY CAST(m.code AS INTEGER)
            """).fetchall()
        except Exception:
            rows = []

        if not rows:

            for code, name in self._DEFAULT_MODULES:
                rows_fake = type('R', (), {
                    '__getitem__': lambda s,k: 0 if k not in ('code','name') else (code if k=='code' else name)
                })()

        for r in rows:
            total = r["total"] or 0
            quota = r["quota_sum"] or 0
            pct   = int(total / quota * 100) if quota > 0 else 0
            loaded = total > 0
            code = r["code"]
            name = r["name"]

            if loaded:
                bar_color = _pick(C_SUCCESS)
                dot_color = _pick(C_SUCCESS)
                brd_color = (_pick(C_SUCCESS), "#334155")
            else:
                bar_color = "#CBD5E1"
                dot_color = "#CBD5E1"
                brd_color = C_BORDER

            card = ctk.CTkFrame(self._mod_list,
                                fg_color=C_CARD,
                                border_width=1,
                                border_color=brd_color,
                                corner_radius=8,
                                cursor="hand2")
            card.pack(fill="x", pady=(0, 5))

            row_inner = ctk.CTkFrame(card, fg_color="transparent")
            row_inner.pack(fill="x", padx=12, pady=8)

            left = ctk.CTkFrame(row_inner, fg_color="transparent")
            left.pack(side="left", fill="x", expand=True)

            code_row = ctk.CTkFrame(left, fg_color="transparent")
            code_row.pack(anchor="w")
            ctk.CTkLabel(code_row, text="●", fg_color="transparent",
                         text_color=dot_color,
                         font=("Helvetica", 10)).pack(side="left")
            ctk.CTkLabel(code_row, text=f" M{code}",
                         fg_color="transparent", text_color=C_TEXT,
                         font=("Helvetica", 12, "bold")).pack(side="left")

            ctk.CTkLabel(left, text=name,
                         fg_color="transparent",
                         text_color=C_TEXT if loaded else C_MUTED,
                         font=("Helvetica", 10),
                         anchor="w").pack(anchor="w")

            pb = ctk.CTkProgressBar(left, height=4, corner_radius=2,
                                    fg_color=C_BORDER,
                                    progress_color=bar_color,
                                    width=200)
            pb.set(min(pct / 100, 1.0))
            pb.pack(anchor="w", pady=(4, 0))

            right = ctk.CTkFrame(row_inner, fg_color="transparent")
            right.pack(side="right", anchor="n")

            if loaded:
                ctk.CTkLabel(right, text=f"{total} пит.",
                             fg_color="transparent",
                             text_color=_pick(C_SUCCESS),
                             font=("Helvetica", 11, "bold")).pack(anchor="e")
                badges = ctk.CTkFrame(right, fg_color="transparent")
                badges.pack(anchor="e", pady=(4, 0))
                for cat_code, cnt, col in [
                    ("B1.1", r["b11"] or 0, "#1D4ED8"),
                    ("B1.3", r["b13"] or 0, "#15803D"),
                    ("B2",   r["b2"]  or 0, "#7C3AED"),
                ]:
                    if cnt > 0:
                        ctk.CTkLabel(badges, text=cat_code,
                                     fg_color=col, text_color="white",
                                     font=("Helvetica", 8, "bold"),
                                     corner_radius=4,
                                     width=32, height=18).pack(
                            side="left", padx=2)
            else:
                ctk.CTkLabel(right, text="не завантажено",
                             fg_color="transparent",
                             text_color="#CBD5E1",
                             font=("Helvetica", 9, "italic")).pack(anchor="e")

            self._module_rows[code] = {"card": card, "loaded": loaded, "name": name}

            def _sel(e=None, c=code, n=name):
                self._on_select(c, n)
            card.bind("<Button-1>", _sel)
            for child in card.winfo_children():
                child.bind("<Button-1>", _sel)
                for gc in child.winfo_children():
                    try:
                        gc.bind("<Button-1>", _sel)
                    except Exception:
                        pass

        if self._selected_code and self._selected_code in self._module_rows:
            self._highlight_row(self._selected_code)

    def _on_select(self, code, name):
        self._selected_code = code
        self._highlight_row(code)
        self._sel_lbl.configure(text=f"M{code} — {name}", text_color=C_TEXT)
        self.module_path.set("")
        self._import_status_lbl.configure(text="")

    def _highlight_row(self, code):
        for c, info in self._module_rows.items():
            if c == code:
                info["card"].configure(
                    border_color=_pick(C_ACCENT),
                    fg_color=("#EBF5FB", "#1E3A5F"))
            else:
                brd = _pick(C_SUCCESS) if info["loaded"] else _pick(C_BORDER)
                info["card"].configure(border_color=brd, fg_color=C_CARD)

    def _file_row(self, parent, var, placeholder, pattern):
        row = ctk.CTkFrame(parent, fg_color="transparent")
        row.pack(fill="x", padx=16, pady=(4, 0))
        ctk.CTkEntry(row, textvariable=var, font=FONT_SM,
                     height=36, corner_radius=8,
                     border_width=1, border_color=C_BORDER).pack(
            side="left", fill="x", expand=True)
        ghost_btn(row, "📂",
                  lambda: self._browse(var, placeholder, pattern)).pack(
            side="left", padx=(6, 0))

    def _browse(self, var, label, pattern):
        f = filedialog.askopenfilename(
            title=f"Оберіть: {label}",
            filetypes=[(label, pattern), ("Всі файли", "*.*")],
        )
        if f:
            var.set(f)

    def _update_norm_status(self, conn):
        try:
            stats = db.get_db_stats(self.app.db_path.get())
            if stats["quotas"] > 0:
                self._norm_status_lbl.configure(
                    text=f"✓ {stats['sections']} розділів, {stats['quotas']} норм",
                    text_color=C_SUCCESS)
            else:
                self._norm_status_lbl.configure(
                    text="Норми не завантажені", text_color=C_WARNING)
        except Exception:
            self._norm_status_lbl.configure(
                text="Помилка читання БД", text_color=C_DANGER)

    def on_show(self):
        dp = self.app.db_path.get()
        if not os.path.exists(dp):
            self._norm_status_lbl.configure(text="БД не знайдена",
                                             text_color=C_DANGER)
            return
        try:
            conn = sqlite3.connect(dp)
            conn.row_factory = sqlite3.Row
            self._build_readiness(conn)
            self._build_module_list(conn)
            self._update_norm_status(conn)
            conn.close()
        except Exception as e:
            self.app.set_status(f"Помилка: {e}", _pick(C_DANGER))

    def _import_quotas(self):
        path = self.quota_path.get().strip()
        if not path or not os.path.exists(path):
            messagebox.showwarning("Увага", "Оберіть файл таблиці норм!")
            return
        dp = self.app.db_path.get()
        self.app.set_status("Завантажуємо норми…", _pick(C_ACCENT))

        def worker():
            try:
                db.init_db(dp)
                ql.load_quotas_from_docx(path, dp)
                self.app.set_status("✓ Норми завантажено", _pick(C_SUCCESS))
                self.after(0, self.on_show)
            except Exception as e:
                self.app.set_status(f"Помилка: {e}", _pick(C_DANGER))

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
        name = self._module_rows[code]["name"] if code in self._module_rows else code
        dp   = self.app.db_path.get()

        self._import_status_lbl.configure(text="Імпортуємо…", text_color=C_ACCENT)
        self.app.set_status(f"Імпорт M{code}…", _pick(C_ACCENT))

        def worker():
            try:
                result = prs.parse_module_file(path, code, name, dp)
                total  = sum(result.values())
                self.after(0, lambda: self._import_status_lbl.configure(
                    text=f"✓ {total} питань імпортовано", text_color=C_SUCCESS))
                self.app.set_status(f"Імпорт завершено: {total} питань",
                                    _pick(C_SUCCESS))
                self.after(0, self.on_show)
            except Exception as e:
                self.after(0, lambda: self._import_status_lbl.configure(
                    text=f"✗ {e}", text_color=C_DANGER))
                self.app.set_status(f"Помилка: {e}", _pick(C_DANGER))

        threading.Thread(target=worker, daemon=True).start()

class StatsFrame(ctk.CTkFrame):

    _CAT_COLORS = {"B1.1": "#1D4ED8", "B1.3": "#15803D", "B2": "#7C3AED"}

    def __init__(self, parent, app):
        super().__init__(parent, fg_color=C_BG, corner_radius=0)
        self.app = app
        self._build()

    def _build(self):
        hdr = ctk.CTkFrame(self, fg_color="transparent")
        hdr.pack(fill="x", padx=24, pady=(20, 10))
        ctk.CTkLabel(hdr, text="Журнал генерації",
                     fg_color="transparent", text_color=C_TEXT,
                     font=FONT_TITLE).pack(side="left")
        ctk.CTkLabel(self,
                     text="Тут зберігається історія всіх згенерованих тестів",
                     fg_color="transparent", text_color=C_MUTED,
                     font=("Helvetica", 11)).pack(
            anchor="w", padx=24, pady=(0, 12))

        self._list = ctk.CTkScrollableFrame(
            self, fg_color=C_BG, corner_radius=0,
            scrollbar_button_color=C_BORDER,
        )
        self._list.pack(fill="both", expand=True, padx=24, pady=(0, 16))
        _bind_mousewheel(self._list)

    def on_show(self):
        for w in self._list.winfo_children():
            w.destroy()

        history = _load_history()

        if not history:

            empty = ctk.CTkFrame(self._list, fg_color="transparent")
            empty.pack(expand=True, pady=60)
            ctk.CTkLabel(empty, text="📋",
                         fg_color="transparent",
                         font=("Helvetica", 48)).pack()
            ctk.CTkLabel(empty, text="Журнал порожній",
                         fg_color="transparent", text_color=C_TEXT,
                         font=("Helvetica", 14, "bold")).pack(pady=(8, 4))
            ctk.CTkLabel(empty,
                         text="Записи з'являться після першої генерації тестів",
                         fg_color="transparent", text_color=C_MUTED,
                         font=("Helvetica", 11)).pack()
            return

        for entry in history:
            self._draw_entry(entry)

    def _draw_entry(self, entry: dict):
        cat   = entry.get("category", "—")
        date  = entry.get("date", "—")
        n     = entry.get("variants", 0)
        d     = entry.get("dir", "")
        color = self._CAT_COLORS.get(cat, "#64748B")

        card = ctk.CTkFrame(self._list, fg_color=C_CARD,
                            border_width=1, border_color=C_BORDER,
                            corner_radius=10)
        card.pack(fill="x", pady=(0, 8))

        row = ctk.CTkFrame(card, fg_color="transparent")
        row.pack(fill="x", padx=16, pady=12)

        ctk.CTkLabel(row, text=cat,
                     fg_color=color, text_color="white",
                     font=("Helvetica", 11, "bold"),
                     corner_radius=6, width=48, height=28).pack(side="left")

        info = ctk.CTkFrame(row, fg_color="transparent")
        info.pack(side="left", padx=12, fill="x", expand=True)

        top_row = ctk.CTkFrame(info, fg_color="transparent")
        top_row.pack(anchor="w")
        ctk.CTkLabel(top_row,
                     text=f"{n} варіант(ів) згенеровано",
                     fg_color="transparent", text_color=C_TEXT,
                     font=("Helvetica", 12, "bold")).pack(side="left")
        ctk.CTkLabel(top_row, text=f"  ·  {date}",
                     fg_color="transparent", text_color=C_MUTED,
                     font=("Helvetica", 11)).pack(side="left")

        short_dir = d if len(d) <= 60 else "…" + d[-57:]
        ctk.CTkLabel(info, text=short_dir,
                     fg_color="transparent", text_color=C_MUTED,
                     font=("Helvetica", 10),
                     anchor="w").pack(anchor="w")

        if os.path.exists(d):
            ghost_btn(row, "📂  Відкрити",
                      lambda p=d: subprocess.Popen(["open", p]),
                      width=110).pack(side="right")

class SettingsFrame(ctk.CTkFrame):
    _SECTIONS = [
        {
            "icon": "📦",
            "title": "Крок 1 — Підготовка бази даних",
            "color": ("#EBF5FB", "#1A2A3A"),
            "border": "#2E86AB",
            "steps": [
                ("1.1", "Перейдіть у вкладку «Модулі»"),
                ("1.2", "У блоці «Таблиця норм» оберіть файл «Таблиця по категоріям.docx» "
                        "та натисніть «Завантажити норми». Це потрібно зробити лише один раз."),
                ("1.3", "Для кожного модуля оберіть його зі списку, потім у блоці "
                        "«Завантаження модуля» вкажіть файл .docx та натисніть «Імпортувати модуль»."),
                ("1.4", "Зелений індикатор ● біля модуля означає, що питання завантажено."),
            ],
        },
        {
            "icon": "✈",
            "title": "Крок 2 — Генерація тестів",
            "color": ("#EAFAF1", "#1A2E20"),
            "border": "#27AE60",
            "steps": [
                ("2.1", "Перейдіть у вкладку «Генерація» та оберіть категорію здобувача: "
                        "B1.1 (ГТД), B1.3 (ПД) або B2 (авіоніка)."),
                ("2.2", "На кроці «Норми» відображаються лише ті розділи, що передбачені "
                        "для обраної категорії. «Норма» — рекомендована кількість, «В БД» — доступно."),
                ("2.3", "На кроці «Варіанти» вкажіть скільки варіантів тесту потрібно."),
                ("2.4", "На кроці «Папка» оберіть директорію для збереження файлів."),
                ("2.5", "Натисніть «Генерувати» — програма створить файли тестів і ключів."),
            ],
        },
        {
            "icon": "📄",
            "title": "Вихідні файли",
            "color": ("#FEF9E7", "#2A2510"),
            "border": "#D4AC0D",
            "steps": [
                ("●", "Файл здобувача — тестові завдання з бланком відповідей."),
                ("●", "Файл викладача — лист відповідей з правильними варіантами."),
                ("●", "Всі файли зберігаються у вибрану папку у форматі .docx."),
            ],
        },
        {
            "icon": "📊",
            "title": "Статистика",
            "color": ("#F4ECF7", "#1E1225"),
            "border": "#7D3C98",
            "steps": [
                ("●", "Вкладка «Статистика» показує кількість завантажених питань "
                        "по кожному модулю та категорії."),
                ("●", "Використовуйте її для перевірки перед генерацією тестів."),
            ],
        },
        {
            "icon": "💡",
            "title": "Корисні поради",
            "color": (("#FDFEFE", "#1A1F2A")),
            "border": "#BDC3C7",
            "steps": [
                ("●", "База даних зберігається у файлі questions.db у папці data/. "
                        "При перенесенні програми скопіюйте разом із нею цей файл."),
                ("●", "Якщо потрібно оновити питання модуля — просто завантажте новий файл. "
                        "Старі питання замінюються новими."),
                ("●", "Категорії B1.1 і B1.3 використовують різні набори модулів."),
            ],
        },
    ]

    def __init__(self, parent, app):
        super().__init__(parent, fg_color=C_BG, corner_radius=0)
        self.app = app
        self._build()

    def _build(self):
        hdr = ctk.CTkFrame(self, fg_color="transparent")
        hdr.pack(fill="x", padx=24, pady=(20, 4))
        ctk.CTkLabel(hdr, text="Інструкція користувача",
                     fg_color="transparent", text_color=C_TEXT,
                     font=FONT_TITLE).pack(side="left")

        ctk.CTkLabel(self,
                     text="Система автоматизованого формування тестів EASA Part 147  ·  "
                          "Підтримувані категорії: B1.1  B1.3  B2",
                     fg_color="transparent", text_color=C_MUTED,
                     font=("Helvetica", 11)).pack(
            anchor="w", padx=24, pady=(0, 12))

        scroll = ctk.CTkScrollableFrame(
            self, fg_color=C_BG, corner_radius=0,
            scrollbar_button_color=C_BORDER,
        )
        scroll.pack(fill="both", expand=True, padx=24, pady=(0, 16))
        _bind_mousewheel(scroll)

        for sec in self._SECTIONS:
            self._build_section(scroll, sec)

        ctk.CTkLabel(scroll,
                     text="Розроблено: Деревецька Поліна Михайлівна  ·  "
                          "Криворізький фаховий коледж КАІ  ·  2026",
                     fg_color="transparent", text_color=C_MUTED,
                     font=("Helvetica", 10, "italic")).pack(
            anchor="center", pady=(16, 8))

    def on_show(self):
        pass

    def _build_section(self, parent, sec):
        bg  = sec["color"]
        brd = sec["border"]

        outer = ctk.CTkFrame(parent, fg_color=brd,
                             corner_radius=8)
        outer.pack(fill="x", pady=(0, 10))

        inner = ctk.CTkFrame(outer, fg_color=bg,
                             corner_radius=6)
        inner.pack(fill="both", expand=True, padx=3, pady=3)

        hdr = ctk.CTkFrame(inner, fg_color="transparent")
        hdr.pack(fill="x", padx=14, pady=(12, 6))

        ctk.CTkLabel(hdr, text=sec["icon"],
                     fg_color="transparent",
                     font=("Helvetica", 20)).pack(
            side="left", padx=(0, 8))
        ctk.CTkLabel(hdr, text=sec["title"],
                     fg_color="transparent", text_color=C_TEXT,
                     font=("Helvetica", 13, "bold")).pack(side="left")

        for num, text in sec["steps"]:
            row = ctk.CTkFrame(inner, fg_color="transparent")
            row.pack(fill="x", padx=14, pady=(0, 6))

            ctk.CTkLabel(row, text=num, fg_color="transparent",
                         text_color=brd,
                         font=("Helvetica", 11, "bold"),
                         width=32, anchor="nw").pack(side="left", pady=2)
            ctk.CTkLabel(row, text=text, fg_color="transparent",
                         text_color=C_TEXT,
                         font=("Helvetica", 11),
                         justify="left", wraplength=700,
                         anchor="nw").pack(side="left", fill="x",
                                           expand=True, pady=2)

        ctk.CTkFrame(inner, fg_color="transparent", height=6).pack()

if __name__ == "__main__":
    app = App()
    app.mainloop()
