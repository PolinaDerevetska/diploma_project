"""
database.py — робота з SQLite базою даних для системи генерації тестів Part 147.
Містить схему БД та всі CRUD-операції.
"""

import sqlite3
import os
from typing import Optional


DB_PATH = os.path.join(os.path.dirname(os.path.dirname(__file__)), "data", "questions.db")


def get_connection(db_path: str = DB_PATH) -> sqlite3.Connection:
    """Повертає з'єднання з базою даних."""
    os.makedirs(os.path.dirname(db_path), exist_ok=True)
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def init_db(db_path: str = DB_PATH) -> None:
    """Створює всі таблиці БД, якщо вони ще не існують."""
    with get_connection(db_path) as conn:
        conn.executescript("""
            -- Модулі (M1, M2, ... M17)
            CREATE TABLE IF NOT EXISTS modules (
                id      INTEGER PRIMARY KEY AUTOINCREMENT,
                code    TEXT    NOT NULL UNIQUE,   -- "1", "2", ... "17"
                name    TEXT    NOT NULL            -- "Математика", "Фізика" ...
            );

            -- Категорії (B1.1, B1.3, B2)
            CREATE TABLE IF NOT EXISTS categories (
                id      INTEGER PRIMARY KEY AUTOINCREMENT,
                code    TEXT    NOT NULL UNIQUE    -- "B1.1", "B1.3", "B2"
            );

            -- Розділи (підтеми всередині модуля)
            CREATE TABLE IF NOT EXISTS sections (
                id          INTEGER PRIMARY KEY AUTOINCREMENT,
                module_id   INTEGER NOT NULL REFERENCES modules(id),
                code        TEXT    NOT NULL,  -- "1.1", "1.2(a)" ...
                name        TEXT    NOT NULL   -- "Арифметика" ...
            );

            -- Питання
            CREATE TABLE IF NOT EXISTS questions (
                id              INTEGER PRIMARY KEY AUTOINCREMENT,
                section_id      INTEGER NOT NULL REFERENCES sections(id),
                category_id     INTEGER NOT NULL REFERENCES categories(id),
                source_number   INTEGER,           -- оригінальний номер у Word-файлі
                question_text   TEXT    NOT NULL,
                option_a        TEXT    NOT NULL,
                option_b        TEXT    NOT NULL,
                option_c        TEXT    NOT NULL,
                correct_answer  TEXT    NOT NULL   -- "A", "B" або "C"
            );
            CREATE INDEX IF NOT EXISTS idx_q_section_cat
                ON questions(section_id, category_id);

            -- Зображення, вбудовані у питання (формули, схеми тощо)
            -- context: "question" | "option_a" | "option_b" | "option_c"
            CREATE TABLE IF NOT EXISTS question_images (
                id           INTEGER PRIMARY KEY AUTOINCREMENT,
                question_id  INTEGER NOT NULL REFERENCES questions(id),
                context      TEXT    NOT NULL,  -- де саме стоїть зображення
                img_index    INTEGER NOT NULL DEFAULT 0,  -- порядок серед кількох зображень
                ext          TEXT    NOT NULL DEFAULT 'png',  -- розширення файлу
                data         BLOB    NOT NULL   -- бінарні дані зображення
            );
            CREATE INDEX IF NOT EXISTS idx_qimg_question
                ON question_images(question_id, context);

            -- Квоти: скільки питань потрібно з кожного розділу для кожної категорії
            CREATE TABLE IF NOT EXISTS quotas (
                section_id      INTEGER NOT NULL REFERENCES sections(id),
                category_id     INTEGER NOT NULL REFERENCES categories(id),
                count           INTEGER NOT NULL DEFAULT 0,
                PRIMARY KEY (section_id, category_id)
            );
        """)
    print(f"[DB] Ініціалізовано: {db_path}")


# ─── Модулі ──────────────────────────────────────────────────────────────────

def insert_module(code: str, name: str, db_path: str = DB_PATH) -> int:
    """Додає модуль або повертає існуючий id."""
    with get_connection(db_path) as conn:
        existing = conn.execute(
            "SELECT id FROM modules WHERE code = ?", (code,)
        ).fetchone()
        if existing:
            return existing["id"]
        cur = conn.execute(
            "INSERT INTO modules (code, name) VALUES (?, ?)", (code, name)
        )
        return cur.lastrowid


def get_all_modules(db_path: str = DB_PATH) -> list[sqlite3.Row]:
    with get_connection(db_path) as conn:
        return conn.execute("SELECT * FROM modules ORDER BY CAST(code AS INTEGER)").fetchall()


# ─── Категорії ───────────────────────────────────────────────────────────────

def insert_category(code: str, db_path: str = DB_PATH) -> int:
    """Додає категорію або повертає існуючий id."""
    with get_connection(db_path) as conn:
        existing = conn.execute(
            "SELECT id FROM categories WHERE code = ?", (code,)
        ).fetchone()
        if existing:
            return existing["id"]
        cur = conn.execute("INSERT INTO categories (code) VALUES (?)", (code,))
        return cur.lastrowid


def get_category_id(code: str, db_path: str = DB_PATH) -> Optional[int]:
    with get_connection(db_path) as conn:
        row = conn.execute(
            "SELECT id FROM categories WHERE code = ?", (code,)
        ).fetchone()
        return row["id"] if row else None


def get_all_categories(db_path: str = DB_PATH) -> list[sqlite3.Row]:
    with get_connection(db_path) as conn:
        return conn.execute("SELECT * FROM categories ORDER BY code").fetchall()


# ─── Розділи ─────────────────────────────────────────────────────────────────

def insert_section(module_id: int, code: str, name: str, db_path: str = DB_PATH) -> int:
    """Додає розділ або повертає існуючий id."""
    with get_connection(db_path) as conn:
        existing = conn.execute(
            "SELECT id FROM sections WHERE module_id = ? AND code = ?",
            (module_id, code)
        ).fetchone()
        if existing:
            return existing["id"]
        cur = conn.execute(
            "INSERT INTO sections (module_id, code, name) VALUES (?, ?, ?)",
            (module_id, code, name)
        )
        return cur.lastrowid


def get_sections_by_module(module_id: int, db_path: str = DB_PATH) -> list[sqlite3.Row]:
    with get_connection(db_path) as conn:
        return conn.execute(
            "SELECT * FROM sections WHERE module_id = ? ORDER BY id",
            (module_id,)
        ).fetchall()


# ─── Питання ─────────────────────────────────────────────────────────────────

def insert_question(
    section_id: int,
    category_id: int,
    source_number: int,
    question_text: str,
    option_a: str,
    option_b: str,
    option_c: str,
    correct_answer: str,
    db_path: str = DB_PATH,
) -> int:
    with get_connection(db_path) as conn:
        cur = conn.execute(
            """INSERT INTO questions
               (section_id, category_id, source_number,
                question_text, option_a, option_b, option_c, correct_answer)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
            (section_id, category_id, source_number,
             question_text, option_a, option_b, option_c, correct_answer),
        )
        return cur.lastrowid


def get_questions(
    section_id: int,
    category_id: int,
    db_path: str = DB_PATH,
) -> list[sqlite3.Row]:
    """Повертає всі питання для заданого розділу та категорії."""
    with get_connection(db_path) as conn:
        return conn.execute(
            """SELECT * FROM questions
               WHERE section_id = ? AND category_id = ?""",
            (section_id, category_id),
        ).fetchall()


def count_questions(section_id: int, category_id: int, db_path: str = DB_PATH) -> int:
    with get_connection(db_path) as conn:
        row = conn.execute(
            "SELECT COUNT(*) as cnt FROM questions WHERE section_id=? AND category_id=?",
            (section_id, category_id),
        ).fetchone()
        return row["cnt"]


# ─── Квоти ───────────────────────────────────────────────────────────────────

def insert_quota(
    section_id: int,
    category_id: int,
    count: int,
    db_path: str = DB_PATH,
) -> None:
    with get_connection(db_path) as conn:
        conn.execute(
            """INSERT INTO quotas (section_id, category_id, count)
               VALUES (?, ?, ?)
               ON CONFLICT(section_id, category_id) DO UPDATE SET count=excluded.count""",
            (section_id, category_id, count),
        )


def get_quotas_for_category(category_id: int, db_path: str = DB_PATH) -> list[sqlite3.Row]:
    """Повертає всі квоти для заданої категорії (тільки де count > 0)."""
    with get_connection(db_path) as conn:
        return conn.execute(
            """SELECT q.section_id, q.count, s.code as section_code,
                      s.name as section_name, m.code as module_code, m.name as module_name
               FROM quotas q
               JOIN sections s ON s.id = q.section_id
               JOIN modules  m ON m.id = s.module_id
               WHERE q.category_id = ? AND q.count > 0
               ORDER BY CAST(m.code AS INTEGER), s.id""",
            (category_id,),
        ).fetchall()


# ─── Зображення ──────────────────────────────────────────────────────────────

def insert_question_image(
    question_id: int,
    context: str,
    img_index: int,
    ext: str,
    data: bytes,
    db_path: str = DB_PATH,
) -> int:
    """Зберігає зображення, прив'язане до питання."""
    with get_connection(db_path) as conn:
        cur = conn.execute(
            """INSERT INTO question_images (question_id, context, img_index, ext, data)
               VALUES (?, ?, ?, ?, ?)""",
            (question_id, context, img_index, ext, data),
        )
        return cur.lastrowid


def get_question_images(
    question_id: int,
    db_path: str = DB_PATH,
) -> list[sqlite3.Row]:
    """Повертає всі зображення для питання, відсортовані за context та img_index."""
    with get_connection(db_path) as conn:
        return conn.execute(
            """SELECT * FROM question_images
               WHERE question_id = ?
               ORDER BY context, img_index""",
            (question_id,),
        ).fetchall()


# ─── Статистика ──────────────────────────────────────────────────────────────

def get_db_stats(db_path: str = DB_PATH) -> dict:
    """Повертає загальну статистику по БД."""
    with get_connection(db_path) as conn:
        stats = {
            "modules":    conn.execute("SELECT COUNT(*) FROM modules").fetchone()[0],
            "categories": conn.execute("SELECT COUNT(*) FROM categories").fetchone()[0],
            "sections":   conn.execute("SELECT COUNT(*) FROM sections").fetchone()[0],
            "questions":  conn.execute("SELECT COUNT(*) FROM questions").fetchone()[0],
            "quotas":     conn.execute("SELECT COUNT(*) FROM quotas WHERE count > 0").fetchone()[0],
        }
    return stats


if __name__ == "__main__":
    init_db()
    print("[DB] Статистика:", get_db_stats())
