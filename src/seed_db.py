"""
seed_db.py — початкове заповнення бази даних нормами з таблиці категорій.

Викликається автоматично при першому запуску програми, якщо БД порожня
(відсутні категорії або розділи).

Таблиця норм шукається в:
  1. <папка_програми>/data/Таблиця по категоріям.docx  (основне місце)
  2. Поряд з виконуваним файлом (для .app / .exe збірок)
"""

import os
import logging

import database as db
import quota_loader as ql

log = logging.getLogger(__name__)

_SRC  = os.path.dirname(os.path.abspath(__file__))
_ROOT = os.path.dirname(_SRC)

# Шляхи пошуку файлу таблиці норм
_QUOTA_CANDIDATES = [
    os.path.join(_ROOT, "data", "Таблиця по категоріям.docx"),
    os.path.join(_ROOT, "data", "Tablytsia_po_katehoriyam.docx"),
    os.path.join(_SRC, "Таблиця по категоріям.docx"),
]


def find_quota_table() -> str | None:
    """Шукає файл таблиці норм у стандартних місцях. Повертає шлях або None."""
    for path in _QUOTA_CANDIDATES:
        if os.path.exists(path):
            log.info("Знайдено таблицю норм: %s", path)
            return path
    return None


def needs_seeding(db_path: str = db.DB_PATH) -> bool:
    """
    Перевіряє чи потрібна ініціалізація:
    - БД не існує, або
    - таблиця categories порожня, або
    - таблиця quotas порожня
    """
    if not os.path.exists(db_path):
        return True
    try:
        stats = db.get_db_stats(db_path)
        return stats["categories"] == 0 or stats["quotas"] == 0
    except Exception:
        return True


def seed(db_path: str = db.DB_PATH, quota_path: str | None = None) -> bool:
    """
    Ініціалізує БД нормами.

    Args:
        db_path:    шлях до SQLite БД
        quota_path: явний шлях до таблиці норм (якщо None — шукається автоматично)

    Returns:
        True — якщо ініціалізація пройшла успішно
        False — якщо файл таблиці не знайдено
    """
    quota_file = quota_path or find_quota_table()
    if not quota_file:
        log.warning(
            "Файл таблиці норм не знайдено. "
            "Покладіть 'Таблиця по категоріям.docx' у папку data/ поряд з програмою."
        )
        return False

    log.info("Ініціалізація БД: %s", db_path)
    db.init_db(db_path)
    ql.load_quotas_from_docx(quota_file, db_path)

    stats = db.get_db_stats(db_path)
    log.info(
        "БД ініціалізовано: модулів=%d, розділів=%d, норм=%d",
        stats["modules"], stats["sections"], stats["quotas"],
    )
    return True


def auto_seed_if_needed(db_path: str = db.DB_PATH) -> bool:
    """
    Перевіряє необхідність ініціалізації і виконує її.
    Безпечно викликати при кожному запуску програми.

    Returns:
        True — якщо ініціалізація була виконана (або вже не потрібна)
        False — якщо потрібна, але файл таблиці не знайдено
    """
    if not needs_seeding(db_path):
        log.debug("БД вже заповнена, seed не потрібен.")
        return True

    log.info("БД порожня — запускаємо початкове заповнення...")
    return seed(db_path)


if __name__ == "__main__":
    import sys
    logging.basicConfig(level=logging.INFO,
                        format="%(levelname)s %(message)s")

    db_path = db.DB_PATH
    if len(sys.argv) >= 2:
        db_path = sys.argv[1]

    quota_path = sys.argv[2] if len(sys.argv) >= 3 else None

    ok = seed(db_path, quota_path)
    if ok:
        print("✓ БД успішно ініціалізована.")
        print(db.get_db_stats(db_path))
    else:
        print("✗ Не вдалося знайти файл таблиці норм.")
        sys.exit(1)
