"""
main.py — точка входу системи автоматичного формування тестових завдань Part 147.

Використання:
  python main.py init                          — ініціалізація БД
  python main.py import <папка_або_файл>       — імпорт Word-документів до БД
  python main.py quota <файл_квот.docx>        — завантаження квот із таблиці категорій
  python main.py generate <категорія> <ПІБ>   — згенерувати один варіант
  python main.py batch <категорія> <файл.txt> — кілька варіантів (ПІБ у файлі)
  python main.py stats                         — статистика БД

Приклад:
  python main.py init
  python main.py import data/source_docs
  python main.py quota "data/Таблиця по категоріям.docx"
  python main.py generate B1.1 "Іваненко Іван Іванович"
"""

import sys
import os
import argparse
import logging

# Додаємо папку src до шляху імпорту
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "src"))

import database as db
from parser import parse_module_file, parse_all_modules
from quota_loader import load_quotas_from_docx
from generator import generate_test, generate_multiple_tests
from exporter_docx import export_student_docx, export_answer_key_docx
from exporter_pdf import export_student_pdf, export_answer_key_pdf

OUTPUT_DIR = os.path.join(os.path.dirname(__file__), "output")

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger(__name__)


def cmd_init(args) -> None:
    """Ініціалізує базу даних."""
    db.init_db()
    print("✓ База даних ініціалізована:", db.DB_PATH)


def cmd_import(args) -> None:
    """Імпортує Word-документи до БД."""
    path = args.path
    if os.path.isfile(path):
        # Один файл
        code = input("Введіть код модуля (наприклад 1): ").strip()
        name = input("Введіть назву модуля (наприклад Математика): ").strip()
        stats = parse_module_file(path, code, name)
        print(f"✓ Імпортовано: {stats}")
    elif os.path.isdir(path):
        # Вся папка
        parse_all_modules(path)
    else:
        print(f"Помилка: '{path}' не знайдено")
        sys.exit(1)


def cmd_quota(args) -> None:
    """Завантажує квоти з файлу таблиці категорій."""
    load_quotas_from_docx(args.file)
    print("✓ Квоти завантажено")


def cmd_generate(args) -> None:
    """Генерує один варіант тесту."""
    category = args.category
    student  = args.name
    fmt      = args.format.lower()

    log.info("Генерація варіанту для '%s', категорія %s...", student, category)

    variant = generate_test(category, student)

    os.makedirs(OUTPUT_DIR, exist_ok=True)

    files = []
    if fmt in ("docx", "all"):
        sf = export_student_docx(variant, OUTPUT_DIR)
        af = export_answer_key_docx(variant, OUTPUT_DIR)
        files += [sf, af]

    if fmt in ("pdf", "all"):
        sf = export_student_pdf(variant, OUTPUT_DIR)
        af = export_answer_key_pdf(variant, OUTPUT_DIR)
        files += [sf, af]

    print(f"\n✓ Згенеровано {variant.total_questions} питань для '{student}'")
    for f in files:
        print(f"  → {f}")


def cmd_batch(args) -> None:
    """Генерує варіанти для кількох здобувачів."""
    category   = args.category
    names_file = args.names_file
    fmt        = args.format.lower()

    if not os.path.exists(names_file):
        print(f"Файл зі списком '{names_file}' не знайдено")
        sys.exit(1)

    with open(names_file, encoding="utf-8") as f:
        names = [line.strip() for line in f if line.strip()]

    if not names:
        print("Список здобувачів порожній")
        sys.exit(1)

    log.info("Генерація для %d здобувачів, категорія %s...", len(names), category)
    variants = generate_multiple_tests(category, names)

    os.makedirs(OUTPUT_DIR, exist_ok=True)

    for variant in variants:
        if fmt in ("docx", "all"):
            export_student_docx(variant, OUTPUT_DIR)
            export_answer_key_docx(variant, OUTPUT_DIR)
        if fmt in ("pdf", "all"):
            export_student_pdf(variant, OUTPUT_DIR)
            export_answer_key_pdf(variant, OUTPUT_DIR)

    print(f"\n✓ Згенеровано {len(variants)} варіантів. Файли збережено в '{OUTPUT_DIR}'")


def cmd_stats(args) -> None:
    """Виводить статистику бази даних."""
    stats = db.get_db_stats()
    print("\n── Статистика бази даних ──────────────────")
    print(f"  Модулів:    {stats['modules']}")
    print(f"  Категорій:  {stats['categories']}")
    print(f"  Розділів:   {stats['sections']}")
    print(f"  Питань:     {stats['questions']}")
    print(f"  Квот:       {stats['quotas']}")
    print("───────────────────────────────────────────\n")

    # Детально по категоріях
    for cat in db.get_all_categories():
        cat_id = cat["id"]
        total_q = db.get_connection().execute(
            "SELECT COUNT(*) FROM questions WHERE category_id=?", (cat_id,)
        ).fetchone()[0]
        print(f"  Категорія {cat['code']}: {total_q} питань")


def main():
    parser = argparse.ArgumentParser(
        description="Система автоматичного формування тестових завдань Part 147",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    sub = parser.add_subparsers(dest="command", required=True)

    # init
    sub.add_parser("init", help="Ініціалізувати базу даних")

    # import
    p_import = sub.add_parser("import", help="Імпортувати Word-документи")
    p_import.add_argument("path", help="Шлях до .docx файлу або папки з файлами")

    # quota
    p_quota = sub.add_parser("quota", help="Завантажити квоти з таблиці категорій")
    p_quota.add_argument("file", help="Шлях до файлу 'Таблиця по категоріям.docx'")

    # generate
    p_gen = sub.add_parser("generate", help="Згенерувати варіант тесту")
    p_gen.add_argument("category", help="Код категорії: B1.1, B1.3 або B2")
    p_gen.add_argument("name", help="ПІБ здобувача у лапках")
    p_gen.add_argument("--format", default="all",
                       choices=["docx", "pdf", "all"],
                       help="Формат вихідного файлу (default: all)")

    # batch
    p_batch = sub.add_parser("batch", help="Згенерувати варіанти для списку здобувачів")
    p_batch.add_argument("category", help="Код категорії: B1.1, B1.3 або B2")
    p_batch.add_argument("names_file", help="Текстовий файл зі списком ПІБ (по одному на рядок)")
    p_batch.add_argument("--format", default="all",
                         choices=["docx", "pdf", "all"],
                         help="Формат вихідного файлу (default: all)")

    # stats
    sub.add_parser("stats", help="Показати статистику БД")

    args = parser.parse_args()

    commands = {
        "init":     cmd_init,
        "import":   cmd_import,
        "quota":    cmd_quota,
        "generate": cmd_generate,
        "batch":    cmd_batch,
        "stats":    cmd_stats,
    }
    commands[args.command](args)


if __name__ == "__main__":
    main()
