"""
parser.py — парсинг Word-документів із банком питань Part 147 та збереження до SQLite.

Структура вхідного файлу (наприклад "1 Модуль №1-В1_В2.docx"):
  - Paragraphs з Heading 1: "Категорія: B1.1" / "Категорія: B1.3" / "Категорія: B2"
  - 3 таблиці (по одній на категорію):
      row 0 — заголовок секції: «{розділ}» Рівень − N кількість − K
      row 1 — заголовки колонок (пропускаємо)
      row 2+ — питання: [num, num, "Текст\nА) ...\nВ) ...\nС) ...", "A"/"B"/"C", ...]
  - Нові секції всередині таблиці починаються рядком, де cell[2] містить "кількість"
"""

import os
import re
import logging
from docx import Document
from docx.oxml.ns import qn

import database as db

logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
log = logging.getLogger(__name__)

# Відповідність позначень варіантів до стандартного A/B/C
ANSWER_MAP = {
    "А": "A", "A": "A",
    "В": "B", "B": "B", "Б": "B",
    "С": "C", "C": "C",
}

# Відповідність букв у тексті питання до ключів option_a/b/c
# Підтримуються обидва формати: А/В/С та А/Б/В
OPTION_KEYS = {
    "А": "option_a", "A": "option_a",
    "В": "option_b", "B": "option_b",
    "Б": "option_b",                   # формат А/Б/В
    "С": "option_c", "C": "option_c",
}


# Namespace-константи для VML (OLE Equation 3.0 прев'ю)
_VML_NS = "urn:schemas-microsoft-com:vml"
_R_NS   = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_W_NS   = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

_EXT_MAP = {"jpeg": "jpg", "tiff": "tif", "x-emf": "emf",
            "x-wmf": "wmf", "x-msmetafile": "wmf"}


def _is_excluded_row(row) -> bool:
    """
    Повертає True, якщо рядок таблиці позначений як «Вилучено».
    Ознаки (достатньо однієї):
      1. Будь-яка клітинка має червоний фон (w:shd fill=FF0000).
      2. Будь-яка клітинка містить слово «вилучено» (незалежно від регістру).
    """
    full_text = " ".join(c.text for c in row.cells).lower()
    if "вилучено" in full_text:
        return True
    for cell in row.cells:
        tcPr = cell._tc.find(qn("w:tcPr"))
        if tcPr is None:
            continue
        shd = tcPr.find(qn("w:shd"))
        if shd is None:
            continue
        fill = (shd.get(qn("w:fill")) or "").upper()
        if fill == "FF0000":
            return True
    return False


def _extract_cell_content(cell, doc: Document) -> tuple[str, list[tuple[bytes, str]]]:
    """
    Зчитує контент клітинки Word, вставляючи маркери [IMG_N] саме там,
    де у тексті стоять зображення або формули (EMF/WMF/PNG тощо).

    Ітерує по XML-дочірніх елементах кожного параграфа, щоб зберегти
    правильний порядок: текст → [IMG_0] → текст → [IMG_1] → ...

    Повертає:
        text_with_markers — повний текст зі вставленими маркерами
        images            — список (data: bytes, ext: str) у порядку появи
    """
    para_texts: list[str] = []
    images: list[tuple[bytes, str]] = []

    for para in cell.paragraphs:
        parts: list[str] = []

        for child in para._element:
            local_tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag

            if local_tag != "r":          # обробляємо лише <w:r>
                continue

            # --- Текстові вузли <w:t> ---
            for t_node in child.findall(f"{{{_W_NS}}}t"):
                parts.append(t_node.text or "")

            # --- DrawingML (<a:blip r:embed="...">) ---
            for blip in child.iter(qn("a:blip")):
                r_embed = blip.get(qn("r:embed"))
                if not r_embed:
                    continue
                try:
                    part = doc.part.related_parts[r_embed]
                    ext = _EXT_MAP.get(
                        part.content_type.split("/")[-1].lower(),
                        part.content_type.split("/")[-1].lower(),
                    )
                    parts.append(f"[IMG_{len(images)}]")
                    images.append((part.blob, ext))
                except (KeyError, AttributeError):
                    log.debug("DrawingML: не вдалося витягти rId=%s", r_embed)

            # --- OLE Equation 3.0 (<v:imagedata r:id="...">) ---
            for imgdata in child.iter(f"{{{_VML_NS}}}imagedata"):
                r_id = imgdata.get(f"{{{_R_NS}}}id")
                if not r_id:
                    continue
                try:
                    part = doc.part.related_parts[r_id]
                    ext = _EXT_MAP.get(
                        part.content_type.split("/")[-1].lower(),
                        part.content_type.split("/")[-1].lower(),
                    )
                    parts.append(f"[IMG_{len(images)}]")
                    images.append((part.blob, ext))
                except (KeyError, AttributeError):
                    log.debug("OLE imagedata: не вдалося витягти rId=%s", r_id)

        para_texts.append("".join(parts))

    return "\n".join(para_texts), images


def _parse_section_header(text: str) -> tuple[str, str] | None:
    """
    Парсить заголовок секції виду «1.1 Арифметика» Рівень − 2 кількість − 6
    або "11.1 Назва\nкількість − 5" (без лапок).
    Повертає (code, name) або None якщо не відповідає шаблону.
    """
    # Спочатку шукаємо текст у лапках «...»
    m = re.search(r'[«"](.*?)[»"]', text, re.DOTALL)
    if m:
        inner = m.group(1).strip()
    else:
        # Формат без лапок: "11.1 Назва\nРівень − 2 кількість − 5"
        # Беремо перший рядок до \n
        first_line = text.split("\n")[0].strip()
        # Перший токен — код секції у форматі X.Y або X.Y.Z
        code_m = re.match(r'^(\d+\.\d+(?:\.\d+)?\.?)\s+(.*)', first_line)
        if code_m:
            code = code_m.group(1).rstrip(".")   # прибираємо: "10.1." → "10.1"
            name = code_m.group(2).strip()
            suffix = re.search(r'\(([a-zA-Zа-яА-ЯіІїЇєЄ])\)\s*$', name)
            if suffix:
                code = f"{code}({_normalize_suffix(suffix.group(1))})"
            return code, name
        return None
    # "1.2 Алгебра(b)" або "1.2. Алгебра"
    # Перший токен — базовий код розділу (напр. "1.2"), решта — назва
    parts = inner.split(None, 1)
    if not parts:
        return None
    code = parts[0].strip().rstrip(".")   # прибираємо зайву крапку: "10.1." → "10.1"
    name = parts[1].strip() if len(parts) > 1 else code
    # Якщо код вже містить суфікс (a)/(b)/(c) — не додаємо повторно
    # Якщо ні, але назва закінчується суфіксом — переносимо його в код
    if not re.search(r'\([a-zA-Z]\)$', code):
        suffix = re.search(r'\(([a-zA-Zа-яА-ЯіІїЇєЄ])\)\s*$', name)
        if suffix:
            code = f"{code}({_normalize_suffix(suffix.group(1))})"
    return code, name


# Нормалізація кирилиця → латиниця для суфіксів (a)/(b)/(c)
_CYR_SUFFIX_MAP = {"а": "a", "б": "b", "в": "c", "г": "d",
                   "с": "c", "е": "e", "d": "d"}

def _normalize_suffix(ch: str) -> str:
    """Нормалізує одну літеру суфіксу до латинського ASCII нижнього регістру."""
    ch = ch.lower()
    return _CYR_SUFFIX_MAP.get(ch, ch)


def _parse_question_row(cells: list[str]) -> dict | None:
    """
    Парсить рядок таблиці з питанням.
    cells: [num, num, "Текст\nА) ...\nВ) ...\nС) ...", correct_answer, ...]
    Повертає словник з полями або None якщо рядок не є питанням.
    """
    if len(cells) < 3:
        return None

    # Перший стовпець — номер питання (ціле число, або порожній/зірочка у деяких файлах)
    num_str = cells[0].strip()
    if num_str.isdigit():
        source_number = int(num_str)
    elif num_str in ("", "*", "–", "-"):
        source_number = 0  # Буде перезаписано лічильником у parse_module_file
    else:
        return None

    # Визначаємо позицію колонок:
    # Структура 1: [num, num_дубль, питання, відповідь, ...] — cells[1] є числом або порожнє
    # Структура 2: [num, питання, відповідь, дата]            — cells[1] є текстом питання
    if len(cells) > 1 and (cells[1].strip().isdigit() or cells[1].strip() == ""):
        text_col, ans_col = 2, 3
    else:
        text_col, ans_col = 1, 2

    if text_col >= len(cells) or ans_col >= len(cells):
        return None

    raw_text   = cells[text_col].strip()
    raw_answer = cells[ans_col].strip()

    if not raw_text or not raw_answer:
        return None

    # Розбиваємо текст питання на саме питання та варіанти відповідей
    lines = [ln.strip() for ln in raw_text.split("\n") if ln.strip()]
    question_text = ""
    options: dict[str, str] = {}

    for line in lines:
        # Шукаємо рядки вигляду "А) текст", "A) текст", "Б) текст"
        m = re.match(r'^([АВСABCавсБб])\s*[)\.]\s*(.*)', line)
        if m:
            letter = m.group(1).upper()
            opt_text = m.group(2).strip()
            key = OPTION_KEYS.get(letter)
            if key:
                options[key] = opt_text
        else:
            if not question_text:
                question_text = line
            else:
                question_text += " " + line

    # Якщо варіанти не знайдені (формат без маркерів А/В/С):
    # рядки після першого вважаємо варіантами A, B, C
    if not options and len(lines) >= 4:
        question_text = lines[0]
        option_lines = [ln for ln in lines[1:] if ln]
        if len(option_lines) >= 3:
            options["option_a"] = option_lines[0]
            options["option_b"] = option_lines[1]
            options["option_c"] = option_lines[2]

    # Перевірка наявності всіх трьох варіантів
    if not all(k in options for k in ("option_a", "option_b", "option_c")):
        log.warning("Неповні варіанти відповідей для питання %d: %s", source_number, raw_text[:60])
        for k in ("option_a", "option_b", "option_c"):
            options.setdefault(k, "")

    # Нормалізуємо відповідь
    correct = ANSWER_MAP.get(raw_answer.upper().strip(), raw_answer.upper().strip())
    if correct not in ("A", "B", "C"):
        log.warning("Незрозуміла правильна відповідь '%s' для питання %d", raw_answer, source_number)
        return None

    return {
        "source_number": source_number,
        "question_text": question_text,
        "option_a": options["option_a"],
        "option_b": options["option_b"],
        "option_c": options["option_c"],
        "correct_answer": correct,
    }


def _extract_category_from_heading(text: str) -> str | None:
    """Витягує код категорії з заголовка "Категорія: B1.1"."""
    m = re.search(r'[Кк]атегор[іi]я\s*[:\-]?\s*([A-ZВ][0-9\.]+)', text)
    if m:
        raw = m.group(1).strip()
        # Нормалізуємо: "В1.1" (кирилиця) → "B1.1"
        raw = raw.replace("В", "B")
        # Прибираємо зайву крапку в кінці (напр. "B1.1." → "B1.1")
        raw = raw.rstrip(".")
        # Додаємо префікс T, якщо ще не є TB-кодом
        if raw.startswith("B1") or raw == "B2":
            raw = "T" + raw
        return raw
    return None


def parse_module_file(
    filepath: str,
    module_code: str,
    module_name: str,
    db_path: str = db.DB_PATH,
) -> dict[str, int]:
    """
    Парсить один Word-файл модуля і зберігає дані до БД.

    Args:
        filepath:    шлях до .docx файлу
        module_code: код модуля (напр. "1")
        module_name: назва модуля (напр. "Математика")
        db_path:     шлях до БД

    Returns:
        Словник {category_code: кількість_збережених_питань}
    """
    log.info("Обробка файлу: %s", os.path.basename(filepath))

    doc = Document(filepath)

    # Зберігаємо / отримуємо модуль
    module_id = db.insert_module(module_code, module_name, db_path)

    # Визначаємо категорії з Heading 1 абзаців та їх порядок
    # Потім зіставляємо їх з таблицями (таблиці йдуть у тому ж порядку)
    category_order: list[str] = []
    for para in doc.paragraphs:
        if para.style.name.startswith("Heading"):
            cat = _extract_category_from_heading(para.text)
            if cat:
                category_order.append(cat)

    if not category_order:
        # Якщо заголовків нема — спробуємо взяти стандартні три
        category_order = ["TB1.1", "TB1.3", "TB2"]
        log.warning("Заголовки категорій не знайдені, використовую стандартні: %s", category_order)

    tables = doc.tables
    if len(tables) != len(category_order):
        log.warning(
            "Кількість таблиць (%d) не збігається з кількістю категорій (%d)",
            len(tables), len(category_order)
        )

    stats: dict[str, int] = {}

    for table_idx, table in enumerate(tables):
        if table_idx >= len(category_order):
            break

        cat_code = category_order[table_idx]

        # TB1 (без підкатегорії) — питання застосовуються до обох підкатегорій
        if cat_code in ("TB1", "B1"):
            target_cats = ["TB1.1", "TB1.3"]
        else:
            target_cats = [cat_code]

        # Отримуємо id для всіх цільових категорій
        target_cat_ids = [db.insert_category(c, db_path) for c in target_cats]

        current_section_id: int | None = None
        saved_count = 0
        skipped_count = 0
        auto_number = 0  # лічильник для рядків без номера

        for row_idx, row in enumerate(table.rows):
            cells      = [cell.text.strip() for cell in row.cells]
            raw_cells  = list(row.cells)  # зберігаємо об'єкти Cell для витягання зображень

            # Рядок-заголовок секції: перша клітинка містить "кількість"
            if "кількість" in cells[0].lower() or (len(cells) > 2 and "кількість" in cells[2].lower()):
                header_text = cells[0] if "кількість" in cells[0].lower() else cells[2]
                parsed = _parse_section_header(header_text)
                if parsed:
                    sec_code, sec_name = parsed
                    current_section_id = db.insert_section(module_id, sec_code, sec_name, db_path)
                    log.debug("  Розділ: %s — %s", sec_code, sec_name)
                continue

            # Рядок заголовків стовпців — пропускаємо
            if cells[0].lower() in ("№ питання", "№"):
                continue
            if len(cells) > 1 and "питання та варіанти" in cells[1].lower():
                continue
            if len(cells) > 2 and "питання та варіанти" in cells[2].lower():
                continue

            # Пропускаємо рядки, позначені як «Вилучено» (червоний фон або текст)
            if _is_excluded_row(row):
                skipped_count += 1
                continue

            # Рядок з питанням
            if current_section_id is None:
                log.warning("Питання поза секцією у рядку %d, пропускаємо", row_idx)
                continue

            # Визначаємо колонку з питанням (та варіантами)
            text_col = 2
            if len(cells) > 1 and not (cells[1].strip().isdigit() or cells[1].strip() == ""):
                text_col = 1

            # Витягуємо текст з маркерами [IMG_N] та самі зображення
            if text_col < len(raw_cells):
                marked_text, cell_images = _extract_cell_content(raw_cells[text_col], doc)
                marked_cells = list(cells)
                marked_cells[text_col] = marked_text
            else:
                marked_cells = cells
                cell_images = []

            q = _parse_question_row(marked_cells)
            if q is None:
                continue

            # Призначаємо автоматичний номер для питань без явного номера
            if q["source_number"] == 0:
                auto_number += 1
                q["source_number"] = auto_number
            else:
                auto_number = q["source_number"]

            # Зберігаємо питання для кожної цільової категорії
            for cat_id in target_cat_ids:
                question_id = db.insert_question(
                    section_id=current_section_id,
                    category_id=cat_id,
                    source_number=q["source_number"],
                    question_text=q["question_text"],
                    option_a=q["option_a"],
                    option_b=q["option_b"],
                    option_c=q["option_c"],
                    correct_answer=q["correct_answer"],
                    db_path=db_path,
                )

                # Зберігаємо зображення лише для першої категорії (щоб не дублювати BLOB)
                if cat_id == target_cat_ids[0]:
                    for context, field in [
                        ("question", q["question_text"]),
                        ("option_a", q["option_a"]),
                        ("option_b", q["option_b"]),
                        ("option_c", q["option_c"]),
                    ]:
                        local_idx = 0
                        for m_img in re.finditer(r'\[IMG_(\d+)\]', field):
                            global_idx = int(m_img.group(1))
                            if global_idx < len(cell_images):
                                img_data, ext = cell_images[global_idx]
                                db.insert_question_image(
                                    question_id, context, local_idx, ext, img_data, db_path
                                )
                                local_idx += 1

            saved_count += 1

        display_code = "/".join(target_cats)
        log.info("  Категорія %s: збережено %d питань (пропущено %d)", display_code, saved_count, skipped_count)
        for tc in target_cats:
            stats[tc] = saved_count

    return stats


def parse_all_modules(source_dir: str, db_path: str = db.DB_PATH) -> None:
    """
    Парсить усі Word-файли модулів із директорії source_dir.

    Очікує файли з іменами вигляду "1 Модуль №1-*.docx" або будь-які .docx файли.
    Назву та код модуля можна передати вручну або витягти з імені файлу.

    Args:
        source_dir: шлях до папки з Word-файлами
        db_path:    шлях до БД
    """
    docx_files = sorted([
        f for f in os.listdir(source_dir)
        if f.endswith(".docx") and not f.startswith("~")
    ])

    if not docx_files:
        log.error("Не знайдено .docx файлів у %s", source_dir)
        return

    log.info("Знайдено %d файлів для обробки", len(docx_files))

    for filename in docx_files:
        filepath = os.path.join(source_dir, filename)

        # Витягуємо номер модуля з імені файлу (наприклад "1 Модуль №1-...")
        m = re.search(r'[Мм]одуль\s*[№#]?\s*(\d+)', filename)
        if m:
            module_code = m.group(1)
        else:
            # Якщо не вдалося — використовуємо перший числовий токен з імені файлу
            nums = re.findall(r'\d+', filename)
            module_code = nums[0] if nums else "0"

        # Назву модуля витягуємо після дефісу або беремо з файлу
        name_m = re.search(r'[-–]\s*([А-ЯҐЄІЇа-яґєії ]+)', filename)
        module_name = name_m.group(1).strip() if name_m else f"Модуль {module_code}"

        parse_module_file(filepath, module_code, module_name, db_path)

    stats = db.get_db_stats(db_path)
    log.info("Імпорт завершено. Статистика БД: %s", stats)


if __name__ == "__main__":
    import sys

    db.init_db()

    if len(sys.argv) >= 3:
        # Запуск: python parser.py <шлях_до_файлу> <код_модуля> [назва_модуля]
        fpath = sys.argv[1]
        mcode = sys.argv[2]
        mname = sys.argv[3] if len(sys.argv) > 3 else f"Модуль {mcode}"
        parse_module_file(fpath, mcode, mname)
    elif len(sys.argv) == 2:
        # Запуск: python parser.py <шлях_до_директорії>
        parse_all_modules(sys.argv[1])
    else:
        # За замовчуванням — папка source_docs поруч з data/
        source = os.path.join(os.path.dirname(os.path.dirname(__file__)), "data", "source_docs")
        parse_all_modules(source)

    print(db.get_db_stats())
