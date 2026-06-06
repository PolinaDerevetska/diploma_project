
import random
import re
import logging
from dataclasses import dataclass, field
from datetime import datetime

_FORMULA_ONLY = re.compile(r'^\s*[АаAaБбВвBbСсCc]\s*\)?\s*$', re.UNICODE)
_IMG_MARKER = re.compile(r'\[IMG_\d+\]')

def _is_real_option(text: str) -> bool:
    t = (text or '').strip()
    if not t:
        return False
    if _FORMULA_ONLY.match(t):
        return False

    if _IMG_MARKER.search(t):
        return True

    t_clean = _IMG_MARKER.sub('', t).strip().strip(',').strip(';').strip('.').strip('?').strip()
    return bool(t_clean)

def _is_usable_question(row: dict) -> bool:
    opt_a = row["option_a"] or ''
    opt_b = row["option_b"] or ''
    opt_c = row["option_c"] or ''

    has_content = (
        _is_real_option(opt_a)
        or _is_real_option(opt_b)
        or _is_real_option(opt_c)
        or _IMG_MARKER.search(row["question_text"] or '')
    )
    if not has_content:
        return False

    ans = (row["correct_answer"] or 'A').upper()
    correct_text = {'A': opt_a, 'B': opt_b, 'C': opt_c}.get(ans, '')

    if not _is_real_option(correct_text) and not _IMG_MARKER.search(row["question_text"] or ''):
        return False

    return True

import database as db

log = logging.getLogger(__name__)

@dataclass
class QuestionImage:
    context: str
    img_index: int
    ext: str
    data: bytes

@dataclass
class Question:
    id: int
    source_number: int
    question_text: str
    option_a: str
    option_b: str
    option_c: str
    correct_answer: str
    section_code: str
    section_name: str
    module_code: str
    module_name: str
    images: list["QuestionImage"] = field(default_factory=list)

    def has_images(self, context: str | None = None) -> bool:
        if context:
            return any(img.context == context for img in self.images)
        return len(self.images) > 0

@dataclass
class TestVariant:
    student_name: str
    category_code: str
    generated_at: str
    questions: list[Question] = field(default_factory=list)

    @property
    def total_questions(self) -> int:
        return len(self.questions)

def generate_test(
    category_code: str,
    student_name: str = "Здобувач",
    seed: int | None = None,
    db_path: str = db.DB_PATH,
) -> TestVariant:
    if seed is not None:
        random.seed(seed)

    cat_id = db.get_category_id(category_code, db_path)
    if cat_id is None:
        raise ValueError(f"Категорія '{category_code}' не знайдена в БД. "
                         f"Спочатку виконайте парсинг файлів.")

    quotas = db.get_quotas_for_category(cat_id, db_path)
    if not quotas:
        raise ValueError(f"Норми для категорії '{category_code}' не знайдені. "
                         f"Перевірте файл 'Таблиця по категоріям.docx'.")

    selected_questions: list[Question] = []
    warnings: list[str] = []

    for quota in quotas:
        section_id = quota["section_id"]
        needed = quota["count"]
        sec_code = quota["section_code"]
        sec_name = quota["section_name"]
        mod_code = quota["module_code"]
        mod_name = quota["module_name"]

        if needed == 0:
            continue

        rows = db.get_questions(section_id, cat_id, db_path)
        available = [r for r in rows if _is_usable_question(r)]

        if len(available) < needed:
            msg = (f"Розділ {sec_code} ({cat_id=}): потрібно {needed}, "
                   f"доступно {len(available)}")
            warnings.append(msg)
            log.warning(msg)

            chosen = available
        else:
            chosen = random.sample(available, needed)

        for row in chosen:

            img_rows = db.get_question_images(row["id"], db_path)
            images = [
                QuestionImage(
                    context=r["context"],
                    img_index=r["img_index"],
                    ext=r["ext"],
                    data=bytes(r["data"]),
                )
                for r in img_rows
            ]
            selected_questions.append(Question(
                id=row["id"],
                source_number=row["source_number"],
                question_text=row["question_text"],
                option_a=row["option_a"],
                option_b=row["option_b"],
                option_c=row["option_c"],
                correct_answer=row["correct_answer"],
                section_code=sec_code,
                section_name=sec_name,
                module_code=mod_code,
                module_name=mod_name,
                images=images,
            ))

    random.shuffle(selected_questions)

    variant = TestVariant(
        student_name=student_name,
        category_code=category_code,
        generated_at=datetime.now().strftime("%d.%m.%Y %H:%M"),
        questions=selected_questions,
    )

    log.info(
        "Варіант для '%s' (категорія %s): %d питань сформовано",
        student_name, category_code, variant.total_questions
    )
    if warnings:
        log.warning("Попередження при генерації: %s", "; ".join(warnings))

    return variant

def generate_descriptive_questions(
    category_code: str,
    exclude_ids: set[int] | None = None,
    count_per_module: int = 5,
    db_path: str = db.DB_PATH,
) -> list[Question]:

    DESCRIPTIVE_MODULE_CODES = {'7', '9', '10'}

    cat_id = db.get_category_id(category_code, db_path)
    if cat_id is None:
        return []

    exclude_ids = exclude_ids or set()
    result: list[Question] = []

    with db.get_connection(db_path) as conn:
        conn.row_factory = __import__('sqlite3').Row
        cur = conn.cursor()

        for mod_code in sorted(DESCRIPTIVE_MODULE_CODES):
            cur.execute(
                """SELECT s.id, s.code, s.name, m.code AS mc, m.name AS mn
                   FROM sections s
                   JOIN modules m ON m.id = s.module_id
                   WHERE m.code = ?""",
                (mod_code,)
            )
            sections = cur.fetchall()
            if not sections:
                continue

            pool: list[dict] = []
            for sec in sections:
                rows = db.get_questions(sec["id"], cat_id, db_path)
                for r in rows:
                    if r["id"] not in exclude_ids and _is_usable_question(r):
                        pool.append({
                            "row": r,
                            "sec_code": sec["code"],
                            "sec_name": sec["name"],
                            "mod_code": mod_code,
                            "mod_name": sec["mn"],
                        })

            if not pool:
                continue

            chosen = random.sample(pool, min(count_per_module, len(pool)))

            for item in chosen:
                r = item["row"]
                img_rows = db.get_question_images(r["id"], db_path)
                images = [
                    QuestionImage(
                        context=ir["context"],
                        img_index=ir["img_index"],
                        ext=ir["ext"],
                        data=bytes(ir["data"]),
                    )
                    for ir in img_rows
                ]
                result.append(Question(
                    id=r["id"],
                    source_number=r["source_number"],
                    question_text=r["question_text"],
                    option_a=r["option_a"],
                    option_b=r["option_b"],
                    option_c=r["option_c"],
                    correct_answer=r["correct_answer"],
                    section_code=item["sec_code"],
                    section_name=item["sec_name"],
                    module_code=item["mod_code"],
                    module_name=item["mod_name"],
                    images=images,
                ))

    return result

def generate_multiple_tests(
    category_code: str,
    student_names: list[str],
    db_path: str = db.DB_PATH,
) -> list[TestVariant]:
    variants: list[TestVariant] = []
    for name in student_names:
        variant = generate_test(category_code, student_name=name, db_path=db_path)
        variants.append(variant)
    return variants

if __name__ == "__main__":

    logging.basicConfig(level=logging.INFO)
    try:
        v = generate_test("B1.1", "Іваненко Іван Іванович")
        print(f"Згенеровано {v.total_questions} питань для {v.student_name}")
        print(f"Перше питання: {v.questions[0].question_text[:80]}...")
    except ValueError as e:
        print(f"Помилка: {e}")
        print("Спочатку запустіть: python parser.py <папка_з_документами>")
