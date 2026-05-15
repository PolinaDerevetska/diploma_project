"""
generator.py — алгоритм формування унікального варіанту тесту.

Логіка:
  1. Отримати з БД норми для заданої категорії (скільки питань з кожного розділу).
  2. Для кожного розділу випадково вибрати потрібну кількість питань.
  3. Перемішати фінальний список і повернути структуру тесту.
"""

import random
import logging
from dataclasses import dataclass, field
from datetime import datetime

import database as db

log = logging.getLogger(__name__)


@dataclass
class QuestionImage:
    """Зображення, вбудоване у питання."""
    context: str      # "question" | "option_a" | "option_b" | "option_c"
    img_index: int
    ext: str          # розширення файлу (png, jpg, emf...)
    data: bytes       # бінарні дані


@dataclass
class Question:
    """Одне питання тесту."""
    id: int
    source_number: int
    question_text: str
    option_a: str
    option_b: str
    option_c: str
    correct_answer: str          # "A", "B" або "C"
    section_code: str
    section_name: str
    module_code: str
    module_name: str
    images: list["QuestionImage"] = field(default_factory=list)  # вбудовані зображення

    def has_images(self, context: str | None = None) -> bool:
        """Перевіряє, чи є зображення (опційно — для конкретного контексту)."""
        if context:
            return any(img.context == context for img in self.images)
        return len(self.images) > 0


@dataclass
class TestVariant:
    """Сформований варіант тесту."""
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
    """
    Генерує унікальний варіант тесту для заданої категорії.

    Args:
        category_code: код категорії ("B1.1", "B1.3" або "B2")
        student_name:  ПІБ здобувача (для титульного аркуша)
        seed:          фіксований seed для відтворюваності (None = справжній random)
        db_path:       шлях до БД

    Returns:
        TestVariant із переліком питань

    Raises:
        ValueError: якщо категорія не знайдена або питань недостатньо
    """
    if seed is not None:
        random.seed(seed)

    # Отримуємо id категорії
    cat_id = db.get_category_id(category_code, db_path)
    if cat_id is None:
        raise ValueError(f"Категорія '{category_code}' не знайдена в БД. "
                         f"Спочатку виконайте парсинг файлів.")

    # Отримуємо норми для цієї категорії
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

        # Отримуємо всі доступні питання для цього розділу та категорії
        rows = db.get_questions(section_id, cat_id, db_path)
        available = list(rows)

        if len(available) < needed:
            msg = (f"Розділ {sec_code} ({cat_id=}): потрібно {needed}, "
                   f"доступно {len(available)}")
            warnings.append(msg)
            log.warning(msg)
            # Беремо стільки, скільки є
            chosen = available
        else:
            chosen = random.sample(available, needed)

        for row in chosen:
            # Завантажуємо зображення для цього питання
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

    # Перемішуємо фінальний список, щоб питання різних модулів чергувалися
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


def generate_multiple_tests(
    category_code: str,
    student_names: list[str],
    db_path: str = db.DB_PATH,
) -> list[TestVariant]:
    """
    Генерує унікальні варіанти для кількох здобувачів одночасно.

    Args:
        category_code: код категорії
        student_names: список ПІБ здобувачів
        db_path:       шлях до БД

    Returns:
        Список TestVariant (по одному на кожного здобувача)
    """
    variants: list[TestVariant] = []
    for name in student_names:
        variant = generate_test(category_code, student_name=name, db_path=db_path)
        variants.append(variant)
    return variants


if __name__ == "__main__":
    # Швидка перевірка генератора
    logging.basicConfig(level=logging.INFO)
    try:
        v = generate_test("B1.1", "Іваненко Іван Іванович")
        print(f"Згенеровано {v.total_questions} питань для {v.student_name}")
        print(f"Перше питання: {v.questions[0].question_text[:80]}...")
    except ValueError as e:
        print(f"Помилка: {e}")
        print("Спочатку запустіть: python parser.py <папка_з_документами>")
