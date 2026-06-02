"""
exporter_docx.py — генерація вихідних файлів у форматі Word (.docx):
  - тест для здобувача (без правильних відповідей)
  - ключ відповідей для викладача
"""

import os
import io
import re
import struct
from docx import Document
from docx.shared import Pt, Cm, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.opc.part import Part
from docx.opc.packuri import PackURI
from lxml import etree

from generator import TestVariant, Question, QuestionImage


# ─── EMF: пряма вставка через XML ────────────────────────────────────────────

_emf_counter = 0


_WMF_MAGIC = 0x9AC6CDD7   # Placeable WMF


def _get_emf_dimensions_emu(data: bytes) -> tuple[int, int]:
    """
    Читає розміри з заголовка EMF або WMF.

    EMF: frame rect @ offset 24, int32 LE, одиниці 0.01 мм → × 360 = EMU.
    WMF Placeable: magic 0x9AC6CDD7, bbox @ offset 6, int16 LE, inch @ offset 14.
    """
    try:
        # --- Placeable WMF ---
        if len(data) >= 22 and struct.unpack_from("<I", data, 0)[0] == _WMF_MAGIC:
            left, top, right, bottom = struct.unpack_from("<hhhh", data, 6)
            inch = struct.unpack_from("<H", data, 14)[0] or 96
            w = (right - left) * 914_400 // inch
            h = (bottom - top) * 914_400 // inch
            if w > 0 and h > 0:
                return w, h
        # --- EMF ---
        left, top, right, bottom = struct.unpack_from("<iiii", data, 24)
        w = (right - left) * 360
        h = (bottom - top) * 360
        if w > 0 and h > 0:
            return w, h
    except Exception:
        pass
    return 2_880_000, 2_160_000   # запасний розмір 8 × 6 см


def _insert_emf_picture(paragraph, emf_data: bytes, doc_part,
                        max_width_emu: int = 2_880_000) -> None:
    """
    Вставляє EMF/WMF-зображення у параграф, оминаючи Pillow.
    Додає бінарні дані як Part пакету та формує <w:drawing> вручну.

    max_width_emu — максимальна ширина (за замовчуванням 8 см).
    """
    global _emf_counter
    _emf_counter += 1

    width_emu, height_emu = _get_emf_dimensions_emu(emf_data)
    if width_emu > max_width_emu:
        height_emu = int(height_emu * max_width_emu / width_emu)
        width_emu = max_width_emu

    # Додаємо EMF як окремий Part у пакет docx
    uri = PackURI(f"/word/media/emf_{_emf_counter}.emf")
    emf_part = Part(uri, "image/x-emf", emf_data, doc_part.package)
    rId = doc_part.relate_to(
        emf_part,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
    )

    # XML для inline-зображення
    xml = (
        '<w:drawing'
        ' xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
        ' xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"'
        ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
        ' xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"'
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f'<wp:inline distT="0" distB="0" distL="0" distR="0">'
        f'<wp:extent cx="{width_emu}" cy="{height_emu}"/>'
        f'<wp:effectExtent l="0" t="0" r="0" b="0"/>'
        f'<wp:docPr id="{_emf_counter}" name="emf{_emf_counter}"/>'
        f'<wp:cNvGraphicFramePr>'
        f'<a:graphicFrameLocks noChangeAspect="1"/>'
        f'</wp:cNvGraphicFramePr>'
        f'<a:graphic>'
        f'<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        f'<pic:pic>'
        f'<pic:nvPicPr>'
        f'<pic:cNvPr id="0" name="emf{_emf_counter}"/>'
        f'<pic:cNvPicPr/>'
        f'</pic:nvPicPr>'
        f'<pic:blipFill>'
        f'<a:blip r:embed="{rId}"/>'
        f'<a:stretch><a:fillRect/></a:stretch>'
        f'</pic:blipFill>'
        f'<pic:spPr>'
        f'<a:xfrm><a:off x="0" y="0"/><a:ext cx="{width_emu}" cy="{height_emu}"/></a:xfrm>'
        f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        f'</pic:spPr>'
        f'</pic:pic>'
        f'</a:graphicData>'
        f'</a:graphic>'
        f'</wp:inline>'
        f'</w:drawing>'
    )
    run = paragraph.add_run()
    run._element.append(etree.fromstring(xml.encode("utf-8")))


# ─── Допоміжні функції ────────────────────────────────────────────────────────

def _set_font(run, size_pt: float, bold: bool = False,
              italic: bool = False, color: tuple | None = None) -> None:
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    run.font.italic = italic
    run.font.name = "Times New Roman"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")
    if color:
        run.font.color.rgb = RGBColor(*color)


def _add_paragraph(doc: Document, text: str = "", align=WD_ALIGN_PARAGRAPH.LEFT,
                   space_before: float = 0, space_after: float = 6) -> object:
    p = doc.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    if text:
        run = p.add_run(text)
        _set_font(run, 12)
    return p


def _set_margins(doc: Document) -> None:
    """Поля сторінки: ліво 3 см, решта 2 см."""
    for section in doc.sections:
        section.left_margin = Cm(3)
        section.right_margin = Cm(2)
        section.top_margin = Cm(2)
        section.bottom_margin = Cm(2)


def _add_title_block(doc: Document, variant: TestVariant, is_answer_key: bool) -> None:
    """Додає шапку документа."""
    doc_type = "КЛЮЧ ВІДПОВІДЕЙ" if is_answer_key else "ТЕСТОВЕ ЗАВДАННЯ"
    title_text = f"ТЕСТ PART 147 — {doc_type}"
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(title_text)
    _set_font(run, 14, bold=True)

    p2 = doc.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run2 = p2.add_run(f"Категорія: {variant.category_code}")
    _set_font(run2, 12, bold=True)

    doc.add_paragraph()  # порожній рядок

    # Номер варіанту + пусте поле ПІБ
    p3 = doc.add_paragraph()
    r1 = p3.add_run(f"{variant.student_name}    ")
    _set_font(r1, 12, bold=True)

    p_pib = doc.add_paragraph()
    rpib1 = p_pib.add_run("ПІБ здобувача: ")
    _set_font(rpib1, 12, bold=True)
    rpib2 = p_pib.add_run("___________________________")
    _set_font(rpib2, 12)

    # Дата генерації
    p4 = doc.add_paragraph()
    r3 = p4.add_run("Дата формування: ")
    _set_font(r3, 12, bold=True)
    r4 = p4.add_run(variant.generated_at)
    _set_font(r4, 12)

    # Кількість питань
    p5 = doc.add_paragraph()
    r5 = p5.add_run("Кількість питань: ")
    _set_font(r5, 12, bold=True)
    r6 = p5.add_run(str(variant.total_questions))
    _set_font(r6, 12)

    if not is_answer_key:
        p6 = doc.add_paragraph()
        r7 = p6.add_run("Час виконання: ")
        _set_font(r7, 12, bold=True)
        r8 = p6.add_run("_______ хв.")
        _set_font(r8, 12)

    doc.add_paragraph()  # відступ


# ─── Допоміжна: вставка зображень із маркерами ───────────────────────────────

_IMG_MARKER_RE = re.compile(r'\[IMG_(\d+)\]')


def _insert_single_image(paragraph, img: QuestionImage, doc_part,
                          is_option: bool) -> None:
    """Вставляє одне зображення як inline-run."""
    max_w = 1_440_000 if is_option else 2_880_000
    if img.ext in ("emf", "wmf"):
        _insert_emf_picture(paragraph, img.data, doc_part, max_width_emu=max_w)
    else:
        run = paragraph.add_run()
        try:
            width = Inches(1.5) if is_option else Inches(3.5)
            run.add_picture(io.BytesIO(img.data), width=width)
        except Exception:
            _set_font(run, 10, italic=True)
            run.text = "[зображення]"


def _add_text_with_images(paragraph, text: str, images: list[QuestionImage],
                           context: str, doc_part,
                           font_size: float = 12, bold: bool = False) -> None:
    """
    Додає текст до параграфа, вставляючи зображення [IMG_N] на потрібних позиціях.

    Якщо текст містить маркери [IMG_N] — зображення вставляються між відрізками тексту.
    Якщо маркерів немає (старий формат) — текст додається цілком, зображення в кінці.
    """
    is_option = context in ("option_a", "option_b", "option_c")
    ctx_imgs = {img.img_index: img for img in images if img.context == context}

    # Зображення контексту відсортовані за локальним індексом
    sorted_imgs = sorted(ctx_imgs.values(), key=lambda i: i.img_index)

    if _IMG_MARKER_RE.search(text):
        # ── Новий формат: маркери є → вставляємо зображення на місцях ──
        # Кожен [IMG_N] означає "наступне зображення для цього контексту"
        # (N ігнорується — використовуємо ітератор)
        img_iter = iter(sorted_imgs)
        parts = _IMG_MARKER_RE.split(text)
        for i, part in enumerate(parts):
            if i % 2 == 0:          # текстовий сегмент
                if part:
                    run = paragraph.add_run(part)
                    _set_font(run, font_size, bold=bold)
            else:                   # маркер зображення
                img = next(img_iter, None)
                if img:
                    _insert_single_image(paragraph, img, doc_part, is_option)
    else:
        # ── Старий формат або без маркерів → текст + зображення в кінці ──
        if text:
            run = paragraph.add_run(text)
            _set_font(run, font_size, bold=bold)
        for img in sorted_imgs:
            _insert_single_image(paragraph, img, doc_part, is_option)


# ─── Тест для здобувача ───────────────────────────────────────────────────────

def _add_question_student(doc: Document, num: int, q: Question) -> None:
    """Додає одне питання у форматі для здобувача."""
    # Текст питання з вбудованими зображеннями на потрібних позиціях
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(2)
    r1 = p.add_run(f"{num}. ")
    _set_font(r1, 12, bold=True)
    _add_text_with_images(p, q.question_text, q.images, "question", doc.part,
                          font_size=12, bold=False)

    # Варіанти відповідей з вбудованими зображеннями
    for label, text, ctx in [
        ("А)", q.option_a or "", "option_a"),
        ("В)", q.option_b or "", "option_b"),
        ("С)", q.option_c or "", "option_c"),
    ]:
        op = doc.add_paragraph()
        op.paragraph_format.left_indent = Cm(1)
        op.paragraph_format.space_before = Pt(1)
        op.paragraph_format.space_after = Pt(1)
        rl = op.add_run(f"  {label} ")
        _set_font(rl, 12, bold=True)
        _add_text_with_images(op, text, q.images, ctx, doc.part,
                              font_size=12, bold=False)


def export_student_docx(
    variant: TestVariant,
    output_dir: str,
    filename: str | None = None,
) -> str:
    """
    Генерує Word-документ із тестом для здобувача.

    Args:
        variant:    сформований варіант тесту
        output_dir: папка для збереження
        filename:   ім'я файлу (без розширення); якщо None — генерується автоматично

    Returns:
        Повний шлях до збереженого файлу
    """
    os.makedirs(output_dir, exist_ok=True)

    safe_name = variant.student_name.replace(" ", "_").replace("/", "-")
    if filename is None:
        filename = f"Тест_{variant.category_code}_{safe_name}"
    filepath = os.path.join(output_dir, f"{filename}.docx")

    doc = Document()
    _set_margins(doc)
    _add_title_block(doc, variant, is_answer_key=False)

    # Інструкція
    instr = doc.add_paragraph()
    instr.paragraph_format.space_after = Pt(10)
    ri = instr.add_run(
        "Інструкція: для кожного питання оберіть одну правильну відповідь "
        "(А, В або С) та позначте її у бланку відповідей."
    )
    _set_font(ri, 11, italic=True)

    # Питання
    for i, q in enumerate(variant.questions, start=1):
        _add_question_student(doc, i, q)

    # Бланк відповідей
    doc.add_page_break()
    p_bl = doc.add_paragraph()
    p_bl.alignment = WD_ALIGN_PARAGRAPH.CENTER
    rb = p_bl.add_run("БЛАНК ВІДПОВІДЕЙ")
    _set_font(rb, 14, bold=True)

    doc.add_paragraph()
    pn = doc.add_paragraph()
    rn1 = pn.add_run(f"{variant.student_name}    ")
    _set_font(rn1, 12, bold=True)
    rn2 = pn.add_run("ПІБ: ___________________________")
    _set_font(rn2, 12)

    doc.add_paragraph()

    # Таблиця бланку відповідей
    cols = 10
    rows_needed = (variant.total_questions + cols - 1) // cols
    table = doc.add_table(rows=rows_needed * 2, cols=cols + 1)
    table.style = "Table Grid"

    for r in range(rows_needed):
        num_row = table.rows[r * 2]
        ans_row = table.rows[r * 2 + 1]
        num_row.cells[0].text = "№"
        ans_row.cells[0].text = "Відп."
        for j in range(cols):
            q_num = r * cols + j + 1
            if q_num <= variant.total_questions:
                num_row.cells[j + 1].text = str(q_num)
            else:
                num_row.cells[j + 1].text = ""
            ans_row.cells[j + 1].text = ""

    doc.save(filepath)
    return filepath


# ─── Ключ відповідей для викладача ────────────────────────────────────────────

def export_answer_key_docx(
    variant: TestVariant,
    output_dir: str,
    filename: str | None = None,
) -> str:
    """
    Генерує Word-документ із ключем відповідей для викладача.

    Args:
        variant:    сформований варіант тесту
        output_dir: папка для збереження
        filename:   ім'я файлу (без розширення)

    Returns:
        Повний шлях до збереженого файлу
    """
    os.makedirs(output_dir, exist_ok=True)

    safe_name = variant.student_name.replace(" ", "_").replace("/", "-")
    if filename is None:
        filename = f"Відповіді_{variant.category_code}_{safe_name}"
    filepath = os.path.join(output_dir, f"{filename}.docx")

    doc = Document()
    _set_margins(doc)
    _add_title_block(doc, variant, is_answer_key=True)

    # Таблиця з відповідями
    table = doc.add_table(rows=1, cols=5)
    table.style = "Table Grid"

    # Заголовок таблиці
    hdr = table.rows[0].cells
    headers = ["№", "Правильна відповідь", "Розділ", "Модуль", "ID у БД"]
    for i, h in enumerate(headers):
        hdr[i].text = h
        run = hdr[i].paragraphs[0].runs
        if run:
            run[0].bold = True

    answer_labels = {"A": "А", "B": "В", "C": "С"}

    for i, q in enumerate(variant.questions, start=1):
        row = table.add_row().cells
        row[0].text = str(i)
        row[1].text = answer_labels.get(q.correct_answer, q.correct_answer)
        row[2].text = f"{q.section_code} {q.section_name}"
        row[3].text = f"М{q.module_code} {q.module_name}"
        row[4].text = str(q.id)

    doc.save(filepath)
    return filepath


if __name__ == "__main__":
    # Швидка перевірка (потребує наповненої БД)
    import sys
    sys.path.insert(0, os.path.dirname(__file__))
    from generator import generate_test

    v = generate_test("TB1.1", "Тестовий Здобувач")
    out = os.path.join(os.path.dirname(os.path.dirname(__file__)), "output")

    student_file = export_student_docx(v, out)
    answer_file = export_answer_key_docx(v, out)

    print(f"Тест здобувача: {student_file}")
    print(f"Ключ відповідей: {answer_file}")
