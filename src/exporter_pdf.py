"""
exporter_pdf.py — генерація тесту та ключа відповідей у форматі PDF
                  за допомогою бібліотеки reportlab.
"""

import os
import io
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.lib import colors
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    PageBreak, HRFlowable, Image as RLImage
)
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

from generator import TestVariant, Question, QuestionImage

# ─── Шрифти ──────────────────────────────────────────────────────────────────
# Реєструємо DejaVu (підтримує кирилицю), якщо доступний
_FONT_REGULAR = "Helvetica"
_FONT_BOLD    = "Helvetica-Bold"
_FONT_ITALIC  = "Helvetica-Oblique"

def _try_register_fonts() -> None:
    """Намагається зареєструвати шрифт з підтримкою кирилиці."""
    global _FONT_REGULAR, _FONT_BOLD, _FONT_ITALIC

    font_paths = [
        # Linux
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
        # macOS
        "/Library/Fonts/Arial Unicode.ttf",
        # Windows
        "C:/Windows/Fonts/arial.ttf",
    ]

    candidates = [
        ("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
         "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
         "/usr/share/fonts/truetype/dejavu/DejaVuSans-Oblique.ttf",
         "DejaVu"),
        ("C:/Windows/Fonts/arial.ttf",
         "C:/Windows/Fonts/arialbd.ttf",
         "C:/Windows/Fonts/ariali.ttf",
         "Arial"),
    ]

    for reg, bold, italic, name in candidates:
        if os.path.exists(reg):
            try:
                pdfmetrics.registerFont(TTFont(f"{name}", reg))
                if os.path.exists(bold):
                    pdfmetrics.registerFont(TTFont(f"{name}-Bold", bold))
                if os.path.exists(italic):
                    pdfmetrics.registerFont(TTFont(f"{name}-Italic", italic))
                _FONT_REGULAR = name
                _FONT_BOLD    = f"{name}-Bold" if os.path.exists(bold) else name
                _FONT_ITALIC  = f"{name}-Italic" if os.path.exists(italic) else name
                return
            except Exception:
                pass

_try_register_fonts()


# ─── Стилі ───────────────────────────────────────────────────────────────────

def _get_styles() -> dict:
    base = getSampleStyleSheet()
    return {
        "title": ParagraphStyle(
            "title", fontName=_FONT_BOLD, fontSize=14,
            alignment=1, spaceAfter=6, leading=18,
        ),
        "subtitle": ParagraphStyle(
            "subtitle", fontName=_FONT_BOLD, fontSize=12,
            alignment=1, spaceAfter=4,
        ),
        "normal": ParagraphStyle(
            "normal", fontName=_FONT_REGULAR, fontSize=11,
            spaceAfter=3, leading=14,
        ),
        "meta": ParagraphStyle(
            "meta", fontName=_FONT_REGULAR, fontSize=11,
            spaceAfter=3,
        ),
        "question": ParagraphStyle(
            "question", fontName=_FONT_BOLD, fontSize=11,
            spaceBefore=8, spaceAfter=2, leading=14,
        ),
        "option": ParagraphStyle(
            "option", fontName=_FONT_REGULAR, fontSize=11,
            leftIndent=20, spaceAfter=2, leading=13,
        ),
        "instr": ParagraphStyle(
            "instr", fontName=_FONT_ITALIC, fontSize=10,
            spaceAfter=8, leading=13,
        ),
        "answer_key_title": ParagraphStyle(
            "answer_key_title", fontName=_FONT_BOLD, fontSize=13,
            alignment=1, spaceAfter=6,
        ),
    }


# ─── Допоміжна: зображення у PDF ─────────────────────────────────────────────

_UNSUPPORTED_FORMATS = {"wmf", "emf"}

def _add_pdf_images(story: list, images: list[QuestionImage], context: str) -> None:
    """Додає зображення заданого контексту до списку Flowable об'єктів."""
    ctx_imgs = sorted(
        [img for img in images if img.context == context],
        key=lambda i: i.img_index
    )
    is_option = context in ("option_a", "option_b", "option_c")
    img_width = 4 * cm if is_option else 8 * cm

    for img in ctx_imgs:
        # WMF/EMF не підтримуються reportlab — показуємо плейсхолдер
        if img.ext in _UNSUPPORTED_FORMATS:
            story.append(Paragraph(
                "<i>[формула]</i>",
                ParagraphStyle("placeholder", fontName=_FONT_ITALIC,
                               fontSize=9, leftIndent=0, spaceAfter=2),
            ))
            continue
        try:
            rl_img = RLImage(io.BytesIO(img.data), width=img_width, height=None)
            story.append(Spacer(1, 0.1 * cm))
            story.append(rl_img)
            story.append(Spacer(1, 0.1 * cm))
        except Exception:
            story.append(Paragraph(
                "<i>[зображення]</i>",
                ParagraphStyle("placeholder", fontName=_FONT_ITALIC,
                               fontSize=9, leftIndent=0, spaceAfter=2),
            ))


# ─── Тест для здобувача ───────────────────────────────────────────────────────

def export_student_pdf(
    variant: TestVariant,
    output_dir: str,
    filename: str | None = None,
) -> str:
    """
    Генерує PDF-файл із тестом для здобувача.

    Returns:
        Шлях до збереженого PDF файлу
    """
    os.makedirs(output_dir, exist_ok=True)

    safe_name = variant.student_name.replace(" ", "_").replace("/", "-")
    if filename is None:
        filename = f"Тест_{variant.category_code}_{safe_name}"
    filepath = os.path.join(output_dir, f"{filename}.pdf")

    doc = SimpleDocTemplate(
        filepath, pagesize=A4,
        leftMargin=3*cm, rightMargin=2*cm,
        topMargin=2*cm, bottomMargin=2*cm,
    )

    st = _get_styles()
    story = []

    # Шапка
    story.append(Paragraph("ТЕСТ PART 147 — ТЕСТОВЕ ЗАВДАННЯ", st["title"]))
    story.append(Paragraph(f"Категорія: {variant.category_code}", st["subtitle"]))
    story.append(HRFlowable(width="100%", thickness=1, color=colors.grey))
    story.append(Spacer(1, 0.3*cm))

    story.append(Paragraph(f"<b>Здобувач:</b> {variant.student_name}", st["meta"]))
    story.append(Paragraph(f"<b>Дата:</b> {variant.generated_at}", st["meta"]))
    story.append(Paragraph(f"<b>Кількість питань:</b> {variant.total_questions}", st["meta"]))
    story.append(Paragraph("<b>Час виконання:</b> _______ хв.", st["meta"]))
    story.append(Spacer(1, 0.3*cm))

    story.append(Paragraph(
        "Інструкція: для кожного питання оберіть одну правильну відповідь "
        "(А, В або С) та позначте її у бланку відповідей.",
        st["instr"]
    ))
    story.append(HRFlowable(width="100%", thickness=0.5, color=colors.lightgrey))
    story.append(Spacer(1, 0.2*cm))

    # Питання
    for i, q in enumerate(variant.questions, start=1):
        story.append(Paragraph(f"{i}. {q.question_text}", st["question"]))
        _add_pdf_images(story, q.images, "question")

        option_defs = [
            ("А)", q.option_a or "", "option_a"),
            ("В)", q.option_b or "", "option_b"),
            ("С)", q.option_c or "", "option_c"),
        ]
        for label, text, ctx in option_defs:
            story.append(Paragraph(f"{label}  {text}", st["option"]))
            _add_pdf_images(story, q.images, ctx)

    # Бланк відповідей
    story.append(PageBreak())
    story.append(Paragraph("БЛАНК ВІДПОВІДЕЙ", st["title"]))
    story.append(Spacer(1, 0.2*cm))
    story.append(Paragraph(f"<b>Здобувач:</b> {variant.student_name}", st["meta"]))
    story.append(Spacer(1, 0.4*cm))

    # Таблиця бланку
    cols = 10
    rows_needed = (variant.total_questions + cols - 1) // cols
    table_data = [["№"] + [str(j) for j in range(1, cols + 1)]]
    for r in range(rows_needed):
        num_row = ["№"]
        ans_row = ["Відп."]
        for j in range(cols):
            q_num = r * cols + j + 1
            num_row.append(str(q_num) if q_num <= variant.total_questions else "")
            ans_row.append("")
        table_data.append(num_row)
        table_data.append(ans_row)

    col_widths = [1.5*cm] + [1.5*cm] * cols
    tbl = Table(table_data, colWidths=col_widths)
    tbl.setStyle(TableStyle([
        ("GRID",      (0, 0), (-1, -1), 0.5, colors.black),
        ("FONTNAME",  (0, 0), (-1, -1), _FONT_REGULAR),
        ("FONTSIZE",  (0, 0), (-1, -1), 9),
        ("ALIGN",     (0, 0), (-1, -1), "CENTER"),
        ("VALIGN",    (0, 0), (-1, -1), "MIDDLE"),
        ("ROWHEIGHT", (0, 0), (-1, -1), 0.7*cm),
        ("BACKGROUND",(0, 0), (-1, 0),  colors.HexColor("#D0D0D0")),
        ("FONTNAME",  (0, 0), (-1, 0),  _FONT_BOLD),
    ]))
    story.append(tbl)

    doc.build(story)
    return filepath


# ─── Ключ відповідей ──────────────────────────────────────────────────────────

def export_answer_key_pdf(
    variant: TestVariant,
    output_dir: str,
    filename: str | None = None,
) -> str:
    """
    Генерує PDF-файл із ключем відповідей для викладача.

    Returns:
        Шлях до збереженого PDF файлу
    """
    os.makedirs(output_dir, exist_ok=True)

    safe_name = variant.student_name.replace(" ", "_").replace("/", "-")
    if filename is None:
        filename = f"Відповіді_{variant.category_code}_{safe_name}"
    filepath = os.path.join(output_dir, f"{filename}.pdf")

    doc = SimpleDocTemplate(
        filepath, pagesize=A4,
        leftMargin=3*cm, rightMargin=2*cm,
        topMargin=2*cm, bottomMargin=2*cm,
    )

    st = _get_styles()
    story = []

    story.append(Paragraph("ТЕСТ PART 147 — КЛЮЧ ВІДПОВІДЕЙ", st["title"]))
    story.append(Paragraph(f"Категорія: {variant.category_code}", st["subtitle"]))
    story.append(HRFlowable(width="100%", thickness=1, color=colors.grey))
    story.append(Spacer(1, 0.3*cm))

    story.append(Paragraph(f"<b>Здобувач:</b> {variant.student_name}", st["meta"]))
    story.append(Paragraph(f"<b>Дата:</b> {variant.generated_at}", st["meta"]))
    story.append(Spacer(1, 0.5*cm))

    # Таблиця відповідей
    answer_labels = {"A": "А", "B": "В", "C": "С"}
    table_data = [["№", "Відповідь", "Розділ", "Модуль", "ID"]]
    for i, q in enumerate(variant.questions, start=1):
        table_data.append([
            str(i),
            answer_labels.get(q.correct_answer, q.correct_answer),
            f"{q.section_code} {q.section_name[:30]}",
            f"М{q.module_code}",
            str(q.id),
        ])

    col_w = [1.2*cm, 2.2*cm, 7*cm, 2.5*cm, 1.5*cm]
    tbl = Table(table_data, colWidths=col_w)
    tbl.setStyle(TableStyle([
        ("GRID",       (0, 0), (-1, -1), 0.5, colors.black),
        ("FONTNAME",   (0, 0), (-1, -1), _FONT_REGULAR),
        ("FONTSIZE",   (0, 0), (-1, -1), 9),
        ("ALIGN",      (0, 0), (1, -1),  "CENTER"),
        ("ALIGN",      (2, 0), (-1, -1), "LEFT"),
        ("VALIGN",     (0, 0), (-1, -1), "MIDDLE"),
        ("ROWHEIGHT",  (0, 0), (-1, -1), 0.65*cm),
        ("BACKGROUND", (0, 0), (-1, 0),  colors.HexColor("#2C5F8A")),
        ("TEXTCOLOR",  (0, 0), (-1, 0),  colors.white),
        ("FONTNAME",   (0, 0), (-1, 0),  _FONT_BOLD),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1),
         [colors.white, colors.HexColor("#F0F4F8")]),
    ]))
    story.append(tbl)

    doc.build(story)
    return filepath


if __name__ == "__main__":
    import sys, os
    sys.path.insert(0, os.path.dirname(__file__))
    from generator import generate_test

    v = generate_test("B1.1", "Тестовий Здобувач")
    out = os.path.join(os.path.dirname(os.path.dirname(__file__)), "output")

    sf = export_student_pdf(v, out)
    af = export_answer_key_pdf(v, out)
    print(f"Тест PDF:     {sf}")
    print(f"Відповіді PDF: {af}")
