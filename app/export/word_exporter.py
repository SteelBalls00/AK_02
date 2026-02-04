from docx import Document
from docx.shared import Inches, Pt
from docx.enum.section import WD_ORIENTATION
from datetime import datetime
import os

from app.export.word_templates import WORD_district_first_TEMPLATES, WORD_TEMPLATES

from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import Pt


def export_model_to_word(model, processor, court, week):
    template_key = processor.word_template_key
    if not template_key:
        raise ValueError("У процессора не задан word_template_key")

    templates = WORD_TEMPLATES.get(template_key)
    if not templates:
        raise ValueError(f"Не найден Word-шаблон: {template_key}")

    specialization = processor.get_specialization()

    tpl = templates.get(specialization)

    if not tpl:
        raise ValueError(
            f"Нет Word-шаблона для specialization={specialization}, "
            f"processor={processor.__class__.__name__}"
        )

    document = Document()

    # --- Альбомный лист ---
    section = document.sections[-1]
    section.orientation = WD_ORIENTATION.LANDSCAPE
    section.page_width, section.page_height = section.page_height, section.page_width

    # --- Узкие поля ---
    for attr in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
        setattr(section, attr, Inches(0.5))

    # --- Заголовок ---
    document.add_paragraph(f"{court} ({specialization}) — {week}")

    rows = model.rowCount()
    cols = model.columnCount()

    table = document.add_table(rows=1, cols=cols)
    table.style = "Table Grid"

    # --- Шапка ---
    for c in range(cols):
        headers = tpl.get("headers") if tpl else None

        for c in range(cols):
            if headers and c < len(headers):
                table.cell(0, c).text = headers[c]
            else:
                table.cell(0, c).text = model.headerData(c, 1)



    # --- Данные ---
    for r in range(rows):
        row = table.add_row().cells
        for c in range(cols):
            row[c].text = str(model.data(model.index(r, c)))

    # --- Объединения столбцов ---
    if tpl:
        for (r1, c1), (r2, c2), text in tpl["merge"]:
            cell = table.cell(r1, c1)
            cell.merge(table.cell(r2, c2))
            cell.text = text

    # Жирная первая строка
    for cell in table.rows[0].cells:
        for p in cell.paragraphs:
            p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            for run in p.runs:
                run.font.size = Pt(9)
                run.bold = True

    # --- Шрифт ---
    for row in table.rows:
        for cell in row.cells:
            for run in cell.paragraphs[0].runs:
                run.font.size = Pt(8)

    # --- Жирная последняя строка ---
    for cell in table.rows[-1].cells:
        for run in cell.paragraphs[0].runs:
            run.bold = True

    # --- Сохранение ---
    filename = f"big_table_{datetime.now():%d.%m.%Y.%H.%M.%S}.docx"
    document.save(filename)
    os.startfile(filename)
