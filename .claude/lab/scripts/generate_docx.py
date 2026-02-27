#!/usr/bin/env python3
"""Генератор ГОСТ-совместимого Word-документа из report.md.

Аргументы:
    --config PATH      config.json (обязательно)
    --report PATH      report.md (обязательно)
    --output PATH      Выходной файл .docx (обязательно)
    --images PATH      Директория с изображениями (опционально)
    --image-map PATH   image_map.json — маппинг номер→файл (опционально)
    --template PATH    Пример .docx — используется как шаблон стилей и титульного листа (опционально)
"""

import argparse
import json
import re
import sys
from pathlib import Path

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn


# ---------------------------------------------------------------------------
# Константы ГОСТ
# ---------------------------------------------------------------------------
FONT_NAME = "Times New Roman"
FONT_SIZE_NORMAL = Pt(14)
FONT_SIZE_HEADING1 = Pt(16)
FONT_SIZE_HEADING2 = Pt(14)
FONT_SIZE_CAPTION = Pt(12)
FONT_SIZE_TITLE = Pt(16)
LINE_SPACING = 1.5
LINE_SPACING_SINGLE = 1.0
FIRST_LINE_INDENT = Cm(1.25)
PAGE_MARGIN = Cm(2)
MAX_IMAGE_WIDTH = Cm(17)

# Регулярные выражения
RE_HEADING1 = re.compile(r"^(\d+)\.\s+(.+)$")
RE_HEADING2 = re.compile(r"^(\d+\.\d+)\.?\s+(.+)$")
RE_PLACEHOLDER = re.compile(r"\[ВСТАВИТЬ РИСУНОК (\d+) ЗДЕСЬ\]")
RE_FIGURE_CAPTION = re.compile(r"^Рисунок (\d+)\.\s+(.+)$")
RE_TABLE_CAPTION = re.compile(r"^Таблица (\d+)\.\s+(.+)$")
RE_BOLD = re.compile(r"\*\*(.+?)\*\*")
RE_ITALIC = re.compile(r"\*(.+?)\*")


# ---------------------------------------------------------------------------
# Настройка стилей
# ---------------------------------------------------------------------------
def setup_styles(doc: Document):
    """Настроить стили документа по ГОСТ."""
    style = doc.styles["Normal"]
    font = style.font
    font.name = FONT_NAME
    font.size = FONT_SIZE_NORMAL
    font.color.rgb = RGBColor(0, 0, 0)

    pf = style.paragraph_format
    pf.line_spacing = LINE_SPACING
    pf.first_line_indent = FIRST_LINE_INDENT
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.left_indent = Cm(0)
    pf.right_indent = Cm(0)
    pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    # Heading 1
    h1 = doc.styles["Heading 1"]
    h1.font.name = FONT_NAME
    h1.font.size = FONT_SIZE_HEADING1
    h1.font.bold = True
    h1.font.color.rgb = RGBColor(0, 0, 0)
    h1.paragraph_format.space_before = Pt(12)
    h1.paragraph_format.space_after = Pt(6)
    h1.paragraph_format.first_line_indent = Cm(0)
    h1.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT

    # Heading 2
    h2 = doc.styles["Heading 2"]
    h2.font.name = FONT_NAME
    h2.font.size = FONT_SIZE_HEADING2
    h2.font.bold = True
    h2.font.color.rgb = RGBColor(0, 0, 0)
    h2.paragraph_format.space_before = Pt(12)
    h2.paragraph_format.space_after = Pt(6)
    h2.paragraph_format.first_line_indent = Cm(0)
    h2.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT

    # FigureCaption
    if "FigureCaption" not in [s.name for s in doc.styles]:
        fc = doc.styles.add_style("FigureCaption", 1)  # 1 = WD_STYLE_TYPE.PARAGRAPH
    else:
        fc = doc.styles["FigureCaption"]
    fc.font.name = FONT_NAME
    fc.font.size = FONT_SIZE_CAPTION
    fc.font.color.rgb = RGBColor(0, 0, 0)
    fc.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
    fc.paragraph_format.space_before = Pt(6)
    fc.paragraph_format.space_after = Pt(0)
    fc.paragraph_format.first_line_indent = Cm(0)

    # TableCaption
    if "TableCaption" not in [s.name for s in doc.styles]:
        tc = doc.styles.add_style("TableCaption", 1)
    else:
        tc = doc.styles["TableCaption"]
    tc.font.name = FONT_NAME
    tc.font.size = FONT_SIZE_CAPTION
    tc.font.color.rgb = RGBColor(0, 0, 0)
    tc.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    tc.paragraph_format.space_before = Pt(0)
    tc.paragraph_format.space_after = Pt(6)
    tc.paragraph_format.first_line_indent = Cm(0)


def setup_page(doc: Document):
    """Настроить поля страницы."""
    for section in doc.sections:
        section.top_margin = PAGE_MARGIN
        section.bottom_margin = PAGE_MARGIN
        section.left_margin = PAGE_MARGIN
        section.right_margin = PAGE_MARGIN


# ---------------------------------------------------------------------------
# Титульный лист
# ---------------------------------------------------------------------------
def add_title_page(doc: Document, config: dict):
    """Создать титульный лист из config.json."""
    student = config.get("student", {})
    institution = config.get("institution", {})
    course = config.get("course", {})

    # Верх: название учебного заведения
    inst_name = institution.get("name", "Учебное заведение")
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = LINE_SPACING_SINGLE
    p.paragraph_format.first_line_indent = Cm(0)
    run = p.add_run(inst_name)
    run.font.name = FONT_NAME
    run.font.size = FONT_SIZE_NORMAL
    run.font.bold = True

    # Пустое пространство
    for _ in range(6):
        sp = doc.add_paragraph()
        sp.paragraph_format.space_before = Pt(0)
        sp.paragraph_format.space_after = Pt(0)
        sp.paragraph_format.first_line_indent = Cm(0)

    # Центр: название работы
    course_name = course.get("name", "Предмет")
    lab_number = course.get("lab_number", "")
    lab_title = course.get("lab_title", "Лабораторная работа")

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.first_line_indent = Cm(0)
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    run = p.add_run(course_name.upper())
    run.font.name = FONT_NAME
    run.font.size = FONT_SIZE_TITLE
    run.font.bold = True

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.first_line_indent = Cm(0)
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(0)
    lab_text = f"Лабораторная работа № {lab_number}" if lab_number else "Лабораторная работа"
    run = p.add_run(lab_text)
    run.font.name = FONT_NAME
    run.font.size = FONT_SIZE_TITLE
    run.font.bold = True

    if lab_title:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.first_line_indent = Cm(0)
        p.paragraph_format.space_before = Pt(6)
        p.paragraph_format.space_after = Pt(0)
        run = p.add_run(f"«{lab_title}»")
        run.font.name = FONT_NAME
        run.font.size = FONT_SIZE_TITLE
        run.font.bold = True

    # Пустое пространство
    for _ in range(6):
        sp = doc.add_paragraph()
        sp.paragraph_format.space_before = Pt(0)
        sp.paragraph_format.space_after = Pt(0)
        sp.paragraph_format.first_line_indent = Cm(0)

    # Правый нижний угол: данные студента
    student_name = student.get("name", "Студент")
    student_group = student.get("group", "Группа")

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p.paragraph_format.first_line_indent = Cm(0)
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = LINE_SPACING_SINGLE
    run = p.add_run(f"Выполнил: {student_name}")
    run.font.name = FONT_NAME
    run.font.size = FONT_SIZE_NORMAL

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p.paragraph_format.first_line_indent = Cm(0)
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = LINE_SPACING_SINGLE
    run = p.add_run(f"Группа: {student_group}")
    run.font.name = FONT_NAME
    run.font.size = FONT_SIZE_NORMAL

    # Пустое пространство до низа
    for _ in range(4):
        sp = doc.add_paragraph()
        sp.paragraph_format.space_before = Pt(0)
        sp.paragraph_format.space_after = Pt(0)
        sp.paragraph_format.first_line_indent = Cm(0)

    # Низ по центру: город, год
    city = institution.get("city", "Город")
    import datetime
    year = datetime.datetime.now().year

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.first_line_indent = Cm(0)
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    run = p.add_run(f"{city}, {year}")
    run.font.name = FONT_NAME
    run.font.size = FONT_SIZE_NORMAL
    run.font.bold = True

    # Разрыв страницы
    doc.add_page_break()


# ---------------------------------------------------------------------------
# Вставка изображения
# ---------------------------------------------------------------------------
def insert_image(doc: Document, image_path: str):
    """Вставить изображение с масштабированием."""
    from PIL import Image

    img = Image.open(image_path)
    width_px, height_px = img.size

    # Масштабирование: max 17 см по ширине
    max_width_cm = 17.0
    dpi = img.info.get("dpi", (96, 96))
    dpi_x = dpi[0] if isinstance(dpi, tuple) else dpi
    if dpi_x == 0:
        dpi_x = 96

    width_cm = width_px / dpi_x * 2.54
    if width_cm > max_width_cm:
        scale = max_width_cm / width_cm
        width_cm = max_width_cm
    else:
        scale = 1.0

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.first_line_indent = Cm(0)
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(0)
    run = p.add_run()
    run.add_picture(image_path, width=Cm(width_cm))


# ---------------------------------------------------------------------------
# Парсинг и добавление текста
# ---------------------------------------------------------------------------
def add_formatted_run(paragraph, text: str):
    """Добавить текст с базовым форматированием (жирный, курсив)."""
    # Разбить текст на сегменты с форматированием
    parts = re.split(r"(\*\*.*?\*\*|\*.*?\*)", text)
    for part in parts:
        if part.startswith("**") and part.endswith("**"):
            run = paragraph.add_run(part[2:-2])
            run.font.name = FONT_NAME
            run.bold = True
        elif part.startswith("*") and part.endswith("*") and len(part) > 2:
            run = paragraph.add_run(part[1:-1])
            run.font.name = FONT_NAME
            run.italic = True
        else:
            run = paragraph.add_run(part)
            run.font.name = FONT_NAME


def process_report(doc: Document, report_text: str, image_map: dict, images_dir: Path | None):
    """Обработать report.md и добавить содержимое в документ."""
    lines = report_text.split("\n")
    i = 0

    while i < len(lines):
        line = lines[i].rstrip()

        # Пустая строка — пропускаем
        if not line:
            i += 1
            continue

        # Подзаголовок (2-й уровень, проверяем до заголовка 1-го)
        m = RE_HEADING2.match(line)
        if m:
            p = doc.add_paragraph(style="Heading 2")
            p.add_run(f"{m.group(1)}. {m.group(2)}")
            for run in p.runs:
                run.font.name = FONT_NAME
            i += 1
            continue

        # Заголовок (1-й уровень)
        m = RE_HEADING1.match(line)
        if m:
            p = doc.add_paragraph(style="Heading 1")
            p.add_run(f"{m.group(1)}. {m.group(2)}")
            for run in p.runs:
                run.font.name = FONT_NAME
            i += 1
            continue

        # Плейсхолдер для рисунка
        m = RE_PLACEHOLDER.search(line)
        if m:
            fig_num = m.group(1)
            img_file = image_map.get(fig_num) or image_map.get(int(fig_num))
            if img_file and images_dir:
                img_path = images_dir / img_file if not Path(img_file).is_absolute() else Path(img_file)
                if img_path.exists():
                    insert_image(doc, str(img_path))
                else:
                    p = doc.add_paragraph()
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    p.paragraph_format.first_line_indent = Cm(0)
                    run = p.add_run(f"[Изображение не найдено: {img_file}]")
                    run.font.name = FONT_NAME
                    run.font.color.rgb = RGBColor(255, 0, 0)
            else:
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.paragraph_format.first_line_indent = Cm(0)
                run = p.add_run(line)
                run.font.name = FONT_NAME
                run.font.color.rgb = RGBColor(128, 128, 128)
            i += 1
            continue

        # Подпись рисунка
        m = RE_FIGURE_CAPTION.match(line)
        if m:
            p = doc.add_paragraph(style="FigureCaption")
            p.add_run(line)
            for run in p.runs:
                run.font.name = FONT_NAME
            i += 1
            continue

        # Подпись таблицы
        m = RE_TABLE_CAPTION.match(line)
        if m:
            p = doc.add_paragraph(style="TableCaption")
            p.add_run(line)
            for run in p.runs:
                run.font.name = FONT_NAME
            i += 1
            continue

        # Обычный абзац
        p = doc.add_paragraph(style="Normal")
        add_formatted_run(p, line)
        i += 1


# ---------------------------------------------------------------------------
# Основная логика
# ---------------------------------------------------------------------------
def main():
    parser = argparse.ArgumentParser(
        description="Генератор ГОСТ-совместимого Word-документа из report.md"
    )
    parser.add_argument("--config", required=True, help="Путь к config.json")
    parser.add_argument("--report", required=True, help="Путь к report.md")
    parser.add_argument("--output", required=True, help="Путь к выходному файлу .docx")
    parser.add_argument("--images", default=None, help="Директория с изображениями")
    parser.add_argument("--image-map", default=None, help="Путь к image_map.json")
    parser.add_argument("--template", default=None, help="Пример .docx — шаблон стилей и титульного листа")

    args = parser.parse_args()

    # Загрузить config
    config_path = Path(args.config)
    if not config_path.exists():
        print(f"Файл конфигурации не найден: {config_path}", file=sys.stderr)
        sys.exit(1)

    with open(config_path, encoding="utf-8") as f:
        config = json.load(f)

    # Загрузить report.md
    report_path = Path(args.report)
    if not report_path.exists():
        print(f"Файл отчёта не найден: {report_path}", file=sys.stderr)
        sys.exit(1)

    report_text = report_path.read_text(encoding="utf-8")

    # Загрузить image_map
    image_map = {}
    if args.image_map:
        map_path = Path(args.image_map)
        if map_path.exists():
            with open(map_path, encoding="utf-8") as f:
                image_map = json.load(f)

    images_dir = Path(args.images) if args.images else None

    # Создать документ
    if args.template:
        template_path = Path(args.template)
        if template_path.exists():
            doc = Document(str(template_path))
            # Пример уже содержит титульный лист — начинаем с нового раздела
            has_title = True
        else:
            print(
                f"Пример не найден: {template_path}. Создаю стандартный титульный лист.",
                file=sys.stderr,
            )
            doc = Document()
            has_title = False
    else:
        doc = Document()
        has_title = False

    # Настройка
    setup_styles(doc)
    setup_page(doc)

    # Титульный лист (если нет примера-шаблона)
    if not has_title:
        add_title_page(doc, config)

    # Содержание отчёта
    process_report(doc, report_text, image_map, images_dir)

    # Сохранить
    output_path = Path(args.output)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(output_path))
    print(f"Документ сохранён: {output_path}")


if __name__ == "__main__":
    main()
