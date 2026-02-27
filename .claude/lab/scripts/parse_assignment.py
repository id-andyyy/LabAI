#!/usr/bin/env python3
"""Парсер заданий лабораторных работ.

Извлекает текст из PDF, DOCX, TXT и MD файлов.
Вывод: plain text в stdout (дословно, без правок).
"""

import sys
from pathlib import Path


def parse_pdf(file_path: Path) -> str:
    """Извлечь текст из PDF-файла."""
    from pypdf import PdfReader

    reader = PdfReader(str(file_path))
    pages = []
    for i, page in enumerate(reader.pages):
        text = page.extract_text()
        if text:
            pages.append(text)
    return "\n\n".join(pages)


def parse_docx(file_path: Path) -> str:
    """Извлечь текст из DOCX-файла."""
    from docx import Document

    doc = Document(str(file_path))
    parts = []

    for para in doc.paragraphs:
        text = para.text.strip()
        if text:
            parts.append(text)

    for table in doc.tables:
        for row in table.rows:
            cells = [cell.text.strip() for cell in row.cells]
            parts.append(" | ".join(cells))

    return "\n\n".join(parts)


def parse_text(file_path: Path) -> str:
    """Прочитать текстовый файл (TXT/MD)."""
    return file_path.read_text(encoding="utf-8")


PARSERS = {
    ".pdf": parse_pdf,
    ".docx": parse_docx,
    ".txt": parse_text,
    ".md": parse_text,
}


def main():
    if len(sys.argv) != 2:
        print("Использование: parse_assignment.py <путь_к_файлу>", file=sys.stderr)
        sys.exit(1)

    file_path = Path(sys.argv[1])

    if not file_path.exists():
        print(f"Файл не найден: {file_path}", file=sys.stderr)
        sys.exit(1)

    ext = file_path.suffix.lower()
    parser = PARSERS.get(ext)

    if parser is None:
        supported = ", ".join(PARSERS.keys())
        print(
            f"Неподдерживаемый формат: {ext}. Поддерживаются: {supported}",
            file=sys.stderr,
        )
        sys.exit(1)

    try:
        text = parser(file_path)
        print(text)
    except Exception as e:
        print(f"Ошибка при обработке файла {file_path}: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
