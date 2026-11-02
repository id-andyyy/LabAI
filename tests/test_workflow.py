"""Проверки CLI и содержимого документов на временных материалах."""

import json
from pathlib import Path
import re
import subprocess
import sys
import tempfile
import unittest

from docx import Document
from PIL import Image
from pypdf import PdfWriter
from pypdf.generic import DictionaryObject, NameObject, DecodedStreamObject


ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / ".lab" / "scripts"


class WorkflowTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory(prefix="labai test ")
        self.addCleanup(self.temp_dir.cleanup)
        self.work = Path(self.temp_dir.name)
        (self.work / "images").mkdir()
        self.image = self.work / "images" / "скриншот 1.png"
        Image.new("RGB", (120, 60), "green").save(self.image)
        self.config = self.work / "config.json"
        self.config.write_text(json.dumps({"preferences": {"no_title_page": True}}))
        self.report = self.work / "report.md"
        self.report.write_text(
            "## Цель работы\nПроверить сборку отчёта.\n\n"
            "## Ход работы\nПолучен проверяемый результат.\n\n"
            "[ВСТАВИТЬ РИСУНОК 1 ЗДЕСЬ]\nРисунок 1. Результат\n\n"
            "## Выводы\nРезультат подтверждён.\n",
            encoding="utf-8",
        )

    def run_script(self, name, *args, success=True):
        result = subprocess.run(
            [sys.executable, str(SCRIPTS / name), *map(str, args)],
            cwd=self.work, text=True, encoding="utf-8", capture_output=True,
            timeout=30,
        )
        if success:
            self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        else:
            self.assertNotEqual(result.returncode, 0, result.stdout + result.stderr)
        return result

    def generate(self, output, *args, success=True):
        return self.run_script(
            "generate_docx.py", "--config", self.config, "--report", self.report,
            "--output", output, "--no-title-page", *args, success=success,
        )

    def test_parse_markdown_and_text_with_unicode_paths(self):
        for extension in ("md", "txt"):
            with self.subTest(extension=extension):
                source = self.work / f"Задание с пробелами.{extension}"
                source.write_text("Задание: выполнить расчёт.\n", encoding="utf-8")
                result = self.run_script("parse_assignment.py", source)
                self.assertEqual(result.stdout.strip(), source.read_text().strip())

    def test_parse_pdf_text(self):
        source = self.work / "task.pdf"
        writer = PdfWriter()
        page = writer.add_blank_page(width=300, height=300)
        font = DictionaryObject({
            NameObject("/Type"): NameObject("/Font"),
            NameObject("/Subtype"): NameObject("/Type1"),
            NameObject("/BaseFont"): NameObject("/Helvetica"),
        })
        page[NameObject("/Resources")] = DictionaryObject({
            NameObject("/Font"): DictionaryObject({NameObject("/F1"): font})
        })
        stream = DecodedStreamObject()
        stream.set_data(b"BT /F1 12 Tf 30 250 Td (LabAI task text) Tj ET")
        page[NameObject("/Contents")] = stream
        writer.write(source)
        for options in ((), ("--full",)):
            with self.subTest(options=options):
                result = self.run_script("parse_assignment.py", source, *options)
                self.assertIn("LabAI task text", result.stdout)

    def test_parse_multiple_docx_preserves_each_image(self):
        extracted = []
        for index, colour in enumerate(("red", "blue"), start=1):
            Image.new("RGB", (120, 60), colour).save(self.image)
            source = self.work / f"Документ {index}.docx"
            document = Document()
            document.add_paragraph(f"Текст задания {index}")
            document.add_picture(str(self.image))
            document.save(source)
            result = self.run_script("parse_assignment.py", "--full", source)
            self.assertIn(f"Текст задания {index}", result.stdout)
            links = re.findall(r"!\[[^\]]+\]\(<([^>]+)>\)", result.stdout)
            self.assertEqual(len(links), 1, result.stdout)
            image_path = self.work / links[0]
            self.assertTrue(image_path.is_file())
            extracted.append(image_path)
        self.assertNotEqual(extracted[0], extracted[1])
        self.assertNotEqual(extracted[0].read_bytes(), extracted[1].read_bytes())

    def test_docx_embeds_project_relative_basename_and_absolute_images(self):
        for index, (mapped_path, use_directory) in enumerate((
            ("images/скриншот 1.png", True),
            ("скриншот 1.png", True),
            (str(self.image), False),
            ("images/скриншот 1.png", False),
        )):
            with self.subTest(mapped_path=mapped_path, use_directory=use_directory):
                image_map = self.work / "image_map.json"
                image_map.write_text(json.dumps({"1": mapped_path}))
                output = self.work / f"result{index}.docx"
                options = ["--image-map", image_map]
                if use_directory:
                    options.extend(["--images", "images"])
                original = self.report.read_bytes()
                self.generate(output, *options)
                document = Document(output)
                self.assertEqual(len(document.inline_shapes), 1)
                text = "\n".join(p.text for p in document.paragraphs)
                self.assertIn("Получен проверяемый результат.", text)
                self.assertNotIn("[ВСТАВИТЬ РИСУНОК", text)
                self.assertEqual(self.report.read_bytes(), original)

    def test_unmatched_placeholder_remains_visible(self):
        output = self.work / "missing-image.docx"
        self.generate(output)
        document = Document(output)
        self.assertEqual(len(document.inline_shapes), 0)
        self.assertIn("[ВСТАВИТЬ РИСУНОК 1 ЗДЕСЬ]",
                      "\n".join(p.text for p in document.paragraphs))

    def test_template_is_preserved_and_content_inserted(self):
        template = self.work / "Шаблон отчёта.docx"
        document = Document()
        document.add_paragraph("Титульный лист пользователя")
        for heading in ("Цель работы", "Ход работы", "Выводы"):
            document.add_paragraph(heading)
        document.save(template)
        original = template.read_bytes()
        output = self.work / "result.docx"
        self.generate(output, "--template", template)
        self.assertEqual(template.read_bytes(), original)
        text = "\n".join(p.text for p in Document(output).paragraphs)
        self.assertIn("Титульный лист пользователя", text)
        self.assertIn("Получен проверяемый результат.", text)
        self.assertIn("Результат подтверждён.", text)

    def test_missing_template_fails_without_creating_output(self):
        output = self.work / "result.docx"
        result = self.generate(output, "--template", "missing.docx", success=False)
        self.assertIn("Шаблон не найден", result.stderr)
        self.assertFalse(output.exists())

    def test_missing_input_has_readable_error(self):
        result = self.run_script("parse_assignment.py", "missing.pdf", success=False)
        self.assertIn("Файл не найден", result.stderr)
        self.assertNotIn("Traceback", result.stderr)


if __name__ == "__main__":
    unittest.main()
