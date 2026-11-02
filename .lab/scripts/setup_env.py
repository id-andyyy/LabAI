#!/usr/bin/env python3
"""Подготовить или проверить локальное Python-окружение LabAI."""

import argparse
from pathlib import Path
import shutil
import subprocess
import sys
import venv


LAB_DIR = Path(__file__).resolve().parents[1]
VENV_DIR = LAB_DIR / "venv"
VENV_PYTHON = VENV_DIR / ("Scripts/python.exe" if sys.platform == "win32" else "bin/python")
REQUIREMENTS = LAB_DIR / "scripts" / "requirements.txt"


def find_libreoffice():
    """Найти CLI LibreOffice, включая стандартную установку на macOS."""
    for command in ("libreoffice", "soffice"):
        executable = shutil.which(command)
        if executable:
            return executable
    macos_path = Path("/Applications/LibreOffice.app/Contents/MacOS/soffice")
    return str(macos_path) if macos_path.is_file() else None


def check_environment():
    if not VENV_PYTHON.is_file():
        print("Окружение не найдено. Запусти setup_env.py без --check.", file=sys.stderr)
        return 1
    subprocess.run([str(VENV_PYTHON), "--version"], check=True)
    subprocess.run(
        [str(VENV_PYTHON), "-c",
         "from importlib.metadata import version; "
         "packages = ('python-docx', 'pypdf', 'Pillow', 'matplotlib'); "
         "print('Зависимости: ' + ', '.join(p + ' ' + version(p) for p in packages))"],
        check=True,
    )
    subprocess.run([str(VENV_PYTHON), "-m", "pip", "check"], check=True)
    print(f"Python LabAI: {VENV_PYTHON}")
    office = find_libreoffice()
    if office:
        print(f"Конвертация в PDF: {office}")
    else:
        print("LibreOffice не найден: доступны DOCX и MD. Для PDF установи LibreOffice.")
    return 0


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--check", action="store_true", help="Только проверить окружение без установки")
    args = parser.parse_args()
    if sys.version_info < (3, 10):
        print("Для LabAI нужен Python 3.10 или новее.", file=sys.stderr)
        return 1
    try:
        if not args.check:
            if not VENV_PYTHON.is_file():
                print(f"Создание окружения: {VENV_DIR}", flush=True)
                venv.EnvBuilder(with_pip=True).create(VENV_DIR)
            subprocess.run(
                [str(VENV_PYTHON), "-m", "pip", "install", "-r", str(REQUIREMENTS)],
                check=True,
            )
        return check_environment()
    except (OSError, subprocess.CalledProcessError) as error:
        print(f"Не удалось подготовить или проверить окружение: {error}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
