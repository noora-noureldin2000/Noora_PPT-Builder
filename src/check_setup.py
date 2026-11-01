"""Environment check: run this before your first deck.

    python src/check_setup.py
"""

from __future__ import annotations

import importlib
import shutil
import sys
from pathlib import Path

if __package__ in (None, ""):
    sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from src.config import (  # noqa: E402
    GEMINI_API_KEY,
    GEMINI_MODEL,
    PROJECT_ROOT,
    resolve_tesseract_cmd,
)


def _line(status: str, label: str, detail: str = "") -> None:
    print(f"[{status:<4}] {label}" + (f"  {detail}" if detail else ""))


def check_python() -> bool:
    ok = sys.version_info >= (3, 10)
    _line("OK" if ok else "FAIL", "Python", sys.version.split()[0] + " (need 3.10+)")
    return ok


def check_packages() -> bool:
    packages = [
        ("pptx", "python-pptx"),
        ("fitz", "PyMuPDF"),
        ("docx", "python-docx"),
        ("PIL", "Pillow"),
        ("pytesseract", "pytesseract"),
    ]
    all_ok = True
    for module_name, label in packages:
        try:
            module = importlib.import_module(module_name)
            version = getattr(module, "__version__", "")
            _line("OK", label, str(version))
        except Exception as exc:
            all_ok = False
            _line("FAIL", label, str(exc))
    return all_ok


def check_optional() -> None:
    try:
        from google import genai  # noqa: F401

        _line("OK", "google-genai", "installed")
    except Exception:
        _line("WARN", "google-genai", "missing - MCQs will not run")

    try:
        import win32com.client  # noqa: F401

        _line("OK", "pywin32", "PowerPoint COM export available")
    except Exception:
        _line("WARN", "pywin32", "missing - slide PNGs fall back to LibreOffice/Pillow")


def check_tesseract() -> bool:
    cmd = resolve_tesseract_cmd()
    if cmd:
        _line("OK", "Tesseract", cmd)
        return True
    _line(
        "FAIL",
        "Tesseract",
        "not found. Install it from https://github.com/UB-Mannheim/tesseract/wiki "
        r"and keep the default path C:\Program Files\Tesseract-OCR\tesseract.exe",
    )
    return False


def check_rendering() -> None:
    from src.render import available_backends

    backends = available_backends()
    if backends.get("powerpoint"):
        _line("OK", "Rendering", "PowerPoint COM")
    elif backends.get("libreoffice"):
        _line("OK", "Rendering", "LibreOffice headless")
    else:
        _line("WARN", "Rendering", "Pillow preview only (approximate)")

    soffice = shutil.which("soffice")
    if soffice:
        _line("OK", "LibreOffice", soffice)


def check_gemini() -> None:
    if GEMINI_API_KEY:
        _line("OK", "GEMINI_API_KEY", f"set (model: {GEMINI_MODEL})")
    else:
        _line(
            "WARN",
            "GEMINI_API_KEY",
            "not set - slide generation works, MCQ generation needs the key",
        )


def check_output() -> None:
    output = PROJECT_ROOT / "output"
    try:
        output.mkdir(parents=True, exist_ok=True)
        probe = output / ".write_test"
        probe.write_text("ok", encoding="utf-8")
        probe.unlink()
        _line("OK", "Output folder", str(output))
    except Exception as exc:
        _line("FAIL", "Output folder", str(exc))


def main() -> int:
    print("Noora PPT Builder - environment check")
    print("=" * 60)
    results = [check_python(), check_packages(), check_tesseract()]
    check_optional()
    check_rendering()
    check_gemini()
    check_output()
    print("=" * 60)

    if all(results):
        print("Core requirements are satisfied.")
        return 0
    print("Fix the FAIL lines above, then run the check again.")
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
