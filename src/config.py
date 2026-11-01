"""Central configuration for Noora PPT Builder.

The module deliberately avoids extra dependencies so the repository stays
standalone.  It reads a local ``.env`` file when present and otherwise uses
normal environment variables.
"""

from __future__ import annotations

import os
import shutil
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]


def _load_dotenv() -> None:
    """Load simple KEY=VALUE pairs from ``.env`` without python-dotenv."""

    env_path = PROJECT_ROOT / ".env"
    if not env_path.exists():
        return

    for raw_line in env_path.read_text(encoding="utf-8", errors="ignore").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#"):
            continue
        if line.lower().startswith("export "):
            line = line[7:].strip()
        if "=" not in line:
            continue

        key, value = line.split("=", 1)
        key = key.strip()
        value = value.strip().strip('"').strip("'")
        if key and key not in os.environ:
            os.environ[key] = value


_load_dotenv()

# ---------------------------------------------------------------------------
# OCR
# ---------------------------------------------------------------------------
# Standard Windows installation path.  Override with the TESSERACT_CMD
# environment variable if Tesseract lives somewhere else.
TESSERACT_CMD = os.environ.get(
    "TESSERACT_CMD", r"C:\Program Files\Tesseract-OCR\tesseract.exe"
)

try:
    import pytesseract

    pytesseract.pytesseract.tesseract_cmd = TESSERACT_CMD
except Exception:  # pragma: no cover - import errors are surfaced by the CLI
    pytesseract = None  # type: ignore[assignment]


def resolve_tesseract_cmd() -> str | None:
    """Return a usable Tesseract executable path, or ``None``."""

    if Path(TESSERACT_CMD).is_file():
        return TESSERACT_CMD

    found = shutil.which("tesseract")
    if found:
        return found

    return None


def tesseract_available() -> bool:
    return resolve_tesseract_cmd() is not None


# ---------------------------------------------------------------------------
# Gemini / MCQ generation
# ---------------------------------------------------------------------------
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "").strip()
GEMINI_MODEL = os.environ.get("GEMINI_MODEL", "gemini-2.5-flash").strip()

# ---------------------------------------------------------------------------
# Output / presentation defaults
# ---------------------------------------------------------------------------
DEFAULT_OUTDIR = PROJECT_ROOT / "output"
SLIDE_WIDTH_IN = 13.333  # 16:9
SLIDE_HEIGHT_IN = 7.5
SLIDE_IMAGE_WIDTH = 1920
SLIDE_IMAGE_HEIGHT = 1080

# Text budgets used by the parser and slide planner.
MAX_OCR_IMAGE_PIXELS = 2600  # downscale very large scans before OCR
MIN_OCR_IMAGE_SIDE = 64  # ignore 1x1 spacers and similar artefacts
MIN_FIGURE_SIDE = 300  # figures smaller than this are not shown on slides
