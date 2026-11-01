"""Slide-image rendering backends.

Order of preference:

1. **Microsoft PowerPoint (Windows COM)** - pixel-perfect, uses the copy of
   PowerPoint already installed on the machine.
2. **LibreOffice headless** - cross-platform fallback; converts the deck to
   PDF and rasterises each page with PyMuPDF.
3. **Pillow preview** - always available, approximate.  It re-draws the text
   of every slide so the pipeline still produces preview images on machines
   with neither PowerPoint nor LibreOffice.

The public entry point is :func:`export_slide_images`.
"""

from __future__ import annotations

import logging
import re
import shutil
import subprocess
import tempfile
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

LOGGER = logging.getLogger(__name__)

DEFAULT_WIDTH = 1920
DEFAULT_HEIGHT = 1080

_BACKGROUND = (248, 249, 251)
_INK = (20, 23, 28)
_MUTED = (94, 104, 116)
_ACCENT = (14, 124, 134)

_FONT_CANDIDATES = {
    "regular": [
        "C:/Windows/Fonts/segoeui.ttf",
        "C:/Windows/Fonts/calibri.ttf",
        "C:/Windows/Fonts/arial.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/System/Library/Fonts/Supplemental/Arial.ttf",
    ],
    "bold": [
        "C:/Windows/Fonts/segoeuib.ttf",
        "C:/Windows/Fonts/calibrib.ttf",
        "C:/Windows/Fonts/arialbd.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
        "/System/Library/Fonts/Supplemental/Arial Bold.ttf",
    ],
}


# ---------------------------------------------------------------------------
# Public entry point
# ---------------------------------------------------------------------------
def export_slide_images(
    pptx_path: str | Path,
    output_dir: str | Path,
    *,
    width: int = DEFAULT_WIDTH,
    backend: str = "auto",
) -> list[Path]:
    """Export ``pptx_path`` to ``slide_XX.png`` files and return their paths."""

    pptx_path = Path(pptx_path).resolve()
    if not pptx_path.is_file():
        raise FileNotFoundError(f"Deck not found: {pptx_path}")

    output_dir = Path(output_dir).expanduser().resolve()
    output_dir.mkdir(parents=True, exist_ok=True)

    requested = (backend or "auto").lower()
    if requested == "auto":
        order = ["powerpoint", "libreoffice", "pillow"]
    else:
        order = [requested, "pillow"] if requested != "pillow" else ["pillow"]

    errors: list[str] = []
    for name in order:
        try:
            if name == "powerpoint":
                files = _export_powerpoint(pptx_path, output_dir, width)
            elif name == "libreoffice":
                files = _export_libreoffice(pptx_path, output_dir, width)
            elif name == "pillow":
                files = _export_pillow_preview(pptx_path, output_dir, width)
            else:
                raise ValueError(f"Unknown backend: {name}")

            if files:
                if name == "pillow":
                    LOGGER.warning(
                        "Slide images are approximate previews. Install PowerPoint "
                        "or LibreOffice for pixel-perfect PNG export."
                    )
                else:
                    LOGGER.info("Rendered %s slides with %s", len(files), name)
                return files
            errors.append(f"{name}: produced no files")
        except Exception as exc:
            LOGGER.debug("Backend %s failed: %s", name, exc)
            errors.append(f"{name}: {exc}")

    raise RuntimeError("Could not render slide images. " + " | ".join(errors))


def available_backends() -> dict[str, bool]:
    """Report which rendering backends are usable on this machine."""

    return {
        "powerpoint": _powerpoint_available(),
        "libreoffice": _find_soffice() is not None,
        "pillow": True,
    }


# ---------------------------------------------------------------------------
# PowerPoint COM
# ---------------------------------------------------------------------------
def _powerpoint_available() -> bool:
    try:
        import win32com.client  # noqa: F401
    except Exception:
        return False
    return True


def _export_powerpoint(pptx_path: Path, output_dir: Path, width: int) -> list[Path]:
    import win32com.client

    height = int(round(width * DEFAULT_HEIGHT / DEFAULT_WIDTH))
    _clear_existing(output_dir)

    app = win32com.client.DispatchEx("PowerPoint.Application")
    presentation = None
    try:
        try:
            app.DisplayAlerts = 1  # ppAlertsNone
        except Exception:
            pass
        presentation = app.Presentations.Open(
            str(pptx_path), ReadOnly=True, Untitled=False, WithWindow=False
        )
        presentation.Export(str(output_dir), "PNG", width, height)
    finally:
        if presentation is not None:
            try:
                presentation.Close()
            except Exception:
                pass
        try:
            app.Quit()
        except Exception:
            pass

    files = sorted(
        [path for path in output_dir.iterdir() if path.suffix.lower() == ".png"],
        key=_slide_sort_key,
    )
    if not files:
        raise RuntimeError("PowerPoint exported no PNG files")
    return [_rename_final(path, output_dir, index) for index, path in enumerate(files, 1)]


def _slide_sort_key(path: Path):
    match = re.search(r"(\d+)", path.stem)
    return (int(match.group(1)) if match else 10**6, path.name.lower())


# ---------------------------------------------------------------------------
# LibreOffice headless
# ---------------------------------------------------------------------------
def _find_soffice() -> str | None:
    found = shutil.which("soffice") or shutil.which("libreoffice")
    if found:
        return found
    common = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        "/usr/bin/soffice",
        "/usr/local/bin/soffice",
    ]
    for candidate in common:
        if Path(candidate).is_file():
            return candidate
    return None


def _export_libreoffice(pptx_path: Path, output_dir: Path, width: int) -> list[Path]:
    soffice = _find_soffice()
    if not soffice:
        raise RuntimeError("LibreOffice (soffice) not found")

    import fitz  # PyMuPDF

    _clear_existing(output_dir)
    with tempfile.TemporaryDirectory(prefix="noora_lo_") as tmp:
        tmp_path = Path(tmp)
        command = [
            soffice,
            "--headless",
            "--norestore",
            "--convert-to",
            "pdf",
            "--outdir",
            str(tmp_path),
            str(pptx_path),
        ]
        process = subprocess.run(
            command,
            capture_output=True,
            text=True,
            timeout=240,
        )
        if process.returncode != 0:
            raise RuntimeError(
                f"LibreOffice failed: {process.stderr.strip()[:300]}"
            )
        pdf_files = list(tmp_path.glob("*.pdf"))
        if not pdf_files:
            raise RuntimeError("LibreOffice produced no PDF")
        pdf_path = pdf_files[0]

        document = fitz.open(pdf_path)
        try:
            files: list[Path] = []
            for index, page in enumerate(document, start=1):
                zoom = width / max(page.rect.width, 1.0)
                pixmap = page.get_pixmap(
                    matrix=fitz.Matrix(zoom, zoom), alpha=False
                )
                target = output_dir / f"slide_{index:02d}.png"
                pixmap.save(str(target))
                files.append(target)
            return files
        finally:
            document.close()


# ---------------------------------------------------------------------------
# Pillow preview fallback
# ---------------------------------------------------------------------------
def _load_font(kind: str, size: int) -> ImageFont.FreeTypeFont | ImageFont.ImageFont:
    for candidate in _FONT_CANDIDATES[kind]:
        if Path(candidate).is_file():
            try:
                return ImageFont.truetype(candidate, size)
            except Exception:
                continue
    try:
        return ImageFont.load_default(size=size)
    except TypeError:  # Pillow < 10
        return ImageFont.load_default()


def _wrap(draw, text, font, max_width):
    words = str(text).split()
    if not words:
        return [""]
    lines: list[str] = []
    current = words[0]
    for word in words[1:]:
        candidate = f"{current} {word}"
        if draw.textlength(candidate, font=font) <= max_width:
            current = candidate
        else:
            lines.append(current)
            current = word
    lines.append(current)
    return lines


def _slide_text_groups(slide):
    """Return (title, body_lines, is_dark) for the Pillow preview."""

    background = _BACKGROUND
    dark = False
    try:
        fill = slide.background.fill
        if fill.type is not None:
            rgb = fill.fore_color.rgb
            background = (rgb[0], rgb[1], rgb[2])
            luminance = 0.299 * background[0] + 0.587 * background[1] + 0.114 * background[2]
            dark = luminance < 110
    except Exception:
        pass

    entries = []
    for shape in slide.shapes:
        if not getattr(shape, "has_text_frame", False):
            continue
        text = shape.text_frame.text.strip()
        if not text:
            continue
        max_size = 0.0
        for paragraph in shape.text_frame.paragraphs:
            for run in paragraph.runs:
                if run.font.size is not None:
                    max_size = max(max_size, run.font.size.pt)
        entries.append((max_size, shape.top or 0, shape.left or 0, text))

    if not entries:
        return background, dark, "", []

    entries.sort(key=lambda item: (-item[0], item[1]))
    title = entries[0][3].replace("\n", " ").strip()
    body: list[str] = []
    for size, _, _, text in entries[1:]:
        # Kickers, footers and page numbers are small; they are furniture, not
        # slide content.
        if size and size < 12:
            continue
        for line in text.splitlines():
            line = line.strip()
            if not line or line == title or line.isdigit():
                continue
            body.append(line)

    # De-duplicate while keeping order.
    seen = set()
    unique = []
    for line in body:
        key = line.lower()
        if key in seen:
            continue
        seen.add(key)
        unique.append(line)
    return background, dark, title, unique[:8]


def _export_pillow_preview(pptx_path: Path, output_dir: Path, width: int) -> list[Path]:
    from pptx import Presentation

    _clear_existing(output_dir)
    presentation = Presentation(str(pptx_path))
    height = int(round(width * DEFAULT_HEIGHT / DEFAULT_WIDTH))
    scale = width / 1920.0

    files: list[Path] = []
    for index, slide in enumerate(presentation.slides, start=1):
        background, dark, title, body = _slide_text_groups(slide)
        image = Image.new("RGB", (width, height), background)
        draw = ImageDraw.Draw(image)

        ink = (255, 255, 255) if dark else _INK
        muted = (176, 186, 198) if dark else _MUTED
        accent = (127, 209, 216) if dark else _ACCENT

        margin = int(120 * scale)
        draw.rectangle(
            [margin, int(250 * scale), margin + int(110 * scale), int(258 * scale)],
            fill=accent,
        )

        title_font = _load_font("bold", max(20, int(64 * scale)))
        body_font = _load_font("regular", max(14, int(30 * scale)))
        small_font = _load_font("regular", max(10, int(18 * scale)))

        y = int(120 * scale)
        draw.text((margin, y), f"SLIDE {index:02d}", font=small_font, fill=accent)
        y = int(290 * scale)

        title_lines = _wrap(draw, title or "Slide", title_font, width - 2 * margin)[:3]
        for line in title_lines:
            draw.text((margin, y), line, font=title_font, fill=ink)
            y += int(title_font.size * 1.22)

        y += int(40 * scale)
        for item in body[:6]:
            lines = _wrap(draw, item, body_font, width - 2 * margin - int(30 * scale))[:4]
            draw.rectangle(
                [margin, y + int(body_font.size * 0.45), margin + int(10 * scale), y + int(body_font.size * 0.62)],
                fill=accent,
            )
            for line in lines:
                draw.text((margin + int(28 * scale), y), line, font=body_font, fill=muted)
                y += int(body_font.size * 1.32)
            y += int(18 * scale)
            if y > height - int(140 * scale):
                break

        draw.text(
            (width - margin - int(60 * scale), height - int(70 * scale)),
            f"{index:02d}",
            font=small_font,
            fill=muted,
        )
        draw.text(
            (margin, height - int(70 * scale)),
            "Approximate preview - install PowerPoint or LibreOffice for exact rendering",
            font=small_font,
            fill=muted,
        )

        target = output_dir / f"slide_{index:02d}.png"
        image.save(target, format="PNG", optimize=True)
        files.append(target)

    return files


# ---------------------------------------------------------------------------
# Utilities
# ---------------------------------------------------------------------------
def _clear_existing(output_dir: Path) -> None:
    for pattern in ("slide_*.png", "Slide*.PNG", "Slide*.png"):
        for path in output_dir.glob(pattern):
            try:
                path.unlink()
            except Exception:
                pass


def _rename_final(source: Path, output_dir: Path, index: int) -> Path:
    target = output_dir / f"slide_{index:02d}.png"
    try:
        if source.resolve() != target.resolve():
            if target.exists():
                target.unlink()
            source.replace(target)
        return target
    except Exception:
        return source
