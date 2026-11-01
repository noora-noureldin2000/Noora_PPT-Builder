"""Document parsing with OCR for PDF and DOCX files.

The parser returns a :class:`ParsedDocument` containing ordered text blocks
(headings, paragraphs, tables and OCR output) plus the extracted figures.
OCR text is inserted into the normal reading flow, so scanned pages and
diagram labels become part of the material used to build slides and MCQs.
"""

from __future__ import annotations

import glob
import hashlib
import io
import logging
import re
from pathlib import Path
from statistics import median

import docx
import fitz  # PyMuPDF
import pytesseract
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph
from PIL import Image, ImageOps

from src.config import (
    MAX_OCR_IMAGE_PIXELS,
    MIN_FIGURE_SIDE,
    MIN_OCR_IMAGE_SIDE,
    resolve_tesseract_cmd,
)
from src.models import DocumentImage, ParsedDocument, TextBlock
from src.textutils import (
    clean_ocr_text,
    looks_meaningful,
    normalize_whitespace,
    short_title,
    strip_page_noise,
    truncate_at_word,
)

LOGGER = logging.getLogger(__name__)

SUPPORTED_SUFFIXES = {".pdf", ".docx"}

_OCR_CACHE: dict[str, str] = {}
_OCR_DISABLED_REASON: str | None = None


# ---------------------------------------------------------------------------
# Input expansion
# ---------------------------------------------------------------------------
def expand_inputs(raw_inputs: list[str]) -> list[Path]:
    """Expand files, directories, and glob patterns into PDF/DOCX paths."""

    found: list[Path] = []
    for raw in raw_inputs:
        candidate = Path(raw).expanduser()
        if candidate.is_dir():
            found.extend(
                path
                for path in candidate.iterdir()
                if path.is_file() and path.suffix.lower() in SUPPORTED_SUFFIXES
            )
            continue

        if candidate.is_file():
            if candidate.suffix.lower() not in SUPPORTED_SUFFIXES:
                raise ValueError(f"Unsupported file type: {candidate}")
            found.append(candidate)
            continue

        matches = [Path(p) for p in glob.glob(raw, recursive=True)]
        matches = [
            path
            for path in matches
            if path.is_file() and path.suffix.lower() in SUPPORTED_SUFFIXES
        ]
        if not matches:
            raise FileNotFoundError(f"No PDF or DOCX file matched: {raw}")
        found.extend(matches)

    unique: list[Path] = []
    seen: set[str] = set()
    for path in found:
        resolved = str(path.resolve())
        if resolved not in seen:
            seen.add(resolved)
            unique.append(path.resolve())
    return sorted(unique, key=lambda p: p.name.lower())


# ---------------------------------------------------------------------------
# OCR helpers
# ---------------------------------------------------------------------------
def _disable_ocr(reason: str) -> None:
    global _OCR_DISABLED_REASON
    if _OCR_DISABLED_REASON is None:
        _OCR_DISABLED_REASON = reason
        LOGGER.warning("OCR disabled: %s", reason)


def _prepare_for_ocr(image: Image.Image) -> Image.Image:
    image = image.convert("RGB")
    width, height = image.size
    longest = max(width, height)
    if longest > MAX_OCR_IMAGE_PIXELS:
        scale = MAX_OCR_IMAGE_PIXELS / longest
        image = image.resize(
            (max(1, int(width * scale)), max(1, int(height * scale))),
            Image.Resampling.LANCZOS,
        )
    elif longest < 900:
        scale = min(3.0, 1800 / max(longest, 1))
        if scale > 1.05:
            image = image.resize(
                (max(1, int(width * scale)), max(1, int(height * scale))),
                Image.Resampling.LANCZOS,
            )
    image = ImageOps.grayscale(image)
    return ImageOps.autocontrast(image)


def _ocr_image(image: Image.Image, lang: str = "eng", psm: int = 6) -> str:
    """Run Tesseract and return cleaned text.  Failures never crash parsing."""

    if _OCR_DISABLED_REASON is not None:
        return ""

    if pytesseract is None:
        _disable_ocr("pytesseract is not installed")
        return ""

    tesseract_cmd = resolve_tesseract_cmd()
    if not tesseract_cmd:
        _disable_ocr(
            "Tesseract executable not found. Install it or set TESSERACT_CMD."
        )
        return ""

    pytesseract.pytesseract.tesseract_cmd = tesseract_cmd
    prepared = _prepare_for_ocr(image)
    digest = hashlib.sha1(prepared.tobytes()).hexdigest()
    cache_key = f"{digest}:{lang}:{psm}"
    if cache_key in _OCR_CACHE:
        return _OCR_CACHE[cache_key]

    try:
        text = pytesseract.image_to_string(
            prepared,
            lang=lang,
            config=f"--oem 3 --psm {psm}",
        )
    except pytesseract.TesseractNotFoundError:
        _disable_ocr("Tesseract executable was not found on this machine")
        return ""
    except pytesseract.TesseractError as exc:
        if lang != "eng":
            LOGGER.warning("OCR failed with lang=%s (%s); retrying with eng", lang, exc)
            return _ocr_image(image, lang="eng", psm=psm)
        LOGGER.warning("Tesseract failed on one image: %s", exc)
        return ""
    except Exception as exc:  # pragma: no cover - defensive
        LOGGER.warning("Unexpected OCR failure: %s", exc)
        return ""

    cleaned = clean_ocr_text(text)
    _OCR_CACHE[cache_key] = cleaned
    return cleaned


def _save_image(image: Image.Image, target: Path) -> bool:
    try:
        target.parent.mkdir(parents=True, exist_ok=True)
        image.convert("RGB").save(target, format="PNG", optimize=True)
        return True
    except Exception as exc:  # pragma: no cover - defensive
        LOGGER.warning("Could not save extracted image %s: %s", target, exc)
        return False


def _ocr_block_text(text: str) -> str:
    text = clean_ocr_text(text)
    if not looks_meaningful(text):
        return ""
    return text


def _guess_caption(texts: list[str], page: int) -> str:
    for text in texts:
        if re.search(r"\b(figure|fig\.|table|chart|diagram|image)\b", text, re.I):
            return truncate_at_word(text.strip(), 160)
    return f"Figure extracted from page {page}"


# ---------------------------------------------------------------------------
# PDF parsing
# ---------------------------------------------------------------------------
def _estimate_body_size(doc: fitz.Document) -> float:
    sizes: list[float] = []
    for page_index in range(min(len(doc), 6)):
        try:
            data = doc.load_page(page_index).get_text("dict")
        except Exception:
            continue
        for block in data.get("blocks", []):
            if block.get("type") != 0:
                continue
            for line in block.get("lines", []):
                for span in line.get("spans", []):
                    if len(span.get("text", "").strip()) > 2:
                        sizes.append(float(span.get("size", 0)))
    sizes = [size for size in sizes if size > 4]
    return median(sizes) if sizes else 11.0


def _page_text_blocks(
    page: fitz.Page,
    page_number: int,
    body_size: float,
    source_name: str,
) -> list[TextBlock]:
    blocks: list[TextBlock] = []
    try:
        data = page.get_text("dict")
    except Exception:
        data = {"blocks": []}

    for raw_block in data.get("blocks", []):
        if raw_block.get("type") != 0:
            continue
        line_texts: list[str] = []
        sizes: list[float] = []
        bold = False
        for line in raw_block.get("lines", []):
            spans = line.get("spans", [])
            text = "".join(span.get("text", "") for span in spans).strip()
            if text:
                line_texts.append(text)
            for span in spans:
                sizes.append(float(span.get("size", 0)))
                bold = bold or bool(span.get("flags", 0) & 16)

        text = normalize_whitespace(" ".join(line_texts))
        text = strip_page_noise(text)
        if not text or len(text) < 2:
            continue

        max_size = max(sizes) if sizes else body_size
        is_first_page = page_number == 1
        kind = "paragraph"
        if is_first_page and max_size >= body_size * 1.55 and len(text) <= 180:
            kind = "title"
        elif max_size >= body_size * 1.18 and len(text) <= 200:
            kind = "heading"
        elif bold and len(text) <= 110 and not re.search(r"[.!?]$", text):
            kind = "heading"

        blocks.append(
            TextBlock(
                text=text,
                page=page_number,
                kind=kind,
                source=source_name,
            )
        )

    if blocks:
        return blocks

    # Fallback for unusual PDFs where get_text("dict") returns nothing useful.
    raw = page.get_text("text") or ""
    for paragraph in re.split(r"\n\s*\n", raw):
        paragraph = strip_page_noise(paragraph)
        if paragraph:
            blocks.append(
                TextBlock(
                    text=normalize_whitespace(paragraph),
                    page=page_number,
                    kind="paragraph",
                    source=source_name,
                )
            )
    return blocks


def _image_rect_area(page: fitz.Page, xref: int) -> float:
    try:
        rects = page.get_image_rects(xref)
    except Exception:
        return 0.0
    page_area = max(page.rect.width * page.rect.height, 1.0)
    return max(
        ((rect.width * rect.height) / page_area for rect in rects),
        default=0.0,
    )


def parse_pdf(
    file_path: str | Path,
    use_ocr: bool = True,
    assets_dir: str | Path | None = None,
    ocr_lang: str = "eng",
) -> ParsedDocument:
    path = Path(file_path)
    LOGGER.info("Parsing PDF: %s (OCR=%s)", path.name, use_ocr)

    doc = fitz.open(path)
    if doc.needs_pass:
        doc.close()
        raise ValueError(f"PDF is password protected: {path.name}")

    page_count = len(doc)
    body_size = _estimate_body_size(doc)
    blocks: list[TextBlock] = []
    images: list[DocumentImage] = []
    seen_xrefs: set[int] = set()
    seen_hashes: set[str] = set()
    asset_root = Path(assets_dir) if assets_dir else None

    try:
        for page_index in range(page_count):
            page = doc.load_page(page_index)
            page_number = page_index + 1
            page_text = page.get_text("text") or ""
            page_blocks = _page_text_blocks(
                page, page_number, body_size, path.name
            )
            blocks.extend(page_blocks)

            page_text_length = len(normalize_whitespace(page_text))
            sparse_page = page_text_length < 80

            if use_ocr and sparse_page:
                pix = page.get_pixmap(dpi=200, alpha=False)
                page_image = Image.open(io.BytesIO(pix.tobytes("png")))
                ocr_text = _ocr_block_text(
                    _ocr_image(page_image, lang=ocr_lang, psm=3)
                )
                if ocr_text:
                    blocks.append(
                        TextBlock(
                            text=ocr_text,
                            page=page_number,
                            kind="ocr",
                            source=path.name,
                        )
                    )

            if not use_ocr:
                continue

            page_caption_texts = [block.text for block in page_blocks]
            for image_index, image_info in enumerate(page.get_images(full=True)):
                xref = int(image_info[0])
                if xref in seen_xrefs:
                    continue
                seen_xrefs.add(xref)

                try:
                    base_image = doc.extract_image(xref)
                    image_bytes = base_image["image"]
                    image = Image.open(io.BytesIO(image_bytes))
                    image.load()
                except Exception as exc:
                    LOGGER.debug("Could not extract image xref=%s: %s", xref, exc)
                    continue

                width, height = image.size
                if width < MIN_OCR_IMAGE_SIDE or height < MIN_OCR_IMAGE_SIDE:
                    continue

                digest = hashlib.md5(image_bytes).hexdigest()
                if digest in seen_hashes:
                    continue
                seen_hashes.add(digest)

                # A full-page scan is already covered by the page-level OCR.
                if sparse_page and _image_rect_area(page, xref) > 0.72:
                    continue

                ocr_text = _ocr_block_text(
                    _ocr_image(image, lang=ocr_lang, psm=6)
                )

                saved_path = ""
                if (
                    asset_root is not None
                    and width >= MIN_FIGURE_SIDE
                    and height >= 200
                ):
                    target = (
                        asset_root
                        / f"{path.stem}_p{page_number:03d}_img{image_index + 1:02d}.png"
                    )
                    if _save_image(image, target):
                        saved_path = str(target)

                if ocr_text:
                    blocks.append(
                        TextBlock(
                            text=ocr_text,
                            page=page_number,
                            kind="ocr",
                            source=path.name,
                            image=saved_path or None,
                        )
                    )

                if saved_path:
                    images.append(
                        DocumentImage(
                            path=saved_path,
                            page=page_number,
                            width=width,
                            height=height,
                            ocr_text=ocr_text,
                            caption=_guess_caption(page_caption_texts, page_number),
                            source=path.name,
                        )
                    )

            LOGGER.debug(
                "Page %s/%s: %s text blocks, %s images so far",
                page_number,
                page_count,
                len(page_blocks),
                len(images),
            )

        metadata_title = normalize_whitespace(doc.metadata.get("title") or "")
        title = ""
        if metadata_title and len(metadata_title) > 3:
            title = metadata_title
        if not title:
            for block in blocks:
                if block.kind in {"title", "heading"} and len(block.text) > 3:
                    title = short_title(block.text, max_words=14, max_chars=110)
                    break
        if not title:
            for block in blocks:
                if len(block.text) > 8:
                    title = short_title(block.text, max_words=14, max_chars=110)
                    break
        title = title or path.stem

        return ParsedDocument(
            path=path,
            title=title,
            blocks=blocks,
            images=images,
            page_count=page_count,
        )
    finally:
        doc.close()


# ---------------------------------------------------------------------------
# DOCX parsing
# ---------------------------------------------------------------------------
def _docx_style_kind(style_name: str) -> tuple[str, int]:
    name = (style_name or "").lower()
    level_match = re.search(r"(\d+)", name)
    level = int(level_match.group(1)) if level_match else 0
    if "title" in name:
        return "title", 0
    if "heading" in name:
        return "heading", level
    if "list" in name or "bullet" in name:
        return "list", level
    if "caption" in name:
        return "caption", 0
    return "paragraph", 0


def parse_docx(
    file_path: str | Path,
    use_ocr: bool = True,
    assets_dir: str | Path | None = None,
    ocr_lang: str = "eng",
) -> ParsedDocument:
    path = Path(file_path)
    LOGGER.info("Parsing DOCX: %s (OCR=%s)", path.name, use_ocr)

    document = docx.Document(path)
    blocks: list[TextBlock] = []
    images: list[DocumentImage] = []
    seen_hashes: set[str] = set()
    asset_root = Path(assets_dir) if assets_dir else None

    for child in document.element.body.iterchildren():
        if child.tag == qn("w:p"):
            paragraph = Paragraph(child, document)
            text = normalize_whitespace(paragraph.text)
            if not text:
                continue
            style_name = paragraph.style.name if paragraph.style is not None else ""
            kind, level = _docx_style_kind(style_name or "")
            blocks.append(
                TextBlock(
                    text=text,
                    page=None,
                    kind=kind,
                    level=level,
                    source=path.name,
                )
            )
        elif child.tag == qn("w:tbl"):
            table = Table(child, document)
            for row_index, row in enumerate(table.rows):
                cells = [normalize_whitespace(cell.text) for cell in row.cells]
                row_text = " | ".join(cell for cell in cells if cell)
                if row_text:
                    blocks.append(
                        TextBlock(
                            text=row_text,
                            page=None,
                            kind="table",
                            level=row_index,
                            source=path.name,
                        )
                    )

    if use_ocr:
        for image_index, shape in enumerate(document.inline_shapes):
            try:
                blip = shape._inline.graphic.graphicData.pic.blipFill.blip
                relationship_id = blip.embed
                image_part = document.part.related_parts[relationship_id]
                image_bytes = image_part.blob
                image = Image.open(io.BytesIO(image_bytes))
                image.load()
            except Exception as exc:
                LOGGER.debug("Could not extract DOCX image %s: %s", image_index, exc)
                continue

            width, height = image.size
            if width < MIN_OCR_IMAGE_SIDE or height < MIN_OCR_IMAGE_SIDE:
                continue

            digest = hashlib.md5(image_bytes).hexdigest()
            if digest in seen_hashes:
                continue
            seen_hashes.add(digest)

            ocr_text = _ocr_block_text(_ocr_image(image, lang=ocr_lang, psm=6))
            saved_path = ""
            if asset_root is not None and width >= MIN_FIGURE_SIDE and height >= 200:
                target = asset_root / f"{path.stem}_img{image_index + 1:02d}.png"
                if _save_image(image, target):
                    saved_path = str(target)

            if ocr_text:
                blocks.append(
                    TextBlock(
                        text=ocr_text,
                        page=None,
                        kind="ocr",
                        source=path.name,
                        image=saved_path or None,
                    )
                )

            if saved_path:
                images.append(
                    DocumentImage(
                        path=saved_path,
                        page=None,
                        width=width,
                        height=height,
                        ocr_text=ocr_text,
                        caption=f"Embedded image {image_index + 1}",
                        source=path.name,
                    )
                )

    title = normalize_whitespace(document.core_properties.title or "")
    if not title:
        for block in blocks:
            if block.kind in {"title", "heading"} and len(block.text) > 3:
                title = short_title(block.text, max_words=14, max_chars=110)
                break
    if not title:
        for block in blocks:
            if len(block.text) > 8:
                title = short_title(block.text, max_words=14, max_chars=110)
                break
    title = title or path.stem

    return ParsedDocument(
        path=path,
        title=title,
        blocks=blocks,
        images=images,
        page_count=0,
    )


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------
def parse_document(
    file_path: str | Path,
    use_ocr: bool = True,
    assets_dir: str | Path | None = None,
    ocr_lang: str = "eng",
) -> ParsedDocument:
    path = Path(file_path)
    suffix = path.suffix.lower()
    if suffix == ".pdf":
        return parse_pdf(path, use_ocr=use_ocr, assets_dir=assets_dir, ocr_lang=ocr_lang)
    if suffix == ".docx":
        return parse_docx(path, use_ocr=use_ocr, assets_dir=assets_dir, ocr_lang=ocr_lang)
    raise ValueError(
        f"Unsupported file format: {path.suffix}. Please use PDF or DOCX."
    )


def best_figures(document: ParsedDocument, limit: int = 2) -> list[DocumentImage]:
    """Return the largest useful figures that contain OCR text."""

    candidates: list[DocumentImage] = []
    for image in document.images:
        if not image.ocr_text or len(image.ocr_text) < 35:
            continue
        if image.width < MIN_FIGURE_SIDE or image.height < 200:
            continue
        aspect = image.width / max(image.height, 1)
        if aspect < 0.45 or aspect > 2.8:
            continue
        candidates.append(image)
    candidates.sort(key=lambda img: img.width * img.height, reverse=True)
    return candidates[:limit]
