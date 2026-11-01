"""The PowerPoint design engine.

Every slide is built from plain shapes and text boxes on a blank 16:9 canvas,
so the result stays fully editable in PowerPoint (no flattened images, no
locked placeholders).  The visual language is deliberately quiet: one accent
colour, generous margins, thin rules and a single "takeaway" panel per slide.

``generate_slide_images`` is the documented hook for turning the saved deck
into PNG files; the real work happens in :mod:`src.render`.
"""

from __future__ import annotations

import logging
import math
from pathlib import Path

from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.oxml.xmlchemy import OxmlElement
from pptx.util import Inches, Pt

from src.config import SLIDE_HEIGHT_IN, SLIDE_WIDTH_IN
from src.models import DeckPlan, DeckResult, ParsedDocument, SlideSpec, TextBlock
from src.slide_planner import plan_deck
from src.textutils import normalize_whitespace, truncate_at_word

LOGGER = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Design tokens
# ---------------------------------------------------------------------------
FONT = "Segoe UI"
FONT_SERIF = "Georgia"

INK = RGBColor(0x14, 0x17, 0x1C)
DARK = RGBColor(0x0F, 0x14, 0x1A)
DARK_SOFT = RGBColor(0x1B, 0x23, 0x2B)
PAPER = RGBColor(0xFF, 0xFF, 0xFF)
CANVAS = RGBColor(0xF8, 0xF9, 0xFB)
MUTED = RGBColor(0x5E, 0x68, 0x74)
MUTED_LIGHT = RGBColor(0xA9, 0xB4, 0xC0)
HAIRLINE = RGBColor(0xDD, 0xE2, 0xE8)
ACCENT = RGBColor(0x0E, 0x7C, 0x86)
ACCENT_DEEP = RGBColor(0x0A, 0x3F, 0x4A)
ACCENT_SOFT = RGBColor(0xE6, 0xF2, 0xF3)
ACCENT_LIGHT = RGBColor(0x7F, 0xD1, 0xD8)

MARGIN_X = 0.85
CONTENT_W = SLIDE_WIDTH_IN - 2 * MARGIN_X


# ---------------------------------------------------------------------------
# Low-level shape helpers
# ---------------------------------------------------------------------------
def _textbox(
    slide,
    left: float,
    top: float,
    width: float,
    height: float,
    *,
    anchor=MSO_ANCHOR.TOP,
    margins=(0.0, 0.0, 0.0, 0.0),
):
    box = slide.shapes.add_textbox(
        Inches(left), Inches(top), Inches(width), Inches(height)
    )
    frame = box.text_frame
    frame.clear()
    frame.word_wrap = True
    frame.auto_size = MSO_AUTO_SIZE.NONE
    frame.vertical_anchor = anchor
    left_m, top_m, right_m, bottom_m = margins
    frame.margin_left = Inches(left_m)
    frame.margin_top = Inches(top_m)
    frame.margin_right = Inches(right_m)
    frame.margin_bottom = Inches(bottom_m)
    return box


def _write(
    frame,
    text: str,
    *,
    size: float,
    color: RGBColor,
    bold: bool = False,
    italic: bool = False,
    font: str = FONT,
    align=PP_ALIGN.LEFT,
    line_spacing: float = 1.0,
    space_before: float = 0.0,
    space_after: float = 0.0,
    tracking: int = 0,
    first: bool = True,
    paragraph=None,
):
    """Write one paragraph into ``frame`` and style it."""

    if first:
        para = frame.paragraphs[0]
        run = para.add_run() if para.runs else para.add_run()
    else:
        para = frame.add_paragraph()
        run = para.add_run()

    run.text = text
    font_obj = run.font
    font_obj.name = font
    font_obj.size = Pt(size)
    font_obj.bold = bold
    font_obj.italic = italic
    font_obj.color.rgb = color
    if tracking:
        run.font._rPr.set("spc", str(int(tracking)))

    para.alignment = align
    para.line_spacing = line_spacing
    para.space_before = Pt(space_before)
    para.space_after = Pt(space_after)
    return run


def _rect(slide, left, top, width, height, *, fill=None, line=None, shape=MSO_SHAPE.RECTANGLE, radius=None):
    box = slide.shapes.add_shape(
        shape, Inches(left), Inches(top), Inches(width), Inches(height)
    )
    if radius is not None:
        try:
            box.adjustments[0] = radius
        except Exception:
            pass
    if fill is None:
        box.fill.background()
    else:
        box.fill.solid()
        box.fill.fore_color.rgb = fill
    if line is None:
        box.line.fill.background()
    else:
        box.line.color.rgb = line
        box.line.width = Pt(1)
    box.shadow.inherit = False
    if box.has_text_frame:
        box.text_frame.word_wrap = True
    return box


def _set_alpha(shape, alpha: float) -> None:
    """Apply fill transparency (0..1) to a shape that already has a solid fill."""

    try:
        solid_fill = shape.fill._xPr.find(qn("a:solidFill"))
        srgb = solid_fill.find(qn("a:srgbClr"))
        if srgb is None:
            return
        for existing in srgb.findall(qn("a:alpha")):
            srgb.remove(existing)
        alpha_el = OxmlElement("a:alpha")
        alpha_el.set("val", str(int(max(0.0, min(1.0, alpha)) * 100_000)))
        srgb.append(alpha_el)
    except Exception:  # pragma: no cover - cosmetic only
        pass


def _set_line_alpha(shape, alpha: float) -> None:
    try:
        line = shape.line._get_or_add_ln()
        solid = line.find(qn("a:solidFill"))
        if solid is None:
            return
        srgb = solid.find(qn("a:srgbClr"))
        if srgb is None:
            return
        alpha_el = OxmlElement("a:alpha")
        alpha_el.set("val", str(int(max(0.0, min(1.0, alpha)) * 100_000)))
        srgb.append(alpha_el)
    except Exception:  # pragma: no cover
        pass


def _bullet_paragraph(paragraph, color: RGBColor, char: str = "▪", indent: float = 0.30):
    """Add a coloured square bullet with a hanging indent via raw XML."""

    p_pr = paragraph._p.get_or_add_pPr()
    p_pr.set("marL", str(int(Inches(indent))))
    p_pr.set("indent", str(-int(Inches(indent))))

    for tag in ("a:buNone", "a:buChar", "a:buAutoNum", "a:buClr", "a:buSzPct", "a:buFont"):
        for element in p_pr.findall(qn(tag)):
            p_pr.remove(element)

    bu_clr = OxmlElement("a:buClr")
    srgb = OxmlElement("a:srgbClr")
    srgb.set("val", "%02X%02X%02X" % (color[0], color[1], color[2]))
    bu_clr.append(srgb)

    bu_sz = OxmlElement("a:buSzPct")
    bu_sz.set("val", "85000")

    bu_font = OxmlElement("a:buFont")
    bu_font.set("typeface", "Arial")

    bu_char = OxmlElement("a:buChar")
    bu_char.set("char", char)

    for element in (bu_clr, bu_sz, bu_font, bu_char):
        p_pr.append(element)


def _footer(slide, source: str, page: int | None, *, dark: bool = False) -> None:
    colour = MUTED_LIGHT if dark else MUTED
    if source:
        box = _textbox(slide, MARGIN_X, SLIDE_HEIGHT_IN - 0.55, CONTENT_W - 1.0, 0.3)
        _write(
            box.text_frame,
            truncate_at_word(source, 90),
            size=9,
            color=colour,
            tracking=40,
        )
    if page is not None:
        box = _textbox(
            slide,
            SLIDE_WIDTH_IN - MARGIN_X - 0.9,
            SLIDE_HEIGHT_IN - 0.55,
            0.9,
            0.3,
        )
        _write(
            box.text_frame,
            f"{page:02d}",
            size=9,
            color=colour,
            align=PP_ALIGN.RIGHT,
            tracking=40,
        )


def _notes(slide, text: str) -> None:
    if not text:
        return
    try:
        frame = slide.notes_slide.notes_text_frame
        frame.text = text[:4000]
    except Exception:  # pragma: no cover - notes are a bonus
        LOGGER.debug("Could not write speaker notes")


def _blank_slide(prs: Presentation, background: RGBColor):
    layout = None
    for candidate in prs.slide_layouts:
        if candidate.name and candidate.name.lower() == "blank":
            layout = candidate
            break
    if layout is None:
        layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(layout)
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = background
    return slide


# ---------------------------------------------------------------------------
# Text fitting
# ---------------------------------------------------------------------------
def _estimate_lines(text: str, width_in: float, size_pt: float) -> int:
    usable = max(width_in * 72.0 - 16.0, 40.0)
    per_line = max(1, int(usable / (size_pt * 0.50)))
    words = max(1, math.ceil(len(text) / per_line))
    # Long words / wide glyphs occasionally need one extra line.
    return max(1, words + (1 if len(text) > per_line * 0.92 * words else 0))


def _fit_size(
    paragraphs: list[str],
    width_in: float,
    height_in: float,
    *,
    base: float,
    minimum: float,
    line_factor: float = 1.32,
    gap_factor: float = 0.55,
) -> float:
    size = base
    while size > minimum:
        total = 0.0
        for text in paragraphs:
            lines = _estimate_lines(text, width_in, size)
            total += lines * size * line_factor / 72.0 + size * gap_factor / 72.0
        if total <= height_in:
            return size
        size -= 1.0
    return minimum


# ---------------------------------------------------------------------------
# Slide builders
# ---------------------------------------------------------------------------
def _cover(prs, spec: SlideSpec, page: int) -> None:
    slide = _blank_slide(prs, DARK)

    # Oversized translucent accent circle bleeding off the right edge.
    circle = _rect(slide, 9.4, -1.9, 7.2, 7.2, fill=ACCENT, shape=MSO_SHAPE.OVAL)
    _set_alpha(circle, 0.16)
    ring = _rect(slide, 10.9, 3.6, 3.6, 3.6, fill=None, line=ACCENT_LIGHT, shape=MSO_SHAPE.OVAL)
    ring.line.width = Pt(1.2)
    _set_line_alpha(ring, 0.28)

    # Brand line
    brand = _textbox(slide, MARGIN_X, 0.62, 6.0, 0.3)
    _write(
        brand.text_frame,
        (spec.kicker or "GENERATED PRESENTATION").upper(),
        size=10.5,
        color=ACCENT_LIGHT,
        bold=True,
        tracking=220,
    )

    # Accent rule above the title
    _rect(slide, MARGIN_X, 1.72, 1.15, 0.055, fill=ACCENT_LIGHT)

    title = truncate_at_word(spec.title or "Untitled Deck", 130)
    title_size = _fit_size(
        [title], 8.6, 2.35, base=50, minimum=30, line_factor=1.18
    )
    title_box = _textbox(slide, MARGIN_X, 1.98, 8.6, 2.45)
    _write(
        title_box.text_frame,
        title,
        size=title_size,
        color=PAPER,
        bold=True,
        line_spacing=1.06,
    )

    if spec.subtitle:
        sub_box = _textbox(slide, MARGIN_X, 4.55, 7.7, 0.95)
        _write(
            sub_box.text_frame,
            spec.subtitle,
            size=15.5,
            color=MUTED_LIGHT,
            line_spacing=1.25,
        )

    if spec.bullets:
        source_text = "  ·  ".join(
            truncate_at_word(item, 78) for item in spec.bullets[:3]
        )
        if len(spec.bullets) > 3:
            source_text += f"  ·  +{len(spec.bullets) - 3} more"
        source_box = _textbox(slide, MARGIN_X, 6.28, 8.4, 0.55)
        _write(
            source_box.text_frame,
            f"SOURCES   {source_text}",
            size=9,
            color=MUTED_LIGHT,
            tracking=60,
        )

    _notes(slide, spec.notes)


def _section(prs, spec: SlideSpec, page: int) -> None:
    slide = _blank_slide(prs, ACCENT_DEEP)

    giant = _textbox(slide, 8.9, 1.15, 4.0, 3.4)
    _write(
        giant.text_frame,
        spec.kicker.replace("SECTION", "").strip() or f"{page:02d}",
        size=190,
        color=ACCENT,
        bold=True,
        align=PP_ALIGN.RIGHT,
        line_spacing=0.9,
    )
    for run in giant.text_frame.paragraphs[0].runs:
        run.font._rPr.set("spc", "-200")

    label = _textbox(slide, MARGIN_X, 1.55, 6.5, 0.3)
    _write(
        label.text_frame,
        (spec.kicker or "SECTION").upper(),
        size=10.5,
        color=ACCENT_LIGHT,
        bold=True,
        tracking=240,
    )

    title = truncate_at_word(spec.title or "Section", 90)
    size = _fit_size([title], 7.4, 2.0, base=42, minimum=26, line_factor=1.15)
    title_box = _textbox(slide, MARGIN_X, 2.05, 7.6, 2.1)
    _write(
        title_box.text_frame,
        title,
        size=size,
        color=PAPER,
        bold=True,
        line_spacing=1.08,
    )

    _rect(slide, MARGIN_X, 4.35, 1.0, 0.05, fill=ACCENT_LIGHT)

    if spec.subtitle:
        sub_box = _textbox(slide, MARGIN_X, 4.65, 6.6, 0.9)
        _write(
            sub_box.text_frame,
            spec.subtitle,
            size=14,
            color=MUTED_LIGHT,
            line_spacing=1.3,
        )

    _footer(slide, spec.source, page, dark=True)
    _notes(slide, spec.notes)


def _agenda(prs, spec: SlideSpec, page: int) -> None:
    slide = _blank_slide(prs, CANVAS)

    label = _textbox(slide, MARGIN_X, 0.62, 6.0, 0.3)
    _write(
        label.text_frame,
        (spec.kicker or "AGENDA").upper(),
        size=10.5,
        color=ACCENT,
        bold=True,
        tracking=240,
    )

    title_box = _textbox(slide, MARGIN_X, 0.95, CONTENT_W, 0.8)
    _write(
        title_box.text_frame,
        spec.title or "Agenda",
        size=36,
        color=INK,
        bold=True,
    )

    _rect(slide, MARGIN_X, 1.86, 1.15, 0.055, fill=ACCENT)

    items = spec.bullets[:6]
    if not items:
        items = ["Overview"]

    top = 2.28
    max_row = 1.25 if len(items) <= 3 else 0.98
    row_h = min(max_row, (6.55 - top) / max(len(items), 1))
    if len(items) <= 3:
        top += max(0.0, (6.4 - top - row_h * len(items)) / 2)
    for index, item in enumerate(items):
        y = top + index * row_h
        number = _textbox(slide, MARGIN_X, y, 0.85, row_h - 0.1)
        _write(
            number.text_frame,
            f"{index + 1:02d}",
            size=22,
            color=ACCENT,
            bold=True,
        )
        text_box = _textbox(slide, MARGIN_X + 0.95, y + 0.05, CONTENT_W - 1.0, row_h - 0.16)
        _write(
            text_box.text_frame,
            truncate_at_word(item, 105),
            size=16,
            color=INK,
            line_spacing=1.25,
        )
        if index < len(items) - 1:
            _rect(
                slide,
                MARGIN_X + 0.95,
                y + row_h - 0.14,
                CONTENT_W - 0.95,
                0.008,
                fill=HAIRLINE,
            )

    _footer(slide, spec.source, page)
    _notes(slide, spec.notes)


def _content(prs, spec: SlideSpec, page: int) -> None:
    slide = _blank_slide(prs, CANVAS)

    label = _textbox(slide, MARGIN_X, 0.55, CONTENT_W, 0.28)
    _write(
        label.text_frame,
        (spec.kicker or "KEY CONTENT").upper(),
        size=9.5,
        color=ACCENT,
        bold=True,
        tracking=200,
    )

    title = truncate_at_word(spec.title or "Key Content", 95)
    title_size = _fit_size([title], CONTENT_W, 1.0, base=31, minimum=22, line_factor=1.1)
    title_box = _textbox(slide, MARGIN_X, 0.85, CONTENT_W, 1.02)
    _write(
        title_box.text_frame,
        title,
        size=title_size,
        color=INK,
        bold=True,
        line_spacing=1.08,
    )

    _rect(slide, MARGIN_X, 1.99, CONTENT_W, 0.008, fill=HAIRLINE)
    _rect(slide, MARGIN_X, 1.965, 1.15, 0.05, fill=ACCENT)

    bullets = [truncate_at_word(b, 260) for b in spec.bullets if b.strip()]
    takeaway = spec.takeaway.strip()

    body_top = 2.26
    body_bottom = 5.92 if takeaway else 6.75
    body_h = body_bottom - body_top

    if bullets:
        base = 27 if len(bullets) <= 2 else 24 if len(bullets) == 3 else 21
        size = _fit_size(bullets, CONTENT_W, body_h, base=base, minimum=13)
        box = _textbox(slide, MARGIN_X, body_top, CONTENT_W, body_h)
        frame = box.text_frame
        for index, bullet in enumerate(bullets):
            if index == 0:
                run = frame.paragraphs[0].add_run()
                para = frame.paragraphs[0]
            else:
                para = frame.add_paragraph()
                run = para.add_run()
            run.text = bullet
            run.font.name = FONT
            run.font.size = Pt(size)
            run.font.color.rgb = INK
            para.line_spacing = 1.28
            para.space_after = Pt(size * 0.55)
            _bullet_paragraph(para, ACCENT, indent=0.30)

    if takeaway:
        _rect(
            slide,
            MARGIN_X,
            6.05,
            CONTENT_W,
            0.86,
            fill=PAPER,
            line=HAIRLINE,
            shape=MSO_SHAPE.ROUNDED_RECTANGLE,
            radius=0.07,
        )
        _rect(slide, MARGIN_X, 6.05, 0.06, 0.86, fill=ACCENT)
        panel_text = _textbox(
            slide,
            MARGIN_X + 0.26,
            6.13,
            CONTENT_W - 0.5,
            0.7,
            anchor=MSO_ANCHOR.MIDDLE,
        )
        _write(
            panel_text.text_frame,
            takeaway,
            size=12.5,
            color=MUTED,
            italic=True,
            line_spacing=1.22,
        )

    _footer(slide, spec.source, page)
    _notes(slide, spec.notes)


def _figure(prs, spec: SlideSpec, page: int) -> None:
    slide = _blank_slide(prs, CANVAS)

    label = _textbox(slide, MARGIN_X, 0.55, 5.8, 0.28)
    _write(
        label.text_frame,
        (spec.kicker or "FIGURE").upper(),
        size=9.5,
        color=ACCENT,
        bold=True,
        tracking=200,
    )

    title = truncate_at_word(spec.title or "Visual Evidence", 70)
    title_size = _fit_size([title], 5.6, 1.0, base=29, minimum=20, line_factor=1.12)
    title_box = _textbox(slide, MARGIN_X, 0.86, 5.6, 1.05)
    _write(
        title_box.text_frame,
        title,
        size=title_size,
        color=INK,
        bold=True,
        line_spacing=1.08,
    )
    _rect(slide, MARGIN_X, 1.99, 5.6, 0.008, fill=HAIRLINE)
    _rect(slide, MARGIN_X, 1.965, 1.0, 0.05, fill=ACCENT)

    # Image panel on the right
    panel_left, panel_top = 6.95, 1.35
    panel_w, panel_h = 5.55, 4.75
    _rect(slide, panel_left, panel_top, panel_w, panel_h, fill=PAPER, line=HAIRLINE)
    _rect(slide, panel_left, panel_top, panel_w, 0.05, fill=ACCENT)

    image_path = Path(spec.image_path) if spec.image_path else None
    rendered = False
    if image_path and image_path.is_file():
        try:
            with Image.open(image_path) as img:
                iw, ih = img.size
            inner_w = panel_w - 0.36
            inner_h = panel_h - 0.62
            scale = min(inner_w / max(iw, 1), inner_h / max(ih, 1))
            draw_w = max(0.4, iw * scale)
            draw_h = max(0.4, ih * scale)
            left = panel_left + (panel_w - draw_w) / 2
            top = panel_top + 0.16 + (inner_h - draw_h) / 2
            slide.shapes.add_picture(
                str(image_path), Inches(left), Inches(top), Inches(draw_w), Inches(draw_h)
            )
            rendered = True
        except Exception as exc:
            LOGGER.warning("Could not place figure %s: %s", image_path, exc)

    if not rendered:
        fallback = _textbox(
            slide,
            panel_left + 0.3,
            panel_top + 1.9,
            panel_w - 0.6,
            0.9,
        )
        _write(
            fallback.text_frame,
            "Figure text extracted with OCR",
            size=14,
            color=MUTED,
            italic=True,
            align=PP_ALIGN.CENTER,
        )

    if spec.caption:
        caption = _textbox(slide, panel_left, panel_top + panel_h - 0.44, panel_w, 0.32)
        _write(
            caption.text_frame,
            truncate_at_word(spec.caption, 90),
            size=8.5,
            color=MUTED,
            align=PP_ALIGN.CENTER,
        )

    notes = [truncate_at_word(b, 210) for b in spec.bullets if b.strip()]
    if notes:
        size = _fit_size(notes, 5.6, 3.0, base=15, minimum=11)
        box = _textbox(slide, MARGIN_X, 2.3, 5.6, 3.0)
        frame = box.text_frame
        for index, note in enumerate(notes[:3]):
            if index == 0:
                para = frame.paragraphs[0]
                run = para.add_run()
            else:
                para = frame.add_paragraph()
                run = para.add_run()
            run.text = note
            run.font.name = FONT
            run.font.size = Pt(size)
            run.font.color.rgb = MUTED
            para.line_spacing = 1.32
            para.space_after = Pt(10)
            _bullet_paragraph(para, ACCENT_LIGHT, indent=0.26)

    _footer(slide, spec.source, page)
    _notes(slide, spec.notes)


def _summary(prs, spec: SlideSpec, page: int) -> None:
    slide = _blank_slide(prs, DARK)
    glow = _rect(slide, -1.6, 4.3, 5.6, 5.6, fill=ACCENT, shape=MSO_SHAPE.OVAL)
    _set_alpha(glow, 0.14)

    label = _textbox(slide, MARGIN_X, 0.62, 6.0, 0.3)
    _write(
        label.text_frame,
        (spec.kicker or "SUMMARY").upper(),
        size=10.5,
        color=ACCENT_LIGHT,
        bold=True,
        tracking=240,
    )

    title_box = _textbox(slide, MARGIN_X, 0.95, 10.5, 0.9)
    _write(
        title_box.text_frame,
        spec.title or "Key Takeaways",
        size=38,
        color=PAPER,
        bold=True,
    )
    if spec.subtitle:
        sub_box = _textbox(slide, MARGIN_X, 1.86, 9.4, 0.5)
        _write(
            sub_box.text_frame,
            spec.subtitle,
            size=14,
            color=MUTED_LIGHT,
        )

    items = spec.bullets[:4] or ["Review the deck for the full narrative."]
    top = 2.62
    row_h = min(1.18, (6.35 - top) / max(len(items), 1))
    for index, item in enumerate(items):
        y = top + index * row_h
        num_box = _textbox(slide, MARGIN_X, y + 0.04, 0.9, 0.5)
        _write(
            num_box.text_frame,
            f"{index + 1:02d}",
            size=17,
            color=ACCENT_LIGHT,
            bold=True,
        )
        text_box = _textbox(slide, MARGIN_X + 0.85, y, CONTENT_W - 1.2, row_h - 0.1)
        size = _fit_size([item], CONTENT_W - 1.2, row_h - 0.12, base=17, minimum=12)
        _write(
            text_box.text_frame,
            truncate_at_word(item, 170),
            size=size,
            color=PAPER,
            line_spacing=1.28,
        )

    _footer(slide, spec.source, page, dark=True)
    _notes(slide, spec.notes)


_BUILDERS = {
    "cover": _cover,
    "agenda": _agenda,
    "section": _section,
    "content": _content,
    "figure": _figure,
    "summary": _summary,
}


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------
def build_presentation(plan: DeckPlan, output_path: str | Path) -> DeckResult:
    """Render a :class:`DeckPlan` to a ``.pptx`` file."""

    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    prs = Presentation()
    prs.slide_width = Inches(SLIDE_WIDTH_IN)
    prs.slide_height = Inches(SLIDE_HEIGHT_IN)

    props = prs.core_properties
    props.title = plan.title
    props.subject = plan.subtitle
    props.author = "Noora PPT Builder"
    props.comments = "Generated from: " + "; ".join(plan.sources)

    slides = plan.slides or [
        SlideSpec(kind="cover", title=plan.title, subtitle=plan.subtitle)
    ]
    for index, spec in enumerate(slides, start=1):
        builder = _BUILDERS.get(spec.kind, _content)
        try:
            builder(prs, spec, index)
        except Exception as exc:  # a broken slide must not kill the deck
            LOGGER.error("Slide %s failed (%s): %s", index, spec.kind, exc)
            slide = _blank_slide(prs, CANVAS)
            box = _textbox(slide, MARGIN_X, 2.5, CONTENT_W, 1.5)
            _write(
                box.text_frame,
                spec.title or "Slide",
                size=30,
                color=INK,
                bold=True,
            )

    prs.save(str(output_path))
    LOGGER.info("Saved deck: %s (%s slides)", output_path, len(prs.slides))
    return DeckResult(path=output_path, slide_count=len(prs.slides), plan=plan)


def create_deck(
    documents: list[ParsedDocument],
    output_path: str | Path,
    *,
    title: str | None = None,
    max_slides: int = 24,
    include_images: bool = True,
    per_slide: int = 3,
) -> DeckResult:
    plan = plan_deck(
        documents,
        title=title,
        max_slides=max_slides,
        include_images=include_images,
        per_slide=per_slide,
    )
    return build_presentation(plan, output_path)


def create_slides(text_content: str, output_path: str | Path, title: str | None = None) -> DeckResult:
    """Backwards-compatible helper: build a deck from a single text blob."""

    paragraphs = [
        line.strip()
        for line in normalize_whitespace(text_content).splitlines()
        if len(line.strip()) > 3
    ]
    blocks = [
        TextBlock(text=paragraph, kind="paragraph", source="input")
        for paragraph in paragraphs
    ]
    document = ParsedDocument(
        path=Path(str(output_path)).with_suffix(".txt"),
        title=title or "Generated Presentation",
        blocks=blocks,
        images=[],
        page_count=0,
    )
    return create_deck([document], output_path, title=title)


def generate_slide_images(
    pptx_path: str | Path,
    output_dir: str | Path,
    *,
    width: int = 1920,
    backend: str = "auto",
) -> list[Path]:
    """Export every slide to PNG.

    True PPTX-to-PNG rendering needs a real layout engine: Microsoft
    PowerPoint via COM automation on Windows, or LibreOffice in headless mode
    on any platform.  This hook delegates to :mod:`src.render`, which tries
    PowerPoint first, then LibreOffice, and finally falls back to an
    approximate Pillow preview.
    """

    from src.render import export_slide_images

    return export_slide_images(
        pptx_path, output_dir, width=width, backend=backend
    )
