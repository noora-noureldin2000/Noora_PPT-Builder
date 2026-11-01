"""Turn parsed documents into a paced, presentation-ready slide plan."""

from __future__ import annotations

import re
from dataclasses import dataclass, field

from src.models import DocumentImage, ParsedDocument, SlideSpec, TextBlock
from src.parser import best_figures
from src.textutils import (
    clean_ocr_text,
    dedupe_keep_order,
    key_sentence,
    normalize_whitespace,
    sentence_chunks,
    short_title,
    strip_page_noise,
    truncate_at_word,
)


@dataclass
class Section:
    title: str
    blocks: list[TextBlock] = field(default_factory=list)
    page: int | None = None


_GENERIC_TITLES = {
    "overview",
    "key concepts",
    "key content",
    "content",
    "introduction",
    "background",
    "general",
    "document",
}


def _page_label(page: int | None) -> str:
    return f"Page {page}" if page else "Document"


def _doc_meta(doc: ParsedDocument) -> str:
    if doc.page_count:
        return f"{doc.page_count} pages"
    return "Word document"


def _plural(count: int, singular: str) -> str:
    return f"{count} {singular}" if count == 1 else f"{count} {singular}s"


def _clean_block_text(block: TextBlock) -> str:
    text = clean_ocr_text(block.text) if block.is_ocr else normalize_whitespace(block.text)
    text = strip_page_noise(text)
    return text.strip()


def _sections_for_document(doc: ParsedDocument) -> list[Section]:
    sections: list[Section] = []
    current_title = "Overview"
    current_blocks: list[TextBlock] = []
    current_page: int | None = None
    doc_title_key = normalize_whitespace(doc.title).lower()

    for block in doc.blocks:
        text = _clean_block_text(block)
        if not text:
            continue

        is_heading = block.kind in {"title", "heading"}
        if is_heading:
            # The document title is already on the cover/section slide.
            if not current_blocks and text.lower().strip() == doc_title_key:
                continue
            if current_blocks:
                sections.append(
                    Section(
                        title=short_title(current_title, max_words=10, max_chars=85),
                        blocks=current_blocks,
                        page=current_page,
                    )
                )
            current_title = text
            current_blocks = []
            current_page = block.page
            continue

        if len(text) < 12 and block.kind not in {"table", "list"}:
            continue
        if current_page is None:
            current_page = block.page
        current_blocks.append(block)

    if current_blocks:
        sections.append(
            Section(
                title=short_title(current_title, max_words=10, max_chars=85),
                blocks=current_blocks,
                page=current_page,
            )
        )

    if not sections and doc.blocks:
        sections.append(
            Section(
                title="Key Concepts",
                blocks=[block for block in doc.blocks if _clean_block_text(block)],
                page=doc.blocks[0].page,
            )
        )

    # Pacing happens per section in ``_content_slides_for_section``; keep the
    # sections intact so their headings stay meaningful.
    return sections


def _split_ocr_lines(text: str) -> list[str]:
    """Split OCR text from a figure into readable, step-sized segments."""

    text = re.sub(
        r"\b(Figure|Table|Chart|Diagram|Fig\.)\s*(\d+)\s*[.)]\s+",
        r"\1 \2: ",
        text,
        flags=re.I,
    )
    text = re.sub(r"(?=\b(?:Step|Phase|Stage)\s+\d+\s*[:.\-])", "\n", text)
    text = re.sub(r"(?=\b(?:Figure|Table|Chart|Diagram)\s+\d+\s*[:.]\s)", "\n", text)
    return [line.strip() for line in text.splitlines() if line.strip()]


def _units_from_blocks(blocks: list[TextBlock], max_chars: int = 190) -> list[str]:
    units: list[str] = []
    for block in blocks:
        text = _clean_block_text(block)
        if not text:
            continue

        if block.kind == "table":
            row_parts = [
                part.strip()
                for part in text.split("|")
                if part.strip() and len(part.strip()) > 1
            ]
            if row_parts:
                units.append(truncate_at_word(" · ".join(row_parts), max_chars))
            continue

        if block.is_ocr:
            for segment in _split_ocr_lines(text):
                units.extend(sentence_chunks(segment, max_chars=max_chars))
            continue

        units.extend(sentence_chunks(text, max_chars=max_chars))

    return dedupe_keep_order(units)


def _pack_chunks(units: list[str], per_slide: int) -> list[list[str]]:
    """Split units into slides, folding a lonely remainder into the last chunk."""

    per_slide = max(2, per_slide)
    chunks = [units[i : i + per_slide] for i in range(0, len(units), per_slide)]
    if len(chunks) >= 2 and len(chunks[-1]) == 1:
        chunks[-2].extend(chunks.pop())
    return chunks


def _content_slides_for_section(
    section: Section,
    doc: ParsedDocument,
    per_slide: int = 3,
) -> list[SlideSpec]:
    units = _units_from_blocks(section.blocks)
    if not units:
        return []

    chunks = _pack_chunks(units, per_slide)
    slides: list[SlideSpec] = []
    section_title = section.title or "Key Concepts"
    generic = section_title.strip().lower() in _GENERIC_TITLES
    page_label = _page_label(section.page)
    has_ocr = any(block.is_ocr for block in section.blocks)

    for index, chunk in enumerate(chunks):
        bullets = list(chunk)
        title = section_title

        if generic and len(bullets) >= 3:
            title = truncate_at_word(bullets[0], 62)
            bullets = bullets[1:]
        elif index > 0:
            title = f"{section_title} — Part {index + 1}"

        # Pull one key sentence out of the body and into the highlight panel.
        # This keeps the panel relevant without repeating a visible bullet.
        takeaway = ""
        if len(bullets) >= 3:
            ordered = bool(
                re.search(
                    r"\b(?:step|phase|stage)\s+\d+\b", " ".join(bullets[:3]), re.I
                )
            )
            if ordered:
                # Preserve the sequence: the last step becomes the highlight.
                takeaway = bullets[-1]
                bullets = bullets[:-1]
            else:
                candidate = key_sentence(" ".join(bullets), max_chars=145)
                if candidate:
                    for position, bullet in enumerate(bullets):
                        if re.sub(r"\W+", " ", candidate.lower()).strip() in re.sub(
                            r"\W+", " ", bullet.lower()
                        ):
                            takeaway = candidate
                            bullets = bullets[:position] + bullets[position + 1 :]
                            break

        slides.append(
            SlideSpec(
                kind="content",
                title=title,
                kicker=(
                    f"{doc.title[:48]} · OCR-assisted"
                    if has_ocr
                    else f"{doc.title[:48]} · {page_label}"
                ),
                bullets=bullets,
                takeaway=takeaway,
                source=page_label,
                notes="\n".join(bullets),
            )
        )
    return slides


def _figure_slide(image: DocumentImage, doc: ParsedDocument) -> SlideSpec | None:
    segments: list[str] = []
    for line in _split_ocr_lines(clean_ocr_text(image.ocr_text)):
        segments.extend(sentence_chunks(line, max_chars=150))
    notes = segments[:3]
    if not notes:
        return None

    caption = image.caption.strip()
    generic_caption = caption.lower().startswith(
        ("figure extracted", "embedded image")
    )
    if generic_caption and re.match(r"^(?:figure|table|chart|diagram)\b", notes[0], re.I):
        title = short_title(notes[0], max_words=9, max_chars=70)
    elif generic_caption:
        title = "Visual Evidence"
    else:
        title = short_title(caption, max_words=9, max_chars=70)
    takeaway = key_sentence(image.ocr_text, max_chars=135)

    return SlideSpec(
        kind="figure",
        title=title,
        kicker=f"{doc.title[:48]} · Figure OCR",
        bullets=notes,
        takeaway=takeaway,
        source=_page_label(image.page),
        image_path=image.path,
        caption=caption,
        notes=image.ocr_text,
    )


def _agenda_slide(documents: list[ParsedDocument], sections: dict[int, list[Section]]) -> SlideSpec:
    bullets = []
    for index, doc in enumerate(documents, start=1):
        section_count = len(sections.get(index - 1, []))
        bullets.append(
            f"{doc.title} — {_doc_meta(doc)} · "
            f"{_plural(section_count, 'section')}"
        )
    return SlideSpec(
        kind="agenda",
        title="What This Deck Covers",
        subtitle="A guided review of the source material",
        kicker="AGENDA",
        bullets=bullets,
        takeaway="Every slide is editable in PowerPoint.",
    )


def _summary_slide(documents: list[ParsedDocument], sections: dict[int, list[Section]]) -> SlideSpec:
    takeaways: list[str] = []
    seen: set[str] = set()

    def add(text: str) -> None:
        text = truncate_at_word(text.strip(), 135)
        key = re.sub(r"\W+", " ", text.lower()).strip()
        if len(key) < 20 or key in seen:
            return
        seen.add(key)
        takeaways.append(text)

    for index in range(len(documents)):
        for section in sections.get(index, []):
            units = _units_from_blocks(section.blocks)
            if not units:
                continue
            # Prefer a substantive but compact sentence: tiny fragments read
            # as labels and very long ones overflow the summary slide.
            sized = [unit for unit in units if 40 <= len(unit) <= 200] or units
            best = min(sized, key=lambda unit: abs(len(unit) - 105))
            add(best)
            if len(takeaways) >= 3:
                break
        if len(takeaways) >= 3:
            break

    while len(takeaways) < 3 and len(takeaways) < len(documents):
        doc = documents[len(takeaways)]
        add(f"{doc.title} — {_doc_meta(doc)}")

    return SlideSpec(
        kind="summary",
        title="Key Takeaways",
        subtitle="A compact review of the material in this deck",
        kicker="SUMMARY",
        bullets=takeaways[:3],
        takeaway="Use the editable slides and speaker notes to adapt this deck.",
    )


def _trim_to_budget(slides: list[SlideSpec], max_slides: int) -> list[SlideSpec]:
    if max_slides <= 0:
        max_slides = 1
    if len(slides) <= max_slides:
        return slides

    indexed = list(enumerate(slides))
    protected_kinds = {"cover", "agenda", "section", "summary"}
    protected = [(i, slide) for i, slide in indexed if slide.kind in protected_kinds]
    optional = [(i, slide) for i, slide in indexed if slide.kind not in protected_kinds]

    if max_slides == 1:
        return [slides[0]]

    if len(protected) >= max_slides:
        # Keep the cover, the first structural slides and the summary.
        keep = protected[: max_slides - 1]
        if protected[-1] not in keep:
            keep.append(protected[-1])
        keep = keep[:max_slides]
        return [slide for _, slide in sorted(keep, key=lambda item: item[0])]

    budget = max_slides - len(protected)
    if budget <= 0:
        selected_optional: list[tuple[int, SlideSpec]] = []
    elif len(optional) <= budget:
        selected_optional = optional
    else:
        step = len(optional) / budget
        picked: dict[int, tuple[int, SlideSpec]] = {}
        for pick_index in range(budget):
            original_index, slide = optional[min(len(optional) - 1, int(pick_index * step))]
            picked[original_index] = (original_index, slide)
        selected_optional = list(picked.values())

    merged = sorted(protected + selected_optional, key=lambda item: item[0])
    return [slide for _, slide in merged][:max_slides]


def plan_deck(
    documents: list[ParsedDocument],
    title: str | None = None,
    max_slides: int = 24,
    include_images: bool = True,
    per_slide: int = 3,
) -> tuple:
    """Build a :class:`DeckPlan` from one or more parsed documents."""

    from src.models import DeckPlan  # local import keeps the module cycle-free

    if not documents:
        plan = DeckPlan(
            title=title or "No Content Found",
            subtitle="No readable source material was supplied.",
            sources=[],
            slides=[
                SlideSpec(
                    kind="cover",
                    title=title or "No Content Found",
                    subtitle="Check the input files and try again.",
                    kicker="NOORA PPT BUILDER",
                )
            ],
        )
        return plan

    total_pages = sum(doc.page_count for doc in documents)
    deck_title = title or (
        documents[0].title if len(documents) == 1 else "Multi-Source Review"
    )
    subtitle = (
        f"A visual, editable synthesis of {len(documents)} source document"
        f"{'s' if len(documents) != 1 else ''}"
    )
    if total_pages:
        subtitle += f" · {total_pages} pages"

    sources = [
        f"{doc.title} ({_doc_meta(doc)})"
        for doc in documents
    ]

    slides: list[SlideSpec] = [
        SlideSpec(
            kind="cover",
            title=deck_title,
            subtitle=subtitle,
            kicker="GENERATED PRESENTATION",
            bullets=sources,
            source=documents[0].source_name,
            notes="\n".join(f"- {source}" for source in sources),
        )
    ]

    sections_by_doc: dict[int, list[Section]] = {}
    for index, doc in enumerate(documents):
        sections_by_doc[index] = _sections_for_document(doc)

    if len(documents) > 1 or sum(len(v) for v in sections_by_doc.values()) > 3:
        slides.append(_agenda_slide(documents, sections_by_doc))

    for index, doc in enumerate(documents):
        sections = sections_by_doc.get(index, [])
        if len(documents) > 1:
            details = [
                _doc_meta(doc),
                _plural(len(sections), "section"),
            ]
            if doc.ocr_block_count:
                details.append("OCR-assisted")
            slides.append(
                SlideSpec(
                    kind="section",
                    title=doc.title,
                    subtitle=" · ".join(details),
                    kicker=f"SECTION {index + 1:02d}",
                    source=doc.source_name,
                    notes=doc.text[:2500],
                )
            )

        for section in sections:
            slides.extend(
                _content_slides_for_section(section, doc, per_slide=per_slide)
            )

        if include_images:
            for image in best_figures(doc, limit=2):
                figure = _figure_slide(image, doc)
                if figure is not None:
                    slides.append(figure)

    slides.append(_summary_slide(documents, sections_by_doc))
    slides = _trim_to_budget(slides, max_slides=max_slides)

    return DeckPlan(
        title=deck_title,
        subtitle=subtitle,
        sources=sources,
        slides=slides,
    )
