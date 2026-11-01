"""Small data objects shared by the parser, planner and writers."""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path


@dataclass
class TextBlock:
    """A readable piece of a source document."""

    text: str
    page: int | None = None
    kind: str = "paragraph"  # title | heading | paragraph | list | table | ocr | caption
    level: int = 0
    source: str = ""
    image: str | None = None

    @property
    def is_ocr(self) -> bool:
        return self.kind == "ocr"


@dataclass
class DocumentImage:
    """An image extracted from a PDF or DOCX, plus its OCR text."""

    path: str
    page: int | None
    width: int
    height: int
    ocr_text: str = ""
    caption: str = ""
    source: str = ""


@dataclass
class ParsedDocument:
    path: Path
    title: str
    blocks: list[TextBlock] = field(default_factory=list)
    images: list[DocumentImage] = field(default_factory=list)
    page_count: int = 0

    @property
    def stem(self) -> str:
        return self.path.stem

    @property
    def source_name(self) -> str:
        return self.path.name

    @property
    def text(self) -> str:
        return "\n\n".join(block.text for block in self.blocks if block.text.strip())

    @property
    def ocr_block_count(self) -> int:
        return sum(1 for block in self.blocks if block.is_ocr)

    @property
    def word_count(self) -> int:
        return len(self.text.split())

    def to_prompt_text(self, include_ocr: bool = True) -> str:
        """A source-labelled representation used by the Gemini prompt."""

        lines = [
            f"### SOURCE: {self.source_name} "
            f"({self.page_count or '?'} pages, {self.word_count} words)"
        ]
        for block in self.blocks:
            if not block.text.strip():
                continue
            if block.is_ocr and not include_ocr:
                continue
            page = f"p.{block.page} " if block.page else ""
            kind = block.kind.upper()
            lines.append(f"[{page}{kind}] {block.text.strip()}")
        return "\n".join(lines)


@dataclass
class SlideSpec:
    """Everything needed to draw one slide, independent of python-pptx."""

    kind: str  # cover | agenda | section | content | figure | summary
    title: str
    subtitle: str = ""
    kicker: str = ""
    bullets: list[str] = field(default_factory=list)
    takeaway: str = ""
    source: str = ""
    image_path: str | None = None
    caption: str = ""
    notes: str = ""


@dataclass
class DeckPlan:
    title: str
    subtitle: str
    sources: list[str]
    slides: list[SlideSpec]


@dataclass
class DeckResult:
    path: Path
    slide_count: int
    plan: DeckPlan


@dataclass
class MCQ:
    number: int
    question: str
    options: dict[str, str]
    explanations: dict[str, str]
    answer: str

    def to_dict(self) -> dict:
        return {
            "number": self.number,
            "question": self.question,
            "options": self.options,
            "explanations": self.explanations,
            "answer": self.answer,
        }


@dataclass
class McqBank:
    questions: list[MCQ]
    raw_text: str
    sources: list[str]
    model: str
    generated_at: str

    @property
    def count(self) -> int:
        return len(self.questions)
