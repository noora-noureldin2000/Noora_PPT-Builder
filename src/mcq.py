"""Gemini-powered MCQ generation with a strict, predictable output format.

The generator sends the parsed source material (raw text *and* OCR text) to
``gemini-2.5-flash`` and then parses the answer back into :class:`MCQ`
objects.  The written ``.txt`` file keeps exactly the structure requested by
the project brief, and the answer key is exported separately so it can be
handed out on its own.

The module works with one document or many; questions are distributed across
the sources so every file is represented.
"""

from __future__ import annotations

import json
import logging
import re
import time
from datetime import datetime, timezone
from pathlib import Path

from src.config import GEMINI_API_KEY, GEMINI_MODEL
from src.models import MCQ, McqBank, ParsedDocument, TextBlock
from src.textutils import normalize_whitespace, truncate_at_word

LOGGER = logging.getLogger(__name__)

DEFAULT_QUESTION_COUNT = 12
MAX_QUESTION_COUNT = 60
CHUNK_CHARS = 14_000
RETRY_ATTEMPTS = 3

FORMAT_BLOCK = """Which of the following statements regarding [Topic] is CORRECT?
A. [Option A]
B. [Option B]
C. [Option C]
D. [Option D]

A. Incorrect: [Explanation]
B. Incorrect: [Explanation]
C. Correct: [Explanation]
D. Incorrect: [Explanation]"""

SYSTEM_PROMPT = f"""You are a senior medical educator writing high-quality,
single-best-answer examination items from supplied source material.

STRICT OUTPUT RULES
1. Return ONLY the questions. No preamble, no commentary, no markdown, no
   numbering other than the question numbers, no "Here are the questions".
2. Every item must contain exactly four options labelled A, B, C and D, then a
   blank line, then four explanation lines labelled A-D.
3. Exactly ONE option is CORRECT. In the explanation block that option must
   start with "Correct:" and the other three must start with "Incorrect:".
4. Use this exact structure for every single question:

{FORMAT_BLOCK}

5. Questions must be answerable from the SOURCE MATERIAL alone. Do not invent
   facts, statistics, drug doses or study results that are not present.
6. Prefer clinically meaningful, discriminating stems ("Which of the following
   statements regarding X is CORRECT?"). Avoid "all of the above", "none of the
   above", negative stems and trivial recall.
7. Distribute the items across the material instead of clustering on the first
   paragraph, and do not repeat the same fact twice.
8. Keep each option under 30 words and each explanation to one or two sentences
   that state why the option is correct or incorrect.
9. If the material supports fewer items than requested, return only the
   high-quality items that are fully supported."""


class McqUnavailable(RuntimeError):
    """Raised when MCQ generation cannot run (missing SDK or API key)."""


# ---------------------------------------------------------------------------
# Prompt assembly
# ---------------------------------------------------------------------------
def estimate_question_count(documents: list[ParsedDocument], requested: int | None = None) -> int:
    if requested and requested > 0:
        return min(int(requested), MAX_QUESTION_COUNT)
    words = sum(max(doc.word_count, 0) for doc in documents)
    if words <= 0:
        return 0
    return max(4, min(30, round(words / 220)))


def _split_text(text: str, max_chars: int = CHUNK_CHARS) -> list[str]:
    lines = text.splitlines()
    chunks: list[str] = []
    buffer: list[str] = []
    size = 0
    for line in lines:
        line_len = len(line) + 1
        if buffer and size + line_len > max_chars:
            chunks.append("\n".join(buffer).strip())
            buffer = []
            size = 0
        buffer.append(line)
        size += line_len
    if buffer:
        chunks.append("\n".join(buffer).strip())
    return [chunk for chunk in chunks if len(chunk) > 80]


def _document_chunks(document: ParsedDocument) -> list[str]:
    chunks = _split_text(document.to_prompt_text(include_ocr=True))
    if not chunks:
        chunks = _split_text(document.text)
    if not chunks:
        return []
    # Re-attach the source header to every chunk after the first.
    header = (
        f"### SOURCE: {document.source_name} "
        f"({document.page_count or '?'} pages, {document.word_count} words)"
    )
    labelled = []
    for chunk in chunks:
        if chunk.startswith("### SOURCE"):
            labelled.append(chunk)
        else:
            labelled.append(f"{header}\n{chunk}")
    return labelled


def _build_prompt(source_text: str, question_count: int, source_name: str) -> str:
    return (
        f"{SYSTEM_PROMPT}\n\n"
        f"Generate exactly {question_count} questions from {source_name}. "
        f"Stay strictly inside the source material below.\n\n"
        f"SOURCE MATERIAL\n"
        f"----------------\n"
        f"{source_text}\n"
        f"----------------\n\n"
        f"Now output the {question_count} questions in the required format."
    )


# ---------------------------------------------------------------------------
# Gemini call
# ---------------------------------------------------------------------------
def _client(api_key: str | None):
    key = (api_key or GEMINI_API_KEY or "").strip()
    if not key:
        raise McqUnavailable(
            "GEMINI_API_KEY is not set. Set it in your environment or in a "
            "local .env file (see .env.example) to generate MCQs."
        )
    try:
        from google import genai
    except Exception as exc:  # pragma: no cover - dependency issue
        raise McqUnavailable(
            "google-genai is not installed. Run: pip install -r requirements.txt"
        ) from exc
    return genai.Client(api_key=key)


def _call_gemini(client, prompt: str, model: str) -> str:
    last_error: Exception | None = None
    for attempt in range(1, RETRY_ATTEMPTS + 1):
        try:
            response = client.models.generate_content(
                model=model,
                contents=prompt,
                config={"temperature": 0.35, "top_p": 0.9},
            )
            text = getattr(response, "text", None)
            if not text and getattr(response, "candidates", None):
                parts = []
                for candidate in response.candidates:
                    content = getattr(candidate, "content", None)
                    for part in getattr(content, "parts", []) or []:
                        if getattr(part, "text", None):
                            parts.append(part.text)
                text = "\n".join(parts)
            if text and text.strip():
                return text
            last_error = RuntimeError("Gemini returned an empty response")
        except Exception as exc:  # network, quota, safety blocks...
            last_error = exc
            message = str(exc).lower()
            if "api key" in message or "unauthenticated" in message or "permission" in message:
                raise McqUnavailable(f"Gemini rejected the API key: {exc}") from exc
            LOGGER.warning(
                "Gemini attempt %s/%s failed: %s", attempt, RETRY_ATTEMPTS, exc
            )
        if attempt < RETRY_ATTEMPTS:
            time.sleep(2 ** attempt)

    raise RuntimeError(f"Gemini request failed after {RETRY_ATTEMPTS} attempts: {last_error}")


# ---------------------------------------------------------------------------
# Parsing the model output
# ---------------------------------------------------------------------------
_OPTION_RE = re.compile(r"^\s*\(?([A-D])[\)\.\:]\s+(.+?)\s*$", re.I)
_EXPL_RE = re.compile(
    r"^\s*\(?([A-D])[\)\.\:]\s*(Correct|Incorrect|True|False)\s*[:\-–]\s*(.*)$",
    re.I,
)
_ANSWER_RE = re.compile(
    r"^\s*(?:correct\s+)?answer\s*[:\-]\s*\(?([A-D])\)?\b", re.I
)
_QUESTION_START_RE = re.compile(
    r"^\s*(?:question|q)?\s*\d+\s*[\.\)\:]\s*", re.I
)


def _clean_response(text: str) -> str:
    text = (text or "").replace("\r\n", "\n").replace("\r", "\n")
    text = re.sub(r"```(?:text|markdown)?", "", text)
    text = text.replace("**", "").replace("__", "").replace("`", "")
    text = re.sub(r"[ \t]+", " ", text)
    # Some responses put every explanation on one line; split them apart.
    text = re.sub(
        r"(?<!^)\s+(?=[A-D][\)\.\:]\s*(?:Correct|Incorrect)\s*[:\-])",
        "\n",
        text,
        flags=re.I,
    )
    return text.strip()


def _strip_question_number(text: str) -> str:
    return _QUESTION_START_RE.sub("", text).strip()


def parse_mcq_text(raw_text: str, start_number: int = 1) -> list[MCQ]:
    """Turn Gemini's plain-text output into structured :class:`MCQ` objects."""

    text = _clean_response(raw_text)
    lines = [line.strip() for line in text.splitlines()]
    questions: list[MCQ] = []

    current: dict | None = None

    def flush() -> None:
        nonlocal current
        if not current:
            return
        options = current["options"]
        explanations = current["explanations"]
        question = normalize_whitespace(current["question"]).strip()
        answer = current.get("answer", "")
        if not answer:
            for letter, value in explanations.items():
                if value.lower().startswith(("correct", "true")):
                    answer = letter
                    break
        if (
            question
            and len(options) == 4
            and all(letter in options for letter in "ABCD")
        ):
            questions.append(
                MCQ(
                    number=start_number + len(questions),
                    question=question,
                    options=options,
                    explanations={
                        letter: explanations.get(letter, "")
                        for letter in "ABCD"
                    },
                    answer=answer if answer in "ABCD" else "",
                )
            )
        current = None

    for line in lines:
        if not line:
            continue

        explanation = _EXPL_RE.match(line)
        if explanation and current is not None:
            letter = explanation.group(1).upper()
            verdict = explanation.group(2).capitalize()
            body = explanation.group(3).strip()
            current["explanations"][letter] = f"{verdict}: {body}".strip()
            if verdict in {"Correct", "True"}:
                current["answer"] = letter
            continue

        answer_match = _ANSWER_RE.match(line)
        if answer_match and current is not None:
            current["answer"] = answer_match.group(1).upper()
            continue

        option = _OPTION_RE.match(line)
        if option and current is not None:
            letter = option.group(1).upper()
            value = option.group(2).strip()
            value = re.sub(
                r"^(?:Correct|Incorrect|True|False)\s*[:\-–]\s*", "", value, flags=re.I
            )
            if letter not in current["options"] and len(current["options"]) < 4:
                current["options"][letter] = value
            elif letter not in current["explanations"]:
                current["explanations"][letter] = value
            continue

        # A question stem is a numbered line, a line ending in "?", or any
        # prose line that arrives after the previous item is complete.
        is_stem = bool(_QUESTION_START_RE.match(line)) or line.rstrip().endswith("?")
        if current is None:
            if is_stem:
                current = {
                    "question": _strip_question_number(line),
                    "options": {},
                    "explanations": {},
                    "answer": "",
                }
            continue

        if not current["options"]:
            # Wrapped stem.
            current["question"] = f"{current['question']} {_strip_question_number(line)}".strip()
            continue

        if len(current["options"]) >= 4:
            flush()
            if is_stem:
                current = {
                    "question": _strip_question_number(line),
                    "options": {},
                    "explanations": {},
                    "answer": "",
                }
        # Otherwise it is stray prose inside an incomplete item: ignore it.

    flush()
    return _dedupe(questions)


def _dedupe(questions: list[MCQ]) -> list[MCQ]:
    seen: set[str] = set()
    result: list[MCQ] = []
    for question in questions:
        key = re.sub(r"\W+", " ", question.question.lower()).strip()
        if not key or key in seen:
            continue
        seen.add(key)
        question.number = len(result) + 1
        result.append(question)
    return result


# ---------------------------------------------------------------------------
# Public generation API
# ---------------------------------------------------------------------------
def generate_mcq_bank(
    documents: list[ParsedDocument],
    *,
    question_count: int | None = None,
    api_key: str | None = None,
    model: str | None = None,
) -> McqBank:
    """Generate a bank of MCQs covering every supplied document."""

    documents = [doc for doc in documents if doc.word_count > 0]
    if not documents:
        raise ValueError("No readable text was found in the supplied documents.")

    total_requested = estimate_question_count(documents, question_count)
    if total_requested <= 0:
        raise ValueError("Could not determine a sensible question count.")

    client = _client(api_key)
    model_name = (model or GEMINI_MODEL or "gemini-2.5-flash").strip()

    total_words = sum(doc.word_count for doc in documents)
    raw_parts: list[str] = []
    collected: list[MCQ] = []

    for doc in documents:
        weight = doc.word_count / max(total_words, 1)
        doc_target = max(2, round(total_requested * weight))
        chunks = _document_chunks(doc)
        if not chunks:
            LOGGER.warning("Skipping %s: no extractable text", doc.source_name)
            continue

        LOGGER.info(
            "Generating ~%s MCQs from %s (%s chunk(s))",
            doc_target,
            doc.source_name,
            len(chunks),
        )

        per_chunk = max(2, round(doc_target / len(chunks)))
        for chunk_index, chunk in enumerate(chunks, start=1):
            if chunk_index > 1:
                time.sleep(1.0)  # be gentle with the free tier
            prompt = _build_prompt(
                chunk,
                per_chunk,
                f"{doc.source_name} (part {chunk_index} of {len(chunks)})",
            )
            try:
                response_text = _call_gemini(client, prompt, model_name)
            except McqUnavailable:
                raise
            except Exception as exc:
                LOGGER.error(
                    "Skipping chunk %s/%s of %s: %s",
                    chunk_index,
                    len(chunks),
                    doc.source_name,
                    exc,
                )
                continue

            raw_parts.append(
                f"### {doc.source_name} - part {chunk_index}\n{response_text.strip()}"
            )
            parsed = parse_mcq_text(response_text, start_number=len(collected) + 1)
            LOGGER.info("  -> %s usable question(s)", len(parsed))
            collected.extend(parsed)

    unique = _dedupe(collected)
    if total_requested and len(unique) > total_requested + 4:
        unique = unique[: total_requested + 4]

    return McqBank(
        questions=unique,
        raw_text="\n\n".join(raw_parts),
        sources=[doc.source_name for doc in documents],
        model=model_name,
        generated_at=datetime.now(timezone.utc).astimezone().isoformat(timespec="seconds"),
    )


def generate_mcqs(text_content: str, question_count: int | None = None) -> str:
    """Backwards-compatible helper returning the formatted text block."""

    document = ParsedDocument(
        path=Path("input.txt"),
        title="Input",
        blocks=[
            TextBlock(text=paragraph, source="input")
            for paragraph in normalize_whitespace(text_content).splitlines()
            if paragraph.strip()
        ],
        page_count=0,
    )
    bank = generate_mcq_bank([document], question_count=question_count)
    return format_mcq_bank(bank)


# ---------------------------------------------------------------------------
# Writers
# ---------------------------------------------------------------------------
def format_question(question: MCQ) -> str:
    options = "\n".join(
        f"{letter}. {question.options.get(letter, '').strip()}" for letter in "ABCD"
    )
    explanations = "\n".join(
        f"{letter}. {question.explanations.get(letter, '').strip() or 'Incorrect: Not supported by the source.'}"
        for letter in "ABCD"
    )
    return (
        f"{question.number}. {question.question.strip()}\n"
        f"{options}\n\n"
        f"{explanations}"
    )


def format_mcq_bank(bank: McqBank) -> str:
    header = (
        "MULTIPLE CHOICE QUESTIONS\n"
        f"Sources: {', '.join(bank.sources)}\n"
        f"Generated: {bank.generated_at}  |  Model: {bank.model}\n"
        + "=" * 72
        + "\n\n"
    )
    body = "\n\n".join(format_question(question) for question in bank.questions)
    return f"{header}{body}\n"


def format_answer_key(bank: McqBank) -> str:
    lines = [
        "ANSWER KEY",
        f"Sources: {', '.join(bank.sources)}",
        f"Generated: {bank.generated_at}",
        "=" * 72,
        "",
    ]
    for question in bank.questions:
        letter = question.answer or "?"
        explanation = question.explanations.get(letter, "")
        explanation = re.sub(r"^(?:Correct|True)\s*[:\-–]\s*", "", explanation, flags=re.I)
        line = f"{question.number:02d}. {letter}"
        if explanation:
            line += f"  —  {truncate_at_word(explanation, 150)}"
        lines.append(line)
    return "\n".join(lines) + "\n"


def write_mcq_outputs(bank: McqBank, output_dir: str | Path, stem: str) -> dict[str, Path]:
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    txt_path = output_dir / f"{stem}_mcqs.txt"
    key_path = output_dir / f"{stem}_answer_key.txt"
    json_path = output_dir / f"{stem}_mcqs.json"

    txt_path.write_text(format_mcq_bank(bank), encoding="utf-8")
    key_path.write_text(format_answer_key(bank), encoding="utf-8")
    json_path.write_text(
        json.dumps(
            {
                "sources": bank.sources,
                "model": bank.model,
                "generated_at": bank.generated_at,
                "count": bank.count,
                "questions": [question.to_dict() for question in bank.questions],
            },
            indent=2,
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )

    LOGGER.info("Saved %s MCQs to %s", bank.count, txt_path)
    return {"mcqs_txt": txt_path, "answer_key": key_path, "mcqs_json": json_path}
