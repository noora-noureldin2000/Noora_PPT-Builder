"""Text cleaning, sentence splitting and chunking helpers."""

from __future__ import annotations

import re
import unicodedata

_WS_RE = re.compile(r"[ \t\f\v]+")
_SENTENCE_RE = re.compile(r"(?<=[.!?])\s+(?=[A-Z0-9(\[\"'])")
_BULLET_RE = re.compile(r"^\s*(?:[-*•·▪◦]|\d+[.)])\s+")
_CLAUSE_RE = re.compile(r"(?<=[;:])\s+")

_KEYWORD_WEIGHTS = {
    "therefore": 3,
    "conclusion": 3,
    "recommend": 3,
    "important": 2,
    "significant": 2,
    "associated": 2,
    "demonstrated": 2,
    "showed": 2,
    "shows": 2,
    "results": 2,
    "finding": 2,
    "mechanism": 2,
    "treatment": 2,
    "diagnosis": 2,
    "risk": 2,
    "however": 1,
    "because": 1,
    "suggest": 1,
    "increase": 1,
    "decrease": 1,
    "improved": 1,
}


def normalize_whitespace(text: str) -> str:
    text = unicodedata.normalize("NFKC", str(text or ""))
    text = text.replace("\r\n", "\n").replace("\r", "\n").replace("\xa0", " ")
    text = _WS_RE.sub(" ", text)
    return "\n".join(line.strip() for line in text.splitlines()).strip()


def clean_ocr_text(text: str) -> str:
    """Clean common OCR artefacts while keeping the words intact."""

    text = normalize_whitespace(text)
    if not text:
        return ""

    # Re-join words that Tesseract split with a hyphen at a line break.
    text = re.sub(r"(\w)-\n(\w)", r"\1\2", text)

    cleaned_lines: list[str] = []
    for line in text.splitlines():
        line = line.strip(" |_~`")
        if not line:
            continue
        alnum = sum(ch.isalnum() for ch in line)
        if len(line) <= 12 and alnum / max(len(line), 1) < 0.25:
            continue
        cleaned_lines.append(line)

    # OCR of a figure is usually one thought; join wrapped lines with spaces.
    joined = " ".join(cleaned_lines)
    joined = re.sub(r"\s+([,.;:!?])", r"\1", joined)
    joined = re.sub(r"([,.;:!?])(?=[A-Za-z])", r"\1 ", joined)
    joined = re.sub(r"\s{2,}", " ", joined)
    return joined.strip()


def split_sentences(text: str) -> list[str]:
    text = normalize_whitespace(text).replace("\n", " ")
    if not text:
        return []

    sentences = _SENTENCE_RE.split(text)
    # A sentence with no terminal punctuation can still be a useful bullet.
    return [sentence.strip() for sentence in sentences if sentence.strip()]


def _split_long_clause(sentence: str, max_chars: int) -> list[str]:
    if len(sentence) <= max_chars:
        return [sentence]

    parts = [part.strip() for part in _CLAUSE_RE.split(sentence) if part.strip()]
    if len(parts) <= 1:
        return [truncate_at_word(sentence, max_chars)]

    chunks: list[str] = []
    current = ""
    for part in parts:
        candidate = f"{current} {part}".strip()
        if current and len(candidate) > max_chars:
            chunks.append(current)
            current = part
        else:
            current = candidate
    if current:
        chunks.append(current)
    return chunks


def sentence_chunks(
    text: str, max_chars: int = 190, min_chars: int = 35
) -> list[str]:
    """Turn prose into slide-sized bullets, preserving meaning."""

    sentences = split_sentences(text)
    if not sentences:
        return []

    pieces: list[str] = []
    for sentence in sentences:
        pieces.extend(_split_long_clause(sentence, max_chars))

    chunks: list[str] = []
    current = ""
    for piece in pieces:
        piece = _BULLET_RE.sub("", piece).strip()
        if not piece:
            continue
        candidate = f"{current} {piece}".strip()
        if current and len(candidate) > max_chars:
            if len(current) >= min_chars:
                chunks.append(current)
                current = piece
            else:
                # Merge a short fragment with the previous chunk when possible.
                if chunks and len(chunks[-1]) + len(candidate) <= max_chars + 25:
                    chunks[-1] = f"{chunks[-1]} {candidate}".strip()
                    current = ""
                else:
                    chunks.append(current)
                    current = piece
        else:
            current = candidate

    if current:
        chunks.append(current)

    cleaned: list[str] = []
    for chunk in chunks:
        chunk = truncate_at_word(chunk.strip(), max_chars + 20)
        if len(chunk) >= 12:
            cleaned.append(chunk)
    return dedupe_keep_order(cleaned)


def truncate_at_word(text: str, max_chars: int) -> str:
    text = normalize_whitespace(text).replace("\n", " ")
    if len(text) <= max_chars:
        return text
    cut = text[:max_chars].rsplit(" ", 1)[0].rstrip(" ,.;:-")
    return f"{cut}…"


def short_title(text: str, max_words: int = 9, max_chars: int = 72) -> str:
    text = normalize_whitespace(text).replace("\n", " ")
    text = re.sub(r"^\s*(?:\d+[.)]|[-*•])\s*", "", text)
    first = split_sentences(text)
    candidate = first[0] if first else text
    candidate = candidate.split(":", 1)[0].strip(" .;:-")

    words = candidate.split()
    if len(words) > max_words:
        candidate = " ".join(words[:max_words])
    candidate = truncate_at_word(candidate, max_chars).rstrip("…").strip()
    return candidate or "Key Concepts"


def dedupe_keep_order(values: list[str]) -> list[str]:
    seen: set[str] = set()
    result: list[str] = []
    for value in values:
        key = re.sub(r"\W+", " ", value.lower()).strip()
        if not key or key in seen:
            continue
        seen.add(key)
        result.append(value)
    return result


def looks_meaningful(text: str, min_alpha_ratio: float = 0.35) -> bool:
    text = normalize_whitespace(text)
    if len(text) < 8:
        return False
    alnum = sum(ch.isalnum() for ch in text)
    return alnum / max(len(text), 1) >= min_alpha_ratio


def key_sentence(text: str, max_chars: int = 150) -> str:
    """Pick a compact sentence suitable for a 'key point' sidebar."""

    sentences = split_sentences(text)
    if not sentences:
        return ""

    best = ""
    best_score = -1.0
    for sentence in sentences:
        if len(sentence) < 35 or len(sentence) > 260:
            continue
        lowered = sentence.lower()
        score = sum(weight for word, weight in _KEYWORD_WEIGHTS.items() if word in lowered)
        # Prefer concise, information-dense sentences.
        score += min(len(sentence), 180) / 180
        if score > best_score:
            best = sentence
            best_score = score

    if not best:
        best = max(sentences, key=len)
    return truncate_at_word(best, max_chars)


def strip_page_noise(text: str) -> str:
    text = normalize_whitespace(text)
    text = re.sub(r"^\s*(?:page\s*)?\d+\s*(?:of\s*\d+)?\s*$", "", text, flags=re.I)
    return text.strip()
