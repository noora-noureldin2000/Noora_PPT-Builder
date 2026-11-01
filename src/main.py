"""Command-line entry point for Noora PPT Builder.

Examples
--------
Single document, slides + MCQs::

    python src/main.py --input sample.pdf --mode all

Several documents (or a whole folder)::

    python src/main.py --input "docs/*.pdf" "notes.docx" --mode all

Slides only, without OCR or slide images::

    python src/main.py --input report.pdf --mode ppt --no-ocr --no-slide-images
"""

from __future__ import annotations

import argparse
import json
import logging
import re
import sys
from datetime import datetime, timezone
from pathlib import Path

# Allow both `python src/main.py` and `python -m src.main`.
if __package__ in (None, ""):
    sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from src.config import (  # noqa: E402
    GEMINI_MODEL,
    SLIDE_IMAGE_WIDTH,
    tesseract_available,
)
from src.mcq import (  # noqa: E402
    McqUnavailable,
    estimate_question_count,
    generate_mcq_bank,
    write_mcq_outputs,
)
from src.models import ParsedDocument  # noqa: E402
from src.parser import expand_inputs, parse_document  # noqa: E402
from src.ppt import create_deck, generate_slide_images  # noqa: E402

LOGGER = logging.getLogger("noora")


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------
def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="noora-ppt-builder",
        description=(
            "Parse PDF/DOCX files (with OCR for embedded images) and generate "
            "an editable 16:9 PowerPoint deck plus formatted MCQs."
        ),
    )
    parser.add_argument(
        "--input",
        nargs="+",
        required=True,
        metavar="PATH",
        help="One or more PDF/DOCX files, directories, or glob patterns.",
    )
    parser.add_argument(
        "--mode",
        choices=["ppt", "mcq", "all"],
        default="all",
        help="What to generate (default: all).",
    )
    parser.add_argument(
        "--outdir",
        default="output",
        help="Directory for all outputs (default: output/).",
    )
    parser.add_argument(
        "--title",
        default=None,
        help="Deck title. Defaults to the document title.",
    )
    parser.add_argument(
        "--max-slides",
        type=int,
        default=24,
        help="Upper bound on deck length (default: 24).",
    )
    parser.add_argument(
        "--per-slide",
        type=int,
        default=3,
        help="Bullets per content slide (default: 3).",
    )
    parser.add_argument(
        "--no-ocr",
        action="store_true",
        help="Skip OCR of embedded images.",
    )
    parser.add_argument(
        "--ocr-lang",
        default="eng",
        help="Tesseract language code (default: eng).",
    )
    parser.add_argument(
        "--no-images",
        action="store_true",
        help="Do not place extracted figures on slides.",
    )
    parser.add_argument(
        "--no-slide-images",
        action="store_true",
        help="Skip PNG rendering of the deck.",
    )
    parser.add_argument(
        "--image-backend",
        choices=["auto", "powerpoint", "libreoffice", "pillow"],
        default="auto",
        help="Rendering backend for slide PNGs (default: auto).",
    )
    parser.add_argument(
        "--image-width",
        type=int,
        default=SLIDE_IMAGE_WIDTH,
        help=f"PNG width in pixels (default: {SLIDE_IMAGE_WIDTH}).",
    )
    parser.add_argument(
        "--mcq-count",
        type=int,
        default=None,
        help="Total number of MCQs to request (default: estimated from length).",
    )
    parser.add_argument(
        "--model",
        default=None,
        help=f"Gemini model for MCQs (default: {GEMINI_MODEL}).",
    )
    parser.add_argument(
        "--quiet",
        action="store_true",
        help="Only show warnings and errors.",
    )
    return parser


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def _slug(text: str, fallback: str = "deck") -> str:
    slug = re.sub(r"[^A-Za-z0-9]+", "_", text or "").strip("_").lower()
    return slug[:60] or fallback


def _unique_stems(paths: list[Path]) -> dict[Path, str]:
    used: dict[str, int] = {}
    mapping: dict[Path, str] = {}
    for path in paths:
        stem = _slug(path.stem, "document")
        count = used.get(stem, 0)
        used[stem] = count + 1
        mapping[path] = stem if count == 0 else f"{stem}_{count + 1}"
    return mapping


def _save_parsed_text(documents: list[ParsedDocument], outdir: Path, stems: dict[Path, str]) -> Path:
    text_dir = outdir / "text"
    text_dir.mkdir(parents=True, exist_ok=True)

    for document in documents:
        stem = stems.get(document.path, _slug(document.path.stem, "document"))
        (text_dir / f"{stem}.txt").write_text(
            document.to_prompt_text(include_ocr=True), encoding="utf-8"
        )

    combined = "\n\n".join(
        document.to_prompt_text(include_ocr=True) for document in documents
    )
    combined_path = text_dir / "_all_sources.txt"
    combined_path.write_text(combined, encoding="utf-8")
    return combined_path


def _document_summary(document: ParsedDocument, path: Path) -> dict:
    return {
        "file": str(path),
        "name": document.source_name,
        "title": document.title,
        "pages": document.page_count,
        "words": document.word_count,
        "blocks": len(document.blocks),
        "ocr_blocks": document.ocr_block_count,
        "images": len(document.images),
    }


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def run(args: argparse.Namespace) -> int:
    outdir = Path(args.outdir).expanduser()
    outdir.mkdir(parents=True, exist_ok=True)

    try:
        inputs = expand_inputs(args.input)
    except Exception as exc:
        LOGGER.error("Could not resolve --input: %s", exc)
        return 2

    if not inputs:
        LOGGER.error("No PDF or DOCX files were found.")
        return 2

    LOGGER.info("Found %s input file(s).", len(inputs))
    if not args.no_ocr and not tesseract_available():
        LOGGER.warning(
            "Tesseract OCR was not found. Install it (Windows: "
            r"C:\Program Files\Tesseract-OCR\tesseract.exe) or pass --no-ocr."
        )

    assets_dir = outdir / "assets"
    assets_dir.mkdir(parents=True, exist_ok=True)

    documents: list[ParsedDocument] = []
    errors: list[str] = []
    for path in inputs:
        try:
            LOGGER.info("Parsing %s", path.name)
            document = parse_document(
                path,
                use_ocr=not args.no_ocr,
                assets_dir=assets_dir,
                ocr_lang=args.ocr_lang,
            )
            if document.word_count == 0 and not document.images:
                LOGGER.warning("No readable text found in %s.", path.name)
            documents.append(document)
        except Exception as exc:
            LOGGER.error("Failed to parse %s: %s", path.name, exc)
            errors.append(f"parse {path.name}: {exc}")

    if not documents:
        LOGGER.error("Nothing could be parsed; aborting.")
        return 1

    stems = _unique_stems([document.path for document in documents])
    combined_text_path = _save_parsed_text(documents, outdir, stems)
    LOGGER.info("Source text saved to %s", combined_text_path)

    for document in documents:
        LOGGER.info(
            "  %-32s %4s pages  %6s words  %3s OCR blocks  %2s figures",
            document.source_name[:32],
            document.page_count or "-",
            document.word_count,
            document.ocr_block_count,
            len(document.images),
        )

    manifest: dict = {
        "generated_at": datetime.now(timezone.utc).astimezone().isoformat(timespec="seconds"),
        "mode": args.mode,
        "inputs": [str(path) for path in inputs],
        "documents": [
            _document_summary(document, document.path) for document in documents
        ],
        "outputs": {},
        "errors": errors,
        "warnings": [],
    }

    deck_stem = (
        stems.get(documents[0].path, "deck")
        if len(documents) == 1
        else _slug(args.title or "combined_deck", "combined_deck")
    )

    # ------------------------------------------------------------------
    # PowerPoint
    # ------------------------------------------------------------------
    if args.mode in {"ppt", "all"}:
        deck_path = outdir / f"{deck_stem}_slides.pptx"
        try:
            result = create_deck(
                documents,
                deck_path,
                title=args.title,
                max_slides=max(1, args.max_slides),
                include_images=not args.no_images,
                per_slide=max(1, args.per_slide),
            )
            manifest["outputs"]["pptx"] = str(result.path)
            manifest["outputs"]["slide_count"] = result.slide_count
            LOGGER.info("Deck written: %s (%s slides)", result.path, result.slide_count)
        except Exception as exc:
            LOGGER.exception("Deck generation failed: %s", exc)
            errors.append(f"deck: {exc}")

        if not args.no_slide_images and (outdir / f"{deck_stem}_slides.pptx").exists():
            try:
                images = generate_slide_images(
                    outdir / f"{deck_stem}_slides.pptx",
                    outdir / "slides",
                    width=max(320, args.image_width),
                    backend=args.image_backend,
                )
                manifest["outputs"]["slide_images"] = [str(path) for path in images]
                LOGGER.info(
                    "Slide images written to %s (%s files)",
                    outdir / "slides",
                    len(images),
                )
            except Exception as exc:
                LOGGER.warning("Slide image rendering failed: %s", exc)
                manifest["warnings"].append(f"slide images: {exc}")

    # ------------------------------------------------------------------
    # MCQs
    # ------------------------------------------------------------------
    if args.mode in {"mcq", "all"}:
        requested = estimate_question_count(documents, args.mcq_count)
        LOGGER.info("Requesting about %s MCQs from %s document(s).", requested, len(documents))
        try:
            bank = generate_mcq_bank(
                documents,
                question_count=args.mcq_count,
                model=args.model,
            )
            mcq_stem = deck_stem
            paths = write_mcq_outputs(bank, outdir, mcq_stem)
            manifest["outputs"]["mcqs"] = {key: str(value) for key, value in paths.items()}
            manifest["outputs"]["mcq_count"] = bank.count
            LOGGER.info("MCQs written: %s (%s questions)", paths["mcqs_txt"], bank.count)
        except McqUnavailable as exc:
            LOGGER.error("%s", exc)
            manifest["warnings"].append(f"mcq: {exc}")
            if args.mode == "mcq":
                return 1
        except Exception as exc:
            LOGGER.exception("MCQ generation failed: %s", exc)
            errors.append(f"mcq: {exc}")
            if args.mode == "mcq":
                return 1

    manifest_path = outdir / "manifest.json"
    manifest["errors"] = errors
    manifest_path.write_text(
        json.dumps(manifest, indent=2, ensure_ascii=False), encoding="utf-8"
    )

    LOGGER.info("")
    LOGGER.info("Done. Outputs in %s", outdir.resolve())
    LOGGER.info("  manifest: %s", manifest_path)
    for key, value in manifest["outputs"].items():
        LOGGER.info("  %-13s %s", key + ":", value)

    return 0 if not errors else 1


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    logging.basicConfig(
        level=logging.WARNING if args.quiet else logging.INFO,
        format="%(message)s",
    )
    try:
        return run(args)
    except KeyboardInterrupt:
        LOGGER.error("Interrupted.")
        return 130


if __name__ == "__main__":
    raise SystemExit(main())
