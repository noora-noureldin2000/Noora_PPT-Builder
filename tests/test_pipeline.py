"""Fast, offline checks for the parser -> planner -> deck -> MCQ pipeline.

Run with::

    python -m unittest discover -s tests -v

OCR is disabled in these tests, so they do not need Tesseract.
"""

from __future__ import annotations

import sys
import tempfile
import unittest
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import docx  # noqa: E402
import fitz  # noqa: E402

from src.mcq import (  # noqa: E402
    format_answer_key,
    parse_mcq_text,
    write_mcq_outputs,
)
from src.models import McqBank, ParsedDocument, TextBlock  # noqa: E402
from src.parser import parse_docx, parse_pdf  # noqa: E402
from src.ppt import create_slides  # noqa: E402
from src.slide_planner import plan_deck  # noqa: E402
from src.textutils import sentence_chunks  # noqa: E402


SAMPLE_MCQ = """
1. Which of the following statements regarding COPD diagnosis is CORRECT?
A. A post-bronchodilator FEV1/FVC below 0.70 confirms persistent airflow limitation.
B. Spirometry is optional when symptoms are typical.
C. A normal chest radiograph excludes COPD.
D. FEV1/FVC above 0.90 is diagnostic.

A. Correct: A ratio below 0.70 after bronchodilator confirms airflow limitation.
B. Incorrect: Spirometry is required to establish the diagnosis.
C. Incorrect: Radiology does not exclude the diagnosis.
D. Incorrect: A high ratio does not confirm obstruction.

2. Which of the following statements regarding asthma controllers is CORRECT?
A. Short-acting beta agonists are the preferred long-term controller.
B. Inhaled corticosteroids are the most effective controller medication.
C. Montelukast replaces inhaled corticosteroids in all adults.
D. Controller therapy is unnecessary for mild symptoms.

A. Incorrect: SABAs relieve symptoms but are not controllers.
B. Correct: Inhaled corticosteroids remain the most effective controller.
C. Incorrect: Montelukast is an add-on therapy.
D. Incorrect: Persistent symptoms warrant controller therapy.
"""


def _sample_document() -> ParsedDocument:
    paragraphs = [
        "Asthma is a chronic inflammatory disorder of the airways.",
        "Common triggers include pollen, cold air and respiratory infections.",
        "Diagnosis relies on a typical history and evidence of variable airflow limitation.",
        "Inhaled corticosteroids are the most effective controller medication.",
        "A written action plan reduces emergency department visits.",
        "Inhaler technique should be reviewed at every visit.",
        "Smoking cessation improves long-term outcomes.",
    ]
    return ParsedDocument(
        path=Path("sample.pdf"),
        title="Asthma Essentials",
        blocks=[TextBlock(text=text, source="sample.pdf") for text in paragraphs],
        page_count=2,
    )


class TextUtilsTests(unittest.TestCase):
    def test_sentence_chunks_returns_readable_bullets(self):
        text = (
            "Inhaled corticosteroids are the most effective controller medication. "
            "A short-acting beta agonist relieves symptoms but is not a controller. "
            "Written action plans reduce emergency department visits and improve control."
        )
        chunks = sentence_chunks(text, max_chars=120)
        self.assertGreaterEqual(len(chunks), 2)
        self.assertTrue(all(len(chunk) <= 140 for chunk in chunks))
        self.assertEqual(len(chunks), len(set(chunks)))


class ParserTests(unittest.TestCase):
    def test_parse_docx(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "notes.docx"
            document = docx.Document()
            document.add_heading("Asthma in Primary Care", level=0)
            document.add_heading("Diagnosis", level=1)
            document.add_paragraph(
                "Spirometry confirms variable expiratory airflow limitation."
            )
            document.add_paragraph("A trial of treatment may support the diagnosis.")
            document.save(path)

            parsed = parse_docx(path, use_ocr=False)

            self.assertEqual(parsed.page_count, 0)
            self.assertGreaterEqual(len(parsed.blocks), 3)
            self.assertIn("Spirometry", parsed.text)
            self.assertEqual(parsed.title, "Asthma in Primary Care")

    def test_parse_pdf(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "guideline.pdf"
            pdf = fitz.open()
            page = pdf.new_page()
            page.insert_text((60, 80), "Asthma in Primary Care", fontsize=20)
            page.insert_text(
                (60, 140),
                "Spirometry confirms variable expiratory airflow limitation.",
                fontsize=12,
            )
            pdf.save(path)
            pdf.close()

            parsed = parse_pdf(path, use_ocr=False)

            self.assertEqual(parsed.page_count, 1)
            self.assertIn("Spirometry", parsed.text)
            self.assertGreater(parsed.word_count, 5)


class PlannerTests(unittest.TestCase):
    def test_deck_structure_and_budget(self):
        documents = [_sample_document()]
        plan = plan_deck(documents, max_slides=8, include_images=False)

        self.assertEqual(plan.slides[0].kind, "cover")
        self.assertEqual(plan.slides[-1].kind, "summary")
        self.assertLessEqual(len(plan.slides), 8)
        self.assertTrue(any(slide.kind == "content" for slide in plan.slides))
        for slide in plan.slides:
            self.assertTrue(slide.title.strip())

    def test_multi_document_plan_has_sections(self):
        second = _sample_document()
        second.path = Path("second.docx")
        second.title = "COPD Essentials"
        second.page_count = 0
        plan = plan_deck([_sample_document(), second], max_slides=20)
        kinds = [slide.kind for slide in plan.slides]
        self.assertIn("agenda", kinds)
        self.assertGreaterEqual(kinds.count("section"), 2)


class DeckTests(unittest.TestCase):
    def test_build_pptx(self):
        with tempfile.TemporaryDirectory() as tmp:
            output = Path(tmp) / "deck.pptx"
            text = (
                "Hypertension is a major modifiable risk factor for cardiovascular disease.\n"
                "Blood pressure should be measured on two occasions before diagnosis.\n"
                "Lifestyle measures include salt restriction, weight loss and exercise.\n"
                "Thiazide diuretics and ACE inhibitors are first-line options.\n"
            )
            result = create_slides(text, output, title="Hypertension Essentials")

            self.assertTrue(output.is_file())
            self.assertGreaterEqual(result.slide_count, 3)

            from pptx import Presentation

            prs = Presentation(str(output))
            self.assertAlmostEqual(prs.slide_width.inches, 13.333, places=2)
            self.assertAlmostEqual(prs.slide_height.inches, 7.5, places=1)
            notes = [
                slide.notes_slide.notes_text_frame.text
                for slide in prs.slides
                if slide.has_notes_slide
            ]
            self.assertTrue(any(text.strip() for text in notes))


class McqTests(unittest.TestCase):
    def test_parse_questions(self):
        questions = parse_mcq_text(SAMPLE_MCQ)
        self.assertEqual(len(questions), 2)
        self.assertEqual(questions[0].answer, "A")
        self.assertEqual(questions[1].answer, "B")
        self.assertEqual(sorted(questions[0].options), ["A", "B", "C", "D"])
        self.assertIn("Correct", questions[0].explanations["A"])

    def test_parse_markdown_and_answer_line(self):
        raw = """
**Which of the following statements regarding inhaler technique is CORRECT?**
**A.** Rinse the mouth after using a corticosteroid inhaler.
**B.** Shake the inhaler only after use.
**C.** A spacer reduces lung deposition.
**D.** Wait ten minutes between two puffs.

**A. Correct:** Rinsing reduces oral candidiasis.
**B. Incorrect:** The inhaler is shaken before use.
**C. Incorrect:** A spacer improves lung deposition.
**D. Incorrect:** About thirty seconds is enough.
Answer: A
"""
        questions = parse_mcq_text(raw)
        self.assertEqual(len(questions), 1)
        self.assertEqual(questions[0].answer, "A")
        self.assertEqual(len(questions[0].options), 4)

    def test_writers(self):
        questions = parse_mcq_text(SAMPLE_MCQ)
        bank = McqBank(
            questions=questions,
            raw_text=SAMPLE_MCQ,
            sources=["sample.pdf"],
            model="gemini-2.5-flash",
            generated_at=datetime(2026, 1, 1).isoformat(),
        )
        with tempfile.TemporaryDirectory() as tmp:
            paths = write_mcq_outputs(bank, tmp, "sample")
            for key in ("mcqs_txt", "answer_key", "mcqs_json"):
                self.assertTrue(paths[key].is_file())

            text = paths["mcqs_txt"].read_text(encoding="utf-8")
            self.assertIn("Which of the following statements", text)
            self.assertIn("A. Correct:", text)
            self.assertIn("B. Incorrect:", text)

            key = format_answer_key(bank)
            self.assertIn("01. A", key)
            self.assertIn("02. B", key)


if __name__ == "__main__":
    unittest.main(verbosity=2)
