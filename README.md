# Noora PPT Builder

Turn PDF and DOCX source material into a **designed, editable 16:9 PowerPoint
deck** and a **formatted MCQ bank with answer key**.

The pipeline reads text page by page, runs OCR on every embedded image so
diagram labels and scanned pages are not lost, plans a paced slide sequence,
draws the deck with native PowerPoint shapes, and exports slide images.

```
PDF / DOCX  ->  text + OCR  ->  slide plan  ->  .pptx  ->  slide_*.png
                            \->  MCQ bank  ->  .txt / .json / answer key
```

---

## 1. What it produces

| Output | Location | Notes |
| --- | --- | --- |
| Deck | `output/<name>_slides.pptx` | 16:9, fully editable text and shapes |
| Slide images | `output/slides/slide_01.png ...` | Rendered with PowerPoint, LibreOffice, or an approximate Pillow preview |
| MCQs | `output/<name>_mcqs.txt` | Questions + per-option explanations |
| Answer key | `output/<name>_answer_key.txt` | One line per question |
| MCQ data | `output/<name>_mcqs.json` | Machine-readable, with metadata |
| Extracted text | `output/text/*.txt` | Full parse including OCR text |
| Figures | `output/assets/*.png` | Images extracted from the sources |
| Manifest | `output/manifest.json` | Inputs, counts, warnings, output paths |

### Deck design

* 16:9 canvas (13.333 × 7.5 in), blank layouts, no locked placeholders.
* Six slide types: cover, agenda, section divider, content, figure, summary.
* One accent colour, generous margins, thin rules, a highlighted key-point
  panel, and speaker notes on every slide.
* Font sizes are fitted to the amount of text, so a dense slide shrinks
  instead of overflowing.
* Content is chunked into 2–3 bullets per slide; a short remainder is folded
  into the previous slide instead of creating a near-empty one.
* OCR text from figures is split into `Step 1 / Step 2 / Figure 2` segments
  and treated as normal content, and selected figures are placed on their own
  "Visual Evidence" slide.

### MCQ format

Every question is generated and written in exactly this structure:

```
Which of the following statements regarding [Topic] is CORRECT?
A. [Option A]
B. [Option B]
C. [Option C]
D. [Option D]

A. Incorrect: [Explanation]
B. Incorrect: [Explanation]
C. Correct: [Explanation]
D. Incorrect: [Explanation]
```

When several documents are supplied, questions are distributed across all of
them, duplicates are removed, and the source file list is stored in the JSON
and text headers.

---

## 2. Installation

### Step 1 — Python

Python **3.10 or newer** (tested on 3.12).

```bash
cd "D:\GitHub\Noora_PPT Builder"
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

### Step 2 — Tesseract OCR (required for image OCR)

> **Important:** text extraction from PDFs and DOCX files works without
> Tesseract, but **OCR of embedded images and scanned pages will not work**
> until the Windows binary is installed.

1. Download the Windows installer from the UB Mannheim build:
   <https://github.com/UB-Mannheim/tesseract/wiki>
2. Install it to the default location:
   `C:\Program Files\Tesseract-OCR\tesseract.exe`
3. If you install it elsewhere, either edit `TESSERACT_CMD` in
   [`src/config.py`](src/config.py) or set an environment variable:

   ```powershell
   setx TESSERACT_CMD "D:\Tools\Tesseract-OCR\tesseract.exe"
   ```

4. Verify:

   ```powershell
   & "C:\Program Files\Tesseract-OCR\tesseract.exe" --version
   ```

### Step 3 — Gemini API key (required for MCQs only)

Slides are generated locally and do **not** need a key. MCQs use
`gemini-2.5-flash` through the `google-genai` SDK.

Get a key from <https://aistudio.google.com/app/apikey>, then either:

```powershell
# current PowerShell session
$env:GEMINI_API_KEY = "your_api_key_here"
```

```bat
:: classic command prompt (persists for new terminals)
setx GEMINI_API_KEY "your_api_key_here"
```

or copy `.env.example` to `.env` and paste the key there. The local `.env`
file is read automatically at startup and is git-ignored.

### Step 4 — Optional: exact slide images

| Backend | Requirement | Quality |
| --- | --- | --- |
| PowerPoint COM | Windows + Microsoft PowerPoint + `pywin32` | Exact |
| LibreOffice | `soffice` available in PATH | Exact |
| Pillow | Always available | Approximate preview |

The renderer tries them in that order. Install `pywin32` (already in
`requirements.txt` for Windows) and keep PowerPoint installed for the best
results.

### Step 5 — Verify the environment

```bash
python src/check_setup.py
```

You should see `OK` for Python, the libraries, Tesseract, the output folder,
and one rendering backend.

---

## 3. Usage

```bash
python src/main.py --input sample.pdf --mode all
```

### Multiple documents

`--input` accepts files, folders, and glob patterns — mix them freely:

```bash
python src/main.py --input "D:\papers\*.pdf" "notes.docx" "handouts\" --mode all
```

The deck gets a cover, an agenda, a section divider per document, and content
from every file. The MCQ bank samples every document too.

### Modes

| Mode | Effect |
| --- | --- |
| `ppt` | Deck + slide images only (no API key needed) |
| `mcq` | MCQ bank only |
| `all` | Both (default) |

### Common options

| Option | Default | Purpose |
| --- | --- | --- |
| `--outdir` | `output` | Where everything is written |
| `--title` | document title | Override the deck title |
| `--max-slides` | `24` | Hard cap on deck length |
| `--per-slide` | `3` | Target bullets per content slide |
| `--no-ocr` | off | Skip image OCR (faster, less content) |
| `--ocr-lang` | `eng` | Tesseract language pack, e.g. `eng+ara` |
| `--no-images` | off | Do not place figures on slides |
| `--no-slide-images` | off | Skip PNG rendering |
| `--image-backend` | `auto` | `powerpoint`, `libreoffice`, `pillow`, `auto` |
| `--image-width` | `1920` | PNG width in pixels |
| `--mcq-count` | estimated | Total questions to request |
| `--model` | `gemini-2.5-flash` | Gemini model override |
| `--quiet` | off | Only warnings and errors |

### Examples

```bash
# Slides only, from every PDF in a folder, 18 slides max
python src/main.py --input "D:\lectures\*.pdf" --mode ppt --max-slides 18

# MCQ exam bank from three DOCX files
python src/main.py --input ch1.docx ch2.docx ch3.docx --mode mcq --mcq-count 30

# Arabic + English OCR, no figure slides
python src/main.py --input scan.pdf --mode all --ocr-lang eng+ara --no-images
```

---

## 4. Project structure

```
Noora_PPT Builder/
├── requirements.txt
├── README.md
├── .env.example
├── .gitignore
├── output/                 # created on first run
├── tests/
│   └── test_pipeline.py    # offline parser/planner/deck/MCQ checks
└── src/
    ├── __init__.py
    ├── config.py           # API key, Tesseract path, design constants
    ├── models.py           # dataclasses shared across modules
    ├── textutils.py        # cleaning, sentence splitting, chunking
    ├── parser.py           # PDF/DOCX parsing + OCR + figure extraction
    ├── slide_planner.py    # sections -> paced slide plan
    ├── ppt.py              # the 16:9 design engine + generate_slide_images()
    ├── render.py           # PowerPoint COM / LibreOffice / Pillow export
    ├── mcq.py              # Gemini prompts, parsing, writers
    ├── check_setup.py      # environment diagnostics
    └── main.py             # CLI and orchestration
```

### How the parser works

1. **PDF** — `PyMuPDF` extracts text blocks with font-size information so
   titles and headings can be recognised. Pages with almost no text layer are
   rasterised and OCR'd as a whole (scanned pages).
2. **Images** — every image on every page is extracted, deduplicated, and
   passed to Tesseract. The OCR text is inserted into the reading order right
   where the figure appeared.
3. **DOCX** — paragraphs and tables are read in document order, with heading
   styles preserved; inline images are OCR'd the same way.

If Tesseract is missing, parsing continues without OCR and the CLI prints a
warning. Use `--no-ocr` to skip it deliberately.

---

## 5. Troubleshooting

**`TesseractNotFoundError` or "OCR disabled"**
Install the Tesseract Windows binary and confirm the path in
`src/config.py` (or set `TESSERACT_CMD`). Run `python src/check_setup.py`.

**`GEMINI_API_KEY is not set`**
Set the environment variable and open a new terminal, or create `.env` from
`.env.example`. `--mode ppt` works without a key.

**"Slide images are approximate previews"**
No PowerPoint COM and no LibreOffice were found. Install `pywin32` (and keep
PowerPoint installed) or install LibreOffice. The deck itself is unaffected.

**PowerPoint export hangs or fails**
Close any open PowerPoint windows, then retry with
`--image-backend libreoffice` or `--no-slide-images`.

**A slide looks too full**
Lower `--per-slide` (e.g. `2`) or raise `--max-slides`. Font sizes adapt
automatically.

**Wrong language from OCR**
Install the matching Tesseract language pack (`tesseract --list-langs`) and
pass `--ocr-lang eng+ara`.

---

## 6. Tests

Offline checks cover the parsers, the slide planner, the deck builder, and the
MCQ parser/writers. They do **not** need Tesseract or an API key:

```bash
python -m unittest discover -s tests -v
```

Expected: `Ran 9 tests ... OK`.

---

## 7. Notes

* The `.pptx` contains real text boxes and shapes: edit colours, fonts, and
  wording directly in PowerPoint.
* Only the MCQ step sends content to Google. Slide generation is entirely
  local.
* Page numbers, kickers, and speaker notes are added automatically; the notes
  hold the source text for the slide.
* Very large documents are capped by `--max-slides`; content is sampled
  evenly rather than truncated from the front.
