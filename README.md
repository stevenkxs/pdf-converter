# PDF Converter

Convert PDFs to Word, Excel and PowerPoint, with tables rebuilt as real tables.
Scanned PDFs are handled by built-in OCR.

## Features

- **Table detection** for fully ruled, horizontal-rules-only and borderless tables.
  Merged cells are preserved, and narrow blank spacer columns are kept.
- **Excel**: one sheet per table, with numbers stored as numbers
  (`$22,809`, `(1.21)`, `+7.3%`, `+87`) and displayed the way the PDF printed them.
  Tables split across pages are joined when the header row repeats.
- **Word**: layout and tables via `pdf2docx`.
- **PowerPoint**: one slide per page, with native tables, positioned text and images.
- **OCR** (Tesseract via PyMuPDF; no separate install). Scanned pages are detected
  automatically. Table lines are erased before recognition and redrawn afterwards,
  so scanned tables come out as tables. 69 languages are available; each language
  file downloads on first use and is verified against a pinned SHA-256.

## Setup (Windows)

Requires Python 3.10+.

```bat
setup.bat
```

## Usage

Desktop app: double-click `PDF Converter.bat`.

Command line:

```bat
.venv\Scripts\python converter.py report.pdf --to xlsx docx pptx --out C:\output
.venv\Scripts\python converter.py scan.pdf --ocr auto --lang eng+chi_sim
```

`--ocr` accepts `auto` (scanned pages only, the default), `force` (every page) or `off`.

## Tests

```bat
.venv\Scripts\python test_converter.py
```

The test builds a sample PDF with several table styles, plus a fully scanned copy
of it. It checks table structure, cell values, merges, number conversion,
cross-page joining, the spacer column, and OCR.

## Known limits

- Scanned pages are not deskewed, so scans tilted more than about 1–2° can lose table lines.
- Cell colours are not copied to Excel.
- Borderless table detection is heuristic.
