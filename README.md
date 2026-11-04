# PDF Converter (for MS Office)

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

## Updating

Nothing updates automatically, and that is on purpose. Library versions are
pinned in `requirements.txt` and OCR language files are pinned by SHA-256 in
`ocr-languages.json`, so conversions behave the same on every machine and a
changed upstream package can't slip in silently. Each kind of update below is a
deliberate step.

### Get the latest version of this app

```bat
git pull
.venv\Scripts\python -m pip install -r requirements.txt
```

Run the second command whenever `requirements.txt` changed in the pull. If you
don't have a `.venv` yet, run `setup.bat` instead.

### Update the Python libraries

1. List what has newer releases:

   ```bat
   .venv\Scripts\python -m pip list --outdated
   ```

2. In `requirements.txt`, raise the versions you want, one or two at a time.
   Keep exact `==` pins.
3. Install and run the self-check:

   ```bat
   .venv\Scripts\python -m pip install -r requirements.txt
   .venv\Scripts\python test_converter.py
   ```

4. Convert one or two real reports you already have outputs for, and compare
   the old and new results: table count, rows and columns, and numbers.
5. If anything regressed, put the old pin back and run
   `pip install -r requirements.txt` again.

Upgrade `pdfplumber` and `pymupdf` with extra care. Table detection depends on
their internals, so check the real-report comparison closely after changing them.

### OCR language files

Language files are downloaded once into `tessdata\` and reused.

- **Re-download a file** (if it's corrupted or was deleted): delete
  `tessdata\<lang>.traineddata`. On the next OCR run it is restored, either
  copied from a local Tesseract install (`TESSDATA_PREFIX` or
  `C:\Program Files\Tesseract-OCR\tessdata`) or, if neither has it, downloaded
  and checked against its pinned SHA-256. Copied files are not hash-checked.
- **Move to a different tessdata release**: update the `url` and each
  language's `sha256` and `size` in `ocr-languages.json`, delete `tessdata\`,
  and rerun the tests. A file that doesn't match its pinned hash is rejected.

### Python itself

After installing a new Python version, rebuild the environment:

```bat
rmdir /s /q .venv
setup.bat
.venv\Scripts\python test_converter.py
```

## Known limits

- Scanned pages are not deskewed, so scans tilted more than about 1–2° can lose table lines.
- Cell colours are not copied to Excel.
- Borderless table detection is heuristic.
