"""PDF -> Word / Excel / PowerPoint conversion with real table reconstruction.

Word uses pdf2docx (layout + table rebuild). Excel and PowerPoint share one
pdfplumber-based table detector that handles ruled, partially ruled and
borderless tables, and preserves merged cells.

CLI:  python converter.py report.pdf --to xlsx docx pptx --out C:\\output
"""
from __future__ import annotations

import argparse
import os
import re

# OpenCV's numpy pulls in OpenBLAS, which can fail to allocate per-core buffers
# on Windows ("Memory allocation still failed"). One thread is plenty here.
os.environ.setdefault("OPENBLAS_NUM_THREADS", "1")
import statistics
from dataclasses import dataclass
from pathlib import Path

import pdfplumber

FORMATS = ("docx", "xlsx", "pptx")


# --------------------------------------------------------------------------- #
# Table detection (shared by Excel and PowerPoint)
# --------------------------------------------------------------------------- #
@dataclass
class Cell:
    row: int
    col: int
    rowspan: int
    colspan: int
    text: str


@dataclass
class Table:
    page: int  # 1-based
    bbox: tuple[float, float, float, float]  # x0, top, x1, bottom (PDF points)
    n_rows: int
    n_cols: int
    cells: list[Cell]
    col_edges: list[float]
    row_edges: list[float]
    font_size: float
    is_last: bool = False  # last table on its page (for cross-page joining)

    def matrix(self) -> list[list[str]]:
        m = [[""] * self.n_cols for _ in range(self.n_rows)]
        for c in self.cells:
            m[c.row][c.col] = c.text
        return m


# Ruled tables first; then tables with only horizontal rules (typical in
# financial reports); then fully borderless tables aligned by whitespace.
# snap_x_tolerance=1: pdfplumber's default 3pt snaps away narrow spacer
# columns (e.g. a 2.4pt gap between weekly and monthly blocks), which must
# survive as blank columns.
_STRATEGIES = [
    ({"vertical_strategy": "lines", "horizontal_strategy": "lines", "snap_x_tolerance": 1}, False),
    ({"vertical_strategy": "text", "horizontal_strategy": "lines"}, True),
    ({"vertical_strategy": "text", "horizontal_strategy": "text",
      "min_words_vertical": 3, "min_words_horizontal": 2}, True),
]


def _cluster(values: list[float], tol: float = 1.0) -> list[float]:
    """Merge near-identical coordinates (pdf rounding noise) into one edge.

    Compared against each group's first value, not the previous value, so a
    run like 494.9, 495.1, 495.3, 497.2 can't chain together and swallow a
    narrow spacer column.
    """
    groups: list[list[float]] = []
    for v in sorted(values):
        if groups and v - groups[-1][0] <= tol:
            groups[-1].append(v)
        else:
            groups.append([v])
    return [statistics.mean(g) for g in groups]


def _index(edges: list[float], v: float) -> int:
    return min(range(len(edges)), key=lambda i: abs(edges[i] - v))


# Fonts often map the hyphen glyph to a soft hyphen or a dash variant; numbers need '-'.
_DASHES = str.maketrans({c: "-" for c in "\u00ad\u2010\u2011\u2012\u2013\u2014\u2212"})


def _clean(text: str | None) -> str:
    return re.sub(r"\s+", " ", (text or "").translate(_DASHES)).strip()


def _looks_like_table(t: Table) -> bool:
    # ponytail: heuristic filter for whitespace-detected tables. Prose paragraphs
    # yield sparse grids with long cells; real tables are dense and short-celled.
    # Upgrade path: a layout model (e.g. table-transformer) if this misfires.
    if t.n_rows < 3 or t.n_cols < 2:
        return False
    texts = [c.text for c in t.cells]
    filled = [s for s in texts if s]
    if len(filled) < 0.5 * t.n_rows * t.n_cols:
        return False
    return statistics.mean(len(s) for s in filled) <= 40


def _build_table(page, ptable, page_no: int) -> Table | None:
    raw_rows = ptable.rows
    values = ptable.extract()
    boxes = [c for r in raw_rows for c in r.cells if c]
    if not boxes:
        return None
    col_edges = _cluster([b[0] for b in boxes] + [b[2] for b in boxes])
    row_edges = _cluster([b[1] for b in boxes] + [b[3] for b in boxes])
    n_cols, n_rows = len(col_edges) - 1, len(row_edges) - 1
    if n_cols < 1 or n_rows < 1:
        return None

    cells: list[Cell] = []
    for r, vals in zip(raw_rows, values):
        for box, val in zip(r.cells, vals):
            if box is None:  # covered by a merged neighbour
                continue
            c0, c1 = _index(col_edges, box[0]), _index(col_edges, box[2])
            r0, r1 = _index(row_edges, box[1]), _index(row_edges, box[3])
            if c1 <= c0 or r1 <= r0:
                continue
            cells.append(Cell(r0, c0, r1 - r0, c1 - c0, _clean(val)))

    sizes = [ch["size"] for ch in page.within_bbox(ptable.bbox).chars] or [10]
    return Table(page_no, tuple(ptable.bbox), n_rows, n_cols, cells,
                 col_edges, row_edges, statistics.median(sizes))


def _overlaps(a, b) -> bool:
    return not (a[2] <= b[0] or b[2] <= a[0] or a[3] <= b[1] or b[3] <= a[1])


def find_tables(page, page_no: int) -> list[Table]:
    found: list[Table] = []
    for settings, needs_filter in _STRATEGIES:
        if settings["horizontal_strategy"] == "lines" and not (page.lines or page.rects):
            continue
        for pt in page.find_tables(settings):
            if any(_overlaps(pt.bbox, t.bbox) for t in found):
                continue
            t = _build_table(page, pt, page_no)
            if not t or t.n_rows < 2 or t.n_cols < 2:
                continue
            if needs_filter and not _looks_like_table(t):
                continue
            found.append(t)
        if found:  # stop at the most reliable strategy that found something
            break
    return sorted(found, key=lambda t: (t.bbox[1], t.bbox[0]))


# --------------------------------------------------------------------------- #
# Excel
# --------------------------------------------------------------------------- #
_NUM = re.compile(r"^\(?[+-]?[$€£¥]?\s?\d{1,3}(?:,\d{3})*(?:\.\d+)?\)?%?$|^\(?[+-]?[$€£¥]?\s?\d+(?:\.\d+)?\)?%?$")


def to_number(s: str):
    """'1,234.5' -> 1234.5, '(12)' -> -12, '15%' -> 0.15. Otherwise the text."""
    if not s or not _NUM.match(s):
        return s
    digits = re.sub(r"[^\d.]", "", s)
    if len(digits) > 1 and digits[0] == "0" and "." not in digits:
        return s  # leading-zero codes (zip, IDs) stay text
    n = float(digits)
    if s.startswith("(") and s.endswith(")") or "-" in s:
        n = -n
    if s.endswith("%"):
        n = round(n / 100, 10)  # avoid 0.07200000000000001
    return int(n) if n.is_integer() and "." not in digits and not s.endswith("%") else n


def number_format(s: str) -> str:
    """Excel format that displays a number the way the PDF printed it.

    '+7.3%' -> '+0.0%;-0.0%;0.0%', '$22,809' -> '$#,##0', '(1.21)' -> '0.00;(0.00)'.
    """
    m = re.search(r"\.(\d+)", s)
    body = ("#,##0" if "," in s else "0") + ("." + "0" * len(m.group(1)) if m else "")
    if s.startswith("$") or s.lstrip("(+-").startswith("$"):
        body = "$" + body
    if s.endswith("%"):
        body += "%"
    if s.startswith("("):
        return f"{body};({body})"
    if s.startswith("+"):
        return f"+{body};-{body};{body}"
    return body


def to_xlsx(pdf_path: Path, out: Path) -> int:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    wb.remove(wb.active)
    count = 0
    prev: tuple[Table, object, int] | None = None  # table, sheet, next free row

    with pdfplumber.open(pdf_path) as pdf:
        for page_no, page in enumerate(pdf.pages, 1):
            tables = find_tables(page, page_no)
            for i, t in enumerate(tables):
                matrix = t.matrix()
                # Continue a table split across pages: same columns, the previous
                # one runs to the page bottom, this one starts at the top, and
                # its header row repeats. Without the header check, one-table-
                # per-page reports (e.g. a country per page) collapse into one sheet.
                if (prev and i == 0 and prev[0].page == page_no - 1
                        and prev[0].n_cols == t.n_cols and prev[0].is_last
                        and prev[0].bbox[3] > 0.8 * page.height
                        and t.bbox[1] < 0.2 * page.height
                        and any(matrix[0]) and matrix[0] == prev[1].first_row):
                    _, ws, row0 = prev
                    skip = 1  # drop the repeated header
                else:
                    count += 1
                    ws = wb.create_sheet(f"P{page_no}_T{i + 1}"[:31])
                    ws.first_row = matrix[0]  # header to recognise a continuation
                    row0, skip = 1, 0

                spans = []
                for c in t.cells:
                    if c.row < skip:
                        continue
                    r = row0 + c.row - skip
                    cell = ws.cell(r, c.col + 1)
                    if type(cell).__name__ == "MergedCell" or cell.value not in (None, ""):
                        continue  # already covered by a span or filled
                    cell.value = to_number(c.text)
                    if isinstance(cell.value, (int, float)):
                        cell.number_format = number_format(c.text)
                    cell.alignment = Alignment(wrap_text=True, vertical="top")
                    if c.rowspan > 1 or c.colspan > 1:
                        spans.append((r, c.col + 1, r + c.rowspan - 1, c.col + c.colspan))
                # Merge after all values are in, and only where nothing would be hidden
                # (detected spans can overlap, or cover text from a neighbouring row).
                for r0, c0, r1, c1 in spans:
                    cells = [ws.cell(r, c) for r in range(r0, r1 + 1) for c in range(c0, c1 + 1)][1:]
                    if all(type(x).__name__ != "MergedCell" and x.value in (None, "") for x in cells):
                        ws.merge_cells(start_row=r0, start_column=c0, end_row=r1, end_column=c1)
                t.is_last = i == len(tables) - 1
                prev = (t, ws, row0 + t.n_rows - skip)

        if count == 0:  # no tables: still give the user the text
            ws = wb.create_sheet("Text")
            ws.append(["Page", "Line"])
            for page_no, page in enumerate(pdf.pages, 1):
                for line in (page.extract_text() or "").splitlines():
                    ws.append([page_no, line])

    for ws in wb.worksheets:
        for col in ws.columns:
            width = max((len(str(c.value)) for c in col if c.value is not None), default=8)
            ws.column_dimensions[get_column_letter(col[0].column)].width = min(max(width + 2, 8), 60)

    wb.save(out)
    return count


# --------------------------------------------------------------------------- #
# PowerPoint
# --------------------------------------------------------------------------- #
def to_pptx(pdf_path: Path, out: Path) -> int:
    import io

    from pptx import Presentation
    from pptx.util import Pt

    emu = lambda pts: int(pts * 12700)  # noqa: E731  (PDF points -> EMU)
    prs = Presentation()
    count = 0
    with pdfplumber.open(pdf_path) as pdf:
        first = pdf.pages[0]
        prs.slide_width, prs.slide_height = emu(first.width), emu(first.height)
        blank = prs.slide_layouts[6]

        for page_no, page in enumerate(pdf.pages, 1):
            slide = prs.slides.add_slide(blank)
            tables = find_tables(page, page_no)
            count += len(tables)
            in_table = lambda b: any(_overlaps(b, t.bbox) for t in tables)  # noqa: E731

            for img in page.images:  # pictures, cropped from the rendered page
                box = (max(img["x0"], 0), max(img["top"], 0),
                       min(img["x1"], page.width), min(img["bottom"], page.height))
                if box[2] - box[0] < 4 or box[3] - box[1] < 4:
                    continue
                try:
                    buf = io.BytesIO()
                    page.crop(box).to_image(resolution=150).original.save(buf, "PNG")
                    buf.seek(0)
                    slide.shapes.add_picture(buf, emu(box[0]), emu(box[1]),
                                             emu(box[2] - box[0]), emu(box[3] - box[1]))
                except Exception:
                    pass  # an unrenderable image should not lose the slide

            for line in page.extract_text_lines():
                box = (line["x0"], line["top"], line["x1"], line["bottom"])
                if in_table(box):
                    continue
                size = statistics.median(ch["size"] for ch in line["chars"])
                tb = slide.shapes.add_textbox(emu(box[0]), emu(box[1]),
                                              emu(box[2] - box[0] + 4), emu(box[3] - box[1]))
                tf = tb.text_frame
                tf.word_wrap = False
                tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
                run = tf.paragraphs[0].add_run()
                run.text = line["text"]
                run.font.size = Pt(size)
                run.font.bold = "Bold" in (line["chars"][0].get("fontname") or "")

            for t in tables:
                x0, top, x1, bottom = t.bbox
                shape = slide.shapes.add_table(t.n_rows, t.n_cols, emu(x0), emu(top),
                                               emu(x1 - x0), emu(bottom - top))
                tbl = shape.table
                tbl.first_row = False  # don't impose a theme header style
                for i in range(t.n_cols):
                    tbl.columns[i].width = emu(t.col_edges[i + 1] - t.col_edges[i])
                for i in range(t.n_rows):
                    tbl.rows[i].height = emu(t.row_edges[i + 1] - t.row_edges[i])
                for c in t.cells:
                    cell = tbl.cell(c.row, c.col)
                    if c.rowspan > 1 or c.colspan > 1:
                        cell.merge(tbl.cell(c.row + c.rowspan - 1, c.col + c.colspan - 1))
                    cell.margin_left = cell.margin_right = Pt(2)
                    cell.margin_top = cell.margin_bottom = Pt(1)
                    cell.text = c.text
                    for p in cell.text_frame.paragraphs:
                        for r in p.runs:
                            r.font.size = Pt(t.font_size)
    prs.save(out)
    return count


# --------------------------------------------------------------------------- #
# Word
# --------------------------------------------------------------------------- #
def to_docx(pdf_path: Path, out: Path) -> int:
    from pdf2docx import Converter

    cv = Converter(str(pdf_path))
    try:
        cv.convert(str(out))
    finally:
        cv.close()
    from docx import Document
    return len(Document(str(out)).tables)


# --------------------------------------------------------------------------- #
# OCR for scanned pages
#
# Same engine and data as Filee: Tesseract with tessdata_fast 4.1.0 language
# files, fetched from the pinned URLs/SHA-256s in ocr-languages.json (taken
# from Filee's embedded manifest). Tesseract runs inside PyMuPDF, so no
# tesseract.exe install is needed. Each scanned page is rebuilt as a
# "native" page: OCR words as real text at their positions, plus the table
# ruling lines found in the image as vector lines. The table detector and
# all three converters then treat it like any digital PDF.
# --------------------------------------------------------------------------- #
HERE = Path(__file__).parent
TESSDATA = HERE / "tessdata"
OCR_DPI = 300
_CJK_FONTS = {"chi_sim": "china-s", "chi_tra": "china-t", "jpn": "japan", "kor": "korea"}


def ocr_languages() -> list[str]:
    import json
    return list(json.loads((HERE / "ocr-languages.json").read_text("utf-8"))["languages"])


def ensure_tessdata(lang: str, log=print) -> Path:
    """Make every language in 'eng+chi_sim' available in ./tessdata (copy or download)."""
    import hashlib
    import json
    import os
    import shutil
    import urllib.request

    manifest = json.loads((HERE / "ocr-languages.json").read_text("utf-8"))
    TESSDATA.mkdir(exist_ok=True)
    for code in lang.split("+"):
        dest = TESSDATA / f"{code}.traineddata"
        if dest.exists():
            continue
        local = [Path(p) / f"{code}.traineddata" for p in
                 (os.environ.get("TESSDATA_PREFIX"), r"C:\Program Files\Tesseract-OCR\tessdata") if p]
        found = next((p for p in local if p.exists()), None)
        if found:
            shutil.copy2(found, dest)
            continue
        info = manifest["languages"].get(code)
        if not info:
            raise ValueError(f"Unknown OCR language '{code}'. Known: {', '.join(manifest['languages'])}")
        log(f"Downloading OCR language '{code}' ({info['size'] // 1024} KB)...")
        data = urllib.request.urlopen(manifest["url"].format(code), timeout=120).read()
        if hashlib.sha256(data).hexdigest().upper() != info["sha256"].upper():
            raise ValueError(f"Checksum mismatch for {code}.traineddata; download rejected")
        dest.write_bytes(data)
    return TESSDATA


def _ruling_lines(pix):
    """Detect table rules in a page image.

    Returns (segments in PDF points, grayscale image with the rules erased).
    Erasing matters: Tesseract's layout analysis drops most text inside a
    gridded table when the lines are still there.
    """
    import cv2
    import numpy as np

    img = np.frombuffer(pix.samples, np.uint8).reshape(pix.h, pix.w, pix.n)
    gray = cv2.cvtColor(img, cv2.COLOR_RGB2GRAY if pix.n == 3 else cv2.COLOR_RGBA2GRAY)
    bw = cv2.adaptiveThreshold(~gray, 255, cv2.ADAPTIVE_THRESH_MEAN_C, cv2.THRESH_BINARY, 15, -2)
    s = 72 / OCR_DPI
    out, all_rules = [], np.zeros_like(bw)
    for horizontal in (True, False):
        k = (max(pix.w // 30, 20), 1) if horizontal else (1, max(pix.h // 50, 20))
        mask = cv2.morphologyEx(bw, cv2.MORPH_OPEN, cv2.getStructuringElement(cv2.MORPH_RECT, k))
        all_rules |= mask
        contours, _ = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        for c in contours:
            x, y, w, h = cv2.boundingRect(c)
            if horizontal:
                out.append((x * s, (y + h / 2) * s, (x + w) * s, (y + h / 2) * s))
            else:
                out.append(((x + w / 2) * s, y * s, (x + w / 2) * s, (y + h) * s))
    clean = gray.copy()
    clean[cv2.dilate(all_rules, np.ones((3, 3), np.uint8)) > 0] = 255
    return out, clean


def _ocr_words(gray, rect, lang: str, tessdata: str):
    """Tesseract words from a grayscale image laid over `rect`, in PDF points.

    The image goes on a scratch page the same size as the original, so word
    coordinates line up 1:1 with the source page.
    """
    import pymupdf

    h, w = gray.shape
    pix = pymupdf.Pixmap(pymupdf.csGRAY, w, h, gray.tobytes(), 0)
    doc = pymupdf.open()
    page = doc.new_page(width=rect.width, height=rect.height)
    page.insert_image(page.rect, pixmap=pix)
    tp = page.get_textpage_ocr(language=lang, dpi=OCR_DPI, full=True, tessdata=tessdata)
    return page.get_text("words", textpage=tp)


def _is_scanned(page) -> bool:
    return len(page.get_text().strip()) < 20 and bool(page.get_images())


def prepare(pdf_path: Path, ocr: str = "auto", lang: str = "eng", log=print) -> tuple[Path, int]:
    """Return (pdf to convert, number of OCR'd pages). ocr: auto | force | off.

    ponytail: no deskew; scans rotated more than ~1-2 degrees lose table rules.
    Upgrade path: estimate skew from the detected lines and rotate before OCR.
    """
    import tempfile

    import pymupdf

    if ocr == "off":
        return pdf_path, 0
    src = pymupdf.open(pdf_path)
    targets = [i for i, p in enumerate(src) if ocr == "force" or _is_scanned(p)]
    if not targets:
        return pdf_path, 0

    tessdata = str(ensure_tessdata(lang, log))
    cjk = next((f for code, f in _CJK_FONTS.items() if code in lang.split("+")), None)
    ttf = Path(r"C:\Windows\Fonts\arial.ttf")  # Latin/Cyrillic/Greek coverage
    font = pymupdf.Font(cjk) if cjk else pymupdf.Font(fontfile=str(ttf)) if ttf.exists() else pymupdf.Font("helv")

    out = pymupdf.open()
    for i, page in enumerate(src):
        if i not in targets:
            out.insert_pdf(src, from_page=i, to_page=i)
            continue
        log(f"OCR page {i + 1}/{len(src)} ({lang})...")
        lines, clean = _ruling_lines(page.get_pixmap(dpi=OCR_DPI))
        words = _ocr_words(clean, page.rect, lang, tessdata)

        new = out.new_page(width=page.rect.width, height=page.rect.height)
        for x0, y0, x1, y1 in lines:
            new.draw_line((x0, y0), (x1, y1), width=0.5)
        tw = pymupdf.TextWriter(new.rect)
        by_line: dict[tuple, list] = {}
        for w in words:  # (x0, y0, x1, y1, text, block, line, word)
            by_line.setdefault((w[5], w[6]), []).append(w)
        for ws in by_line.values():  # one font size per line keeps lines even in Word
            size = statistics.median(w[3] - w[1] for w in ws) * 0.8
            for x0, y0, x1, y1, text, *_ in ws:
                fit = (x1 - x0) / max(font.text_length(text, 1), 1e-3)
                tw.append((x0, y1 - (y1 - y0) * 0.2), text, font=font, fontsize=max(min(size, fit), 4))
        tw.write_text(new)

    tmp = Path(tempfile.mkdtemp()) / pdf_path.name
    out.save(tmp)
    return tmp, len(targets)


# --------------------------------------------------------------------------- #
CONVERTERS = {"docx": to_docx, "xlsx": to_xlsx, "pptx": to_pptx}


def convert(pdf_path: str | Path, formats=FORMATS, out_dir: str | Path | None = None,
            log=print, ocr: str = "auto", lang: str = "eng"):
    """Convert one PDF to each requested format. Returns {fmt: output path}."""
    import shutil

    pdf_path = Path(pdf_path)
    out_dir = Path(out_dir) if out_dir else pdf_path.parent
    out_dir.mkdir(parents=True, exist_ok=True)
    try:
        src, n_ocr = prepare(pdf_path, ocr, lang, log)
    except Exception as e:
        log(f"FAIL {pdf_path.name} OCR: {type(e).__name__}: {e}")
        return {}
    if n_ocr:
        log(f"OCR  {n_ocr} scanned page{'s' if n_ocr != 1 else ''} recognised")
    results = {}
    try:
        for fmt in formats:
            out = out_dir / f"{pdf_path.stem}.{fmt}"
            try:
                n = CONVERTERS[fmt](src, out)
                log(f"OK   {out.name}  ({n} table{'s' if n != 1 else ''})")
                results[fmt] = out
            except Exception as e:  # report and continue with the next format
                log(f"FAIL {pdf_path.name} -> {fmt}: {type(e).__name__}: {e}")
    finally:
        if src != pdf_path:
            shutil.rmtree(src.parent, ignore_errors=True)
    return results


def main(argv=None):
    ap = argparse.ArgumentParser(description="Convert PDF to Word, Excel and PowerPoint.")
    ap.add_argument("pdfs", nargs="+", help="PDF files")
    ap.add_argument("--to", nargs="+", choices=FORMATS, default=list(FORMATS))
    ap.add_argument("--out", help="output folder (default: next to each PDF)")
    ap.add_argument("--ocr", choices=("auto", "force", "off"), default="auto",
                    help="auto: OCR only scanned pages (default); force: OCR every page")
    ap.add_argument("--lang", default="eng", help="Tesseract language(s), e.g. eng or eng+chi_sim")
    a = ap.parse_args(argv)
    for p in a.pdfs:
        convert(p, a.to, a.out, ocr=a.ocr, lang=a.lang)


if __name__ == "__main__":
    main()
