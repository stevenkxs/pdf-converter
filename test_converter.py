"""Self-check: build a PDF with 3 table styles, convert, assert structure.

Run:  .venv\\Scripts\\python test_converter.py
"""
import tempfile
from pathlib import Path

import pymupdf as fitz  # installed with pdf2docx

import converter

ROWS = [["Region", "Q1", "Q2", "Growth"],
        ["North", "1,200", "1,350", "12.5%"],
        ["South", "980", "(40)", "-4.1%"],
        ["West", "2,045.50", "2,300", "8%"]]


def grid(page, x0, y0, colw, rowh, rows, v_lines=True, h_lines=True, header_merge=False):
    n_r, n_c = len(rows), len(rows[0])
    x1, y1 = x0 + colw * n_c, y0 + rowh * n_r
    if h_lines:
        for r in range(n_r + 1):
            page.draw_line((x0, y0 + r * rowh), (x1, y0 + r * rowh))
    if v_lines:
        for c in range(n_c + 1):
            # merged header: skip interior vertical lines in the top row
            top = y0 + rowh if header_merge and 0 < c < n_c else y0
            page.draw_line((x0 + c * colw, top), (x0 + c * colw, y1))
    for r, row in enumerate(rows):
        for c, txt in enumerate(row):
            page.insert_text((x0 + c * colw + 4, y0 + r * rowh + rowh - 6), txt, fontsize=10)


def make_pdf(path):
    doc = fitz.open()
    p = doc.new_page()  # 1: fully ruled, merged title row
    p.insert_text((50, 60), "Quarterly report", fontsize=16)
    grid(p, 50, 100, 110, 22, [["Sales by region", "", "", ""]] + ROWS, header_merge=True)
    p.insert_text((50, 260), "Figures are unaudited and subject to change.", fontsize=10)

    p = doc.new_page()  # 2: borderless, whitespace-aligned
    p.insert_text((50, 60), "Borderless table", fontsize=14)
    grid(p, 50, 90, 110, 20, ROWS, v_lines=False, h_lines=False)

    p = doc.new_page()  # 3: horizontal rules only (financial style)
    grid(p, 50, 90, 110, 20, ROWS, v_lines=False)

    p = doc.new_page()  # 4+5: one ruled table split across pages, header repeated
    grid(p, 50, 720, 110, 20, ROWS[:3])
    p = doc.new_page()
    grid(p, 50, 40, 110, 20, [ROWS[0], ROWS[3], ["East", "500", "600", "20%"]])

    p = doc.new_page()  # 6: WBR-style narrow (2.4pt) blank spacer column between blocks
    xs = [50, 150, 230, 232.4, 312]  # label | Wk | spacer | Month
    rows = [["Metric", "Wk 44", "Oct-26"], ["YoY (%)", "+7.3%", "-13.4%"], ["GMS", "$22,809", "(1.21)"]]
    for r in range(len(rows) + 1):
        p.draw_line((xs[0], 90 + r * 20), (xs[-1], 90 + r * 20), width=0.5)
    for x in xs:
        p.draw_line((x, 90), (x, 90 + 20 * len(rows)), width=0.5)
    for r, row in enumerate(rows):
        for x, txt in zip((xs[0], xs[1], xs[3]), row):
            p.insert_text((x + 4, 90 + r * 20 + 14), txt, fontsize=10)
    doc.save(path)


def main():
    out = Path(tempfile.mkdtemp())
    pdf = out / "sample.pdf"
    make_pdf(pdf)

    import pdfplumber
    with pdfplumber.open(pdf) as d:
        t1 = converter.find_tables(d.pages[0], 1)
        t2 = converter.find_tables(d.pages[1], 2)
        t3 = converter.find_tables(d.pages[2], 3)
    assert len(t1) == 1 and (t1[0].n_rows, t1[0].n_cols) == (5, 4), t1
    assert any(c.colspan == 4 and c.text == "Sales by region" for c in t1[0].cells), "merged header lost"
    assert t1[0].matrix()[2] == ROWS[1], t1[0].matrix()
    for t in (t2, t3):
        assert len(t) == 1 and t[0].n_cols == 4 and t[0].n_rows >= 4, t
        assert ROWS[2] in t[0].matrix(), t[0].matrix()

    assert converter.to_number("1,200") == 1200
    assert converter.to_number("(40)") == -40
    assert abs(converter.to_number("12.5%") - 0.125) < 1e-9
    assert converter.to_number("00123") == "00123"
    assert converter.to_number("+7.3%") == 0.073 and converter.to_number("-7.2%") == -0.072
    assert converter.to_number("+87") == 87 and converter.to_number("$22,809") == 22809
    assert converter.number_format("+7.3%") == "+0.0%;-0.0%;0.0%"
    assert converter.number_format("$22,809") == "$#,##0"
    assert converter.number_format("(1.21)") == "0.00;(0.00)"

    res = converter.convert(pdf, out_dir=out)
    assert set(res) == {"docx", "xlsx", "pptx"}, res

    from openpyxl import load_workbook
    wb = load_workbook(res["xlsx"])
    assert wb.sheetnames == ["P1_T1", "P2_T1", "P3_T1", "P4_T1", "P6_T1"], wb.sheetnames
    sp = [list(r) for r in wb["P6_T1"].iter_rows(values_only=True)]
    assert sp[0] == ["Metric", "Wk 44", None, "Oct-26"], sp  # spacer kept as a blank column
    assert sp[1] == ["YoY (%)", 0.073, None, -0.134] and sp[2] == ["GMS", 22809, None, -1.21], sp
    assert wb["P6_T1"]["B2"].number_format == "+0.0%;-0.0%;0.0%"
    split = wb["P4_T1"]  # pages 4+5 joined, repeated header dropped
    assert [r[0] for r in split.iter_rows(values_only=True)] == ["Region", "North", "South", "West", "East"]
    ws = wb[wb.sheetnames[0]]
    assert ws["A1"].value == "Sales by region" and "A1:D1" in [str(m) for m in ws.merged_cells.ranges]
    assert ws["B3"].value == 1200 and ws["C4"].value == -40

    from docx import Document
    assert len(Document(res["docx"]).tables) >= 1, "Word output has no tables"

    from pptx import Presentation
    prs = Presentation(res["pptx"])
    assert sum(s.has_table for sl in prs.slides for s in sl.shapes) >= 3

    # Scanned PDF: rasterise the digital pages into image-only pages, then OCR.
    scan = out / "scanned.pdf"
    src, dst = fitz.open(pdf), fitz.open()
    for p in src:
        sp = dst.new_page(width=p.rect.width, height=p.rect.height)
        sp.insert_image(sp.rect, pixmap=p.get_pixmap(dpi=200))
    dst.save(scan)
    assert fitz.open(scan)[0].get_text().strip() == "", "scan fixture should have no text"

    ocr_pdf, n = converter.prepare(scan)
    assert n == 6, n
    with pdfplumber.open(ocr_pdf) as d:
        s1 = converter.find_tables(d.pages[0], 1)
        s3 = converter.find_tables(d.pages[2], 3)
    assert len(s1) == 1 and (s1[0].n_rows, s1[0].n_cols) == (5, 4), s1
    assert any(c.colspan == 4 and c.text == "Sales by region" for c in s1[0].cells), s1[0].matrix()
    assert s1[0].matrix()[1:] == ROWS, s1[0].matrix()
    assert len(s3) == 1 and ROWS[2] in s3[0].matrix(), s3[0].matrix() if s3 else s3

    res = converter.convert(scan, out_dir=out / "scan_out")
    assert set(res) == {"docx", "xlsx", "pptx"}, res
    ws = load_workbook(res["xlsx"])["P1_T1"]
    assert ws["A1"].value == "Sales by region" and ws["B3"].value == 1200 and ws["C4"].value == -40

    print("All checks passed. Outputs in", out)


if __name__ == "__main__":
    main()
