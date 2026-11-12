# -*- coding: utf-8 -*-
"""Shared python-docx helpers: consistent styles, tables, figures, footer with page numbers."""
from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

NAVY = RGBColor(0x1F, 0x3A, 0x4D)
RED = RGBColor(0x9B, 0x2C, 0x2C)
GREY = RGBColor(0x55, 0x55, 0x55)
FONT = "Arial"
BAND_FILL = {"Low": "C6E0B4", "Moderate": "FFE699", "High": "F4B183", "Extreme": "E06666"}


def new_doc(footer_text, margins=2.0):
    d = Document()
    cp = d.core_properties; cp.author = 'Moshood'; cp.last_modified_by = 'Moshood'; cp.comments = ''; cp.title = 'CNG Safety Risk Intelligence'
    s = d.sections[0]
    s.page_width, s.page_height = Cm(21.0), Cm(29.7)
    s.left_margin = s.right_margin = Cm(margins); s.top_margin = Cm(2.0); s.bottom_margin = Cm(2.0)
    st = d.styles["Normal"]; st.font.name = FONT; st.font.size = Pt(10)
    st.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    st.paragraph_format.space_after = Pt(6); st.paragraph_format.line_spacing = 1.12
    for name, size, before, after in (("Heading 1", 16, 18, 6), ("Heading 2", 12.5, 12, 4), ("Heading 3", 10.5, 8, 3)):
        h = d.styles[name]; h.font.name = FONT; h.font.size = Pt(size); h.font.bold = True; h.font.color.rgb = NAVY
        h.element.rPr.rFonts.set(qn("w:eastAsia"), FONT); h.element.rPr.rFonts.set(qn("w:ascii"), FONT); h.element.rPr.rFonts.set(qn("w:hAnsi"), FONT)
        h.paragraph_format.space_before = Pt(before); h.paragraph_format.space_after = Pt(after); h.paragraph_format.keep_with_next = True
    for name in ("List Bullet", "List Number"):
        d.styles[name].font.name = FONT; d.styles[name].font.size = Pt(10)
        d.styles[name].paragraph_format.space_after = Pt(3)
    # footer
    p = s.footer.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(footer_text + "   |   Page "); r.font.size = Pt(7.5); r.font.color.rgb = RED
    _field(p, "PAGE", 7.5)
    return d


def _field(p, code, size):
    r = p.add_run(); r.font.size = Pt(size); r.font.color.rgb = RED
    for t, txt in (("begin", None), (None, code), ("end", None)):
        if t:
            e = OxmlElement("w:fldChar"); e.set(qn("w:fldCharType"), t); r._r.append(e)
        else:
            e = OxmlElement("w:instrText"); e.set(qn("xml:space"), "preserve"); e.text = txt; r._r.append(e)


def para(d, text, bold=False, italic=False, size=None, color=None, align=None, after=None, keep=False):
    p = d.add_paragraph()
    _runs(p, text, bold, italic, size, color)
    if align == "center": p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if after is not None: p.paragraph_format.space_after = Pt(after)
    if keep: p.paragraph_format.keep_with_next = True
    return p


def _runs(p, text, bold=False, italic=False, size=None, color=None):
    """Supports **bold** segments inside text."""
    parts = text.split("**")
    for i, seg in enumerate(parts):
        if not seg: continue
        r = p.add_run(seg); r.bold = bold or (i % 2 == 1); r.italic = italic
        if size: r.font.size = Pt(size)
        if color: r.font.color.rgb = color


def bullet(d, text, style="List Bullet"):
    p = d.add_paragraph(style=style); _runs(p, text); return p


def h(d, text, lvl=1):
    return d.add_heading(text, level=lvl)


def shade(cell, hexfill):
    tcPr = cell._tc.get_or_add_tcPr(); sh = OxmlElement("w:shd")
    sh.set(qn("w:val"), "clear"); sh.set(qn("w:color"), "auto"); sh.set(qn("w:fill"), hexfill); tcPr.append(sh)


def _cell_margins(table, top=40, bottom=40, left=70, right=70):
    tblPr = table._tbl.tblPr; m = OxmlElement("w:tblCellMar")
    for k, v in (("top", top), ("left", left), ("bottom", bottom), ("right", right)):
        e = OxmlElement(f"w:{k}"); e.set(qn("w:w"), str(v)); e.set(qn("w:type"), "dxa"); m.append(e)
    tblPr.append(m)


def fix_widths(t, widths):
    tblPr = t._tbl.tblPr
    lay = OxmlElement("w:tblLayout"); lay.set(qn("w:type"), "fixed"); tblPr.append(lay)
    grid = t._tbl.tblGrid
    for i, gc in enumerate(grid.findall(qn("w:gridCol"))):
        gc.set(qn("w:w"), str(int(widths[i] * 567)))
    for row in t.rows:
        for i, c in enumerate(row.cells):
            c.width = Cm(widths[i])


def table(d, header, rows, widths, size=8.5, band_col=None, header_fill="1F3A4D", align_center=(), zebra=True):
    t = d.add_table(rows=1, cols=len(header)); t.style = "Table Grid"; t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False; _cell_margins(t)
    for i, txt in enumerate(header):
        c = t.rows[0].cells[i]; c.width = Cm(widths[i]); shade(c, header_fill)
        p = c.paragraphs[0]; p.paragraph_format.space_after = Pt(0)
        r = p.add_run(str(txt)); r.bold = True; r.font.size = Pt(size); r.font.color.rgb = RGBColor(255, 255, 255)
    trPr = t.rows[0]._tr.get_or_add_trPr(); e = OxmlElement("w:tblHeader"); e.set(qn("w:val"), "true"); trPr.append(e)
    for ri, row in enumerate(rows):
        cells = t.add_row().cells
        cant = OxmlElement("w:cantSplit"); cant.set(qn("w:val"), "true"); t.rows[-1]._tr.get_or_add_trPr().append(cant)
        for i, v in enumerate(row):
            c = cells[i]; c.width = Cm(widths[i]); p = c.paragraphs[0]; p.paragraph_format.space_after = Pt(0); p.paragraph_format.line_spacing = 1.0
            _runs(p, str(v), size=size)
            if i in align_center: p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            if band_col is not None and i in (band_col if isinstance(band_col, (list, tuple)) else [band_col]) and str(v) in BAND_FILL:
                shade(c, BAND_FILL[str(v)])
            elif zebra and ri % 2 == 1:
                shade(c, "F3F6F8")
    fix_widths(t, widths)
    d.add_paragraph().paragraph_format.space_after = Pt(2)
    return t


def kv_table(d, rows, widths=(4.0, 13.0), size=9):
    t = d.add_table(rows=0, cols=2); t.style = "Table Grid"; t.autofit = False; t.alignment = WD_TABLE_ALIGNMENT.CENTER; _cell_margins(t)
    for k, v in rows:
        cells = t.add_row().cells
        cant = OxmlElement("w:cantSplit"); cant.set(qn("w:val"), "true"); t.rows[-1]._tr.get_or_add_trPr().append(cant)
        cells[0].width = Cm(widths[0]); cells[1].width = Cm(widths[1]); shade(cells[0], "E8EDF1")
        p = cells[0].paragraphs[0]; p.paragraph_format.space_after = Pt(0); r = p.add_run(k); r.bold = True; r.font.size = Pt(size)
        p = cells[1].paragraphs[0]; p.paragraph_format.space_after = Pt(0); _runs(p, v, size=size)
    fix_widths(t, widths)
    d.add_paragraph().paragraph_format.space_after = Pt(2)
    return t


def figure(d, path, width_cm, caption):
    p = d.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.keep_with_next = True; p.paragraph_format.space_after = Pt(2)
    p.add_run().add_picture(path, width=Cm(width_cm))
    c = d.add_paragraph(); c.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = c.add_run(caption); r.italic = True; r.font.size = Pt(8.5); r.font.color.rgb = GREY


def page_break(d):
    d.add_paragraph().add_run().add_break(WD_BREAK.PAGE)


def callout(d, text, fill="FDF2E9", size=9):
    t = d.add_table(rows=1, cols=1); t.style = "Table Grid"; t.autofit = False; t.alignment = WD_TABLE_ALIGNMENT.CENTER; _cell_margins(t, 80, 80, 120, 120)
    c = t.rows[0].cells[0]; c.width = Cm(17.0); shade(c, fill)
    p = c.paragraphs[0]; p.paragraph_format.space_after = Pt(0); _runs(p, text, size=size)
    fix_widths(t, (17.0,))
    d.add_paragraph().paragraph_format.space_after = Pt(2)
