from __future__ import annotations

from pathlib import Path
from datetime import date
from io import BytesIO

from PIL import Image, ImageDraw, ImageFont
import arabic_reshaper
from bidi.algorithm import get_display
from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Inches, Pt, RGBColor


ROOT = Path(r"D:\project\innovation")
OUT = ROOT / "proposal_hamrah_entekhab_Expanded_25pages.docx"
ASSET_DIR = ROOT / "proposal_assets"
ASSET_DIR.mkdir(exist_ok=True)
LOGO_SCI = Path(r"C:\Users\xio\Downloads\logo.jpg")
LOGO_ICDS = Path(r"C:\Users\xio\Downloads\logo_icds.png")

# narrative_proposal preset + named Iranian institutional overrides:
# A4 page, Tahoma Persian typography, RTL layout, brand palette, 18 mm margins.
NAVY = "211A3D"
RED = "92251F"
BLUE = "4F92EA"
GOLD = "F4C842"
GREEN = "4D9B5F"
INK = "262433"
MUTED = "66636F"
PALE = "F5F4F8"
PALE_RED = "F8ECEB"
PALE_BLUE = "EDF4FD"
WHITE = "FFFFFF"
LINE = "D8D5DE"
FONT = "B Nazanin"
HEADING_FONT = "B Zar"


def set_cell_shading(cell, fill: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)
    shd.set(qn("w:fill"), fill)


def set_cell_margins(cell, top=90, start=120, bottom=90, end=120) -> None:
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()
    tc_mar = tc_pr.first_child_found_in("w:tcMar")
    if tc_mar is None:
        tc_mar = OxmlElement("w:tcMar")
        tc_pr.append(tc_mar)
    for m, v in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        node = tc_mar.find(qn(f"w:{m}"))
        if node is None:
            node = OxmlElement(f"w:{m}")
            tc_mar.append(node)
        node.set(qn("w:w"), str(v))
        node.set(qn("w:type"), "dxa")


def set_cell_borders(cell, color=LINE, size="6") -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    borders = tc_pr.find(qn("w:tcBorders"))
    if borders is None:
        borders = OxmlElement("w:tcBorders")
        tc_pr.append(borders)
    for edge in ("top", "start", "bottom", "end", "insideH", "insideV"):
        tag = borders.find(qn(f"w:{edge}"))
        if tag is None:
            tag = OxmlElement(f"w:{edge}")
            borders.append(tag)
        tag.set(qn("w:val"), "single")
        tag.set(qn("w:sz"), size)
        tag.set(qn("w:color"), color)


def set_table_geometry(table, widths_cm: list[float], indent_dxa: int = 120) -> None:
    table.autofit = False
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_pr = table._tbl.tblPr
    if len(widths_cm) >= 3:
        bidi_visual = tbl_pr.find(qn("w:bidiVisual"))
        if bidi_visual is None:
            bidi_visual = OxmlElement("w:bidiVisual")
            # Put table direction before width/alignment properties to keep
            # strict Microsoft Word XML ordering stable.
            insert_at = 1 if tbl_pr.find(qn("w:tblStyle")) is not None else 0
            tbl_pr.insert(insert_at, bidi_visual)
        bidi_visual.set(qn("w:val"), "1")
    tbl_w = tbl_pr.find(qn("w:tblW"))
    if tbl_w is None:
        tbl_w = OxmlElement("w:tblW")
        tbl_pr.append(tbl_w)
    total = sum(int(Cm(w).emu / 635) for w in widths_cm)
    tbl_w.set(qn("w:w"), str(total))
    tbl_w.set(qn("w:type"), "dxa")
    tbl_ind = tbl_pr.find(qn("w:tblInd"))
    if tbl_ind is None:
        tbl_ind = OxmlElement("w:tblInd")
        tbl_pr.append(tbl_ind)
    tbl_ind.set(qn("w:w"), str(indent_dxa))
    tbl_ind.set(qn("w:type"), "dxa")
    grid = table._tbl.tblGrid
    for child in list(grid):
        grid.remove(child)
    for w in widths_cm:
        dxa = int(Cm(w).emu / 635)
        gc = OxmlElement("w:gridCol")
        gc.set(qn("w:w"), str(dxa))
        grid.append(gc)
    for row in table.rows:
        for idx, cell in enumerate(row.cells):
            w = widths_cm[min(idx, len(widths_cm) - 1)]
            cell.width = Cm(w)
            tc_w = cell._tc.get_or_add_tcPr().find(qn("w:tcW"))
            if tc_w is None:
                tc_w = OxmlElement("w:tcW")
                cell._tc.get_or_add_tcPr().append(tc_w)
            tc_w.set(qn("w:w"), str(int(Cm(w).emu / 635)))
            tc_w.set(qn("w:type"), "dxa")
            set_cell_margins(cell)
            set_cell_borders(cell)
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER


def rtl_paragraph(paragraph, align=WD_ALIGN_PARAGRAPH.RIGHT) -> None:
    paragraph.alignment = align
    p_pr = paragraph._p.get_or_add_pPr()
    bidi = p_pr.find(qn("w:bidi"))
    if bidi is None:
        bidi = OxmlElement("w:bidi")
        p_pr.append(bidi)
    bidi.set(qn("w:val"), "1")


def set_run(run, size=11.5, bold=False, color=INK, italic=False, font_name=None) -> None:
    selected_font = font_name or FONT
    run.font.name = selected_font
    run._element.get_or_add_rPr().rFonts.set(qn("w:ascii"), selected_font)
    run._element.get_or_add_rPr().rFonts.set(qn("w:hAnsi"), selected_font)
    run._element.get_or_add_rPr().rFonts.set(qn("w:cs"), selected_font)
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = RGBColor.from_string(color)
    rtl = run._element.get_or_add_rPr().find(qn("w:rtl"))
    if rtl is None:
        rtl = OxmlElement("w:rtl")
        run._element.get_or_add_rPr().append(rtl)
    rtl.set(qn("w:val"), "1")


def set_image_alt(inline_shape, description: str) -> None:
    inline_shape._inline.docPr.set("descr", description)
    inline_shape._inline.docPr.set("title", description)


def style_paragraph(p, after=6, before=0, line=1.25, keep_next=False) -> None:
    rtl_paragraph(p)
    pf = p.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.line_spacing = line
    pf.keep_with_next = keep_next
    pf.widow_control = True


def add_p(doc, text="", *, after=6, before=0, line=1.25, bold=False,
          color=INK, size=11.5, align=WD_ALIGN_PARAGRAPH.RIGHT,
          keep_next=False, italic=False):
    p = doc.add_paragraph()
    rtl_paragraph(p, align)
    pf = p.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.line_spacing = line
    pf.keep_with_next = keep_next
    pf.widow_control = True
    if text:
        r = p.add_run(text)
        set_run(r, size=size, bold=bold, color=color, italic=italic)
    return p


def add_heading(doc, text, level=1):
    p = doc.add_paragraph(style=f"Heading {level}")
    rtl_paragraph(p)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(text)
    if level == 1:
        set_run(r, size=17, bold=True, color=NAVY, font_name=HEADING_FONT)
        p.paragraph_format.space_before = Pt(12)
        p.paragraph_format.space_after = Pt(7)
    elif level == 2:
        set_run(r, size=14, bold=True, color=RED, font_name=HEADING_FONT)
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.space_after = Pt(4)
    else:
        set_run(r, size=12.5, bold=True, color=NAVY, font_name=HEADING_FONT)
        p.paragraph_format.space_before = Pt(6)
        p.paragraph_format.space_after = Pt(3)
    return p


def add_bullet(doc, text, level=0, color=INK):
    p = doc.add_paragraph(style="List Bullet" if level == 0 else "List Bullet 2")
    rtl_paragraph(p)
    p.paragraph_format.left_indent = Cm(0)
    p.paragraph_format.right_indent = Cm(0.55 + level * 0.45)
    p.paragraph_format.first_line_indent = Cm(-0.25)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.line_spacing = 1.22
    p.paragraph_format.widow_control = True
    set_run(p.add_run(text), size=11.2, color=color)
    return p


def add_number(doc, text):
    p = doc.add_paragraph(style="List Number")
    rtl_paragraph(p)
    p.paragraph_format.right_indent = Cm(0.58)
    p.paragraph_format.first_line_indent = Cm(-0.28)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.line_spacing = 1.22
    set_run(p.add_run(text), size=11.2)
    return p


def add_callout(doc, label, text, fill=PALE_BLUE, accent=BLUE):
    table = doc.add_table(rows=1, cols=1)
    tr_pr = table.rows[0]._tr.get_or_add_trPr()
    cant_split = OxmlElement("w:cantSplit")
    cant_split.set(qn("w:val"), "1")
    tr_pr.append(cant_split)
    set_table_geometry(table, [16.65], indent_dxa=120)
    cell = table.cell(0, 0)
    set_cell_shading(cell, fill)
    set_cell_borders(cell, color=accent, size="10")
    p = cell.paragraphs[0]
    rtl_paragraph(p)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1.25
    set_run(p.add_run(label + "  "), size=10.5, bold=True, color=accent)
    set_run(p.add_run(text), size=10.25, color=INK)
    doc.add_paragraph().paragraph_format.space_after = Pt(1)


def add_table(doc, headers: list[str], rows: list[list[str]], widths_cm: list[float],
              header_fill=NAVY, font_size=9.2):
    table = doc.add_table(rows=1, cols=len(headers))
    set_table_geometry(table, widths_cm, indent_dxa=120)
    table.rows[0]._tr.get_or_add_trPr().append(OxmlElement("w:tblHeader"))
    for i, h in enumerate(headers):
        c = table.rows[0].cells[i]
        set_cell_shading(c, header_fill)
        p = c.paragraphs[0]
        rtl_paragraph(p, WD_ALIGN_PARAGRAPH.CENTER)
        p.paragraph_format.space_after = Pt(0)
        set_run(p.add_run(h), size=font_size, bold=True, color=WHITE)
    for ridx, row in enumerate(rows):
        cells = table.add_row().cells
        if ridx % 2 == 1:
            for c in cells:
                set_cell_shading(c, PALE)
        for i, val in enumerate(row):
            p = cells[i].paragraphs[0]
            rtl_paragraph(p, WD_ALIGN_PARAGRAPH.RIGHT)
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = 1.16
            set_run(p.add_run(val), size=font_size, color=INK)
    set_table_geometry(table, widths_cm, indent_dxa=120)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)
    return table


def add_page_break(doc):
    doc.add_page_break()


def add_hyperlink(paragraph, text, url, size=8.5):
    part = paragraph.part
    rid = part.relate_to(url, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink", is_external=True)
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), rid)
    new_run = OxmlElement("w:r")
    r_pr = OxmlElement("w:rPr")
    color = OxmlElement("w:color")
    color.set(qn("w:val"), BLUE)
    underline = OxmlElement("w:u")
    underline.set(qn("w:val"), "single")
    sz = OxmlElement("w:sz")
    sz.set(qn("w:val"), str(int(size * 2)))
    sz_cs = OxmlElement("w:szCs")
    sz_cs.set(qn("w:val"), str(int(size * 2)))
    rtl = OxmlElement("w:rtl")
    rtl.set(qn("w:val"), "1")
    r_fonts = OxmlElement("w:rFonts")
    for attr in ("ascii", "hAnsi", "cs"):
        r_fonts.set(qn(f"w:{attr}"), FONT)
    r_pr.extend([r_fonts, color, underline, rtl, sz, sz_cs])
    new_run.append(r_pr)
    text_node = OxmlElement("w:t")
    text_node.text = text
    new_run.append(text_node)
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)


def add_page_number(paragraph):
    rtl_paragraph(paragraph, WD_ALIGN_PARAGRAPH.CENTER)
    set_run(paragraph.add_run("صفحه "), size=8.5, color=MUTED)
    run = paragraph.add_run()
    set_run(run, size=8.5, color=MUTED)
    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = " PAGE "
    fld_sep = OxmlElement("w:fldChar")
    fld_sep.set(qn("w:fldCharType"), "separate")
    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")
    run._r.extend([fld_begin, instr, fld_sep, fld_end])


def make_architecture_diagram(path: Path):
    w, h = 1800, 1050
    img = Image.new("RGB", (w, h), "white")
    d = ImageDraw.Draw(img)
    fpath = r"C:\Windows\Fonts\tahoma.ttf"
    fbpath = r"C:\Windows\Fonts\tahomabd.ttf"
    title = ImageFont.truetype(fbpath, 52)
    body = ImageFont.truetype(fpath, 32)
    small = ImageFont.truetype(fpath, 27)
    fa = lambda text: get_display(arabic_reshaper.reshape(text))
    d.text((w/2, 38), fa("معماری مفهومی سامانه همراهِ انتخاب"), font=title, fill="#211A3D", anchor="ma")
    boxes = [
        ("درگاه کاربر و مشاور", "وب و موبایل | احراز هویت | رضایت‌نامه | پنل پرونده", "#EDF4FD", "#4F92EA"),
        ("لایه خدمات کسب‌وکار", "مدیریت پرونده | گردش‌کار تأیید | مشاوره | خدمات پس از ازدواج", "#F8ECEB", "#92251F"),
        ("لایه هوش مصنوعی کنترل‌شده", "استخراج سند | تطبیق تناقض | خلاصه‌سازی مستند | توصیه قابل توضیح", "#FFF8D8", "#B88B00"),
        ("درگاه تعامل‌پذیری", "API رسمی دستگاه‌ها | ثبت رویداد | محدودسازی دسترسی | مدیریت خطا", "#EEF8F0", "#4D9B5F"),
        ("خزانه امن داده و حاکمیت", "رمزنگاری | تفکیک شناسه‌ها | نگهداشت حداقلی | ممیزی | حذف و اعتراض", "#F5F4F8", "#66636F"),
    ]
    x0, x1 = 170, 1630
    y = 125
    for i, (head, detail, fill, stroke) in enumerate(boxes):
        y1 = y + 145
        d.rounded_rectangle((x0, y, x1, y1), radius=28, fill=fill, outline=stroke, width=5)
        d.text((w/2, y+34), fa(head), font=body, fill="#211A3D", anchor="ma")
        d.text((w/2, y+91), fa(detail), font=small, fill="#34313D", anchor="ma")
        if i < len(boxes)-1:
            cx = w//2
            d.line((cx, y1+5, cx, y1+32), fill="#8A8792", width=6)
            d.polygon([(cx-12, y1+26), (cx+12, y1+26), (cx, y1+42)], fill="#8A8792")
        y = y1 + 45
    img.save(path, quality=95)


def configure_document(doc: Document):
    sec = doc.sections[0]
    sec.page_width = Cm(21.0)
    sec.page_height = Cm(29.7)
    sec.top_margin = Cm(1.8)
    sec.bottom_margin = Cm(1.7)
    sec.left_margin = Cm(1.8)
    sec.right_margin = Cm(1.8)
    sec.header_distance = Cm(0.75)
    sec.footer_distance = Cm(0.75)
    sec.gutter = Cm(0)

    normal = doc.styles["Normal"]
    normal.font.name = FONT
    normal._element.rPr.rFonts.set(qn("w:ascii"), FONT)
    normal._element.rPr.rFonts.set(qn("w:hAnsi"), FONT)
    normal._element.rPr.rFonts.set(qn("w:cs"), FONT)
    normal.font.size = Pt(11.5)
    pf = normal.paragraph_format
    pf.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    pf.space_after = Pt(6)
    pf.line_spacing = 1.25
    bidi = normal._element.get_or_add_pPr().find(qn("w:bidi"))
    if bidi is None:
        bidi = OxmlElement("w:bidi")
        normal._element.get_or_add_pPr().append(bidi)
    bidi.set(qn("w:val"), "1")

    for style_name, size, color, before, after in [
        ("Heading 1", 17, NAVY, 12, 7),
        ("Heading 2", 14, RED, 8, 4),
        ("Heading 3", 12.5, NAVY, 6, 3),
    ]:
        st = doc.styles[style_name]
        st.font.name = HEADING_FONT
        st._element.rPr.rFonts.set(qn("w:ascii"), HEADING_FONT)
        st._element.rPr.rFonts.set(qn("w:hAnsi"), HEADING_FONT)
        st._element.rPr.rFonts.set(qn("w:cs"), HEADING_FONT)
        st.font.size = Pt(size)
        st.font.bold = True
        st.font.color.rgb = RGBColor.from_string(color)
        st.paragraph_format.space_before = Pt(before)
        st.paragraph_format.space_after = Pt(after)
        st.paragraph_format.keep_with_next = True
        p_pr = st._element.get_or_add_pPr()
        b = p_pr.find(qn("w:bidi"))
        if b is None:
            b = OxmlElement("w:bidi")
            p_pr.append(b)
        b.set(qn("w:val"), "1")

    # Header on every page, including the cover.
    header = sec.header
    header.is_linked_to_previous = False
    hp = header.paragraphs[0]
    rtl_paragraph(hp)
    table = header.add_table(rows=1, cols=3, width=Cm(17.4))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    widths = [3.2, 10.9, 3.2]
    for idx, w in enumerate(widths):
        table.cell(0, idx).width = Cm(w)
        set_cell_margins(table.cell(0, idx), top=0, bottom=0, start=30, end=30)
        set_cell_borders(table.cell(0, idx), color=WHITE, size="0")
    p0 = table.cell(0, 0).paragraphs[0]
    p0.alignment = WD_ALIGN_PARAGRAPH.LEFT
    pic = p0.add_run().add_picture(str(LOGO_SCI), width=Cm(1.05))
    set_image_alt(pic, "نشان آکادمی علوم و فناوری‌های نوین Sciencinnov")
    pc = table.cell(0, 1).paragraphs[0]
    rtl_paragraph(pc, WD_ALIGN_PARAGRAPH.CENTER)
    set_run(pc.add_run("پروپوزال سامانه هوشمند همراهِ انتخاب"), size=8.8, bold=True, color=NAVY)
    p2 = table.cell(0, 2).paragraphs[0]
    p2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    pic = p2.add_run().add_picture(str(LOGO_ICDS), width=Cm(2.35))
    set_image_alt(pic, "نشان Innovation Center for Data Science")
    # subtle rule
    rule_p = header.add_paragraph()
    rule_p.paragraph_format.space_after = Pt(0)
    p_pr = rule_p._p.get_or_add_pPr()
    p_bdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "5")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), LINE)
    p_bdr.append(bottom)
    p_pr.append(p_bdr)

    footer = sec.footer
    fp = footer.paragraphs[0]
    add_page_number(fp)
    fp.paragraph_format.space_before = Pt(2)


def add_cover(doc):
    add_p(doc, "پیشنهاد فنی، اجرایی و امکان‌سنجی", size=11, bold=True, color=RED,
          align=WD_ALIGN_PARAGRAPH.CENTER, before=12, after=16)
    logos = doc.add_table(rows=1, cols=2)
    logos.alignment = WD_TABLE_ALIGNMENT.CENTER
    logos.autofit = False
    for c in logos.rows[0].cells:
        set_cell_borders(c, color=WHITE, size="0")
        set_cell_margins(c, top=0, bottom=0, start=80, end=80)
    p = logos.cell(0, 0).paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pic = p.add_run().add_picture(str(LOGO_SCI), width=Cm(4.0))
    set_image_alt(pic, "نشان آکادمی علوم و فناوری‌های نوین Sciencinnov")
    p = logos.cell(0, 1).paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pic = p.add_run().add_picture(str(LOGO_ICDS), width=Cm(5.0))
    set_image_alt(pic, "نشان Innovation Center for Data Science")
    add_p(doc, "سامانه هوشمند «همراهِ انتخاب»", size=25, bold=True, color=NAVY,
          align=WD_ALIGN_PARAGRAPH.CENTER, before=20, after=7)
    add_p(doc, "پلتفرم حمایت از تصمیم آگاهانه برای ازدواج، احرازپذیری اطلاعات و خدمات پس از ازدواج",
          size=13.2, color=RED, align=WD_ALIGN_PARAGRAPH.CENTER, after=20, line=1.25)
    add_callout(doc, "چشم‌انداز پروژه", "ایجاد یک زیرساخت مورد اعتماد برای کاهش عدم‌تقارن اطلاعات، تقویت گفت‌وگوی آگاهانه و هدایت زوجین به خدمات معتبر؛ با رضایت دوطرفه، دسترسی رسمی، حداقل‌گرایی داده و نظارت انسانی.", fill=PALE_BLUE, accent=BLUE)
    add_p(doc, "مخاطب: نهاد سفارش‌دهنده و شرکای اجرایی بین‌دستگاهی", size=10.3, bold=True,
          align=WD_ALIGN_PARAGRAPH.CENTER, before=16, after=5)
    add_p(doc, "نسخه ۱.۱ توسعه‌یافته | تیر ۱۴۰۵ | طبقه‌بندی: پیشنهادی - نیازمند بررسی حقوقی و فنی کارفرما",
          size=9.2, color=MUTED, align=WD_ALIGN_PARAGRAPH.CENTER, after=0)
    add_page_break(doc)


def build():
    arch = ASSET_DIR / "architecture.png"
    make_architecture_diagram(arch)
    doc = Document()
    configure_document(doc)
    doc.core_properties.title = "پروپوزال سامانه هوشمند همراه انتخاب"
    doc.core_properties.subject = "امکان‌سنجی و طراحی پلتفرم حمایت از تصمیم آگاهانه برای ازدواج"
    doc.core_properties.author = "تیم تدوین پروپوزال"
    doc.core_properties.keywords = "ازدواج، هوش مصنوعی، حاکمیت داده، احراز اطلاعات، مشاوره"

    add_cover(doc)

    add_heading(doc, "خلاصه مدیریتی", 1)
    add_p(doc, "سامانه «همراهِ انتخاب» به‌عنوان یک خدمت عمومی-فناورانه پیشنهاد می‌شود تا افراد در آستانه ازدواج بتوانند با رضایت متقابل، اطلاعات ادعاشده را از مسیرهای معتبر بررسی کنند، پرسش‌های مهم را ساختارمند پیش ببرند و در صورت نیاز به مشاور انسانی یا خدمت حمایتی هدایت شوند. ارزش اصلی سامانه در گردآوری پنهانی اطلاعات نیست؛ بلکه در تبدیل داده‌های مجاز و قابل استناد به تصویری روشن، قابل توضیح و قابل اعتراض برای خود افراد است.")
    add_p(doc, "این پروپوزال درخواست اولیه را به یک طرح قابل دفاع حقوقی و اجرایی تبدیل می‌کند. اتصال به داده‌های قضایی، تحصیلی و شغلی تنها در صورت وجود مجوز، تفاهم‌نامه، API رسمی، احراز هویت قوی و رضایت معتبر قابل اجراست. برای سابقه قضایی نیز الگوی پیشنهادی، دریافت «گواهی رسمی ارائه‌شده توسط خود فرد» یا استعلام مجاز با توکن یک‌بارمصرف است؛ نه بازیابی تاریخچه کامل پرونده‌ها یا دسترسی به حساب ثنای اشخاص.")
    add_callout(doc, "جمع‌بندی تصمیم", "شروع پروژه توصیه می‌شود، اما فقط با یک فاز صفر حقوقی-تعامل‌پذیری و سپس MVP کم‌ریسک. قابلیت‌های پرریسک مانند دسترسی به تاریخچه قضایی اشخاص ثالث، تحلیل لایک‌ها یا حساب‌های خصوصی و برچسب‌زنی عقیدتی باید خارج از دامنه باقی بمانند مگر اینکه مبنای قانونی صریح و سازوکار رضایت معتبر ایجاد شود.", fill=PALE_RED, accent=RED)
    add_heading(doc, "نتایج مورد انتظار", 2)
    for x in [
        "کاهش عدم‌تقارن اطلاعات از طریق مدارک قابل راستی‌آزمایی و ثبت منشأ هر داده.",
        "افزایش کیفیت گفت‌وگوهای پیش از ازدواج با پرسشنامه و مسیرهای مشاوره مبتنی بر شواهد.",
        "کاهش جعل، تحریف و برداشت اشتباه با اعتبارسنجی سند، کنترل تناقض و نمایش سطح اطمینان.",
        "ایجاد مسیر حمایتی پس از ازدواج شامل مشاوره، مهارت‌آموزی و معرفی داوطلبانه به فرصت‌های شغلی.",
        "ایجاد الگوی حاکمیت داده قابل ممیزی برای همکاری ایمن با دستگاه‌های عمومی و خصوصی.",
    ]:
        add_bullet(doc, x)

    add_heading(doc, "فهرست مطالب", 1)
    toc = [
        "۱. بیان مسئله و ضرورت", "۲. چشم‌انداز، اصول و اهداف", "۳. ذی‌نفعان و تجربه کاربر",
        "۴. دامنه محصول و ماژول‌ها", "۵. ماتریس امکان‌پذیری قابلیت‌ها", "۶. معماری فنی و هوش مصنوعی",
        "۷. حاکمیت داده، حریم خصوصی و الزامات حقوقی", "۸. امنیت و تاب‌آوری", "۹. نقشه راه اجرا",
        "۱۰. تیم، اقلام تحویلی و مدل هزینه", "۱۱. شاخص‌های موفقیت و پذیرش", "۱۲. ریسک‌ها و کنترل‌ها",
        "۱۳. جمع‌بندی و تصمیمات مورد نیاز", "پیوست‌های اجرایی، نمونه گزارش و منابع",
    ]
    for item in toc:
        p = add_bullet(doc, item, color=NAVY)
        p.paragraph_format.space_after = Pt(1.5)
        p.paragraph_format.line_spacing = 1.05
        for run in p.runs:
            set_run(run, size=9.2, color=NAVY)

    add_page_break(doc)
    add_heading(doc, "۱. بیان مسئله و ضرورت", 1)
    add_heading(doc, "۱-۱. مسئله محوری", 2)
    add_p(doc, "تصمیم برای ازدواج با داده‌های پراکنده، روایت‌های شفاهی، پرس‌وجوهای غیررسمی و محدودیت زمان مواجه است. در چنین محیطی، هم احتمال پنهان‌ماندن واقعیت‌های مهم وجود دارد و هم احتمال شکل‌گیری قضاوت نادرست درباره فرد بر پایه شایعه، تشابه اسمی یا داده بدون زمینه. بنابراین مسئله صرفاً «کمبود داده» نیست؛ مسئله اصلی، کمبود اطلاعات معتبر، متناسب، رضایت‌محور و قابل توضیح است.")
    add_heading(doc, "۱-۲. فرصت مداخله", 2)
    add_p(doc, "یک سامانه هوشمند می‌تواند نقش هماهنگ‌کننده را بر عهده گیرد: دریافت ادعاها و مدارک، ثبت رضایت، استعلام مجاز، تشخیص ناسازگاری‌های آشکار، خلاصه‌سازی با ذکر منبع، طرح پرسش‌های تکمیلی و ارجاع موارد حساس به انسان. این سامانه جایگزین شناخت، مشاوره یا تصمیم فردی نیست؛ بلکه زیرساخت تصمیم بهتر است.")
    add_p(doc, "ارزش‌آفرینی سامانه در سه لایه شکل می‌گیرد: نخست «راستی‌آزمایی» ادعاهای مشخص از منبع معتبر؛ دوم «تفسیر محتاطانه» شواهد با نمایش محدودیت و سطح اطمینان؛ و سوم «مداخله حمایتی» از طریق گفت‌وگوی هدایت‌شده، مشاوره و ارجاع. این تفکیک مانع آن می‌شود که گردآوری داده به‌تنهایی با شناخت کامل فرد یا پیش‌بینی موفقیت ازدواج اشتباه گرفته شود.")
    add_heading(doc, "۱-۳. مرز مسئله", 2)
    add_table(doc, ["موضوع", "رویکرد پیشنهادی", "رویکرد مردود یا پرریسک"], [
        ["اطلاعات قضایی", "گواهی رسمی با مشارکت خود فرد یا استعلام مجاز", "ورود به حساب ثنا، دریافت تاریخچه پرونده‌ها یا جست‌وجوی اشخاص ثالث"],
        ["ردپای وب", "تحلیل محتوای عمومیِ معرفی‌شده توسط صاحب حساب با ذکر شواهد", "خزش حساب خصوصی، تحلیل مخفیانه لایک‌ها یا برچسب‌گذاری عقیدتی"],
        ["تحصیل و اشتغال", "مدرک دیجیتال، QR/امضای معتبر یا API رسمی", "اتکا به جست‌وجوی نام یا داده خریداری‌شده از واسطه"],
        ["سازگاری", "توصیه قابل توضیح و موضوعات قابل گفت‌وگو", "نمره قطعی همسر مناسب/نامناسب یا تصمیم خودکار"],
    ], [3.1, 6.8, 6.75], font_size=8.8)

    add_heading(doc, "۲. چشم‌انداز، اصول و اهداف", 1)
    add_heading(doc, "۲-۱. چشم‌انداز", 2)
    add_p(doc, "«هر فرد پیش از ازدواج بتواند با اختیار و آگاهی، اطلاعات مهم را از مسیر قابل اعتماد بررسی کند، ابهام‌ها را به گفت‌وگو تبدیل کند و بدون واگذاری اختیار تصمیم به الگوریتم، از حمایت حرفه‌ای بهره‌مند شود.»")
    add_heading(doc, "۲-۲. اصول طراحی غیرقابل مذاکره", 2)
    principles = [
        ("رضایت دوطرفه و قابل پس‌گرفتن", "هر منبع داده، هدف استفاده، مدت نگهداری و گیرنده نتیجه جداگانه تأیید شود."),
        ("حداقل‌گرایی داده", "فقط داده لازم برای هدف مشخص دریافت شود؛ نه هر داده‌ای که از نظر فنی قابل جمع‌آوری است."),
        ("اصل ادعا + شاهد", "هر نتیجه باید منشأ، زمان، سطح اطمینان و محدودیت خود را نشان دهد."),
        ("نظارت انسانی", "موارد پراثر، تناقض‌ها و اعتراض‌ها به کارشناس مجاز ارجاع شود."),
        ("عدم تبعیض", "استنباط قومیت، مذهب، گرایش سیاسی، وضعیت سلامت یا سایر صفات حساس از رفتار آنلاین ممنوع باشد."),
        ("حق توضیح، اصلاح و حذف", "کاربر بتواند داده را ببیند، منبع را بشناسد، خطا را اصلاح و در حدود قانون حذف کند."),
    ]
    for head, detail in principles:
        p = add_p(doc, after=5, line=1.25)
        set_run(p.add_run(head + ": "), size=10.4, bold=True, color=RED)
        set_run(p.add_run(detail), size=10.25, color=INK)
    add_heading(doc, "۲-۳. اهداف قابل سنجش", 2)
    for x in [
        "ایجاد پرونده رضایت‌محور و قابل ممیزی برای هر فرایند آشنایی.",
        "راستی‌آزمایی حداقل سه گروه اطلاعات در پایلوت: هویت، تحصیل و اشتغال/گواهی‌های رسمی.",
        "ارائه خلاصه‌ای حداکثر در چند صفحه با پیوند به اصل شواهد و پرهیز از افشای داده غیرضروری.",
        "کاهش زمان بررسی مدارک و افزایش نرخ کشف تناقض‌های قابل اثبات در مقایسه با فرایند دستی.",
        "ایجاد سازوکار اعتراض و اصلاح با زمان پاسخ مشخص.",
    ]:
        add_bullet(doc, x)

    add_heading(doc, "۳. ذی‌نفعان و تجربه کاربر", 1)
    add_heading(doc, "۳-۱. ذی‌نفعان", 2)
    add_table(doc, ["ذی‌نفع", "نیاز اصلی", "نقش در سامانه"], [
        ["فرد متقاضی", "آگاهی، کنترل و امنیت", "مالک رضایت و ارائه‌کننده مدارک خود"],
        ["طرف مقابل", "تقارن اطلاعات و امکان پاسخ", "مشاهده محدود، تأیید متقابل و اعتراض"],
        ["مشاور خانواده", "اطلاعات ساختاریافته و زمینه‌دار", "تفسیر انسانی، ارجاع و ثبت یادداشت حرفه‌ای"],
        ["نهاد سفارش‌دهنده", "اثربخشی، پاسخ‌گویی و انطباق", "مالک محصول و سیاست‌های حاکمیت"],
        ["دستگاه‌های داده‌دار", "استفاده مجاز و کنترل‌شده", "ارائه API/گواهی و ثبت رویداد"],
        ["نهاد ناظر", "شفافیت و امکان حسابرسی", "بازرسی فرایند، مدل و رخدادهای امنیتی"],
    ], [3.3, 6.1, 7.25], font_size=9.0)
    add_heading(doc, "۳-۲. سفر نمونه کاربر", 2)
    steps = [
        "ثبت‌نام و احراز هویت قوی؛ معرفی هدف استفاده از سامانه با زبان ساده.",
        "ایجاد «پرونده آشنایی» و دعوت طرف مقابل؛ پرونده تا پذیرش او فعال نمی‌شود.",
        "انتخاب بسته‌های بررسی؛ برای هر بسته رضایت جداگانه و زمان‌دار ثبت می‌شود.",
        "ارائه مدارک یا صدور توکن استعلام رسمی؛ نتیجه خام در خزانه امن و خلاصه حداقلی در پرونده ذخیره می‌شود.",
        "هوش مصنوعی تناقض‌های قابل اثبات و پرسش‌های نیازمند گفت‌وگو را پیشنهاد می‌دهد؛ هیچ حکم قطعی صادر نمی‌کند.",
        "مشاهده مشترک گزارش، پاسخ صاحب داده و در صورت لزوم جلسه با مشاور.",
        "بستن پرونده، تعیین نگهداشت یا حذف؛ در صورت ازدواج، فعال‌سازی اختیاری خدمات پس از ازدواج.",
    ]
    for s in steps:
        add_number(doc, s)
    add_callout(doc, "نقطه کنترل", "هیچ کاربری صرفاً با وارد کردن نام و نام خانوادگی فرد دیگر نباید بتواند گزارش تولید کند. حداقل احراز هویت، تطبیق شناسه یکتا، دعوت معتبر و رضایت صاحب داده الزامی است.", fill=PALE_RED, accent=RED)

    add_heading(doc, "۳-۳. سناریوهای کاربردی تفصیلی", 2)
    add_heading(doc, "سناریوی اول: تشکیل پرونده مشترک پیش از ازدواج", 3)
    add_p(doc, "کاربر نخست پس از احراز هویت، یک پرونده آشنایی ایجاد و طرف مقابل را از طریق دعوت زمان‌دار وارد می‌کند. پیش از هر اقدام، سامانه هدف خدمت، انواع داده، گیرندگان و مدت نگهداشت را با زبان ساده نمایش می‌دهد. هر دو نفر می‌توانند بسته‌های بررسی را جداگانه انتخاب کنند؛ برای مثال یک زوج ممکن است فقط هویت و تحصیل را بررسی کند و زوج دیگر علاوه بر آن، گواهی اشتغال و گواهی رسمی عدم سوءپیشینه را نیز در دامنه رضایت قرار دهد.")
    add_p(doc, "پس از پذیرش دعوت، هر فرد ادعاهای مربوط به خود را ثبت و مدارک لازم را ارائه می‌کند. سامانه اصالت فنی سند، سازگاری فیلدها و پاسخ مرجع رسمی را بررسی می‌کند و نتیجه را ابتدا به صاحب داده نشان می‌دهد. صاحب داده فرصت دارد خطای انتساب، قدیمی‌بودن مدرک یا توضیح زمینه‌ای را ثبت کند. تنها پس از این مرحله، خلاصه حداقلی و متقارن برای طرف مقابل قابل مشاهده می‌شود. این ترتیب، اصل شنیده‌شدن فرد و جلوگیری از غافلگیری را در طراحی محصول تثبیت می‌کند.")
    add_heading(doc, "سناریوی دوم: مشاهده تناقض و ارجاع به گفت‌وگو", 3)
    add_p(doc, "فرض شود بازه اشتغال درج‌شده در رزومه با بازه گواهی ارائه‌شده یکسان نیست. سامانه نباید این اختلاف را دروغ یا سوءنیت تفسیر کند؛ بلکه دو تاریخ، منبع هر تاریخ و سطح اطمینان استخراج را کنار هم قرار می‌دهد و از صاحب داده می‌خواهد توضیح یا مدرک تکمیلی ارائه کند. اگر اختلاف با یک خطای OCR یا تفاوت تاریخ صدور برطرف شود، گزارش اصلاح و سابقه تغییر برای ممیزی حفظ می‌شود.")
    add_p(doc, "اگر اختلاف پس از توضیح نیز باقی بماند، خروجی به صورت «نیازمند گفت‌وگو» نمایش داده می‌شود. سامانه چند سؤال خنثی پیشنهاد می‌کند؛ مانند اینکه ماهیت همکاری تمام‌وقت، پروژه‌ای یا بدون بیمه بوده است. در موارد حساس، کاربر می‌تواند جلسه‌ای با مشاور رزرو کند. هدف این مسیر، تبدیل تناقض قابل مشاهده به گفت‌وگوی ساختاریافته است، نه تولید امتیاز منفی یا قضاوت شخصیت.")
    add_heading(doc, "سناریوی سوم: فعال‌سازی خدمت پس از ازدواج", 3)
    add_p(doc, "پس از ثبت ازدواج یا اعلام داوطلبانه زوجین، پرونده آشنایی به‌طور خودکار به خدمت پس از ازدواج تبدیل نمی‌شود. هر فرد یک دعوت جدید با توضیح هدف تازه دریافت می‌کند و می‌تواند فقط خدمات مورد نیاز مانند آموزش مهارت ارتباطی، مشاوره مالی یا مسیر اشتغال را فعال کند. داده‌هایی که برای راستی‌آزمایی پیش از ازدواج جمع‌آوری شده‌اند، بدون رضایت تازه وارد این خدمت نمی‌شوند.")
    add_p(doc, "برای معرفی شغلی، سامانه یک پروفایل مهارتی مستقل ایجاد می‌کند که شامل مهارت‌ها، سوابق قابل انتشار، ترجیح شغلی و محدوده جغرافیایی است. نهاد پذیرنده فقط همین پروفایل را دریافت می‌کند و هیچ دسترسی به پاسخ‌های مشاوره‌ای، اطلاعات قضایی، گزارش سازگاری یا داده طرف مقابل ندارد. نتیجه معرفی نیز به صورت وضعیت فرایند ثبت می‌شود تا از استفاده ثانویه و نامتناسب جلوگیری شود.")
    add_heading(doc, "۳-۴. سناریوهای استثنا و توقف", 2)
    for x in [
        "رد دعوت یا لغو رضایت: پرونده مشترک متوقف می‌شود و هیچ استعلام تازه‌ای انجام نمی‌گیرد؛ وضعیت داده موجود طبق سیاست نگهداشت تعیین می‌شود.",
        "عدم دسترسی مرجع رسمی: نتیجه «نامشخص/سرویس در دسترس نیست» نمایش داده می‌شود و با نتیجه منفی اشتباه گرفته نمی‌شود.",
        "ادعای انتساب اشتباه حساب وب: تحلیل پنهان و حساب تا بررسی انسانی از گزارش کنار گذاشته می‌شود.",
        "وجود نشانه تهدید یا خشونت: سامانه از تحلیل خودکار فراتر نمی‌رود و مطابق پروتکل مصوب، مسیر کمک حرفه‌ای و ایمن را نمایش می‌دهد.",
        "شکایت درباره افشای داده: اشتراک نتیجه متوقف، شواهد ممیزی حفظ و رخداد به مسئول حریم خصوصی و امنیت ارجاع می‌شود.",
    ]:
        add_bullet(doc, x)

    add_heading(doc, "۴. دامنه محصول و ماژول‌ها", 1)
    modules = [
        ("۴-۱. هویت، رضایت و مدیریت پرونده", [
            "احراز هویت چندمرحله‌ای و تطبیق شماره همراه با صاحب شناسه.",
            "رضایت‌نامه دانه‌ریز برای هر منبع، هدف، گیرنده و دوره زمانی.",
            "داشبورد پرونده، ثبت تاریخچه دسترسی، امکان لغو رضایت و بستن پرونده.",
        ]),
        ("۴-۲. راستی‌آزمایی قضاییِ محدود و مجاز", [
            "پشتیبانی از بارگذاری گواهی عدم سوءپیشینه و کنترل اصالت آن از مسیر رسمی.",
            "در صورت عقد تفاهم‌نامه، استعلام مبتنی بر توکن و نمایش نتیجه حداقلی مورد توافق.",
            "عدم ذخیره رمز ثنا، عدم ورود نیابتی به حساب فرد و عدم نمایش جزئیات پرونده‌های غیرمرتبط.",
        ]),
        ("۴-۳. تحصیلات و گواهی‌های حرفه‌ای", [
            "ثبت ادعاهای تحصیلی به‌صورت ساختاریافته: دانشگاه، رشته، مقطع و بازه زمانی.",
            "اعتبارسنجی با API رسمی، QR، امضای دیجیتال یا تأیید مؤسسه صادرکننده.",
            "نمایش وضعیت «تأییدشده، تأییدنشده، نیازمند بررسی» به‌جای نتیجه‌گیری درباره توانایی فرد.",
        ]),
        ("۴-۴. اشتغال و سابقه حرفه‌ای", [
            "گواهی اشتغال یا سابقه بیمه/کار با رضایت و دسترسی قانونی؛ تطبیق بازه زمانی و عنوان شغلی.",
            "ارائه معرفی‌نامه داوطلبانه و تماس با مرجع معرفی‌شده در قالب گردش‌کار ثبت‌شده.",
            "تفکیک «عدم امکان تأیید» از «نادرست بودن ادعا» برای جلوگیری از برچسب‌زنی ناعادلانه.",
        ]),
        ("۴-۵. ردپای عمومی وب و شبکه‌های اجتماعی", [
            "دریافت شناسه حساب از خود فرد و تحلیل صرفاً محتوای عمومی در چارچوب شرایط پلتفرم.",
            "خلاصه موضوعات تکرارشونده، لحن عمومی و علایق خوداظهاری‌شده همراه با نمونه و تاریخ.",
            "ممنوعیت استنباط صفات حساس، دسترسی به پیام خصوصی، دورزدن محدودیت‌ها یا ساخت امتیاز عقیدتی.",
        ]),
        ("۴-۶. گفت‌وگوی هدایت‌شده و مشاوره", [
            "پرسشنامه درباره ارزش‌ها، اهداف، مالی، خانواده، فرزندآوری، محل زندگی و سبک حل تعارض.",
            "تولید موضوعات پیشنهادی برای گفت‌وگو و ارجاع موارد حساس به مشاور دارای صلاحیت.",
            "ثبت رضایت جداگانه برای اشتراک پاسخ‌ها و جلوگیری از نمایش پاسخ‌های خصوصی بدون اجازه.",
        ]),
        ("۴-۷. خدمات پس از ازدواج و توانمندسازی", [
            "ادامه پرونده فقط با رضایت تازه و هدف جدید؛ داده دوره آشنایی خودکار به دوره پس از ازدواج منتقل نشود.",
            "مسیرهای مشاوره، آموزش مهارت‌های ارتباطی، سواد مالی و حل تعارض.",
            "پروفایل مهارتی جداگانه برای معرفی داوطلبانه به آموزش یا فرصت شغلی؛ بدون افشای داده‌های پرونده ازدواج.",
        ]),
    ]
    for head, bullets in modules:
        add_heading(doc, head, 2)
        for x in bullets:
            add_bullet(doc, x)
        if head.startswith("۴-۲"):
            add_p(doc, "خروجی این ماژول باید تا حد ممکن به وضعیت گواهی، تاریخ صدور و مرجع صادرکننده محدود باشد. جزئیات نامرتبط پرونده، اطلاعات اشخاص ثالث و داده‌ای که صاحب آن امکان مشاهده و پاسخ به آن را ندارد، وارد گزارش تصمیم‌یار نمی‌شود.", size=10.8, color=MUTED)
        elif head.startswith("۴-۵"):
            add_p(doc, "تحلیل وب صرفاً یک مرور زمینه‌دار از محتوای عمومی است و نباید به «پرونده شخصیت» تبدیل شود. هر گزاره باید با نمونه قابل مشاهده، تاریخ جمع‌آوری و توضیح محدودیت همراه باشد و صاحب حساب بتواند انتساب حساب یا برداشت سامانه را رد و اصلاح کند.", size=10.8, color=MUTED)
        elif head.startswith("۴-۷"):
            add_p(doc, "فعال‌سازی خدمات پس از ازدواج یک فرایند مستقل با رضایت، هدف و دوره نگهداشت تازه است. معرفی شغلی نیز فقط بر پایه پروفایل مهارتی‌ای انجام می‌شود که فرد برای همین هدف تکمیل کرده است و داده‌های حساس دوره آشنایی در اختیار کارفرما یا نهاد معرفی‌کننده قرار نمی‌گیرد.", size=10.8, color=MUTED)

    add_heading(doc, "۴-۸. ساختار گزارش نهایی برای کاربران", 2)
    for x in [
        "خلاصه ادعاهای تأییدشده، تأییدنشده و در انتظار بررسی، همراه با تاریخ و مرجع هر شاهد.",
        "تناقض‌های قابل مشاهده با توضیح خنثی و فضای پاسخ صاحب داده؛ بدون تبدیل اختلاف به اتهام.",
        "محدودیت‌های پوشش داده و سطح اطمینان هر نتیجه، به‌ویژه در تحلیل وب و اسناد کم‌کیفیت.",
        "پرسش‌های پیشنهادی برای گفت‌وگوی دوطرفه و مسیر ارجاع به مشاور در موضوعات حساس.",
        "کنترل‌های مشاهده، اعتراض، اصلاح، لغو رضایت و بستن پرونده در دسترس هر دو طرف.",
    ]:
        add_bullet(doc, x)

    add_heading(doc, "۴-۹. بسته‌های خدمت و سطح نمایش اطلاعات", 2)
    add_p(doc, "برای کنترل دامنه و هزینه، پیشنهاد می‌شود قابلیت‌ها در قالب بسته‌های مستقل عرضه شوند. فعال‌سازی هر بسته به معنای رضایت برای سایر بسته‌ها نیست و کاربر باید بتواند سطح جزئیات را پیش از صدور گزارش مشاهده کند.")
    add_table(doc, ["بسته خدمت", "داده ورودی", "خروجی قابل نمایش", "کنترل ویژه"], [
        ["هویت و تماس", "شناسه هویتی و شماره همراه", "تطبیق هویت و وضعیت احراز", "عدم نمایش شناسه کامل به طرف مقابل"],
        ["تحصیلات", "ادعا، مدرک و پاسخ مرجع", "مقطع، رشته، مرجع و وضعیت تأیید", "تفکیک عدم پاسخ از رد ادعا"],
        ["اشتغال", "گواهی/معرفی‌نامه و بازه زمانی", "وضعیت اشتغال و سابقه منتخب", "عدم نمایش درآمد مگر انتخاب صریح"],
        ["گواهی قضایی", "گواهی رسمی ارائه‌شده توسط فرد", "اعتبار گواهی و تاریخ صدور", "عدم نمایش جزئیات پرونده‌های نامرتبط"],
        ["وب عمومی", "حساب معرفی‌شده و محتوای عمومی", "موضوعات مستند و محدودیت پوشش", "بدون امتیاز عقیدتی یا صفت حساس"],
        ["گفت‌وگو و مشاوره", "پاسخ‌های انتخاب‌شده و ارجاع", "موضوعات قابل گفت‌وگو و برنامه جلسه", "پاسخ خصوصی فقط با مجوز صاحب آن"],
    ], [3.4, 4.7, 5.1, 3.45], font_size=8.5)
    add_p(doc, "این بسته‌بندی علاوه بر شفافیت برای کاربران، مدیریت قرارداد و نقشه راه فنی را ساده می‌کند. کارفرما می‌تواند MVP را با سه بسته کم‌ریسک آغاز کند و هر اتصال تازه را پس از ارزیابی حقوقی، امنیتی و عملکردی به‌صورت مستقل به محصول بیفزاید.")

    add_heading(doc, "۵. ماتریس امکان‌پذیری قابلیت‌ها", 1)
    add_p(doc, "امکان‌پذیری به چهار عامل وابسته است: مبنای قانونی، رضایت معتبر، وجود رابط رسمی و تناسب داده با هدف. جدول زیر وضعیت پیشنهادی برای شروع مذاکره با کارفرماست و باید در فاز صفر با مالکان داده نهایی شود.")
    add_table(doc, ["قابلیت", "وضعیت", "پیش‌نیاز", "نسخه پیشنهادی"], [
        ["احراز هویت و رضایت", "قابل اجرا", "ارائه‌دهنده احراز معتبر، طراحی حقوقی", "MVP"],
        ["گواهی عدم سوءپیشینه", "قابل اجرا با مشارکت فرد", "گواهی رسمی/QR؛ API در صورت تفاهم‌نامه", "MVP محدود"],
        ["تاریخچه پرونده‌های ثنا", "فعلاً خارج از دامنه", "مجوز قانونی صریح و ضرورت اثبات‌شده", "عدم اجرا"],
        ["تأیید مدرک تحصیلی", "مشروط", "API یا اعتبارسنجی صادرکننده", "پایلوت"],
        ["تأیید اشتغال", "مشروط", "گواهی/مرجع رسمی و رضایت", "پایلوت"],
        ["تحلیل وب عمومی", "قابل اجرا محدود", "حساب معرفی‌شده، شرایط پلتفرم، ممیزی مدل", "MVP آزمایشی"],
        ["تحلیل لایک/فالو خصوصی", "غیرقابل‌پذیرش", "فاقد مسیر عادی رضایت و دسترسی", "عدم اجرا"],
        ["پرسشنامه و مشاوره", "قابل اجرا", "طراحی متخصص و کنترل بالینی/اخلاقی", "MVP"],
        ["تطبیق و توصیه", "قابل اجرا غیرقطعی", "توضیح‌پذیری، عدم تبعیض، نظارت انسانی", "نسخه ۲"],
        ["معرفی شغلی پس از ازدواج", "قابل اجرا جداگانه", "رضایت تازه، تفکیک پایگاه داده و شرکای شغلی", "نسخه ۳"],
    ], [4.2, 3.0, 5.4, 4.05], font_size=8.6)
    add_callout(doc, "تفسیر وضعیت", "«مشروط» به معنای دشوار اما ممکن است؛ تا پیش از امضای تفاهم‌نامه و آزمون API نباید در زمان‌بندی یا قیمت ثابت تعهد داده شود. «عدم اجرا» نیز یک کنترل محصول است، نه صرفاً محدودیت فنی.", fill=PALE_BLUE, accent=BLUE)

    add_heading(doc, "۶. معماری فنی و هوش مصنوعی", 1)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pic = p.add_run().add_picture(str(arch), width=Cm(16.5))
    set_image_alt(pic, "نمودار پنج‌لایه معماری مفهومی سامانه همراه انتخاب")
    cp = add_p(doc, "شکل ۱ - معماری مفهومی پیشنهادی (جزئیات پس از فاز صفر و شناخت زیرساخت کارفرما نهایی می‌شود).", size=8.5, color=MUTED, align=WD_ALIGN_PARAGRAPH.CENTER, after=8)
    add_heading(doc, "۶-۱. مؤلفه‌های کلیدی", 2)
    for x in [
        "درگاه کاربر و مشاور: وب واکنش‌گرا، اعلان، مشاهده مجوزها و دسترسی به گزارش.",
        "مدیریت هویت و رضایت: احراز قوی، توکن یک‌بارمصرف، ثبت نسخه متن رضایت و زمان لغو.",
        "لایه تعامل‌پذیری: اتصال جداگانه و قابل قطع به هر دستگاه، محدودسازی نرخ و ثبت پاسخ خام امضاشده.",
        "خزانه امن: تفکیک اطلاعات هویتی از محتوای پرونده، رمزنگاری و سیاست نگهداشت بر حسب نوع داده.",
        "خدمات هوش مصنوعی: OCR و استخراج، تطبیق ادعا و مدرک، خلاصه‌سازی مستند، پیشنهاد سؤال و کشف تناقض.",
        "پنل حاکمیت: ممیزی دسترسی، مدیریت اعتراض، پایش سوگیری، کنترل نسخه مدل و گزارش رخداد.",
    ]:
        add_bullet(doc, x)
    add_heading(doc, "۶-۲. الگوی استفاده از هوش مصنوعی", 2)
    add_table(doc, ["کارکرد", "خروجی", "کنترل الزامی"], [
        ["استخراج سند", "فیلدهای ساختاریافته و ناحیه شاهد", "نمایش تصویر منبع و تأیید کاربر"],
        ["کشف تناقض", "اختلاف تاریخ/عنوان/مرجع", "عدم تبدیل تناقض به اتهام؛ ارجاع برای توضیح"],
        ["خلاصه وب عمومی", "موضوعات و نمونه‌های مستند", "فقط محتوای عمومی، ذکر تاریخ و امکان اعتراض"],
        ["پیشنهاد گفت‌وگو", "سؤال‌های اولویت‌دار", "عدم تشخیص اختلال یا صدور حکم شخصیتی"],
        ["توصیه خدمت", "مشاور/آموزش/فرصت مناسب", "قابل توضیح، قابل رد و مستقل از داده حساس"],
    ], [4.0, 6.1, 6.55], font_size=8.9)
    add_heading(doc, "۶-۳. اصل «مدل تصمیم‌یار، نه تصمیم‌گیر»", 2)
    add_p(doc, "مدل مجاز است شواهد را مرتب کند، اختلاف‌ها را نشان دهد و موضوعات قابل گفت‌وگو را پیشنهاد کند؛ اما مجاز نیست نتیجه «این فرد مناسب/نامناسب است» صادر کند. گزارش نهایی باید زبان احتمالی، سطح اطمینان، منشأ داده و مسیر اعتراض داشته باشد.")
    add_p(doc, "برای هر کارکرد، آستانه اطمینان و قاعده توقف جداگانه تعریف می‌شود: نتیجه کم‌اطمینان نمایش داده نمی‌شود یا با برچسب «نیازمند بررسی» به کارشناس ارجاع می‌گردد؛ نتیجه پراثر نیز حتی در اطمینان بالا بدون شاهد قابل مشاهده و امکان بازبینی انسانی مبنای اقدام قرار نمی‌گیرد. نسخه مدل، داده آزمون و تغییرات پرامپت باید در چرخه انتشار ثبت و قابل بازگشت باشند.")

    add_heading(doc, "۶-۴. جریان داده از دریافت تا گزارش", 2)
    flow_steps = [
        "دریافت و ثبت رضایت: درخواست کاربر به یک شناسه پرونده و هدف پردازش متصل می‌شود. پیش از ورود هر داده، سامانه اعتبار رضایت، گیرنده مجاز، محدوده زمانی و منبع انتخابی را کنترل می‌کند.",
        "دریافت شاهد: سند بارگذاری‌شده، پاسخ API یا محتوای عمومیِ معرفی‌شده با مهر زمان، شناسه منبع و هش رمزنگاری‌شده ثبت می‌شود. داده خام در خزانه حساس قرار می‌گیرد و در رابط کاربر فقط پیش‌نمایش حداقلی دیده می‌شود.",
        "استخراج و استانداردسازی: OCR و قواعد ساختاری فیلدهایی مانند نام مرجع، تاریخ، عنوان و شناسه سند را استخراج می‌کنند. هر فیلد به ناحیه شاهد متصل است تا کاربر یا کارشناس بتواند آن را با اصل سند تطبیق دهد.",
        "تطبیق و ارزیابی: موتور قواعد ادعا را با شاهد مقایسه می‌کند و مدل زبانی فقط برای خلاصه‌سازی مستند یا پیشنهاد سؤال به‌کار می‌رود. تعارض، کمبود داده و عدم دسترسی سرویس سه وضعیت متفاوت‌اند و نباید در یک برچسب ادغام شوند.",
        "انتشار کنترل‌شده: نتیجه ابتدا در صف کنترل کیفیت یا پنل صاحب داده قرار می‌گیرد. سپس خلاصه متناسب با رضایت به طرف مقابل یا مشاور منتشر و هر مشاهده در لاگ غیرقابل‌انکار ثبت می‌شود.",
    ]
    for s in flow_steps:
        add_number(doc, s)
    add_callout(doc, "اصل زنجیره شواهد", "هر گزاره قابل مشاهده در گزارش باید به یک منبع، زمان دریافت، نسخه پردازش و مسئول تأیید متصل باشد. اگر این زنجیره ناقص است، خروجی نباید به صورت گزاره قطعی منتشر شود.", fill=PALE_BLUE, accent=BLUE)

    add_heading(doc, "۶-۵. چرخه عمر مدل و ارزیابی هوش مصنوعی", 2)
    add_p(doc, "چرخه توسعه مدل از تعریف کارکرد محدود آغاز می‌شود. برای هر کارکرد، داده آزمون مستقل، معیار دقت، خطاهای غیرقابل‌قبول و گروه‌های نیازمند پایش تعیین می‌گردد. مجموعه آزمون باید نمونه‌های واقعیِ ناشناس یا ساختگیِ معتبر از کیفیت‌های مختلف سند، نگارش فارسی، تاریخ شمسی، نام‌های مشابه و پاسخ‌های ناقص API را پوشش دهد.")
    add_p(doc, "مدل پیش از استقرار در سه سطح ارزیابی می‌شود: آزمون فنی برای صحت استخراج و انتساب، آزمون تخصصی با حضور حقوقی و مشاور خانواده برای مناسب‌بودن زبان خروجی، و آزمون کاربر برای سنجش فهم، اعتماد و احتمال اتکای بیش از حد. هر نسخه‌ای که حداقل معیارها را نداشته باشد در محیط آزمایش باقی می‌ماند.")
    add_p(doc, "پس از انتشار، نمونه‌برداری منظم از خروجی‌ها، پایش نرخ اعتراض و مقایسه خطا میان گروه‌های مختلف انجام می‌شود. تغییر مدل پایه، پرامپت اصلی، منبع دانش یا آستانه تصمیم یک تغییر کنترل‌شده است و باید گزارش اثر، برنامه بازگشت و تأیید مالک محصول داشته باشد. داده پرونده‌های کاربران بدون مبنای صریح برای آموزش عمومی مدل استفاده نمی‌شود.")
    add_table(doc, ["مرحله", "شاهد لازم", "معیار عبور", "مالک تأیید"], [
        ["تعریف کارکرد", "شرح دامنه و خطاهای ممنوع", "تأیید محصول، حقوقی و متخصص حوزه", "هیئت راهبری"],
        ["ساخت نمونه", "کارت مدل و داده آزمون", "پوشش سناریوهای بحرانی", "رهبر هوش مصنوعی"],
        ["ارزیابی آفلاین", "گزارش دقت، سوگیری و امنیت", "عبور از آستانه مصوب", "QA و امنیت"],
        ["پایلوت محدود", "لاگ، بازخورد و اعتراض", "عدم رخداد بحرانی و اثربخشی قابل قبول", "مالک محصول"],
        ["انتشار", "نسخه، برنامه بازگشت و راهنمای کاربر", "تصویب تغییر و آمادگی پشتیبانی", "کمیته انتشار"],
        ["پایش", "داشبورد کیفیت و گزارش دوره‌ای", "اقدام اصلاحی در مهلت تعیین‌شده", "عملیات مدل"],
    ], [3.1, 5.4, 5.1, 3.05], font_size=8.5)

    add_heading(doc, "۶-۶. الزامات غیرکارکردی و سطح خدمت", 2)
    add_p(doc, "علاوه بر قابلیت‌های محصول، کیفیت عملیاتی برای اعتماد کاربران تعیین‌کننده است. مقادیر زیر نقطه شروع مذاکره‌اند و باید با ظرفیت زیرساخت، تعداد پرونده هم‌زمان و سطح حساسیت اتصال‌ها نهایی شوند.")
    add_table(doc, ["حوزه", "هدف پیشنهادی پایلوت", "روش سنجش"], [
        ["دسترس‌پذیری", "حداقل ۹۹٫۵٪ برای هسته خدمت", "پایش ماهانه خارج از توقف برنامه‌ریزی‌شده"],
        ["کارایی", "پاسخ صفحات اصلی کمتر از ۳ ثانیه", "صدک ۹۵ در بار پایلوت"],
        ["زمان استعلام", "طبق SLA هر مرجع؛ نمایش وضعیت انتظار", "ثبت زمان درخواست و پاسخ"],
        ["بازیابی", "RPO حداکثر ۲۴ ساعت و RTO حداکثر ۸ ساعت", "آزمون بازیابی دوره‌ای"],
        ["پشتیبانی", "پاسخ اولیه رخداد بحرانی کمتر از ۳۰ دقیقه", "سامانه تیکت و گزارش رخداد"],
        ["مقیاس", "پشتیبانی از ۱۰۰۰ پرونده پایلوت بدون افت محسوس", "آزمون بار و ظرفیت"],
        ["قابلیت ممیزی", "ثبت ۱۰۰٪ دسترسی‌های داده حساس", "نمونه‌برداری و تطبیق لاگ"],
    ], [3.4, 7.2, 6.05], font_size=8.7)

    add_heading(doc, "۷. حاکمیت داده، حریم خصوصی و الزامات حقوقی", 1)
    add_heading(doc, "۷-۱. مبنای حقوقی و قراردادی", 2)
    add_p(doc, "قوانین و مقررات ایران، اطلاعات شخصی و حریم خصوصی را از اطلاعات عمومی تفکیک می‌کنند. قانون مدیریت داده‌ها و اطلاعات ملی نیز تعامل‌پذیری دستگاه‌ها را در چارچوب تکالیف و سازوکارهای رسمی دنبال می‌کند. بنابراین، نام پروژه یا رضایت کلی کاربر به‌تنهایی مجوز دسترسی به هر پایگاه داده نیست؛ برای هر اتصال باید مالک داده، مبنای پردازش، سطح پاسخ، نقش طرفین و مسئولیت رخداد روشن شود.")
    add_p(doc, "پیش از هر اتصال، یک «برگه مجوز و مسئولیت داده» تهیه می‌شود که در آن هدف پردازش، فیلدهای مجاز، دوره نگهداشت، محل میزبانی، گیرندگان، زمان پاسخ، مسئول اصلاح خطا و سناریوی رخداد امنیتی مشخص است. این برگه به‌همراه ارزیابی اثرات حریم خصوصی، ضمیمه قرارداد یا تفاهم‌نامه با مالک داده خواهد بود.")
    add_callout(doc, "هشدار حقوقی", "این بخش تحلیل اولیه برای طراحی محصول است و جایگزین نظر رسمی اداره حقوقی کارفرما، مالک پایگاه داده یا مرجع قضایی نیست. پیش از پایلوت واقعی، ارزیابی اثرات حریم خصوصی و تأیید کتبی دسترسی‌ها الزامی است.", fill=PALE_RED, accent=RED)
    add_heading(doc, "۷-۲. سیاست پیشنهادی چرخه عمر داده", 2)
    add_table(doc, ["مرحله", "قاعده پیشنهادی", "شاهد ممیزی"], [
        ["جمع‌آوری", "هدف مشخص، داده حداقلی، رضایت دانه‌ریز", "نسخه رضایت و شناسه منبع"],
        ["استفاده", "فقط برای پرونده و گیرنده تعیین‌شده", "لاگ دسترسی و هدف"],
        ["اشتراک", "خلاصه حداقلی؛ نه داده خام مگر ضرورت", "رسید اشتراک و دامنه فیلدها"],
        ["نگهداشت", "دوره کوتاه پیش‌فرض و تمدید با انتخاب کاربر", "برچسب انقضا و سیاست حذف"],
        ["اعتراض", "توقف اثر نتیجه تا بررسی موارد پراثر", "تیکت، پاسخ و نسخه اصلاح‌شده"],
        ["حذف/آرشیو", "حذف امن یا نگهداشت قانونی تفکیک‌شده", "گواهی عملیات و استثنای قانونی"],
    ], [3.3, 7.2, 6.15], font_size=9.0)
    add_heading(doc, "۷-۳. طبقه‌بندی داده", 2)
    for x in [
        "بسیار حساس: داده قضایی، سلامت، مشاوره، شناسه‌های هویتی و پاسخ‌های خصوصی.",
        "حساس: تحصیل، اشتغال، درآمد خوداظهاری‌شده و نتایج راستی‌آزمایی.",
        "محدود: محتوای عمومی وب که برای پرونده پردازش شده و می‌تواند پروفایل ایجاد کند.",
        "عمومی/عملیاتی: محتوای راهنما، فهرست خدمات و داده‌های ناشناس آماری.",
    ]:
        add_bullet(doc, x)
    add_heading(doc, "۷-۴. ممنوعیت‌های صریح", 2)
    for x in [
        "خرید یا دریافت بانک اطلاعاتی غیررسمی، جمع‌آوری پنهانی داده از حساب خصوصی، دورزدن محدودیت فنی شبکه اجتماعی یا استفاده از رمز/نشست کاربری اشخاص.",
        "استنباط یا امتیازدهی مذهب، قومیت، گرایش سیاسی، سلامت، گرایش جنسی یا سایر صفات حساس.",
        "استفاده ثانویه از داده پرونده ازدواج برای استخدام، تبلیغات، اعتبارسنجی مالی یا تنبیه اداری.",
        "تصمیم کاملاً خودکار با اثر مهم بر فرد، بدون توضیح، اعتراض و مداخله انسانی.",
    ]:
        add_bullet(doc, x, color=RED)

    add_heading(doc, "۷-۵. فرایند ارزیابی اثرات حریم خصوصی", 2)
    add_p(doc, "ارزیابی اثرات حریم خصوصی پیش از پایلوت و پیش از هر تغییر عمده انجام می‌شود. این ارزیابی از توصیف جریان داده آغاز و مشخص می‌کند چه داده‌ای، از چه فردی، برای چه هدفی، توسط چه نقشی و تا چه زمانی پردازش می‌شود. سپس ضرورت و تناسب هر فیلد بررسی می‌گردد؛ اگر هدف با داده کمتر، پاسخ بله/خیر یا مشاهده گواهی بدون ذخیره نسخه قابل تحقق باشد، گزینه کم‌داده باید انتخاب شود.")
    add_p(doc, "در مرحله بعد، پیامدهای احتمالی برای فرد ارزیابی می‌شود: افشای ناخواسته، قضاوت نادرست، فشار برای رضایت، دشواری اعتراض، تبعیض، استفاده ثانویه و آسیب ناشی از نقض امنیت. برای هر خطر، احتمال و شدت قبل و بعد از کنترل‌ها ثبت می‌شود. خطر باقیمانده بحرانی باید به هیئت راهبری ارجاع شود و تا پذیرش رسمی یا تغییر طراحی، قابلیت مربوطه اجازه ورود به پایلوت ندارد.")
    add_p(doc, "خروجی ارزیابی یک سند ایستا نیست. رخداد امنیتی، افزایش نرخ اعتراض، افزودن منبع داده، تغییر مدل یا توسعه جامعه هدف می‌تواند ارزیابی را بی‌اعتبار کند. بنابراین نسخه DPIA، تصمیمات، استثناها و تاریخ بازبینی در پنل حاکمیت ثبت می‌شود و خلاصه قابل فهمی از کنترل‌های اصلی در اختیار کاربران قرار می‌گیرد.")

    add_heading(doc, "۷-۶. طراحی رضایت آگاهانه و بدون اجبار", 2)
    for x in [
        "تفکیک رضایت‌ها: هویت، تحصیل، اشتغال، گواهی قضایی، وب عمومی، اشتراک با مشاور و خدمت پس از ازدواج هرکدام انتخاب مستقل داشته باشند.",
        "شفافیت پیش از اقدام: قبل از تأیید، نمونه خروجی، فیلدهای قابل مشاهده، مدت اعتبار و پیامد لغو نمایش داده شود.",
        "تقارن: هر دو طرف قواعد یکسانی برای ارائه، مشاهده، اعتراض و اصلاح داشته باشند؛ دسترسی یک‌طرفه پیش‌فرض ایجاد نشود.",
        "امکان بازگشت: لغو رضایت ساده و در همان سطح دسترسی باشد؛ اثر آن بر داده ذخیره‌شده و گزارش‌های قبلی به‌روشنی توضیح داده شود.",
        "مدیریت فشار: سامانه عدم پذیرش یک بسته اختیاری را به عنوان نشانه منفی، امتیاز پایین یا هشدار به طرف مقابل نمایش ندهد.",
        "ثبت شاهد رضایت: نسخه متن، زمان، کانال، هویت تأییدشده و محدوده انتخاب در سابقه ممیزی نگهداری شود.",
    ]:
        add_bullet(doc, x)
    add_callout(doc, "قاعده طراحی", "رضایت باید یک انتخاب واقعی و قابل فهم باشد، نه پذیرش یک متن طولانی برای دسترسی به خدمت. بسته‌های ضروری و اختیاری، پیامد رد و راه لغو باید جدا و روشن نمایش داده شوند.", fill=PALE_BLUE, accent=BLUE)

    add_heading(doc, "۷-۷. پیشنهاد دوره نگهداشت و حذف", 2)
    add_table(doc, ["نوع داده", "دوره پایه پیشنهادی", "رویداد حذف/بازبینی", "استثنا"], [
        ["دعوت ردشده", "حداکثر ۳۰ روز", "انقضای دعوت یا درخواست کاربر", "لاگ امنیتی حداقلی"],
        ["مدارک خام", "تا پایان بررسی + ۹۰ روز", "بستن پرونده یا لغو بسته", "دستور قانونی یا اعتراض باز"],
        ["خلاصه گزارش", "۶ ماه پس از بستن پرونده", "اعلان و انتخاب حذف/تمدید", "پرونده اختلاف باز"],
        ["رضایت و لاگ دسترسی", "طبق نیاز ممیزی و نظر حقوقی", "بازبینی دوره‌ای ضرورت", "الزام قانونی نگهداشت"],
        ["داده مشاوره", "مستقل و طبق سیاست حرفه‌ای", "خاتمه خدمت و نظر متخصص", "الزام حرفه‌ای/قانونی"],
        ["پروفایل شغلی", "تا زمان فعال‌بودن انتخاب کاربر", "لغو معرفی یا ۱۲ ماه عدم فعالیت", "درخواست در حال رسیدگی"],
    ], [3.3, 4.4, 5.6, 3.35], font_size=8.5)

    add_heading(doc, "۸. امنیت و تاب‌آوری", 1)
    add_p(doc, "ماهیت داده‌ها ایجاب می‌کند امنیت از آغاز معماری شود، نه در پایان پروژه. چارچوب کنترل باید محرمانگی، صحت و دسترس‌پذیری را پوشش دهد و برای زنجیره شواهد، ثبت منشأ و تمامیت داده نیز سازوکار مشخص داشته باشد.")
    add_heading(doc, "۸-۱. کنترل‌های پایه", 2)
    for x in [
        "رمزنگاری داده در انتقال و در حالت سکون؛ مدیریت کلید مستقل و چرخش دوره‌ای.",
        "کنترل دسترسی مبتنی بر نقش و ویژگی، اصل کمترین دسترسی و تأیید دومرحله‌ای برای نقش‌های حساس.",
        "تفکیک محیط توسعه، آزمون و تولید؛ ممنوعیت استفاده از داده واقعی در توسعه مگر با ناشناس‌سازی مصوب.",
        "ثبت غیرقابل‌انکار دسترسی‌ها، پایش رفتار غیرعادی و هشدار بر مشاهده انبوه یا نامتناسب.",
        "آزمون نفوذ، تحلیل کد، مدیریت آسیب‌پذیری زنجیره تأمین و برنامه پاسخ به رخداد.",
        "پشتیبان‌گیری رمزنگاری‌شده، سناریوی بازیابی و آزمون دوره‌ای تداوم خدمت.",
        "دروازه امنیتی برای ورودی مدل، جلوگیری از نشت داده در پرامپت و ارزیابی حملات تزریق دستور.",
    ]:
        add_bullet(doc, x)
    add_heading(doc, "۸-۲. کنترل‌های اختصاصی هوش مصنوعی", 2)
    add_table(doc, ["ریسک", "کنترل پیشگیرانه", "کنترل کشف/پاسخ"], [
        ["توهم یا انتساب نادرست", "تولید مبتنی بر منبع و قالب محدود", "نمونه‌برداری کیفی و توقف نمایش نتیجه کم‌اطمینان"],
        ["سوگیری", "حذف ویژگی‌های حساس و سناریوهای آزمون", "پایش اختلاف نرخ خطا و بازبینی مستقل"],
        ["نشت داده", "مدل خصوصی/قرارداد پردازش، کمینه‌سازی ورودی", "پایش خروجی و رخداد محرمانگی"],
        ["تزریق دستور", "جداسازی محتوا از دستور و فیلتر منبع", "ثبت حمله، قرنطینه سند و اصلاح قاعده"],
        ["تغییر رفتار مدل", "قفل نسخه و ارزیابی قبل از انتشار", "بازگشت نسخه و گزارش تغییر"],
    ], [4.1, 6.25, 6.3], font_size=8.7)

    add_heading(doc, "۸-۳. مدیریت رخداد امنیتی و نقض داده", 2)
    add_p(doc, "برنامه پاسخ به رخداد باید پیش از ورود داده واقعی تصویب و تمرین شود. رخدادها در چهار سطح طبقه‌بندی می‌شوند: اختلال عملیاتی، دسترسی مشکوک، افشای داده حساس و رخداد بحرانی با اثر گسترده. برای هر سطح، مسئول فرماندهی، کانال ارتباطی، زمان واکنش، حفظ شواهد و اختیار توقف سرویس مشخص است.")
    add_p(doc, "در رخداد مرتبط با محرمانگی، نخست دسترسی یا جریان آسیب‌دیده مهار و هم‌زمان شواهد فنی بدون دستکاری حفظ می‌شود. سپس دامنه افراد و داده‌های درگیر، زمان شروع، مسیر نفوذ و اثر احتمالی بر گزارش‌ها تعیین می‌گردد. اطلاع‌رسانی به مدیریت، مالک داده و افراد متأثر باید بر پایه پروتکل مصوب و با زبان دقیق انجام شود؛ پنهان‌کردن عدم قطعیت یا ارائه اطمینان بی‌پشتوانه می‌تواند اعتماد را بیشتر آسیب بزند.")
    add_p(doc, "پس از مهار، تحلیل ریشه‌ای و برنامه اصلاح شامل رفع آسیب‌پذیری، تغییر کلید یا مجوز، بازبینی خروجی‌های متأثر و پایش تکمیلی تهیه می‌شود. درس‌آموخته‌ها باید به کنترل‌های فنی، آموزش کارکنان و سناریوهای آزمون بازگردند. بستن رخداد فقط پس از تأیید امنیت، حریم خصوصی و مالک خدمت انجام می‌شود.")
    add_table(doc, ["سطح رخداد", "نمونه", "زمان واکنش اولیه", "تصمیم کلیدی"], [
        ["۱ - عملیاتی", "خطای محدود بدون افشای داده", "تا ۴ ساعت", "اصلاح و پایش"],
        ["۲ - مهم", "دسترسی غیرعادی یا خطای تکرارشونده", "تا ۱ ساعت", "محدودسازی دسترسی"],
        ["۳ - شدید", "افشای داده حساس یا دستکاری گزارش", "تا ۳۰ دقیقه", "توقف بخشی از خدمت و اطلاع داخلی"],
        ["۴ - بحرانی", "افشای گسترده یا تهدید فعال", "فوری", "فعال‌سازی تیم بحران و برنامه اطلاع‌رسانی"],
    ], [3.1, 6.0, 3.45, 4.1], font_size=8.6)

    add_heading(doc, "۸-۴. تداوم خدمت و بازیابی", 2)
    add_p(doc, "سامانه باید بتواند خرابی زیرساخت، قطع اتصال دستگاه همکار، خطای نسخه مدل و از دسترس‌رفتن مرکز داده را بدون از دست‌دادن کنترل رضایت مدیریت کند. هر اتصال خارجی با صف، زمان‌سنج، تلاش مجدد محدود و مدارشکن طراحی می‌شود تا اختلال یک دستگاه کل پرونده را متوقف نکند. وضعیت انتظار و آخرین زمان تلاش به کاربر نمایش داده می‌شود.")
    add_p(doc, "پشتیبان‌گیری رمزنگاری‌شده، نگهداری نسخه در محل تفکیک‌شده و آزمون بازیابی حداقل فصلی پیشنهاد می‌شود. بازگشت از نسخه پشتیبان باید هم داده پرونده و هم وضعیت رضایت، لغوها، کلیدهای معتبر و لاگ‌های ممیزی را سازگار بازیابی کند. سناریوی بازیابی ناقص که رضایت لغوشده را دوباره فعال کند یک خطای بحرانی محسوب می‌شود.")

    add_heading(doc, "۹. نقشه راه اجرا", 1)
    add_heading(doc, "۹-۱. فازها و دروازه‌های تصمیم", 2)
    add_table(doc, ["فاز", "مدت برآوردی", "خروجی کلیدی", "دروازه تصمیم"], [
        ["فاز صفر: امکان‌سنجی", "۶ تا ۸ هفته", "نقشه داده، نظر حقوقی، توافق اولیه API، طراحی خدمت", "Go/No-Go برای هر اتصال"],
        ["فاز ۱: MVP کم‌ریسک", "۱۲ تا ۱۶ هفته", "هویت، رضایت، پرونده، بارگذاری/اعتبارسنجی سند، پرسشنامه", "پذیرش امنیت و تجربه کاربر"],
        ["فاز ۲: پایلوت کنترل‌شده", "۱۰ تا ۱۲ هفته", "پایلوت محدود، پنل مشاور، اعتراض، پایش مدل", "اثربخشی و عدم آسیب"],
        ["فاز ۳: اتصال‌های رسمی", "۴ تا ۹ ماه", "اتصال مرحله‌ای تحصیل/اشتغال/قضاییِ مجاز", "توافق سطح خدمت هر دستگاه"],
        ["فاز ۴: خدمات پس از ازدواج", "۳ تا ۵ ماه", "مشاوره، آموزش و مسیر شغلی جداگانه", "رضایت تازه و جداسازی داده"],
    ], [3.2, 2.8, 6.6, 4.05], font_size=8.7)
    add_heading(doc, "۹-۲. محدوده پیشنهادی MVP", 2)
    for x in [
        "ثبت‌نام، احراز هویت، دعوت طرف مقابل و مدیریت رضایت.",
        "پرونده ادعاها و مدارک با کنترل اصالت پایه و ثبت منشأ.",
        "پشتیبانی از گواهی عدم سوءپیشینه‌ای که خود فرد از مسیر رسمی ارائه می‌کند.",
        "راستی‌آزمایی تحصیل/اشتغال از طریق مدارک و مراجع منتخب پایلوت.",
        "پرسشنامه گفت‌وگوی پیش از ازدواج، خلاصه مستند و پنل مشاور.",
        "اعتراض، اصلاح، حذف، لاگ ممیزی و داشبورد شاخص‌های ناشناس.",
    ]:
        add_bullet(doc, x)
    add_heading(doc, "۹-۳. برنامه پایلوت", 2)
    add_p(doc, "پایلوت باید با جامعه محدود، داوطلب و دارای پشتیبانی مشاور اجرا شود. پیش از ورود، رضایت آگاهانه و معیار خروج روشن ارائه می‌شود. هیچ قابلیت آزمایشی بدون برچسب، هیچ تصمیم خودکار و هیچ اتصال تولیدیِ تأییدنشده در پایلوت فعال نمی‌شود. ارزیابی شامل دقت راستی‌آزمایی، تجربه کاربر، زمان پاسخ به اعتراض، رخدادهای حریم خصوصی و آثار ناخواسته است.")
    add_heading(doc, "۹-۴. وابستگی‌های اجرایی و تعهدات کارفرما", 2)
    for x in [
        "معرفی نماینده تام‌الاختیار محصول، حقوقی، امنیت و تعامل‌پذیری برای تصمیم‌گیری هفتگی.",
        "فراهم‌کردن محیط آزمون، داده ساختگی/ناشناس و دسترسی مستند به APIهای مورد توافق.",
        "تأمین شبکه مشاوران و مسئول پاسخ‌گویی به اعتراض‌ها و رخدادهای مرتبط با کاربران پایلوت.",
        "تصویب سیاست حریم خصوصی، دوره نگهداشت، معیارهای توقف پایلوت و برنامه اطلاع‌رسانی رخداد.",
    ]:
        add_bullet(doc, x)

    add_heading(doc, "۹-۵. برنامه کاری دوازده‌ماهه پیشنهادی", 2)
    add_table(doc, ["بازه", "محور فعالیت", "تحویل قابل پذیرش", "وابستگی"], [
        ["ماه ۱ تا ۲", "فاز صفر، حقوقی، داده و طراحی خدمت", "تصمیم اتصال‌ها، دامنه MVP و معماری مصوب", "نمایندگان کارفرما و مالکان داده"],
        ["ماه ۳", "نمونه تجربه کاربر و زیرساخت پایه", "نمونه آزمون‌شده، مدل تهدید و محیط توسعه", "داده ساختگی و زیرساخت"],
        ["ماه ۴ تا ۵", "هویت، رضایت، پرونده و اسناد", "نسخه آزمایشی هسته و پنل مدیریت", "ارائه‌دهنده احراز"],
        ["ماه ۶", "هوش مصنوعی و گزارش مستند", "OCR، تطبیق، خلاصه و مجموعه ارزیابی", "نمونه اسناد و تأیید متخصص"],
        ["ماه ۷", "امنیت، حریم خصوصی و آمادگی پایلوت", "آزمون نفوذ، DPIA و راهنمای عملیات", "تصویب سیاست‌ها"],
        ["ماه ۸ تا ۹", "پایلوت موج اول", "۱۰۰ تا ۲۰۰ پرونده، داشبورد کیفیت و گزارش اولیه", "مشاوران و پشتیبانی"],
        ["ماه ۱۰", "اصلاح و پایلوت موج دوم", "رفع خطا، بهبود رضایت و گسترش کنترل‌شده", "تصمیم دروازه پایلوت"],
        ["ماه ۱۱", "اتصال رسمی منتخب", "حداقل یک اتصال تولیدی با SLA", "تفاهم‌نامه و محیط تولید"],
        ["ماه ۱۲", "جمع‌بندی و برنامه مقیاس", "گزارش اثربخشی، هزینه واقعی و نقشه نسخه بعد", "تصمیم هیئت راهبری"],
    ], [2.5, 4.8, 6.4, 2.95], font_size=8.25)
    add_p(doc, "این برنامه از نظر تقویمی محافظه‌کارانه است و فرض می‌کند تصمیمات حقوقی و دسترسی به محیط آزمون در زمان مقرر انجام شود. تأخیر در تفاهم‌نامه‌های بین‌دستگاهی می‌تواند اتصال رسمی را عقب بیندازد، اما نباید توسعه هسته رضایت، پرونده، سند و مشاوره را متوقف کند؛ به همین دلیل معماری و قرارداد به‌صورت بسته‌های مستقل طراحی شده‌اند.")

    add_heading(doc, "۹-۶. ریتم مدیریت و گزارش‌دهی", 2)
    for x in [
        "جلسه هفتگی محصول: تصمیم درباره دامنه، موانع، تجربه کاربر و اولویت تحویل با ثبت صورت‌جلسه.",
        "کمیته دوهفته‌ای داده و حقوق: بررسی اتصال‌ها، رضایت، اعتراض‌های مهم و تغییرات سیاست نگهداشت.",
        "بازبینی ماهانه امنیت و مدل: آسیب‌پذیری‌ها، کیفیت خروجی، سوگیری، رخدادها و برنامه اصلاح.",
        "دروازه پایان هر فاز: نمایش محصول، گزارش معیارها، وضعیت ریسک و تصمیم کتبی ادامه/اصلاح/توقف.",
        "گزارش فصلی هیئت راهبری: هزینه، زمان، اثربخشی، شکایات و انطباق با مأموریت مصوب.",
    ]:
        add_bullet(doc, x)

    add_heading(doc, "۱۰. تیم، اقلام تحویلی و مدل هزینه", 1)
    add_heading(doc, "۱۰-۱. تیم حداقلی پیشنهادی", 2)
    add_table(doc, ["نقش", "مسئولیت اصلی", "مشارکت"], [
        ["مدیر محصول/پروژه", "دامنه، ذی‌نفعان، برنامه و پذیرش", "تمام‌وقت"],
        ["معمار راهکار", "معماری، تعامل‌پذیری و تصمیمات فنی", "تمام‌وقت در طراحی"],
        ["متخصص حقوق و حریم خصوصی", "مبنای پردازش، رضایت، قرارداد و ارزیابی اثر", "مستمر"],
        ["متخصص خانواده/روان‌شناسی", "پرسشنامه، مرز توصیه و ارجاع", "مستمر"],
        ["توسعه‌دهندگان بک‌اند/فرانت‌اند", "محصول، API و پنل‌ها", "تمام‌وقت"],
        ["مهندس داده و هوش مصنوعی", "استخراج، تطبیق، RAG و ارزیابی", "تمام‌وقت"],
        ["امنیت و DevSecOps", "کنترل امنیتی، زیرساخت و پایش", "مستمر"],
        ["طراح UX و پژوهشگر", "سفر، زبان رضایت و آزمون کاربردپذیری", "پاره‌وقت/موجی"],
        ["QA و ارزیابی مدل", "آزمون کارکرد، داده، سوگیری و رگرسیون", "تمام‌وقت در تحویل"],
    ], [4.0, 8.4, 4.25], font_size=8.8)
    add_heading(doc, "۱۰-۲. اقلام تحویلی", 2)
    for x in [
        "سند نیازمندی‌ها، نقشه سفر و مدل عملیاتی خدمت.",
        "گزارش امکان‌سنجی حقوقی و تعامل‌پذیری برای هر پایگاه داده.",
        "معماری مرجع، مدل داده، طبقه‌بندی اطلاعات و مدل تهدید.",
        "MVP وب/موبایل، پنل مشاور و پنل حاکمیت با مستندات استقرار.",
        "مدل/خدمات هوش مصنوعی به‌همراه کارت مدل، مجموعه آزمون و گزارش ارزیابی.",
        "سیاست‌های رضایت، نگهداشت، اعتراض، رخداد امنیتی و بهره‌برداری.",
        "گزارش پایلوت، درس‌آموخته‌ها و برنامه مقیاس‌پذیری.",
    ]:
        add_bullet(doc, x)
    add_heading(doc, "۱۰-۳. مدل برآورد هزینه", 2)
    add_p(doc, "برآورد زیر یک سناریوی بودجه‌گذاری اولیه برای پایلوت سازمانی با ۵۰۰ تا ۱۰۰۰ پرونده، استقرار خصوصی، حداکثر سه اتصال رسمی و الزامات امنیتی داده حساس است. ارقام قطعی پس از فاز صفر، دریافت مستندات API و تعیین سطح خدمت در پیشنهاد مالی تفصیلی تثبیت می‌شوند.")
    add_table(doc, ["بسته کاری", "شرح مختصر", "حداقل", "حداکثر"], [
        ["فاز صفر و معماری", "حقوقی، نیازسنجی، DPIA، معماری و توافق داده", "۱٫۵", "۲٫۲"],
        ["تجربه کاربر", "پژوهش، طراحی خدمت، نمونه‌سازی و آزمون", "۱٫۲", "۱٫۸"],
        ["هسته سامانه", "هویت، رضایت، پرونده، گردش‌کار و پنل‌ها", "۷٫۰", "۱۰٫۰"],
        ["هوش مصنوعی", "OCR، RAG، تطبیق، خلاصه‌سازی و ارزیابی", "۳٫۵", "۵٫۵"],
        ["امنیت و حریم خصوصی", "DevSecOps، رمزنگاری، ممیزی و آزمون نفوذ", "۲٫۰", "۳٫۵"],
        ["پایلوت و آموزش", "استقرار محدود، آموزش، پشتیبانی و ارزیابی", "۲٫۰", "۳٫۰"],
        ["اتصال‌های رسمی", "طراحی و پیاده‌سازی حداکثر سه اتصال", "۴٫۵", "۸٫۰"],
        ["خدمات پس از ازدواج", "مشاوره، توانمندسازی و مسیر شغلی مستقل", "۳٫۰", "۵٫۰"],
        ["مدیریت و انتقال دانش", "مدیریت پروژه، مستندسازی و تحویل", "۱٫۵", "۲٫۵"],
        ["جمع پایه", "مجموع هزینه ساخت و پایلوت", "۲۶٫۲", "۴۱٫۵"],
        ["ذخیره ریسک ۱۵٪", "تغییر دامنه، تأخیر API و الزامات پیش‌بینی‌نشده", "۳٫۹", "۶٫۲"],
        ["جمع کل پیشنهادی", "بودجه اولیه اجرای سناریوی پایلوت سازمانی", "۳۰٫۱", "۴۷٫۷"],
    ], [4.0, 7.1, 2.75, 2.8], font_size=8.4)
    add_callout(doc, "واحد و دامنه برآورد", "تمام ارقام جدول به میلیارد تومان و بدون مالیات بر ارزش افزوده است. تعرفه مصرف API دستگاه‌ها، سخت‌افزار اختصاصی، مجوز نرم‌افزار، پیامک/احراز و ممیزی مستقل در صورت نیاز جداگانه محاسبه می‌شود. هزینه بهره‌برداری و پشتیبانی سالانه نیز حدود ۴٫۵ تا ۸ میلیارد تومان برآورد اولیه دارد.", fill=PALE_BLUE, accent=BLUE)
    add_p(doc, "مبنای روش برآورد، شکستن پروژه به بسته‌های کاری و انطباق با تعرفه‌های خدمات فنی-تخصصی انفورماتیک است؛ ضرایب و نرخ نفر-ماه باید در فاز صفر با آخرین نسخه تعرفه و استعلام واقعی تأمین‌کنندگان به‌روز شود.", size=9.8, color=MUTED, italic=True)

    add_heading(doc, "۱۰-۴. فرض‌های مالی، پرداخت و کنترل تغییر", 2)
    add_p(doc, "برآورد ارائه‌شده بر مبنای تیم چندتخصصی، توسعه تدریجی و پایلوت محدود است. اگر کارفرما استقرار کاملاً درون‌سازمانی، سخت‌افزار شتاب‌دهنده هوش مصنوعی، مرکز عملیات امنیت اختصاصی، پشتیبانی شبانه‌روزی یا تعداد اتصال بیشتری بخواهد، هزینه زیرساخت و عملیات باید جداگانه افزوده شود. در مقابل، استفاده از زیرساخت موجود کارفرما و محدودکردن دامنه MVP می‌تواند هزینه را کاهش دهد.")
    add_p(doc, "پیشنهاد می‌شود پرداخت هر بسته به تحویل قابل سنجش وابسته باشد، نه صرفاً گذشت زمان. برای نمونه، فاز صفر پس از تحویل نقشه داده، نظر حقوقی، معماری و برنامه MVP تسویه می‌شود؛ هسته محصول پس از آزمون پذیرش؛ و پایلوت پس از ارائه گزارش شاخص‌ها و رفع خطاهای بحرانی. بخشی از مبلغ هر فاز می‌تواند تا تأیید مستندات و انتقال دانش نگه داشته شود.")
    add_table(doc, ["نقطه پرداخت", "سهم راهنما", "شاهد پذیرش"], [
        ["شروع و تجهیز تیم", "۱۰٪", "برنامه تفصیلی، تیم و محیط‌های کاری"],
        ["پایان فاز صفر", "۱۵٪", "بسته تصمیم حقوقی/داده و دامنه مصوب"],
        ["تحویل نسخه آلفا", "۲۰٪", "هویت، رضایت، پرونده و طراحی گزارش"],
        ["تحویل MVP امن", "۲۵٪", "آزمون کارکرد، امنیت و مستندات استقرار"],
        ["پایان پایلوت", "۲۰٪", "گزارش شاخص‌ها و رفع خطاهای بحرانی"],
        ["انتقال دانش و پذیرش نهایی", "۱۰٪", "مستندات، آموزش و صورت‌جلسه تحویل"],
    ], [4.1, 3.1, 9.45], font_size=8.7)
    add_p(doc, "هر درخواست تغییر باید اثر خود بر دامنه، زمان، هزینه، ریسک حریم خصوصی و مجموعه آزمون را نشان دهد. تغییرهای دارای منبع داده تازه یا اثر مهم بر فرد، حتی اگر از نظر توسعه کوچک باشند، بدون بازبینی حقوقی و امنیتی به بک‌لاگ اجرایی وارد نمی‌شوند.")

    add_heading(doc, "۱۱. شاخص‌های موفقیت و معیار پذیرش", 1)
    add_heading(doc, "۱۱-۱. شاخص‌های محصول و خدمت", 2)
    add_table(doc, ["شاخص", "تعریف", "هدف اولیه پیشنهادی"], [
        ["نرخ تکمیل رضایت", "درصد پرونده‌های دارای رضایت معتبر هر دو طرف", "بیش از ۸۵٪ در پایلوت"],
        ["نرخ راستی‌آزمایی", "سهم ادعاهای منتخب با شاهد معتبر", "بیش از ۷۰٪ در دامنه پایلوت"],
        ["دقت استخراج سند", "انطباق فیلد با تأیید انسانی", "حداقل ۹۵٪ برای فیلدهای بحرانی"],
        ["نرخ انتساب معتبر", "خلاصه‌های دارای منبع قابل مشاهده", "۱۰۰٪ خروجی‌های پراثر"],
        ["زمان پاسخ اعتراض", "از ثبت تا تصمیم انسانی", "حداکثر ۳ روز کاری"],
        ["رخداد افشای شدید", "افشای داده بسیار حساس", "صفر"],
        ["رضایت کاربر", "امتیاز تجربه و احساس کنترل", "حداقل ۴ از ۵"],
        ["ارجاع موفق", "تکمیل جلسه مشاوره پس از ارجاع", "خط مبنا در پایلوت"],
    ], [4.0, 8.0, 4.65], font_size=8.8)
    add_heading(doc, "۱۱-۲. معیارهای پذیرش فنی", 2)
    for x in [
        "همه دسترسی‌ها دارای شناسه کاربر، هدف، زمان و نتیجه باشند و قابل گزارش‌گیری شوند.",
        "لغو رضایت، دسترسی‌های آتی را بلافاصله متوقف و وضعیت داده موجود را طبق سیاست تعیین کند.",
        "گزارش هیچ نتیجه پراثری را بدون منبع، تاریخ و سطح اطمینان نمایش ندهد.",
        "آزمون نفوذ بحرانی حل‌نشده، دسترسی پیش‌فرض گسترده یا داده واقعی در لاگ کاربردی وجود نداشته باشد.",
        "مدل در مجموعه آزمون مصوب، حد دقت و تفاوت نرخ خطای گروه‌ها را رعایت کند.",
        "سناریوی اصلاح داده، اعتراض، بستن پرونده و بازیابی از پشتیبان با موفقیت آزمون شود.",
    ]:
        add_bullet(doc, x)

    add_heading(doc, "۱۱-۳. روش ارزیابی اثربخشی پایلوت", 2)
    add_p(doc, "ارزیابی پایلوت باید فراتر از تعداد ثبت‌نام باشد. در آغاز، خط مبنا برای زمان بررسی مدارک، میزان فهم کاربران از رضایت، نرخ اختلاف قابل حل و کیفیت گفت‌وگو ثبت می‌شود. سپس همان شاخص‌ها در پایان پایلوت اندازه‌گیری و با خط مبنا مقایسه می‌گردد. داده‌های آماری تا حد امکان ناشناس یا تجمیعی گزارش می‌شوند و از تبدیل شاخص عملکرد به ابزار ارزیابی افراد پرهیز می‌شود.")
    add_p(doc, "روش پیشنهادی ترکیبی است: تحلیل کمی لاگ و زمان فرایند؛ آزمون دقت روی نمونه برچسب‌خورده؛ مصاحبه نیمه‌ساختاریافته با کاربران و مشاوران؛ و مرور کیفی همه اعتراض‌ها و موارد پراثر. گزارش نهایی باید هم موفقیت و هم آثار ناخواسته را نشان دهد. برای مثال کاهش زمان بررسی اگر با افت فهم رضایت یا افزایش اتکای بیش از حد همراه باشد، موفقیت کامل محسوب نمی‌شود.")
    add_p(doc, "معیار ادامه پس از پایلوت شامل عبور از آستانه‌های امنیت و دقت، رضایت قابل قبول کاربران، توان پاسخ‌گویی به اعتراض در مهلت تعیین‌شده و نبود شواهد آسیب نظام‌مند است. اگر یک ماژول این شرایط را برآورده نکند، می‌توان همان ماژول را متوقف یا محدود کرد، بدون آنکه کل محصول کنار گذاشته شود.")
    add_callout(doc, "اصل ارزیابی", "هدف پایلوت اثبات امکان و ایمنی خدمت در دامنه محدود است، نه اثبات قطعی کاهش طلاق. سنجش اثر اجتماعی بلندمدت نیازمند طراحی پژوهشی مستقل، دوره پیگیری طولانی‌تر و کنترل عوامل مداخله‌گر است.", fill=PALE_RED, accent=RED)

    add_heading(doc, "۱۲. ریسک‌ها و کنترل‌ها", 1)
    add_table(doc, ["ریسک", "شدت", "کنترل پیشنهادی", "مالک"], [
        ["عدم دسترسی به API رسمی", "زیاد", "فاز صفر، گزینه گواهی کاربر و دامنه انعطاف‌پذیر", "مدیر برنامه"],
        ["نقض حریم خصوصی", "بحرانی", "حداقل‌گرایی، رضایت، تفکیک داده و ممیزی", "مسئول حریم خصوصی"],
        ["انتساب حساب به فرد اشتباه", "زیاد", "تأیید صاحب حساب، چند شاهد و حق اعتراض", "مالک محصول"],
        ["توهم یا خلاصه نادرست", "زیاد", "RAG، نمایش شاهد، حد اطمینان و بازبینی انسانی", "رهبر AI"],
        ["سوگیری و برچسب‌زنی", "بحرانی", "ممنوعیت صفات حساس، آزمون و کمیته نظارت", "کمیته حاکمیت"],
        ["نشت داخلی داده", "بحرانی", "کمترین دسترسی، DLP، لاگ و پاسخ به رخداد", "امنیت"],
        ["اتکای بیش از حد کاربر", "متوسط", "زبان غیرقطعی، آموزش و ارجاع مشاور", "تجربه کاربر"],
        ["گسترش بی‌ضابطه دامنه", "زیاد", "دروازه تغییر، DPIA و تصویب مالک داده", "هیئت راهبری"],
        ["استفاده ثانویه برای استخدام", "بحرانی", "جداسازی حقوقی/فنی و رضایت تازه", "کارفرما"],
    ], [4.0, 2.2, 7.1, 3.35], font_size=8.4)
    add_heading(doc, "مدل حاکمیت پیشنهادی", 2)
    add_p(doc, "هیئت راهبری شامل نماینده کارفرما، حقوقی، امنیت، متخصص خانواده، محصول و نماینده کاربران باشد. تغییر در منبع داده، هدف استفاده، ویژگی پراثر یا مدل اصلی بدون ثبت ارزیابی اثر و تصویب هیئت راهبری وارد تولید نشود. گزارش دوره‌ای باید کیفیت، شکایات، سوگیری، رخدادهای امنیتی و درخواست‌های حذف را پوشش دهد.")

    add_heading(doc, "۱۲-۱. تحلیل چند سناریوی پرخطر", 2)
    add_heading(doc, "انتساب اشتباه داده به فرد هم‌نام", 3)
    add_p(doc, "جست‌وجوی نام در وب یا پایگاه‌های فاقد شناسه یکتا می‌تواند اطلاعات یک فرد را به فرد دیگری نسبت دهد. کنترل پیشنهادی، منع تولید نتیجه بر پایه نام به‌تنهایی، تأیید حساب توسط صاحب آن، استفاده از چند شاهد مستقل و نمایش مرحله پیش‌نمایش به صاحب داده است. هر اعتراض انتساب باید فوراً اثر نتیجه را متوقف کند تا بررسی انسانی تکمیل شود.")
    add_heading(doc, "فشار اجتماعی برای پذیرش رضایت", 3)
    add_p(doc, "ممکن است یک طرف، رد بسته اختیاری را نشانه پنهان‌کاری تلقی کند. طراحی محصول باید از نمایش پیام‌های قضاوتی خودداری و به‌روشنی بیان کند که رضایت اختیاری است. مشاوران نیز باید آموزش ببینند که عدم انتخاب یک بسته به‌تنهایی مبنای نتیجه‌گیری نیست. در صورت مشاهده الگوی اجبار یا سوءاستفاده، سامانه باید امکان خروج ایمن و ارتباط محرمانه با پشتیبانی را فراهم کند.")
    add_heading(doc, "برداشت قطعی از تحلیل وب", 3)
    add_p(doc, "محتوای آنلاین گزینشی، وابسته به زمان و گاهی نمایشی است. حتی تحلیل درست موضوعات یک حساب عمومی نمی‌تواند ارزش‌ها، رفتار خصوصی یا آینده رابطه را به‌طور قطعی نشان دهد. کنترل محصول شامل محدودکردن بازه زمانی، ذکر نمونه، نمایش پوشش ناقص، منع صفات حساس و تبدیل خروجی به سؤال گفت‌وگو است. اگر شواهد کافی نیست، بهترین خروجی «عدم امکان جمع‌بندی» خواهد بود.")
    add_heading(doc, "استفاده ثانویه برای استخدام یا ارزیابی اداری", 3)
    add_p(doc, "ارزش داده‌های این سامانه برای اهداف دیگر می‌تواند انگیزه دسترسی نامتناسب ایجاد کند. جداسازی حقوقی و فنی پایگاه‌ها، نقش‌های مستقل، ممنوعیت صادرات انبوه، واترمارک گزارش‌های حساس و پایش دسترسی کارکنان ضروری است. هر خدمت شغلی باید پروفایل تازه و رضایت تازه داشته باشد و دریافت‌کننده هیچ نشانه‌ای از محتوای پرونده ازدواج نبیند.")

    add_heading(doc, "۱۳. جمع‌بندی و تصمیمات مورد نیاز", 1)
    add_p(doc, "سامانه همراهِ انتخاب از نظر فنی قابل ساخت است، اما موفقیت آن بیش از مدل هوش مصنوعی به اعتماد، حاکمیت داده و همکاری رسمی دستگاه‌ها وابسته است. نسخه قابل دفاع پروژه باید از یک پرونده رضایت‌محور و راستی‌آزمایی حداقلی شروع کند و سپس تنها اتصال‌هایی را توسعه دهد که مبنای قانونی، کیفیت داده و سطح خدمت روشن دارند.")
    add_heading(doc, "تصمیمات لازم برای آغاز فاز صفر", 2)
    decisions = [
        "تعیین مالک محصول، نهاد متولی و ترکیب هیئت راهبری.",
        "تأیید نام و دامنه مأموریت: تصمیم‌یار ازدواج و خدمات حمایتی، نه سامانه تحقیق مخفیانه.",
        "معرفی دستگاه‌های هدف و نمایندگان رسمی برای مذاکره داده و API.",
        "تعیین جامعه پایلوت، شهر/استان، نقش مشاوران و ظرفیت پشتیبانی.",
        "انتخاب مدل استقرار: زیرساخت کارفرما، ابر خصوصی یا مرکز داده مورد تأیید.",
        "تخصیص بودجه فاز صفر و پذیرش اصل قیمت‌گذاری مرحله‌ای اتصال‌ها.",
    ]
    for d in decisions:
        add_number(doc, d)
    add_callout(doc, "پیشنهاد نهایی", "قرارداد نخست صرفاً برای فاز صفر ۶ تا ۸ هفته‌ای منعقد شود. خروجی این فاز باید برای هر قابلیت یک تصمیم روشن «اجرا، اجرای مشروط یا عدم اجرا» و برای MVP دامنه، زمان و بودجه قابل اتکا تولید کند.", fill=PALE_BLUE, accent=BLUE)

    add_heading(doc, "پیوست الف - فهرست پرسش‌های فاز صفر", 1)
    questions = [
        "کارفرما دقیقاً چه مسئولیت قانونی و چه جامعه هدفی دارد؟",
        "آیا مالک هر پایگاه داده، ارائه نتیجه برای این هدف را مجاز می‌داند؟",
        "کمینه پاسخ لازم چیست: تأیید/عدم تأیید، وضعیت گواهی یا جزئیات بیشتر؟",
        "رضایت چگونه احراز، ثبت، تمدید و لغو می‌شود؟",
        "داده چه مدت نگهداری و پس از خاتمه آشنایی چگونه حذف می‌شود؟",
        "چه کسی حق مشاهده داده خام را دارد و چه کسی فقط خلاصه را می‌بیند؟",
        "فرایند اعتراض، تصحیح منبع و توقف اثر نتیجه چیست؟",
        "چه سناریوهایی نیازمند ارجاع فوری به مشاور یا مرجع صلاحیت‌دار است؟",
        "حد قابل قبول خطای مدل و مسئولیت هر خطا چگونه تعریف می‌شود؟",
        "آیا جداسازی کامل خدمات پس از ازدواج و اشتغال از پرونده آشنایی پذیرفته می‌شود؟",
    ]
    for q in questions:
        p = add_bullet(doc, q)
        p.paragraph_format.space_after = Pt(1)
        p.paragraph_format.line_spacing = 1.08
        for run in p.runs:
            set_run(run, size=9.1, color=INK)

    add_page_break(doc)
    add_heading(doc, "پیوست ب - نمونه ساختار گزارش کاربر", 1)
    add_p(doc, "این پیوست قالب مفهومی گزارش است و محتوای واقعی ندارد. هدف آن نشان‌دادن زبان، سطح جزئیات و مرز میان شاهد، نتیجه و توصیه گفت‌وگو است. نسخه نهایی باید با آزمون کاربر، نظر حقوقی و متخصص خانواده اصلاح شود.")
    add_heading(doc, "ب-۱. مشخصات و وضعیت پرونده", 2)
    add_table(doc, ["مولفه", "نمونه نمایش"], [
        ["شناسه پرونده", "کد غیرقابل حدس؛ بدون نمایش شناسه ملی"],
        ["طرفین", "نام نمایشی تأییدشده هر دو طرف"],
        ["دامنه رضایت", "هویت، تحصیل و اشتغال؛ وب عمومی غیرفعال"],
        ["اعتبار گزارش", "تاریخ صدور و تاریخ انقضای هر شاهد"],
        ["وضعیت کلی", "۴ ادعا تأییدشده، ۱ مورد در انتظار پاسخ، بدون حکم سازگاری"],
    ], [5.0, 11.65], font_size=8.8)
    add_heading(doc, "ب-۲. خلاصه شواهد", 2)
    add_table(doc, ["ادعا", "وضعیت", "شاهد و تاریخ", "توضیح/محدودیت"], [
        ["مدرک کارشناسی", "تأییدشده", "پاسخ مرجع منتخب - تاریخ نمونه", "اعتبار فقط برای مدرک ذکرشده"],
        ["اشتغال فعلی", "نیازمند بررسی", "گواهی ارائه‌شده - تاریخ نمونه", "پاسخ مرجع هنوز دریافت نشده"],
        ["گواهی عدم سوءپیشینه", "گواهی معتبر", "گواهی رسمی ارائه‌شده توسط فرد", "نمایش جزئیات خارج از دامنه نیست"],
    ], [3.6, 2.8, 5.2, 5.05], font_size=8.4)
    add_heading(doc, "ب-۳. تناقض‌ها و پاسخ صاحب داده", 2)
    add_p(doc, "نمونه متن خنثی: «تاریخ شروع اشتغال در ادعای ثبت‌شده فروردین و در گواهی ارائه‌شده خرداد درج شده است. این اختلاف می‌تواند ناشی از دوره آزمایشی، همکاری پروژه‌ای یا خطای ثبت باشد. صاحب داده توضیح داده است که همکاری از فروردین به‌صورت پروژه‌ای آغاز و قرارداد رسمی از خرداد منعقد شده است. وضعیت برای مشاهده مدرک تکمیلی باز است.»")
    add_heading(doc, "ب-۴. موضوعات پیشنهادی برای گفت‌وگو", 2)
    for x in [
        "تعریف هر دو نفر از ثبات شغلی و میزان پذیرش تغییر شغل یا شهر محل کار چیست؟",
        "چه اطلاعات مالی باید پیش از تصمیم مشترک شفاف شود و چه مواردی خصوصی باقی می‌ماند؟",
        "در اختلاف نظر درباره خانواده‌ها، تصمیم‌گیری و مرزبندی چگونه انجام خواهد شد؟",
        "آیا موضوع باقی‌مانده به گفت‌وگوی دوطرفه پاسخ می‌دهد یا جلسه مشاوره مفیدتر است؟",
    ]:
        add_bullet(doc, x)
    add_callout(doc, "عبارت ثابت گزارش", "این گزارش تنها وضعیت شواهد منتخب در زمان مشخص را نشان می‌دهد و پیش‌بینی موفقیت ازدواج، تشخیص شخصیت یا جایگزین مشاوره و شناخت متقابل نیست.", fill=PALE_RED, accent=RED)

    add_heading(doc, "پیوست پ - ماتریس مسئولیت‌های اجرایی (RACI)", 1)
    add_p(doc, "در جدول زیر: «مسئول» انجام‌دهنده کار، «پاسخ‌گو» صاحب تصمیم نهایی، «مشورتی» مشارکت‌کننده تخصصی و «مطلع» دریافت‌کننده گزارش است. نام واحدها باید در فاز صفر با ساختار واقعی کارفرما جایگزین شود.")
    add_table(doc, ["فعالیت", "پاسخ‌گو", "مسئول", "مشورتی/مطلع"], [
        ["تصویب دامنه و بودجه", "هیئت راهبری", "مدیر برنامه", "مالی، حقوقی و واحدهای خدمت"],
        ["تأیید مبنای پردازش", "اداره حقوقی", "مسئول حریم خصوصی", "مالک داده و امنیت"],
        ["طراحی تجربه رضایت", "مالک محصول", "طراح UX", "حقوقی، مشاور خانواده و کاربران"],
        ["معماری و اتصال API", "معمار ارشد", "تیم بک‌اند/تعامل‌پذیری", "امنیت و مالک داده"],
        ["ارزیابی مدل", "رهبر AI", "QA مدل", "حقوقی، متخصص حوزه و امنیت"],
        ["پذیرش امنیت", "مدیر امنیت", "DevSecOps", "معمار و ممیز مستقل"],
        ["عملیات پایلوت", "مالک محصول", "مدیر پایلوت", "مشاوران، پشتیبانی و تحلیل داده"],
        ["رسیدگی به اعتراض", "مسئول حریم خصوصی", "کارشناس پرونده", "حقوقی، مالک منبع و محصول"],
        ["مدیریت رخداد", "مدیر امنیت", "تیم پاسخ به رخداد", "حقوقی، روابط عمومی و مدیریت ارشد"],
    ], [4.1, 3.2, 4.0, 5.35], font_size=8.3)

    add_heading(doc, "پیوست ت - چک‌لیست فنی اتصال به دستگاه‌ها", 1)
    add_heading(doc, "ت-۱. پیش‌نیازهای قراردادی و داده", 2)
    for x in [
        "مالک رسمی سرویس، نماینده فنی و نماینده حقوقی مشخص و دامنه مجاز استفاده مکتوب شده است.",
        "فهرست دقیق فیلدهای درخواست و پاسخ، کدهای خطا و معنای هر وضعیت تأیید شده است.",
        "کمینه پاسخ بررسی شده و ضرورت دریافت داده خام یا جزئیات اضافی مستند است.",
        "دوره نگهداشت، محدودیت بازنشر، حق ممیزی و مسئول اصلاح داده در توافق درج شده است.",
        "SLA، ساعات پشتیبانی، فرایند رخداد و شرایط قطع اضطراری سرویس تعیین شده است.",
    ]:
        add_bullet(doc, x)
    add_heading(doc, "ت-۲. پیش‌نیازهای فنی و امنیتی", 2)
    for x in [
        "محیط Sandbox با داده ساختگی، مستندات نسخه‌بندی‌شده و نمونه درخواست/پاسخ در دسترس است.",
        "احراز سرویس متقابل، مدیریت کلید، محدودسازی نرخ و فهرست سفید شبکه پیاده‌سازی شده است.",
        "هر درخواست دارای شناسه یکتا، هدف، شناسه رضایت و مهر زمان است و پاسخ امضاشده یا قابل اثبات دریافت می‌شود.",
        "Timeout، Retry، مدارشکن، صف و رفتار سامانه در خطای مرجع آزمایش شده است.",
        "لاگ‌ها فاقد داده حساس غیرضروری‌اند و هم‌بستگی رخداد میان دو طرف اتصال امکان‌پذیر است.",
        "آزمون نفوذ اتصال، بازبینی مجوزها و سناریوی لغو دسترسی پیش از تولید تکمیل شده است.",
    ]:
        add_bullet(doc, x)
    add_heading(doc, "ت-۳. آزمون پذیرش هر اتصال", 2)
    add_table(doc, ["سناریو", "نتیجه مورد انتظار"], [
        ["رضایت معتبر", "ارسال درخواست و ثبت رسید منبع"],
        ["رضایت منقضی/لغوشده", "رد درخواست پیش از تماس با مرجع"],
        ["شناسه نامعتبر", "خطای قابل فهم بدون افشای داده اضافی"],
        ["مرجع خارج از دسترس", "وضعیت انتظار؛ عدم تبدیل به نتیجه منفی"],
        ["پاسخ دیرهنگام/تکراری", "پردازش یک‌باره و حفظ سازگاری پرونده"],
        ["اعتراض و اصلاح منبع", "نسخه‌بندی نتیجه و ثبت توضیح"],
        ["قطع اضطراری", "توقف اتصال بدون اختلال کل سامانه"],
    ], [5.4, 11.25], font_size=8.7)

    add_heading(doc, "پیوست ث - الگوی محتوای رضایت‌نامه", 1)
    add_p(doc, "رضایت‌نامه نهایی باید توسط حقوقی کارفرما تصویب شود. ساختار زیر یک الگوی محتوایی برای طراحی رابط است و متن قراردادی نهایی محسوب نمی‌شود.")
    consent_parts = [
        ("هدف", "می‌دانم این بسته برای بررسی ادعای مشخص در پرونده آشنایی من استفاده می‌شود و برای تبلیغات، استخدام یا اعتبارسنجی مالی به‌کار نمی‌رود."),
        ("داده و منبع", "می‌دانم چه فیلدهایی از چه مرجعی دریافت یا از چه مدرکی استخراج می‌شود و می‌توانم قبل از تأیید، نمونه خروجی را ببینم."),
        ("گیرنده", "انتخاب می‌کنم نتیجه را خودم، طرف مقابل یا مشاور پرونده مشاهده کند. دسترسی هر گیرنده جداگانه قابل مدیریت است."),
        ("مدت", "از تاریخ شروع و پایان اعتبار رضایت و دوره نگهداشت داده مطلع هستم و پیش از تمدید دوباره از من سؤال می‌شود."),
        ("لغو", "می‌توانم رضایت را برای پردازش‌های آینده لغو کنم و می‌دانم لغو چه اثری بر نتایج قبلاً مشاهده‌شده و الزامات قانونی نگهداشت دارد."),
        ("اعتراض", "می‌توانم منبع، برداشت سامانه یا انتساب داده را به چالش بکشم و در موارد پراثر تا پایان بررسی، نمایش نتیجه متوقف می‌شود."),
    ]
    for label, detail in consent_parts:
        p = add_p(doc, after=5, line=1.22)
        set_run(p.add_run(label + ": "), size=11.2, bold=True, color=RED)
        set_run(p.add_run(detail), size=11.0, color=INK)
    add_heading(doc, "کنترل فهم رضایت", 2)
    add_p(doc, "برای بسته‌های حساس، یک سؤال کوتاه فهم می‌تواند پیش از تأیید نمایش داده شود؛ برای مثال «آیا جزئیات پرونده‌های قضایی در گزارش طرف مقابل نمایش داده می‌شود؟» پاسخ صحیح باید نشان دهد که فقط وضعیت گواهی مجاز و حداقلی در دامنه است. این روش جای متن روشن را نمی‌گیرد، اما می‌تواند ابهام‌های مهم را پیش از اقدام آشکار کند.")

    add_heading(doc, "پیوست ج - واژه‌نامه تخصصی", 1)
    glossary = [
        ("ادعا", "اطلاعاتی که فرد درباره خود ثبت می‌کند و هنوز ممکن است شاهد مستقل نداشته باشد."),
        ("شاهد", "سند، پاسخ مرجع یا محتوای عمومی مشخصی که منشأ، تاریخ و محدوده آن ثبت شده است."),
        ("راستی‌آزمایی", "بررسی انطباق یک ادعای معین با شاهد معتبر؛ نه ارزیابی کلی شخصیت یا شایستگی."),
        ("رضایت دانه‌ریز", "انتخاب مستقل برای منبع، هدف، گیرنده و مدت هر پردازش."),
        ("DPIA", "ارزیابی اثرات حریم خصوصی برای شناسایی ضرورت، خطر و کنترل پیش از پردازش."),
        ("RAG", "الگوی تولید پاسخ با تکیه بر منابع بازیابی‌شده و قابل ارجاع، برای کاهش خروجی بی‌پشتوانه."),
        ("OCR", "استخراج متن و فیلد از تصویر سند؛ خروجی آن نیازمند کنترل کیفیت است."),
        ("مدل تصمیم‌یار", "سامانه‌ای که شواهد و سؤال‌ها را سامان می‌دهد اما تصمیم نهایی را به انسان واگذار می‌کند."),
        ("سطح اطمینان", "برآورد کیفیت و کفایت خروجی فنی؛ به معنای صحت قطعی واقعیت یا مناسب‌بودن فرد نیست."),
        ("زنجیره شواهد", "پیوند قابل ممیزی میان گزاره گزارش، منبع، زمان، نسخه پردازش و تأییدکننده."),
        ("SLA", "توافق سطح خدمت شامل دسترس‌پذیری، زمان پاسخ، پشتیبانی و مسئولیت طرفین."),
        ("RPO/RTO", "حد قابل قبول از دست‌رفتن داده و زمان هدف برای بازیابی خدمت پس از اختلال."),
    ]
    add_table(doc, ["اصطلاح", "تعریف کاربردی در این پروپوزال"], [[a, b] for a, b in glossary], [4.0, 12.65], font_size=8.6)

    add_heading(doc, "پیوست چ - منابع و چارچوب‌های مرجع", 1)
    sources = [
        ("قانون انتشار و دسترسی آزاد به اطلاعات", "https://nezamat.ir/post-37380/"),
        ("آیین‌نامه اجرایی قانون انتشار و دسترسی آزاد به اطلاعات (تعریف حریم خصوصی)", "https://nezamat.ir/post-39675/"),
        ("شیوه‌نامه تشخیص و تفکیک اطلاعات مربوط به حریم خصوصی از اطلاعات عمومی", "https://nezamat.ir/post-41049/"),
        ("قانون مدیریت داده‌ها و اطلاعات ملی", "https://nezamat.ir/post-44082/"),
        ("درگاه خدمات الکترونیک قضایی - خدمت درخواست گواهی عدم سوءپیشینه", "https://adliran.ir/"),
        ("سامانه تعرفه خدمات فنی-تخصصی انفورماتیک سازمان نظام صنفی رایانه‌ای کشور", "https://tariffs.irannsr.org/"),
        ("NIST AI Risk Management Framework 1.0", "https://www.nist.gov/publications/artificial-intelligence-risk-management-framework-ai-rmf-10"),
        ("ISO/IEC 27001:2022 - سیستم مدیریت امنیت اطلاعات", "https://www.iso.org/standard/27001"),
        ("ISO/IEC 27701:2025 - سیستم مدیریت اطلاعات حریم خصوصی", "https://www.iso.org/standard/27701"),
    ]
    for idx, (name, url) in enumerate(sources, 1):
        p = add_p(doc, after=2, line=1.05, align=WD_ALIGN_PARAGRAPH.RIGHT)
        set_run(p.add_run(f"{idx}. {name}: "), size=8.5, color=INK)
        add_hyperlink(p, "مشاهده منبع", url, size=8.5)
    doc.save(OUT)
    print(OUT)


if __name__ == "__main__":
    build()
