from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile
from lxml import etree


SRC = Path(r"D:\project\innovation\proposal_hamrah_entekhab.docx")
OUT = Path(r"D:\project\innovation\proposal_hamrah_entekhab_Word_RTL.docx")

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W}


def qn(tag: str) -> str:
    return f"{{{W}}}{tag}"


def first(parent, tag: str):
    return parent.find(qn(tag))


def ensure(parent, tag: str, first_child: bool = False):
    node = first(parent, tag)
    if node is None:
        node = etree.Element(qn(tag))
        if first_child:
            parent.insert(0, node)
        else:
            parent.append(node)
    return node


def set_val(node, value: str = "1"):
    node.set(qn("val"), value)


def harden_run(run):
    rpr = first(run, "rPr")
    if rpr is None:
        rpr = etree.Element(qn("rPr"))
        run.insert(0, rpr)
    set_val(ensure(rpr, "rtl"), "1")
    fonts = ensure(rpr, "rFonts")
    for key in ("ascii", "hAnsi", "cs"):
        fonts.set(qn(key), "Tahoma")
    lang = ensure(rpr, "lang")
    lang.set(qn("val"), "fa-IR")
    lang.set(qn("bidi"), "fa-IR")


def harden_paragraph(p, force_right=False):
    ppr = first(p, "pPr")
    if ppr is None:
        ppr = etree.Element(qn("pPr"))
        p.insert(0, ppr)
    set_val(ensure(ppr, "bidi"), "1")
    jc = ensure(ppr, "jc")
    old = jc.get(qn("val"))
    has_text = bool("".join(p.xpath(".//w:t/text()", namespaces=NS)).strip())
    if force_right and has_text:
        jc.set(qn("val"), "right")
    elif old != "center" and has_text:
        # A true right-aligned body is more stable across Word installations
        # than renderer-dependent RTL justification.
        jc.set(qn("val"), "right")
    for run in p.findall(".//w:r", NS):
        harden_run(run)


def patch_story(xml: bytes, part_name: str) -> bytes:
    root = etree.fromstring(xml)
    in_header_or_footer = part_name.startswith("word/header") or part_name.startswith("word/footer")

    for p in root.findall(".//w:p", NS):
        harden_paragraph(p)

    # Explicitly switch actual body data tables to RTL visual column order.
    # Header/footer furniture and the two-logo cover table retain their layout.
    for index, table in enumerate(root.findall(".//w:tbl", NS)):
        rows = table.findall("w:tr", NS)
        col_count = max((len(r.findall("w:tc", NS)) for r in rows), default=0)
        if not in_header_or_footer and col_count >= 3:
            tbl_pr = first(table, "tblPr")
            if tbl_pr is None:
                tbl_pr = etree.Element(qn("tblPr"))
                table.insert(0, tbl_pr)
            set_val(ensure(tbl_pr, "bidiVisual"), "1")
            for p in table.findall(".//w:tc//w:p", NS):
                harden_paragraph(p, force_right=True)

    for sect in root.findall(".//w:sectPr", NS):
        set_val(ensure(sect, "bidi"), "1")
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone="yes")


def patch_styles(xml: bytes) -> bytes:
    root = etree.fromstring(xml)
    defaults = ensure(root, "docDefaults", first_child=True)
    rpr_default = ensure(defaults, "rPrDefault")
    rpr = ensure(rpr_default, "rPr")
    set_val(ensure(rpr, "rtl"), "1")
    fonts = ensure(rpr, "rFonts")
    for key in ("ascii", "hAnsi", "cs"):
        fonts.set(qn(key), "Tahoma")
    lang = ensure(rpr, "lang")
    lang.set(qn("val"), "fa-IR")
    lang.set(qn("bidi"), "fa-IR")

    ppr_default = ensure(defaults, "pPrDefault")
    default_ppr = ensure(ppr_default, "pPr")
    set_val(ensure(default_ppr, "bidi"), "1")
    set_val(ensure(default_ppr, "jc"), "right")

    for style in root.findall("w:style", NS):
        if style.get(qn("type")) != "paragraph":
            continue
        ppr = ensure(style, "pPr")
        set_val(ensure(ppr, "bidi"), "1")
        style_id = style.get(qn("styleId"), "")
        if style_id in {"Normal", "Heading1", "Heading2", "Heading3", "ListBullet", "ListBullet2", "ListNumber"}:
            set_val(ensure(ppr, "jc"), "right")
        rpr = ensure(style, "rPr")
        set_val(ensure(rpr, "rtl"), "1")
        fonts = ensure(rpr, "rFonts")
        for key in ("ascii", "hAnsi", "cs"):
            fonts.set(qn(key), "Tahoma")
        lang = ensure(rpr, "lang")
        lang.set(qn("val"), "fa-IR")
        lang.set(qn("bidi"), "fa-IR")
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone="yes")


def patch_numbering(xml: bytes) -> bytes:
    root = etree.fromstring(xml)
    for lvl in root.findall(".//w:lvl", NS):
        set_val(ensure(lvl, "lvlJc"), "right")
        ppr = ensure(lvl, "pPr")
        set_val(ensure(ppr, "bidi"), "1")
        ind = ensure(ppr, "ind")
        ind.attrib.pop(qn("left"), None)
        ind.attrib.pop(qn("firstLine"), None)
        ind.set(qn("right"), "720")
        ind.set(qn("hanging"), "360")
        rpr = ensure(lvl, "rPr")
        set_val(ensure(rpr, "rtl"), "1")
        fonts = ensure(rpr, "rFonts")
        for key in ("ascii", "hAnsi", "cs"):
            fonts.set(qn(key), "Tahoma")
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone="yes")


def patch_settings(xml: bytes) -> bytes:
    root = etree.fromstring(xml)
    lang = ensure(root, "themeFontLang")
    lang.set(qn("bidi"), "fa-IR")
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone="yes")


with ZipFile(SRC, "r") as zin, ZipFile(OUT, "w", ZIP_DEFLATED) as zout:
    for info in zin.infolist():
        data = zin.read(info.filename)
        if info.filename == "word/styles.xml":
            data = patch_styles(data)
        elif info.filename == "word/numbering.xml":
            data = patch_numbering(data)
        elif info.filename == "word/settings.xml":
            data = patch_settings(data)
        elif info.filename == "word/document.xml" or info.filename.startswith("word/header") or info.filename.startswith("word/footer"):
            if info.filename.endswith(".xml"):
                data = patch_story(data, info.filename)
        zout.writestr(info, data)

print(OUT)
