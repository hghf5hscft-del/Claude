#!/usr/bin/env python3
"""
hd_to_docx.py — Human Design Thai Report → DOCX
Pure Python stdlib (zipfile, struct, xml.sax.saxutils). No pip install.

Usage:
    python3 hd_to_docx.py report.txt [output.docx]

Input format — one tag per line:
    [COVER_NAME]    ชื่อผู้รับการวิเคราะห์
    [COVER_DATE]    วันเกิด
    [COVER_SUB]     โครงสร้างหลัก • Variables • Shadow Chart
    [PAGEBREAK]
    [H1]  หัวข้อหลัก
    [H2]  หัวข้อรอง
    [H3]  หัวข้อย่อย
    [INTRO]      ย่อหน้าแนะนำ (italic, gray, left border)
    [BODY]       ย่อหน้าปกติ
    [BULLET]     รายการ bullet
    [SUBBULLET]  รายการย่อย
    [TABLE_HEADER]  คอล1 | คอล2 | คอล3
    [TABLE_ROW]     ค่า1  | ค่า2  | ค่า3
    [IMAGE]      /absolute/path/to/photo.jpg
    [IMAGE_CAPTION]  คำอธิบายรูป (optional, place right after [IMAGE])
    (blank line = spacer)
"""

import sys
import os
import struct
import zipfile
import datetime
from xml.sax.saxutils import escape as xmlesc

# ── Fonts & Colors ──────────────────────────────────────────────────
FONT   = "TH Sarabun New"
C_H1   = "1F497D"
C_H2   = "2E74B5"
C_H3   = "4472C4"
C_INT  = "595959"
C_THBG = "1F497D"
C_THFG = "FFFFFF"

MAX_IMG_W_EMU = 5_760_000   # ~6 inches — fits safely inside A4 margins
PX_TO_EMU     = 9525        # at 96 DPI: 914400 / 96


# ══════════════════════════════════════════════════════════════════════
#  Image helpers (no Pillow — pure struct parsing)
# ══════════════════════════════════════════════════════════════════════

def _png_size(data):
    """Return (w, h) from PNG binary data, or None."""
    if len(data) < 24 or data[:8] != b'\x89PNG\r\n\x1a\n':
        return None
    return struct.unpack('>II', data[16:24])


def _jpeg_size(data):
    """Return (w, h) from JPEG binary data, or None."""
    i = 2
    while i + 3 < len(data):
        if data[i] != 0xFF:
            break
        marker = data[i + 1]
        if marker in (0xC0, 0xC1, 0xC2, 0xC3):   # SOF markers
            if i + 9 <= len(data):
                h, w = struct.unpack('>HH', data[i + 5:i + 9])
                return w, h
            break
        if i + 4 > len(data):
            break
        seg_len = struct.unpack('>H', data[i + 2:i + 4])[0]
        i += 2 + seg_len
    return None


def _load_image(path):
    """
    Load image and return (data_bytes, zip_filename, content_type, w_emu, h_emu).
    Raises FileNotFoundError or ValueError on bad input.
    """
    with open(path, 'rb') as fh:
        data = fh.read()

    ext = os.path.splitext(path)[1].lower().lstrip('.')

    if data[:8] == b'\x89PNG\r\n\x1a\n':
        dims = _png_size(data)
        ctype = 'image/png'
        zip_ext = 'png'
    elif data[:2] == b'\xff\xd8':
        dims = _jpeg_size(data)
        ctype = 'image/jpeg'
        zip_ext = 'jpeg'
    elif ext in ('jpg', 'jpeg'):
        dims = _jpeg_size(data)
        ctype = 'image/jpeg'
        zip_ext = 'jpeg'
    elif ext == 'png':
        dims = _png_size(data)
        ctype = 'image/png'
        zip_ext = 'png'
    else:
        raise ValueError(f"Unsupported image format: {path}")

    if dims is None:
        dims = (800, 600)       # fallback if parsing fails

    w_px, h_px = dims
    w_emu = w_px * PX_TO_EMU
    h_emu = h_px * PX_TO_EMU

    if w_emu > MAX_IMG_W_EMU:
        scale = MAX_IMG_W_EMU / w_emu
        w_emu = int(MAX_IMG_W_EMU)
        h_emu = int(h_emu * scale)

    return data, zip_ext, ctype, w_emu, h_emu


def _drawing_xml(rel_id, img_id, w_emu, h_emu):
    """Inline image drawing XML referencing a relationship ID."""
    name = f"Image{img_id}"
    # Namespaces are declared here so this XML fragment is self-contained
    # even if the document root omits some of them.
    return (
        f'<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:drawing>'
        f'<wp:inline'
        f' xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"'
        f' distT="0" distB="114300" distL="0" distR="0">'
        f'<wp:extent cx="{w_emu}" cy="{h_emu}"/>'
        f'<wp:effectExtent l="0" t="0" r="0" b="0"/>'
        f'<wp:docPr id="{img_id}" name="{name}"/>'
        f'<wp:cNvGraphicFramePr>'
        f'<a:graphicFrameLocks'
        f' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
        f' noChangeAspect="1"/>'
        f'</wp:cNvGraphicFramePr>'
        f'<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        f'<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        f'<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        f'<pic:nvPicPr>'
        f'<pic:cNvPr id="{img_id}" name="{name}"/>'
        f'<pic:cNvPicPr/>'
        f'</pic:nvPicPr>'
        f'<pic:blipFill>'
        f'<a:blip'
        f' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
        f' r:embed="{rel_id}"/>'
        f'<a:stretch><a:fillRect/></a:stretch>'
        f'</pic:blipFill>'
        f'<pic:spPr>'
        f'<a:xfrm><a:off x="0" y="0"/><a:ext cx="{w_emu}" cy="{h_emu}"/></a:xfrm>'
        f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        f'</pic:spPr>'
        f'</pic:pic>'
        f'</a:graphicData>'
        f'</a:graphic>'
        f'</wp:inline>'
        f'</w:drawing></w:r></w:p>'
    )


# ══════════════════════════════════════════════════════════════════════
#  Static ZIP members
# ══════════════════════════════════════════════════════════════════════

def _content_types_xml(img_exts):
    """Build [Content_Types].xml, adding image MIME types as needed."""
    extra = ""
    if 'png'  in img_exts: extra += '<Default Extension="png"  ContentType="image/png"/>\n  '
    if 'jpeg' in img_exts: extra += '<Default Extension="jpeg" ContentType="image/jpeg"/>\n  '
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">\n'
        '  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>\n'
        '  <Default Extension="xml"  ContentType="application/xml"/>\n'
        f'  {extra}'
        '  <Override PartName="/word/document.xml"\n'
        '    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>\n'
        '  <Override PartName="/word/styles.xml"\n'
        '    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>\n'
        '  <Override PartName="/word/settings.xml"\n'
        '    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>\n'
        '</Types>'
    ).encode('utf-8')


ROOT_RELS = b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    Target="word/document.xml"/>
</Relationships>"""

SETTINGS = b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:defaultTabStop w:val="720"/>
  <w:compat>
    <w:compatSetting w:name="compatibilityMode"
      w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>
  </w:compat>
</w:settings>"""


def _word_rels_xml(image_rels):
    """
    Build word/_rels/document.xml.rels.
    image_rels: list of (rel_id, zip_filename) e.g. [("rId3","media/image1.jpeg"), ...]
    """
    img_entries = "".join(
        f'  <Relationship Id="{rid}"\n'
        f'    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"\n'
        f'    Target="{target}"/>\n'
        for rid, target in image_rels
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n'
        '  <Relationship Id="rId1"\n'
        '    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"\n'
        '    Target="styles.xml"/>\n'
        '  <Relationship Id="rId2"\n'
        '    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"\n'
        '    Target="settings.xml"/>\n'
        f'{img_entries}'
        '</Relationships>'
    ).encode('utf-8')


def _styles_xml():
    f = FONT
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault><w:rPr>
      <w:rFonts w:ascii="{f}" w:hAnsi="{f}" w:cs="{f}" w:eastAsia="{f}"/>
      <w:sz w:val="26"/><w:szCs w:val="26"/>
      <w:lang w:val="th-TH" w:eastAsia="th-TH"/>
    </w:rPr></w:rPrDefault>
    <w:pPrDefault><w:pPr>
      <w:spacing w:after="80"/><w:jc w:val="both"/>
    </w:pPr></w:pPrDefault>
  </w:docDefaults>

  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:rPr>
      <w:rFonts w:ascii="{f}" w:hAnsi="{f}" w:cs="{f}" w:eastAsia="{f}"/>
      <w:sz w:val="26"/><w:szCs w:val="26"/>
    </w:rPr>
  </w:style>

  <w:style w:type="paragraph" w:styleId="HDH1">
    <w:name w:val="HD Heading 1"/><w:basedOn w:val="Normal"/>
    <w:pPr>
      <w:spacing w:before="280" w:after="120"/>
      <w:jc w:val="left"/>
      <w:pBdr><w:bottom w:val="single" w:sz="8" w:space="4" w:color="{C_H1}"/></w:pBdr>
    </w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="{f}" w:hAnsi="{f}" w:cs="{f}" w:eastAsia="{f}"/>
      <w:b/><w:bCs/><w:color w:val="{C_H1}"/>
      <w:sz w:val="40"/><w:szCs w:val="40"/>
    </w:rPr>
  </w:style>

  <w:style w:type="paragraph" w:styleId="HDH2">
    <w:name w:val="HD Heading 2"/><w:basedOn w:val="Normal"/>
    <w:pPr><w:spacing w:before="180" w:after="80"/><w:jc w:val="left"/></w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="{f}" w:hAnsi="{f}" w:cs="{f}" w:eastAsia="{f}"/>
      <w:b/><w:bCs/><w:color w:val="{C_H2}"/>
      <w:sz w:val="32"/><w:szCs w:val="32"/>
    </w:rPr>
  </w:style>

  <w:style w:type="paragraph" w:styleId="HDH3">
    <w:name w:val="HD Heading 3"/><w:basedOn w:val="Normal"/>
    <w:pPr><w:spacing w:before="120" w:after="60"/><w:jc w:val="left"/></w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="{f}" w:hAnsi="{f}" w:cs="{f}" w:eastAsia="{f}"/>
      <w:b/><w:bCs/><w:color w:val="{C_H3}"/>
      <w:sz w:val="28"/><w:szCs w:val="28"/>
    </w:rPr>
  </w:style>

  <w:style w:type="paragraph" w:styleId="HDIntro">
    <w:name w:val="HD Intro"/><w:basedOn w:val="Normal"/>
    <w:pPr>
      <w:ind w:left="400"/>
      <w:spacing w:before="60" w:after="60"/>
      <w:jc w:val="both"/>
      <w:pBdr><w:left w:val="single" w:sz="12" w:space="8" w:color="{C_H3}"/></w:pBdr>
    </w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="{f}" w:hAnsi="{f}" w:cs="{f}" w:eastAsia="{f}"/>
      <w:i/><w:iCs/><w:color w:val="{C_INT}"/>
      <w:sz w:val="24"/><w:szCs w:val="24"/>
    </w:rPr>
  </w:style>

  <w:style w:type="paragraph" w:styleId="HDBullet">
    <w:name w:val="HD Bullet"/><w:basedOn w:val="Normal"/>
    <w:pPr>
      <w:ind w:left="420" w:hanging="220"/>
      <w:spacing w:after="60"/>
    </w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="{f}" w:hAnsi="{f}" w:cs="{f}" w:eastAsia="{f}"/>
      <w:sz w:val="26"/><w:szCs w:val="26"/>
    </w:rPr>
  </w:style>

  <w:style w:type="paragraph" w:styleId="HDSubBullet">
    <w:name w:val="HD Sub Bullet"/><w:basedOn w:val="Normal"/>
    <w:pPr>
      <w:ind w:left="780" w:hanging="220"/>
      <w:spacing w:after="40"/>
    </w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="{f}" w:hAnsi="{f}" w:cs="{f}" w:eastAsia="{f}"/>
      <w:sz w:val="24"/><w:szCs w:val="24"/>
    </w:rPr>
  </w:style>

  <w:style w:type="paragraph" w:styleId="HDImgCaption">
    <w:name w:val="HD Image Caption"/><w:basedOn w:val="Normal"/>
    <w:pPr><w:spacing w:before="40" w:after="120"/><w:jc w:val="center"/></w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="{f}" w:hAnsi="{f}" w:cs="{f}" w:eastAsia="{f}"/>
      <w:i/><w:iCs/><w:color w:val="888888"/>
      <w:sz w:val="22"/><w:szCs w:val="22"/>
    </w:rPr>
  </w:style>

  <w:style w:type="paragraph" w:styleId="HDCoverTitle">
    <w:name w:val="HD Cover Title"/><w:basedOn w:val="Normal"/>
    <w:pPr><w:spacing w:before="560" w:after="80"/><w:jc w:val="center"/></w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="{f}" w:hAnsi="{f}" w:cs="{f}" w:eastAsia="{f}"/>
      <w:b/><w:bCs/><w:color w:val="{C_H1}"/>
      <w:sz w:val="56"/><w:szCs w:val="56"/>
    </w:rPr>
  </w:style>

  <w:style w:type="paragraph" w:styleId="HDCoverSub">
    <w:name w:val="HD Cover Sub"/><w:basedOn w:val="Normal"/>
    <w:pPr><w:spacing w:before="60" w:after="60"/><w:jc w:val="center"/></w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="{f}" w:hAnsi="{f}" w:cs="{f}" w:eastAsia="{f}"/>
      <w:color w:val="{C_H2}"/>
      <w:sz w:val="30"/><w:szCs w:val="30"/>
    </w:rPr>
  </w:style>

  <w:style w:type="paragraph" w:styleId="HDCoverDetail">
    <w:name w:val="HD Cover Detail"/><w:basedOn w:val="Normal"/>
    <w:pPr><w:spacing w:before="30" w:after="30"/><w:jc w:val="center"/></w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="{f}" w:hAnsi="{f}" w:cs="{f}" w:eastAsia="{f}"/>
      <w:color w:val="888888"/>
      <w:sz w:val="22"/><w:szCs w:val="22"/>
    </w:rPr>
  </w:style>
</w:styles>""".encode('utf-8')


# ══════════════════════════════════════════════════════════════════════
#  Paragraph / table XML builders
# ══════════════════════════════════════════════════════════════════════

def _rpr(bold=False, italic=False, color=None, size=26):
    f = FONT
    b  = "<w:b/><w:bCs/>" if bold else ""
    i  = "<w:i/><w:iCs/>" if italic else ""
    c  = f'<w:color w:val="{color}"/>' if color else ""
    return (f'<w:rPr>'
            f'<w:rFonts w:ascii="{f}" w:hAnsi="{f}" w:cs="{f}" w:eastAsia="{f}"/>'
            f'{b}{i}{c}'
            f'<w:sz w:val="{size}"/><w:szCs w:val="{size}"/>'
            f'</w:rPr>')


def _run(text, bold=False, italic=False, color=None, size=26):
    return (f'<w:r>{_rpr(bold,italic,color,size)}'
            f'<w:t xml:space="preserve">{xmlesc(str(text))}</w:t></w:r>')


def _par(style, text, bold=False, italic=False, color=None, size=26):
    return (f'<w:p><w:pPr><w:pStyle w:val="{style}"/></w:pPr>'
            f'{_run(text, bold, italic, color, size)}</w:p>')


def _page_break():
    return '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'


def _spacer():
    return '<w:p><w:pPr><w:spacing w:after="60"/></w:pPr></w:p>'


def _bullet(text, sub=False):
    style = "HDSubBullet" if sub else "HDBullet"
    char  = "  ◦ " if sub else "• "
    sz    = 24 if sub else 26
    return (f'<w:p><w:pPr><w:pStyle w:val="{style}"/></w:pPr>'
            f'{_run(char, size=sz)}{_run(text, size=sz)}</w:p>')


def _table(headers, rows):
    col_n = max(len(headers), 1)
    col_w = max(1, 9000 // col_n)

    borders = "".join(
        f'<w:{s} w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>'
        for s in ["top","left","bottom","right","insideH","insideV"]
    )

    def _cell(text, header=False):
        fill  = f'<w:shd w:val="clear" w:color="auto" w:fill="{C_THBG}"/>' if header else ""
        color = C_THFG if header else None
        align = "center" if header else "left"
        return (f'<w:tc>'
                f'<w:tcPr><w:tcW w:w="{col_w}" w:type="dxa"/>{fill}</w:tcPr>'
                f'<w:p><w:pPr><w:jc w:val="{align}"/></w:pPr>'
                f'{_run(str(text).strip(), bold=header, color=color, size=22)}'
                f'</w:p></w:tc>')

    hdr_row   = "<w:tr>" + "".join(_cell(h, True) for h in headers) + "</w:tr>"
    data_rows = "".join(
        "<w:tr>" + "".join(
            _cell(c) for c in (list(row) + [""] * (col_n - len(row)))[:col_n]
        ) + "</w:tr>"
        for row in rows
    )

    return (f'<w:tbl>'
            f'<w:tblPr>'
            f'<w:tblW w:w="0" w:type="auto"/>'
            f'<w:tblBorders>{borders}</w:tblBorders>'
            f'<w:tblCellMar>'
            f'<w:top w:w="80" w:type="dxa"/><w:left w:w="100" w:type="dxa"/>'
            f'<w:bottom w:w="80" w:type="dxa"/><w:right w:w="100" w:type="dxa"/>'
            f'</w:tblCellMar>'
            f'</w:tblPr>'
            f'{hdr_row}{data_rows}</w:tbl>' + _spacer())


# ══════════════════════════════════════════════════════════════════════
#  Cover
# ══════════════════════════════════════════════════════════════════════

def _cover_xml(name, date, subs, details):
    parts = [_par("HDCoverTitle", "Human Design Chart", bold=True)]
    parts.append(_par("HDCoverSub", "การแปลผลฉบับสมบูรณ์"))
    if name:
        label = name + (f"  •  {date}" if date else "")
        parts.append(_par("HDCoverSub", label))
    for s in subs:
        parts.append(_par("HDCoverSub", s))
    for d in details:
        parts.append(_par("HDCoverDetail", d, color="888888", size=22))
    today = datetime.date.today().strftime("%d/%m/%Y")
    parts.append(_par("HDCoverDetail", f"วันที่จัดทำ: {today}", color="888888", size=22))
    parts.append(_page_break())
    return parts


# ══════════════════════════════════════════════════════════════════════
#  Parser — converts tagged lines → (xml_parts, image_registry)
# ══════════════════════════════════════════════════════════════════════

def _parse(lines):
    """
    Returns:
        xml_parts   : list[str]  — XML elements for document body
        img_registry: list[dict] — {path, rel_id, zip_name, data, ctype}
    """
    cover   = dict(name="", date="", subs=[], details=[])
    body    = []
    img_reg = []          # collected image entries
    seen_paths = {}       # deduplicate: path → rel_id

    tbl_h = None
    tbl_r = []

    def flush_table():
        nonlocal tbl_h, tbl_r
        if tbl_h is not None:
            body.append(_table(tbl_h, tbl_r))
        tbl_h = None
        tbl_r = []

    def add(xml):
        flush_table()
        body.append(xml)

    def _add_image(path, caption=""):
        """Load image, register it, emit drawing XML."""
        path = path.strip()
        if not os.path.isfile(path):
            add(_par("Normal",
                     f"[ไม่พบรูปภาพ: {path}]", italic=True, color="CC0000", size=22))
            return

        if path in seen_paths:
            rel_id = seen_paths[path]
            # find existing entry for dimensions
            entry = next(e for e in img_reg if e["rel_id"] == rel_id)
            w_emu, h_emu = entry["w_emu"], entry["h_emu"]
        else:
            idx    = len(img_reg) + 1
            rel_id = f"rIdIMG{idx}"
            try:
                data, zip_ext, ctype, w_emu, h_emu = _load_image(path)
            except Exception as exc:
                add(_par("Normal",
                         f"[ไม่สามารถโหลดรูปภาพ: {exc}]",
                         italic=True, color="CC0000", size=22))
                return
            zip_name = f"media/image{idx}.{zip_ext}"
            img_reg.append(dict(
                path=path, rel_id=rel_id, zip_name=zip_name,
                data=data, ctype=ctype, w_emu=w_emu, h_emu=h_emu
            ))
            seen_paths[path] = rel_id

        img_id = img_reg.index(
            next(e for e in img_reg if e["rel_id"] == rel_id)
        ) + 1
        flush_table()
        body.append(_drawing_xml(rel_id, img_id, w_emu, h_emu))
        if caption:
            body.append(_par("HDImgCaption", caption,
                             italic=True, color="888888", size=22))
        body.append(_spacer())

    pending_caption = None

    for raw in lines:
        line = raw.rstrip("\n")

        # Cover metadata
        if   line.startswith("[COVER_NAME]"):   cover["name"] = line[12:].strip()
        elif line.startswith("[COVER_DATE]"):   cover["date"] = line[12:].strip()
        elif line.startswith("[COVER_SUB]"):    cover["subs"].append(line[11:].strip())
        elif line.startswith("[COVER_DETAIL]"): cover["details"].append(line[14:].strip())

        # Page / headings
        elif line.startswith("[PAGEBREAK]"):    add(_page_break())
        elif line.startswith("[H1]"):
            t = line[4:].strip()
            if t: add(_par("HDH1", t, bold=True, color=C_H1, size=40))
        elif line.startswith("[H2]"):
            t = line[4:].strip()
            if t: add(_par("HDH2", t, bold=True, color=C_H2, size=32))
        elif line.startswith("[H3]"):
            t = line[4:].strip()
            if t: add(_par("HDH3", t, bold=True, color=C_H3, size=28))

        # Text styles
        elif line.startswith("[INTRO]"):
            t = line[7:].strip()
            if t: add(_par("HDIntro", t, italic=True, color=C_INT, size=24))
        elif line.startswith("[BODY]"):
            t = line[6:].strip()
            if t: add(_par("Normal", t))
        elif line.startswith("[BULLET]"):
            t = line[8:].strip()
            if t:
                flush_table()
                body.append(_bullet(t))
        elif line.startswith("[SUBBULLET]"):
            t = line[11:].strip()
            if t:
                flush_table()
                body.append(_bullet(t, sub=True))

        # Tables
        elif line.startswith("[TABLE_HEADER]"):
            flush_table()
            tbl_h = [c.strip() for c in line[14:].split("|")]
            tbl_r = []
        elif line.startswith("[TABLE_ROW]"):
            if tbl_h is not None:
                tbl_r.append([c.strip() for c in line[11:].split("|")])

        # Images
        elif line.startswith("[IMAGE]"):
            img_path = line[7:].strip()
            _add_image(img_path)

        elif line.startswith("[IMAGE_CAPTION]"):
            # Caption replaces last spacer with captioned paragraph
            caption_text = line[15:].strip()
            if caption_text and body:
                # Replace trailing spacer if present
                if body and body[-1] == _spacer():
                    body.pop()
                body.append(_par("HDImgCaption", caption_text,
                                 italic=True, color="888888", size=22))
                body.append(_spacer())

        # Blank line
        elif line.strip() == "":
            flush_table()
            body.append(_spacer())

        # Untagged → body text
        else:
            t = line.strip()
            if t: add(_par("Normal", t))

    flush_table()

    cover_parts = _cover_xml(
        cover["name"], cover["date"], cover["subs"], cover["details"]
    )
    return cover_parts + body, img_reg


# ══════════════════════════════════════════════════════════════════════
#  Document XML wrapper
# ══════════════════════════════════════════════════════════════════════

_DOC_NS = (
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
)

_SECT_PR = (
    '<w:sectPr>'
    '<w:pgSz w:w="11906" w:h="16838"/>'
    '<w:pgMar w:top="1080" w:right="1080" w:bottom="1080" w:left="1080"'
    ' w:header="709" w:footer="709" w:gutter="0"/>'
    '</w:sectPr>'
)


def _build_document_xml(elements):
    body_content = "\n".join(elements)
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<w:document {_DOC_NS}>\n'
        f'<w:body>\n{body_content}\n{_SECT_PR}\n'
        f'</w:body>\n</w:document>'
    ).encode("utf-8")


# ══════════════════════════════════════════════════════════════════════
#  Main entry point
# ══════════════════════════════════════════════════════════════════════

def create_docx(input_path, output_path):
    with open(input_path, encoding="utf-8") as fh:
        lines = fh.readlines()

    elements, img_reg = _parse(lines)
    doc_xml   = _build_document_xml(elements)
    styles    = _styles_xml()

    img_exts  = {e["zip_name"].rsplit(".", 1)[-1] for e in img_reg}
    ct_xml    = _content_types_xml(img_exts)
    img_rels  = [(e["rel_id"], e["zip_name"]) for e in img_reg]
    word_rels = _word_rels_xml(img_rels)

    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml",           ct_xml)
        zf.writestr("_rels/.rels",                   ROOT_RELS)
        zf.writestr("word/document.xml",             doc_xml)
        zf.writestr("word/styles.xml",               styles)
        zf.writestr("word/settings.xml",             SETTINGS)
        zf.writestr("word/_rels/document.xml.rels",  word_rels)
        for entry in img_reg:
            zf.writestr(f'word/{entry["zip_name"]}', entry["data"])

    size_kb = os.path.getsize(output_path) // 1024
    n_imgs  = len(img_reg)
    print(f"SAVED:{output_path} ({size_kb} KB, {n_imgs} image(s))")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python3 hd_to_docx.py <input.txt> [output.docx]", file=sys.stderr)
        sys.exit(1)
    inp = sys.argv[1]
    out = sys.argv[2] if len(sys.argv) >= 3 else os.path.splitext(inp)[0] + ".docx"
    create_docx(inp, out)
