#!/usr/bin/env python3
"""
hd_to_docx.py — Human Design Thai Report → DOCX
Pure Python stdlib (zipfile only). No pip install required.

Usage:
    python3 hd_to_docx.py report.txt [output.docx]

Input format — one tag per line:
    [COVER_NAME]   ชื่อผู้รับการวิเคราะห์
    [COVER_DATE]   วันเกิด
    [COVER_SUB]    โครงสร้างหลัก • Variables • Shadow Chart
    [PAGEBREAK]
    [H1]  บทที่ 1: หัวข้อหลัก
    [H2]  หัวข้อรอง
    [H3]  หัวข้อย่อย
    [INTRO]  ข้อความแนะนำสำหรับผู้ใหม่ (italic, gray, left border)
    [BODY]   ย่อหน้าปกติ
    [BULLET] รายการ bullet
    [SUBBULLET] รายการย่อย
    [TABLE_HEADER] คอลัมน์1 | คอลัมน์2 | คอลัมน์3
    [TABLE_ROW]    ค่า1 | ค่า2 | ค่า3
    (blank line = small spacer)
"""

import sys
import os
import zipfile
import datetime
from xml.sax.saxutils import escape as xmlesc

# ── Fonts & Colors ──────────────────────────────────────────────────
FONT   = "TH Sarabun New"
C_H1   = "1F497D"   # dark navy
C_H2   = "2E74B5"   # medium blue
C_H3   = "4472C4"   # lighter blue
C_INT  = "595959"   # intro gray
C_THBG = "1F497D"   # table header bg
C_THFG = "FFFFFF"   # table header text

# ── Static ZIP entries ───────────────────────────────────────────────

CONTENT_TYPES = b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml"  ContentType="application/xml"/>
  <Override PartName="/word/document.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/settings.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
</Types>"""

ROOT_RELS = b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    Target="word/document.xml"/>
</Relationships>"""

WORD_RELS = b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
    Target="styles.xml"/>
  <Relationship Id="rId2"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"
    Target="settings.xml"/>
</Relationships>"""

SETTINGS = b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:defaultTabStop w:val="720"/>
  <w:compat>
    <w:compatSetting w:name="compatibilityMode"
      w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>
  </w:compat>
</w:settings>"""

def _styles():
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
</w:styles>""".encode("utf-8")

# ── XML builders ─────────────────────────────────────────────────────

def _rpr(bold=False, italic=False, color=None, size=26):
    f = FONT
    parts = [f'<w:rFonts w:ascii="{f}" w:hAnsi="{f}" w:cs="{f}" w:eastAsia="{f}"/>']
    if bold:   parts += ["<w:b/>", "<w:bCs/>"]
    if italic: parts += ["<w:i/>", "<w:iCs/>"]
    if color:  parts.append(f'<w:color w:val="{color}"/>')
    parts.append(f'<w:sz w:val="{size}"/><w:szCs w:val="{size}"/>')
    return "<w:rPr>" + "".join(parts) + "</w:rPr>"

def _run(text, bold=False, italic=False, color=None, size=26):
    t = xmlesc(str(text))
    return f'<w:r>{_rpr(bold,italic,color,size)}<w:t xml:space="preserve">{t}</w:t></w:r>'

def _par(style, text, bold=False, italic=False, color=None, size=26, extra_ppr=""):
    return (f'<w:p><w:pPr><w:pStyle w:val="{style}"/>{extra_ppr}</w:pPr>'
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

    def _borders():
        sides = ["top", "left", "bottom", "right", "insideH", "insideV"]
        return "<w:tblBorders>" + "".join(
            f'<w:{s} w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>'
            for s in sides
        ) + "</w:tblBorders>"

    def _cell(text, header=False):
        fill   = f'<w:shd w:val="clear" w:color="auto" w:fill="{C_THBG}"/>' if header else ""
        color  = C_THFG if header else None
        align  = "center" if header else "left"
        sz     = 22
        return (f'<w:tc>'
                f'<w:tcPr><w:tcW w:w="{col_w}" w:type="dxa"/>{fill}</w:tcPr>'
                f'<w:p><w:pPr><w:jc w:val="{align}"/></w:pPr>'
                f'{_run(text.strip(), bold=header, color=color, size=sz)}</w:p>'
                f'</w:tc>')

    tbl_pr = (f'<w:tblPr>'
              f'<w:tblW w:w="0" w:type="auto"/>'
              f'{_borders()}'
              f'<w:tblCellMar>'
              f'<w:top w:w="80" w:type="dxa"/><w:left w:w="100" w:type="dxa"/>'
              f'<w:bottom w:w="80" w:type="dxa"/><w:right w:w="100" w:type="dxa"/>'
              f'</w:tblCellMar>'
              f'</w:tblPr>')

    hdr_row = "<w:tr>" + "".join(_cell(h, header=True) for h in headers) + "</w:tr>"

    data_rows = ""
    for row in rows:
        padded = list(row) + [""] * (col_n - len(row))
        data_rows += "<w:tr>" + "".join(_cell(c) for c in padded[:col_n]) + "</w:tr>"

    return f"<w:tbl>{tbl_pr}{hdr_row}{data_rows}</w:tbl>" + _spacer()

# ── Cover builder ────────────────────────────────────────────────────

def _cover(name, date, subs, details):
    parts = []
    parts.append(_par("HDCoverTitle", "Human Design Chart", bold=True))
    parts.append(_par("HDCoverSub", "การแปลผลฉบับสมบูรณ์"))
    if name:
        line = name + (f"  •  {date}" if date else "")
        parts.append(_par("HDCoverSub", line))
    for s in subs:
        parts.append(_par("HDCoverSub", s))
    for d in details:
        parts.append(_par("HDCoverDetail", d, color="888888", size=22))
    today = datetime.date.today().strftime("%d/%m/%Y")
    parts.append(_par("HDCoverDetail", f"วันที่จัดทำ: {today}", color="888888", size=22))
    parts.append(_page_break())
    return parts

# ── Parser ───────────────────────────────────────────────────────────

def _parse(lines):
    cover = dict(name="", date="", subs=[], details=[])
    body  = []
    pending_headers = None
    pending_rows    = []

    def flush_table():
        nonlocal pending_headers, pending_rows
        if pending_headers is not None:
            body.append(_table(pending_headers, pending_rows))
        pending_headers = None
        pending_rows    = []

    def add(xml):
        flush_table()
        body.append(xml)

    for raw in lines:
        line = raw.rstrip("\n")

        if   line.startswith("[COVER_NAME]"):   cover["name"] = line[12:].strip()
        elif line.startswith("[COVER_DATE]"):   cover["date"] = line[12:].strip()
        elif line.startswith("[COVER_SUB]"):    cover["subs"].append(line[11:].strip())
        elif line.startswith("[COVER_DETAIL]"): cover["details"].append(line[14:].strip())

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

        elif line.startswith("[TABLE_HEADER]"):
            flush_table()
            pending_headers = [c.strip() for c in line[14:].split("|")]
            pending_rows    = []

        elif line.startswith("[TABLE_ROW]"):
            if pending_headers is not None:
                pending_rows.append([c.strip() for c in line[11:].split("|")])

        elif line.strip() == "":
            flush_table()
            body.append(_spacer())

        else:
            # untagged line → treat as body text
            t = line.strip()
            if t: add(_par("Normal", t))

    flush_table()

    cover_parts = _cover(cover["name"], cover["date"], cover["subs"], cover["details"])
    return cover_parts + body

# ── Document XML wrapper ─────────────────────────────────────────────

_DOC_NS = (
    'xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" '
    'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" '
    'xmlns:v="urn:schemas-microsoft-com:vml" '
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    'mc:Ignorable="w14 wp14"'
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
        f'<w:body>\n'
        f'{body_content}\n'
        f'{_SECT_PR}\n'
        f'</w:body>\n'
        f'</w:document>'
    ).encode("utf-8")

# ── Main ─────────────────────────────────────────────────────────────

def create_docx(input_path, output_path):
    with open(input_path, encoding="utf-8") as fh:
        lines = fh.readlines()

    elements   = _parse(lines)
    doc_xml    = _build_document_xml(elements)
    styles_xml = _styles()

    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml",           CONTENT_TYPES)
        zf.writestr("_rels/.rels",                   ROOT_RELS)
        zf.writestr("word/document.xml",             doc_xml)
        zf.writestr("word/styles.xml",               styles_xml)
        zf.writestr("word/settings.xml",             SETTINGS)
        zf.writestr("word/_rels/document.xml.rels",  WORD_RELS)

    size_kb = os.path.getsize(output_path) // 1024
    print(f"SAVED:{output_path} ({size_kb} KB)")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python3 hd_to_docx.py <input.txt> [output.docx]", file=sys.stderr)
        sys.exit(1)
    inp = sys.argv[1]
    out = sys.argv[2] if len(sys.argv) >= 3 else os.path.splitext(inp)[0] + ".docx"
    create_docx(inp, out)
