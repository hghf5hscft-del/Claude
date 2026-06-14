#!/usr/bin/env python3
"""
hd_to_docx.py — Human Design Thai Report → DOCX
stdlib only: zipfile + xml.sax.saxutils. No pip install needed.

Usage:
  python3 hd_to_docx.py report.txt [output.docx]

Input: plain text file with one tag per line.
Tags:
  [COVER_NAME] ชื่อ
  [COVER_DATE] วันเกิด
  [COVER_SUB]  ข้อความบรรทัดรอง (ใส่ได้หลายบรรทัด)
  [PAGEBREAK]
  [H1] หัวข้อหลัก
  [H2] หัวข้อรอง
  [H3] หัวข้อย่อย
  [INTRO] ย่อหน้าแนะนำ (italic, gray, left border)
  [BODY] ย่อหน้าปกติ
  [BULLET] รายการ
  [SUBBULLET] รายการย่อย
  [TABLE_HEADER] คอล1 | คอล2 | คอล3
  [TABLE_ROW]    ค่า1  | ค่า2  | ค่า3
  (blank line = spacer)
  Lines without a tag are treated as [BODY].
"""

import sys, os, zipfile, datetime
from xml.sax.saxutils import escape as xe

# ── Fonts & colours ────────────────────────────────────────────────
FONT   = "TH Sarabun New"
C_H1   = "1F497D"
C_H2   = "2E74B5"
C_H3   = "4472C4"
C_GRY  = "595959"
C_TBG  = "1F497D"
C_TFG  = "FFFFFF"

# ── Static ZIP members (bytes) ─────────────────────────────────────
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

# ── styles.xml (static, Thai font throughout) ──────────────────────
def make_styles():
    def fnt(sz, bold=False, italic=False, color=None, extra=""):
        b = "<w:b/><w:bCs/>" if bold else ""
        i = "<w:i/><w:iCs/>" if italic else ""
        c = f'<w:color w:val="{color}"/>' if color else ""
        return (f'<w:rFonts w:ascii="{FONT}" w:hAnsi="{FONT}" w:cs="{FONT}" w:eastAsia="{FONT}"/>'
                f'{b}{i}{c}<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/>{extra}')

    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault><w:rPr>
      {fnt(26)}
      <w:lang w:val="th-TH" w:eastAsia="th-TH"/>
    </w:rPr></w:rPrDefault>
    <w:pPrDefault><w:pPr>
      <w:spacing w:after="80"/><w:jc w:val="both"/>
    </w:pPr></w:pPrDefault>
  </w:docDefaults>

  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:rPr>{fnt(26)}</w:rPr>
  </w:style>

  <w:style w:type="paragraph" w:styleId="HDH1">
    <w:name w:val="HD Heading 1"/><w:basedOn w:val="Normal"/>
    <w:pPr>
      <w:spacing w:before="280" w:after="120"/><w:jc w:val="left"/>
      <w:pBdr><w:bottom w:val="single" w:sz="8" w:space="4" w:color="{C_H1}"/></w:pBdr>
    </w:pPr>
    <w:rPr>{fnt(40, bold=True, color=C_H1)}</w:rPr>
  </w:style>

  <w:style w:type="paragraph" w:styleId="HDH2">
    <w:name w:val="HD Heading 2"/><w:basedOn w:val="Normal"/>
    <w:pPr><w:spacing w:before="160" w:after="80"/><w:jc w:val="left"/></w:pPr>
    <w:rPr>{fnt(32, bold=True, color=C_H2)}</w:rPr>
  </w:style>

  <w:style w:type="paragraph" w:styleId="HDH3">
    <w:name w:val="HD Heading 3"/><w:basedOn w:val="Normal"/>
    <w:pPr><w:spacing w:before="100" w:after="60"/><w:jc w:val="left"/></w:pPr>
    <w:rPr>{fnt(28, bold=True, color=C_H3)}</w:rPr>
  </w:style>

  <w:style w:type="paragraph" w:styleId="HDIntro">
    <w:name w:val="HD Intro"/><w:basedOn w:val="Normal"/>
    <w:pPr>
      <w:ind w:left="360"/>
      <w:spacing w:before="60" w:after="60"/>
      <w:jc w:val="both"/>
      <w:pBdr><w:left w:val="single" w:sz="12" w:space="8" w:color="{C_H3}"/></w:pBdr>
    </w:pPr>
    <w:rPr>{fnt(24, italic=True, color=C_GRY)}</w:rPr>
  </w:style>

  <w:style w:type="paragraph" w:styleId="HDBullet">
    <w:name w:val="HD Bullet"/><w:basedOn w:val="Normal"/>
    <w:pPr>
      <w:ind w:left="400" w:hanging="200"/>
      <w:spacing w:after="60"/>
    </w:pPr>
    <w:rPr>{fnt(26)}</w:rPr>
  </w:style>

  <w:style w:type="paragraph" w:styleId="HDSubBullet">
    <w:name w:val="HD Sub Bullet"/><w:basedOn w:val="Normal"/>
    <w:pPr>
      <w:ind w:left="720" w:hanging="200"/>
      <w:spacing w:after="40"/>
    </w:pPr>
    <w:rPr>{fnt(24)}</w:rPr>
  </w:style>

  <w:style w:type="paragraph" w:styleId="HDCoverTitle">
    <w:name w:val="HD Cover Title"/><w:basedOn w:val="Normal"/>
    <w:pPr><w:spacing w:before="600" w:after="100"/><w:jc w:val="center"/></w:pPr>
    <w:rPr>{fnt(56, bold=True, color=C_H1)}</w:rPr>
  </w:style>

  <w:style w:type="paragraph" w:styleId="HDCoverSub">
    <w:name w:val="HD Cover Sub"/><w:basedOn w:val="Normal"/>
    <w:pPr><w:spacing w:before="80" w:after="80"/><w:jc w:val="center"/></w:pPr>
    <w:rPr>{fnt(30, color=C_H2)}</w:rPr>
  </w:style>

  <w:style w:type="paragraph" w:styleId="HDCoverDetail">
    <w:name w:val="HD Cover Detail"/><w:basedOn w:val="Normal"/>
    <w:pPr><w:spacing w:before="40" w:after="40"/><w:jc w:val="center"/></w:pPr>
    <w:rPr>{fnt(22, color="888888")}</w:rPr>
  </w:style>
</w:styles>""".encode("utf-8")


# ── XML paragraph / table builders ────────────────────────────────
def _rpr(size=26, bold=False, italic=False, color=None):
    b = "<w:b/><w:bCs/>" if bold else ""
    i = "<w:i/><w:iCs/>" if italic else ""
    c = f'<w:color w:val="{color}"/>' if color else ""
    return (f'<w:rPr>'
            f'<w:rFonts w:ascii="{FONT}" w:hAnsi="{FONT}" w:cs="{FONT}" w:eastAsia="{FONT}"/>'
            f'{b}{i}{c}'
            f'<w:sz w:val="{size}"/><w:szCs w:val="{size}"/>'
            f'</w:rPr>')

def _run(text, size=26, bold=False, italic=False, color=None):
    return f'<w:r>{_rpr(size,bold,italic,color)}<w:t xml:space="preserve">{xe(text)}</w:t></w:r>'

def p_styled(style, text, size=26, bold=False, italic=False, color=None):
    return (f'<w:p><w:pPr><w:pStyle w:val="{style}"/></w:pPr>'
            f'{_run(text, size, bold, italic, color)}</w:p>')

def p_break():
    return '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'

def p_space():
    return '<w:p><w:pPr><w:spacing w:after="40"/></w:pPr></w:p>'

def p_bullet(text, sub=False):
    style = "HDSubBullet" if sub else "HDBullet"
    char  = "   ◦ " if sub else "• "
    sz    = 24 if sub else 26
    return (f'<w:p><w:pPr><w:pStyle w:val="{style}"/></w:pPr>'
            f'{_run(char, sz)}{_run(text, sz)}</w:p>')

def p_table(headers, rows):
    n   = max(1, len(headers))
    cw  = str(max(1, 9000 // n))
    bdr = "".join(
        f'<w:{b} w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>'
        for b in ["top","left","bottom","right","insideH","insideV"]
    )

    def cell(txt, hdr=False):
        fill = f'<w:shd w:val="clear" w:color="auto" w:fill="{C_TBG}"/>' if hdr else ""
        fg   = C_TFG if hdr else None
        jc   = "center" if hdr else "left"
        sz   = 22
        return (f'<w:tc>'
                f'<w:tcPr><w:tcW w:w="{cw}" w:type="dxa"/>{fill}'
                f'<w:tcMar>'
                f'<w:top w:w="60" w:type="dxa"/><w:left w:w="80" w:type="dxa"/>'
                f'<w:bottom w:w="60" w:type="dxa"/><w:right w:w="80" w:type="dxa"/>'
                f'</w:tcMar></w:tcPr>'
                f'<w:p><w:pPr><w:jc w:val="{jc}"/></w:pPr>'
                f'{_run(str(txt).strip(), sz, bold=hdr, color=fg)}</w:p></w:tc>')

    hrow = "<w:tr>" + "".join(cell(h, True) for h in headers) + "</w:tr>"
    drows = ""
    for row in rows:
        padded = list(row) + [""] * (n - len(row))
        drows += "<w:tr>" + "".join(cell(c) for c in padded[:n]) + "</w:tr>"

    return (f'<w:tbl>'
            f'<w:tblPr>'
            f'<w:tblW w:w="0" w:type="auto"/>'
            f'<w:tblBorders>{bdr}</w:tblBorders>'
            f'</w:tblPr>'
            f'{hrow}{drows}</w:tbl>'
            f'{p_space()}')


# ── Parser ─────────────────────────────────────────────────────────
def parse_lines(lines):
    cover = dict(name="", date="", subs=[], details=[])
    body  = []
    tbl_h = None
    tbl_r = []

    def flush_tbl():
        nonlocal tbl_h, tbl_r
        if tbl_h is not None:
            body.append(p_table(tbl_h, tbl_r))
        tbl_h = None
        tbl_r = []

    def emit(xml):
        flush_tbl()
        body.append(xml)

    for raw in lines:
        ln = raw.rstrip("\n")

        if   ln.startswith("[COVER_NAME]"):   cover["name"] = ln[12:].strip()
        elif ln.startswith("[COVER_DATE]"):   cover["date"] = ln[12:].strip()
        elif ln.startswith("[COVER_SUB]"):    cover["subs"].append(ln[11:].strip())
        elif ln.startswith("[COVER_DETAIL]"): cover["details"].append(ln[14:].strip())

        elif ln.startswith("[PAGEBREAK]"):  emit(p_break())
        elif ln.startswith("[H1]"):
            t = ln[4:].strip()
            if t: emit(p_styled("HDH1", t))
        elif ln.startswith("[H2]"):
            t = ln[4:].strip()
            if t: emit(p_styled("HDH2", t))
        elif ln.startswith("[H3]"):
            t = ln[4:].strip()
            if t: emit(p_styled("HDH3", t))
        elif ln.startswith("[INTRO]"):
            t = ln[7:].strip()
            if t: emit(p_styled("HDIntro", t, size=24, italic=True, color=C_GRY))
        elif ln.startswith("[BODY]"):
            t = ln[6:].strip()
            if t: emit(p_styled("Normal", t))
        elif ln.startswith("[BULLET]"):
            t = ln[8:].strip()
            if t:
                flush_tbl()
                body.append(p_bullet(t))
        elif ln.startswith("[SUBBULLET]"):
            t = ln[11:].strip()
            if t:
                flush_tbl()
                body.append(p_bullet(t, sub=True))
        elif ln.startswith("[TABLE_HEADER]"):
            flush_tbl()
            tbl_h = [c.strip() for c in ln[14:].split("|")]
            tbl_r = []
        elif ln.startswith("[TABLE_ROW]"):
            if tbl_h is not None:
                tbl_r.append([c.strip() for c in ln[11:].split("|")])
        elif ln.strip() == "":
            flush_tbl()
            body.append(p_space())
        else:
            t = ln.strip()
            if t: emit(p_styled("Normal", t))

    flush_tbl()

    # ── Cover page ──
    today = datetime.date.today().strftime("%d/%m/%Y")
    cover_xml = [p_styled("HDCoverTitle", "Human Design Chart")]
    cover_xml.append(p_styled("HDCoverSub", "การแปลผลฉบับสมบูรณ์"))
    if cover["name"]:
        sub = cover["name"] + (f"  •  {cover['date']}" if cover["date"] else "")
        cover_xml.append(p_styled("HDCoverSub", sub))
    for s in cover["subs"]:
        cover_xml.append(p_styled("HDCoverSub", s))
    for d in cover["details"]:
        cover_xml.append(p_styled("HDCoverDetail", d))
    cover_xml.append(p_styled("HDCoverDetail", f"วันที่จัดทำ: {today}"))
    cover_xml.append(p_break())

    return cover_xml + body


# ── Document XML wrapper ───────────────────────────────────────────
DOC_NS = (
    'xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" '
    'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" '
    'xmlns:v="urn:schemas-microsoft-com:vml" '
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    'xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" '
    'mc:Ignorable="w14 wp14"'
)
SECTPR = (
    '<w:sectPr>'
    '<w:pgSz w:w="11906" w:h="16838"/>'
    '<w:pgMar w:top="1080" w:right="1080" w:bottom="1080" w:left="1080"'
    ' w:header="709" w:footer="709" w:gutter="0"/>'
    '</w:sectPr>'
)

def build_doc_xml(parts):
    body = "\n".join(parts)
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            f'<w:document {DOC_NS}>\n'
            f'<w:body>\n{body}\n{SECTPR}\n'
            f'</w:body>\n</w:document>').encode("utf-8")


# ── Main ───────────────────────────────────────────────────────────
def create_docx(input_path, output_path):
    with open(input_path, encoding="utf-8") as f:
        lines = f.readlines()

    parts   = parse_lines(lines)
    doc_xml = build_doc_xml(parts)
    styles  = make_styles()

    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml",          CONTENT_TYPES)
        zf.writestr("_rels/.rels",                  ROOT_RELS)
        zf.writestr("word/document.xml",            doc_xml)
        zf.writestr("word/styles.xml",              styles)
        zf.writestr("word/settings.xml",            SETTINGS)
        zf.writestr("word/_rels/document.xml.rels", WORD_RELS)

    kb = os.path.getsize(output_path) // 1024
    print(f"SAVED:{output_path} ({kb} KB)")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python3 hd_to_docx.py <input.txt> [output.docx]", file=sys.stderr)
        sys.exit(1)
    inp = sys.argv[1]
    out = sys.argv[2] if len(sys.argv) >= 3 else os.path.splitext(inp)[0] + ".docx"
    create_docx(inp, out)
