import zipfile, os, re, textwrap
from datetime import datetime

MD = "/home/user/Claude/HD_Reading_PFai.md"
OUT = "/home/user/Claude/HD_Reading_PFai.docx"

with open(MD, encoding="utf-8") as f:
    raw = f.read()

# ── helpers ──────────────────────────────────────────────────────────────────
def esc(t):
    return t.replace("&","&amp;").replace("<","&lt;").replace(">","&gt;").replace('"',"&quot;")

def run(text, bold=False, italic=False, color=None, sz=None):
    rpr = "<w:rPr>"
    if bold:   rpr += "<w:b/>"
    if italic: rpr += "<w:i/>"
    if color:  rpr += f'<w:color w:val="{color}"/>'
    if sz:     rpr += f'<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/>'
    rpr += "</w:rPr>"
    return f"<w:r>{rpr}<w:t xml:space=\"preserve\">{esc(text)}</w:t></w:r>"

def para(runs_xml, style="Normal", spacing_before=0, spacing_after=120, indent=0):
    ind = f'<w:ind w:left="{indent}"/>' if indent else ""
    return (f'<w:p><w:pPr><w:pStyle w:val="{style}"/>'
            f'<w:spacing w:before="{spacing_before}" w:after="{spacing_after}"/>'
            f'{ind}</w:pPr>{runs_xml}</w:p>')

def tbl_cell(content_xml, color_fill=None, bold=False, width=2000):
    fill = f'<w:shd w:val="clear" w:color="auto" w:fill="{color_fill}"/>' if color_fill else ""
    return (f'<w:tc><w:tcPr><w:tcW w:w="{width}" w:type="dxa"/>{fill}</w:tcPr>'
            f'<w:p><w:pPr><w:spacing w:before="40" w:after="40"/></w:pPr>'
            f'{run(content_xml, bold=bold)}</w:p></w:tc>')

# ── parse markdown → docx XML ─────────────────────────────────────────────
paras = []

# Title block
paras.append(para(
    run("การอ่านแผนภูมิ Human Design ของ พี่ฝ้าย", bold=True, sz=36, color="1F4E79"),
    style="Normal", spacing_before=0, spacing_after=80))
paras.append(para(
    run("P'Fai S.  |  Splenic Projector  |  Profile 4/6  |  RAX Service 3", sz=22, color="2E74B5"),
    spacing_after=60))
paras.append(para(
    run("วิเคราะห์จาก geneticmatrix.com — แผนภูมิ 24 หน้า", italic=True, sz=20, color="595959"),
    spacing_after=200))

# ── table-parsing state
in_table = False
table_rows = []

def flush_table():
    global table_rows
    if not table_rows: return
    xml = "<w:tbl><w:tblPr><w:tblW w:w=\"9072\" w:type=\"dxa\"/><w:tblBorders>"
    for side in ("top","left","bottom","right","insideH","insideV"):
        xml += f'<w:{side} w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>'
    xml += "</w:tblBorders></w:tblPr>"
    for ri, row in enumerate(table_rows):
        xml += "<w:tr>"
        cells = [c.strip() for c in row.strip("|").split("|")]
        is_header = ri == 0
        fill = "2E74B5" if is_header else ("F2F7FF" if ri % 2 == 0 else "FFFFFF")
        fc = "FFFFFF" if is_header else None
        for ci, cell in enumerate(cells):
            w = 9072 // max(len(cells), 1)
            inner = run(cell, bold=is_header, color="FFFFFF" if is_header else None)
            fill2 = "2E74B5" if is_header else ("EEF4FB" if ri % 2 == 1 else "FFFFFF")
            xml += (f'<w:tc><w:tcPr><w:tcW w:w="{w}" w:type="dxa"/>'
                    f'<w:shd w:val="clear" w:color="auto" w:fill="{fill2}"/></w:tcPr>'
                    f'<w:p><w:pPr><w:spacing w:before="60" w:after="60"/></w:pPr>{inner}</w:p></w:tc>')
        xml += "</w:tr>"
    xml += "</w:tbl>"
    xml += '<w:p><w:pPr><w:spacing w:before="0" w:after="120"/></w:pPr></w:p>'
    paras.append(xml)
    table_rows = []

lines = raw.split("\n")
i = 0
while i < len(lines):
    line = lines[i]

    # Table row
    if line.strip().startswith("|") and "|" in line[1:]:
        in_table = True
        # skip separator rows
        if re.match(r'^\s*\|[-| :]+\|\s*$', line):
            i += 1
            continue
        table_rows.append(line)
        i += 1
        continue
    else:
        if in_table:
            flush_table()
            in_table = False

    # Skip horizontal rules and empty lines within table
    if re.match(r'^---+$', line.strip()):
        paras.append('<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="2E74B5"/></w:pBdr><w:spacing w:before="120" w:after="120"/></w:pPr></w:p>')
        i += 1
        continue

    if not line.strip():
        paras.append('<w:p><w:pPr><w:spacing w:before="0" w:after="80"/></w:pPr></w:p>')
        i += 1
        continue

    # H1
    if line.startswith("# ") and not line.startswith("## "):
        txt = line[2:].strip()
        paras.append(para(run(txt, bold=True, sz=32, color="1F4E79"), spacing_before=240, spacing_after=120))
        i += 1
        continue

    # H2
    if line.startswith("## "):
        txt = line[3:].strip()
        num = re.match(r'^(\d+)\.\s*(.*)', txt)
        if num:
            paras.append(para(
                run(f"  {num.group(1)}.  ", bold=True, sz=26, color="FFFFFF") +
                run(f" {num.group(2)}", bold=True, sz=26, color="FFFFFF"),
                spacing_before=280, spacing_after=100))
            # heading background via shd — simplify to colored text
            paras.pop()
            paras.append(para(run(f"{num.group(1)}. {num.group(2)}", bold=True, sz=28, color="1F4E79"),
                               spacing_before=280, spacing_after=100))
        else:
            paras.append(para(run(txt, bold=True, sz=28, color="1F4E79"),
                               spacing_before=280, spacing_after=100))
        i += 1
        continue

    # H3
    if line.startswith("### "):
        txt = line[4:].strip()
        paras.append(para(run(txt, bold=True, sz=24, color="2E74B5"), spacing_before=180, spacing_after=80))
        i += 1
        continue

    # blockquote
    if line.startswith("> "):
        txt = line[2:].strip()
        # collect multi-line
        bq = txt
        while i+1 < len(lines) and lines[i+1].startswith("> "):
            i += 1
            bq += " " + lines[i][2:].strip()
        inner = run(bq, italic=True, color="1F4E79")
        paras.append(para(inner, indent=720, spacing_before=120, spacing_after=120))
        i += 1
        continue

    # bullet
    if line.startswith("- "):
        txt = line[2:].strip()
        # inline bold/italic
        def inline(t):
            parts = re.split(r'(\*\*[^*]+\*\*|\*[^*]+\*)', t)
            out = ""
            for p in parts:
                if p.startswith("**") and p.endswith("**"):
                    out += run(p[2:-2], bold=True)
                elif p.startswith("*") and p.endswith("*"):
                    out += run(p[1:-1], italic=True)
                else:
                    out += run(p) if p else ""
            return out
        bullet_run = "•  "
        paras.append(para(run(bullet_run, bold=True, color="2E74B5") + inline(txt), indent=360, spacing_before=40, spacing_after=40))
        i += 1
        continue

    # normal paragraph — handle inline **bold** and *italic*
    def inline_para(t):
        parts = re.split(r'(\*\*[^*]+\*\*|\*[^*]+\*)', t)
        out = ""
        for p in parts:
            if p.startswith("**") and p.endswith("**"):
                out += run(p[2:-2], bold=True)
            elif p.startswith("*") and p.endswith("*"):
                out += run(p[1:-1], italic=True)
            else:
                out += run(p) if p else ""
        return out

    paras.append(para(inline_para(line.strip()), spacing_before=0, spacing_after=100))
    i += 1

if in_table:
    flush_table()

# ── assemble document.xml ─────────────────────────────────────────────────
body = "\n".join(paras)
doc_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"
  xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
  xmlns:o="urn:schemas-microsoft-com:office:office"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
  xmlns:v="urn:schemas-microsoft-com:vml"
  xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
  xmlns:w10="urn:schemas-microsoft-com:office:word"
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
  xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
  xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk"
  xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"
  xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
  mc:Ignorable="w14 wp14">
<w:body>
{body}
<w:sectPr>
  <w:pgSz w:w="12240" w:h="15840"/>
  <w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134"
           w:header="708" w:footer="708" w:gutter="0"/>
</w:sectPr>
</w:body>
</w:document>'''

styles_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:docDefaults>
    <w:rPrDefault><w:rPr>
      <w:rFonts w:ascii="TH Sarabun New" w:hAnsi="TH Sarabun New" w:cs="TH Sarabun New"/>
      <w:sz w:val="24"/><w:szCs w:val="24"/>
      <w:lang w:val="th-TH" w:eastAsia="th-TH" w:bidi="th-TH"/>
    </w:rPr></w:rPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:styleId="Normal" w:default="1">
    <w:name w:val="Normal"/>
    <w:pPr><w:spacing w:after="120"/></w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="TH Sarabun New" w:hAnsi="TH Sarabun New" w:cs="TH Sarabun New"/>
      <w:sz w:val="24"/><w:szCs w:val="24"/>
    </w:rPr>
  </w:style>
</w:styles>'''

settings_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:defaultTabStop w:val="720"/>
  <w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>
</w:settings>'''

ct_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
</Types>'''

rels_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''

word_rels_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
</Relationships>'''

with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", ct_xml)
    z.writestr("_rels/.rels", rels_xml)
    z.writestr("word/document.xml", doc_xml)
    z.writestr("word/styles.xml", styles_xml)
    z.writestr("word/settings.xml", settings_xml)
    z.writestr("word/_rels/document.xml.rels", word_rels_xml)

print(f"Created: {OUT}  ({os.path.getsize(OUT):,} bytes)")
