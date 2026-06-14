#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import zipfile, os

OUTPUT = "/home/user/Claude/JJZ_Human_Design_Analysis.docx"
UPLOAD = "/root/.claude/uploads/81a8195e-f9a8-5b45-8103-2143cdddc351/"
INCH = 914400
W6 = int(6 * INCH)
H4 = int(4 * INCH)
H35 = int(3.5 * INCH)
H42 = int(4.2 * INCH)

IMAGES = [
    ("rId10", "5486667a-IMG_2031.jpeg", "img001.jpeg"),
    ("rId11", "fc0d968b-IMG_2034.jpeg", "img002.jpeg"),
    ("rId12", "653a0c62-IMG_2039.jpeg", "img003.jpeg"),
    ("rId13", "a84f8787-IMG_2037.jpeg", "img004.jpeg"),
    ("rId14", "f1c79847-IMG_2051.jpeg", "img005.jpeg"),
    ("rId15", "b8261110-IMG_2050.jpeg", "img006.jpeg"),
]

def xe(t):
    return str(t).replace('&','&amp;').replace('<','&lt;').replace('>','&gt;').replace('"','&quot;')

def p(text="", bold=False, size=22, color="000000", before=60, after=60, indent=0, center=False, italic=False):
    jc = '<w:jc w:val="center"/>' if center else ''
    ind = f'<w:ind w:left="{indent}"/>' if indent else ''
    b = '<w:b/>' if bold else ''
    it = '<w:i/>' if italic else ''
    runs = []
    for i, part in enumerate(str(text).split('\n')):
        if i > 0:
            runs.append('<w:r><w:br/></w:r>')
        if part:
            runs.append(f'<w:r><w:rPr>{b}{it}<w:color w:val="{color}"/><w:sz w:val="{size}"/><w:szCs w:val="{size}"/></w:rPr><w:t xml:space="preserve">{xe(part)}</w:t></w:r>')
    return f'<w:p><w:pPr><w:spacing w:before="{before}" w:after="{after}"/>{jc}{ind}</w:pPr>{"".join(runs)}</w:p>'

def h(text, level=1):
    cfg = {1:("36","1E5CBF",240,120), 2:("28","2D6B4A",180,90), 3:("24","8B4000",140,70), 4:("22","555555",100,50)}
    sz,color,before,after = cfg.get(level, ("22","000000",80,40))
    return f'<w:p><w:pPr><w:spacing w:before="{before}" w:after="{after}"/><w:keepNext/></w:pPr><w:r><w:rPr><w:b/><w:color w:val="{color}"/><w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/></w:rPr><w:t>{xe(text)}</w:t></w:r></w:p>'

def bullet(text, size=22):
    return p(f"•  {text}", size=size, indent=360)

def img(rid, doc_id, w=W6, ht=H4):
    return f'''<w:p><w:pPr><w:jc w:val="center"/><w:spacing w:before="120" w:after="120"/></w:pPr><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><wp:extent cx="{w}" cy="{ht}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="{doc_id}" name="Pic{doc_id}"/><wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="{doc_id}" name="Pic{doc_id}"/><pic:cNvPicPr><a:picLocks noChangeAspect="1"/></pic:cNvPicPr></pic:nvPicPr><pic:blipFill><a:blip r:embed="{rid}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{w}" cy="{ht}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>'''

def sep():
    return '<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="2D6B4A"/></w:pBdr><w:spacing w:before="100" w:after="100"/></w:pPr></w:p>'

def pb():
    return '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'

def tbl(headers, rows, col_widths=None):
    n = len(headers)
    if not col_widths:
        w_each = 9000 // n
        col_widths = [w_each] * n
    x = f'<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="AAAAAA"/><w:left w:val="single" w:sz="4" w:space="0" w:color="AAAAAA"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="AAAAAA"/><w:right w:val="single" w:sz="4" w:space="0" w:color="AAAAAA"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="AAAAAA"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="AAAAAA"/></w:tblBorders></w:tblPr><w:tblGrid>'
    for cw in col_widths:
        x += f'<w:gridCol w:w="{cw}"/>'
    x += '</w:tblGrid><w:tr>'
    for i,cell in enumerate(headers):
        x += f'<w:tc><w:tcPr><w:tcW w:w="{col_widths[i]}" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="1E5CBF"/></w:tcPr><w:p><w:r><w:rPr><w:b/><w:color w:val="FFFFFF"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr><w:t xml:space="preserve">{xe(str(cell))}</w:t></w:r></w:p></w:tc>'
    x += '</w:tr>'
    for ri,row in enumerate(rows):
        fill = "F2F2F2" if ri%2==0 else "FFFFFF"
        x += '<w:tr>'
        for i,cell in enumerate(row):
            x += f'<w:tc><w:tcPr><w:tcW w:w="{col_widths[i]}" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="{fill}"/></w:tcPr><w:p><w:r><w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr><w:t xml:space="preserve">{xe(str(cell))}</w:t></w:r></w:p></w:tc>'
        x += '</w:tr>'
    x += '</w:tbl>'
    return x

# ── CONTENT ──────────────────────────────────────────────

def build_body():
    c = []

    # COVER
    c += [p(""), p("วิเคราะห์ Human Design", bold=True, size=40, color="1E5CBF", center=True, before=200, after=60),
          p("JJ Z", bold=True, size=52, color="2D6B4A", center=True, before=60, after=60),
          p("Foundation Chart  |  Genetic Matrix  |  System: Tropical", size=20, color="777777", center=True, before=60, after=200),
          sep(), img("rId10", 1, W6, H4), p(""),
          tbl(["หัวข้อ","ข้อมูล"],[
              ["ชื่อ","JJ Z"],
              ["วันเกิด (ท้องถิ่น)","22 กุมภาพันธ์ 2534, 14:35 น. (UTC+07:00)"],
              ["สถานที่เกิด","พอมปราบสัตรูพ่าย กรุงเทพมหานคร ประเทศไทย"],
              ["อายุ","35 ปี"],
              ["วันออกแบบ (Design)","27 พฤศจิกายน 2533, 16:36:35 UTC"],
          ],[2500,6500]), pb()]

    # CORE IDENTITY
    c += [h("ข้อมูลแกนหลัก (Core Identity)", 1), sep(),
          tbl(["หัวข้อ","ค่า","ความสำคัญ"],[
              ["Type","Pure Manifesting Generator","ประเภทที่มีพลังงานสูงที่สุด"],
              ["Profile","4/6 — Opportunistic / Role Model","เส้นทางชีวิตและบทบาทต่อสังคม"],
              ["Definition","Single","พลังงานครบในตัวเอง ไม่ขึ้นกับผู้อื่น"],
              ["Inner Authority","Sacral","ตัดสินใจด้วยเสียงท้อง ไม่ใช่หัว"],
              ["Strategy","Respond — ตอบสนอง","รอให้ชีวิตนำเสนอ แล้วค่อยตอบสนอง"],
              ["Incarnation Cross","RAX The Sleeping Phoenix 1","ธีมชีวิต: ฟื้นคืนชีพ เปลี่ยนแปลง"],
              ["Themes","Satisfaction / Frustration (Anger)","สัญญาณถูก/ผิดทาง"],
              ["Brain","Active","สมองทำงานตลอดเวลา"],
              ["Determination","Hot","ร่างกายทำงานดีในสภาพแวดล้อมอุ่น"],
              ["Cognition","Taste","รับรู้โลกผ่านการลองด้วยตัวเอง"],
              ["Environment","Kitchens - Dry","พื้นที่อากาศแห้ง มีกิจกรรม"],
              ["Motivation","Hope","ขับเคลื่อนด้วยความหวัง"],
              ["Sense","Meditation","ดึงดูดผ่านความสงบนิ่งภายใน"],
              ["Trajectory","Anti-Theist","ท้าทายความเชื่อที่ตายตัว"],
              ["View","Possibility","มองเห็นความเป็นไปได้เสมอ"],
              ["Transferred Motivation","Guilt","ระวัง: อย่าตัดสินใจจากความรู้สึกผิด"],
              ["Transferred View","Probability","ระวัง: อย่าให้ความน่าจะเป็นมาแทน Possibility"],
          ],[2200,2800,4000]), pb()]

    # TYPE & STRATEGY
    c += [h("ประเภท: Pure Manifesting Generator", 1), sep(),
          p("Pure Manifesting Generator คือประเภทที่มีพลังงานสูงที่สุดในระบบ Human Design มีทั้งพลังในการริเริ่ม (Manifest) และพลังสร้างสม่ำเสมอ (Generate) ในตัวเดียวกัน", size=22, before=80, after=60),
          h("กลยุทธ์: Respond (ตอบสนอง)", 2),
          bullet("รอให้ชีวิต สถานการณ์ หรือคนอื่นนำเสนอสิ่งต่างๆ มา ก่อนตัดสินใจ"),
          bullet("ฟังเสียงจาก Sacral (ท้อง) — 'อืม' หมายถึงใช่  /  'อ้า' หมายถึงไม่ใช่"),
          bullet("ห้ามตัดสินใจด้วยหัวหรือใช้ความกดดันบังคับ Sacral"),
          bullet("ความเร็วสูง ชอบทำหลายอย่างพร้อมกัน และมักข้ามขั้นตอนได้โดยธรรมชาติ"),
          h("สัญญาณบอกทาง", 2),
          tbl(["สถานการณ์","สัญญาณ","ความหมาย"],[
              ["เดินถูกทาง","Satisfaction — พึงพอใจ มีชีวิตชีวา","กำลังทำในสิ่งที่ถูกต้อง"],
              ["เดินผิดทาง","Frustration / Anger — หงุดหงิด โกรธ ติดขัด","กำลังฝืนธรรมชาติตัวเอง"],
          ],[2000,4000,3000])]

    # PROFILE
    c += [p(""), h("โปรไฟล์ 4/6 — Opportunistic / Role Model", 1), sep(),
          h("เส้น 4 — Opportunistic", 2),
          bullet("โอกาสและความสำเร็จมาจากเครือข่ายความสัมพันธ์ คนรู้จัก ไม่ใช่ Cold Approach"),
          bullet("ต้องการพื้นฐานมั่นคงก่อนจะก้าวไปที่ใหม่เสมอ"),
          bullet("ลงทุนกับความสัมพันธ์ระยะยาว — นั่นคือทรัพย์สินที่แท้จริงของคุณ"),
          h("เส้น 6 — Role Model (3 เฟสของชีวิต)", 2),
          tbl(["เฟส","ช่วงอายุ","ลักษณะ","บทเรียน"],[
              ["เฟส 1","0-30 ปี","ลองผิดลองถูก (Line 3 energy)","เจ็บปวด เรียนรู้ สะสมประสบการณ์"],
              ["เฟส 2 ★ (ปัจจุบัน)","30-50 ปี","'บนหลังคา' — สังเกตชีวิต","ถอยออก เลือกสรร สะสมความรู้"],
              ["เฟส 3","50+ ปี","Role Model แท้จริง","กลายเป็นแบบอย่างที่คนนับถือ"],
          ],[900,1200,3200,3700]),
          p("JJ อายุ 35 ปี กำลังอยู่ในช่วงเปลี่ยนผ่าน เฟส 2 → เฟส 3 นี่คือช่วงสะสมและเตรียมความพร้อม ไม่ใช่การแสดงออก", bold=True, color="1E5CBF", size=22, before=100, after=60),
          pb()]

    # CHANNELS
    c += [h("Channels ที่กำหนดไว้ (Defined Channels)", 1), sep(),
          img("rId11", 2, W6, H4),
          tbl(["Channel","ชื่อ","Gate","ความหมายเชิงลึก"],[
              ["11-56","Curiosity","Gate 11 (Ideas) + Gate 56 (Stimulation)","นักเล่าเรื่อง แบ่งปันไอเดียได้ดีมาก คนได้รับ 'ความหมาย' จากการคุยกับคุณ  — 'A Seeker'"],
              ["20-34","Charisma","Gate 20 (Now) + Gate 34 (Power)","'Thoughts must become Deeds' มีเสน่ห์สูงเมื่อลงมือทำทันที ความน่าเชื่อถือมาจากการกระทำ"],
              ["07-31","The Alpha","Gate 7 (Self) + Gate 31 (Influence)","'Leadership — Good or Bad' ผู้นำที่คนเลือกตาม ไม่ใช่บังคับ ระวัง: อาจกลายเป็น The Authoritarian"],
          ],[900,1500,2600,4000]),
          p("หมายเหตุ: Throat Center ถูก Define ด้วย Channel ทั้ง 3 — คุณมีพลังการสื่อสารระดับสูงมาก", italic=True, size=20, color="555555", before=80, after=60),
          pb()]

    # SHADOW CHART
    c += [h("Shadow Chart — ศูนย์ที่ไม่ได้กำหนด (Open Centers)", 1), sep(),
          img("rId12", 3, W6, H4),
          tbl(["ศูนย์","เงา (Shadow)","อาการ","วิธีรับมือ"],[
              ["Heart Center","พยายามพิสูจน์คุณค่าตัวเอง","ทำทุกอย่างเพื่อให้คนอื่นรัก ทำสัญญาเกินความสามารถ","หยุดถามว่า 'ฉันดีพอไหม?' คุณค่าของคุณไม่ขึ้นกับคนอื่น"],
              ["Solar Plexus","หลีกเลี่ยงความจริงและความขัดแย้ง","ยอมเพื่อรักษาความสงบ ดูดซับอารมณ์คนรอบข้าง","รู้ว่าอารมณ์ที่รู้สึกอยู่ — เป็นของคุณหรือของคนอื่น?"],
              ["Spleen Center","ยึดติดสิ่งที่ไม่ดีต่อตัวเอง","กลัวการปล่อยวาง ยึดนิสัยหรือความสัมพันธ์ที่ไม่ดี","ฝึกปล่อยวางตามสัญชาตญาณ ไม่ใช่ตามความกลัว"],
              ["Root Center","รีบร้อนอยากหลุดพ้นแรงกดดัน","ตัดสินใจผิดเพราะความเร่งรีบ ทำงานเพื่อหนีความกดดัน","หยุด — ความกดดันจะผ่านไปเอง ไม่ต้องรีบแก้ทันที"],
              ["Head Center","คิดเรื่องที่ไม่เกี่ยวกับตัวเอง","รับความกังวลของคนอื่นมาเป็นของตัว","ถามตัวเองว่า 'ฉันต้องแก้ปัญหานี้จริงๆ ไหม?'"],
          ],[1600,1900,2700,2800]),
          pb()]

    # LINES & ENERGY FAMILIES
    c += [h("Lines Summary & Energy Families", 1), sep(),
          img("rId13", 4, W6, H35),
          h("Lines Summary", 2),
          tbl(["เส้น","ชื่อ","Design","Personality","Total","%"],[
              ["Line 4","Fixed Externalizing","4","4","8","30.77% ★"],
              ["Line 6","Role Model","4","2","6","23.08%"],
              ["Line 2","Natural Projection","2","4","6","23.08%"],
              ["Line 1","Introspective","2","1","3","11.54%"],
              ["Line 3","Adapting","0","2","2","7.69%"],
              ["Line 5","Universalizing Projection","1","0","1","3.85%"],
          ],[1200,2500,900,1100,700,1600]),
          p(""),
          h("Energy Families", 2),
          tbl(["กลุ่มพลังงาน","Design","Personality","Total","%"],[
              ["Individual - Empowering","6","5","11","37.93% ★"],
              ["Logical - Sharing","3","3","6","20.69%"],
              ["Abstract - Sharing","2","2","4","13.79%"],
              ["Integration - Survival","3","0","3","10.34%"],
              ["Centering - Self-Empowering","2","1","3","10.34%"],
              ["Tribal - Support","0","1","1","3.45%"],
              ["Defense - Support","0","1","1","3.45%"],
          ],[2500,900,1100,700,1800]),
          p(""),
          h("Incarnation Quarters", 2),
          tbl(["ไตรมาส","Design","Personality","Total","%","ธีม"],[
              ["Mutation","8","5","13","50.00% ★","การเปลี่ยนแปลงและวิวัฒนาการ"],
              ["Civilization","3","4","7","26.92%","การสร้างและจัดระเบียบสังคม"],
              ["Initiation","1","3","4","15.38%","การเริ่มต้นและจุดประกาย"],
              ["Duality","1","1","2","7.69%","ความสมดุลระหว่างสองขั้ว"],
          ],[1500,800,1100,700,1100,3800]),
          p("Individual 37.93% + Mutation Quarter 50% = ถูกออกแบบมาเพื่อนำการเปลี่ยนแปลงและสร้างแรงบันดาลใจผ่านความเป็นปัจเจก", bold=True, color="1E5CBF", size=22, before=100),
          pb()]

    # ENERGY CIRCUITS
    c += [h("Energy Circuit", 1), sep(),
          img("rId14", 5, W6, H42),
          tbl(["วงจร","Gates ที่มี","Channels สมบูรณ์","%","หมายเหตุ"],[
              ["Knowing (Individual)","7/18 = 38%","0/9 = 0%","37.93% ★","ศักยภาพสูง รอคนมาเชื่อมวงจร"],
              ["Understanding (Logical)","5/14 = 35%","1/7 = 14%","20.69%","Channel 07-31 ทำงานที่นี่"],
              ["Sensing (Abstract)","3/14 = 21%","1/7 = 14%","13.79%","Channel 11-56 ทำงานที่นี่"],
              ["Integration","4/8 = 50%","1/4 = 25%","10.34%","Channel 20-34 ทำงานที่นี่"],
              ["Centering","2/4 = 50%","0/2 = 0%","10.34%","ศูนย์กลางพลังงานตัวเอง"],
              ["Tribal","1/10 = 10%","0/5 = 0%","3.45%","พลังงานกลุ่มน้อย"],
              ["Defense","1/4 = 25%","0/2 = 0%","3.45%","พลังงานป้องกันน้อย"],
          ],[2000,1500,1700,900,2900]),
          pb()]

    # VARIABLE & ADVANCED
    c += [h("Variable System & Advanced Elements", 1), sep(),
          img("rId15", 6, W6, H42),
          tbl(["Variable","ค่า","ความหมาย"],[
              ["Design Dependent","Cognition: Taste | Line: 6th | Determination: Hot","รับรู้โลกผ่านการลองด้วยตัวเอง / ร่างกายต้องการสภาพแวดล้อมอุ่น"],
              ["Design Independent","Environment: Kitchens-Dry | Tone: Touch","พื้นที่อากาศแห้ง / ตอบสนองต่อการสัมผัส"],
              ["Personality Dependent","Sense: Meditation | Line: 4th | Motivation: Hope | Trajectory: Anti-Theist","ดึงดูดคนผ่านความสงบ / ท้าทายความเชื่อที่ตายตัว"],
              ["Personality Independent","View: Possibility | Tone: Security","มองเห็นความเป็นไปได้ / ต้องการความรู้สึกปลอดภัยเบื้องหลัง"],
          ],[2000,3500,3500]),
          p(""),
          tbl(["ชั้น","Orientation","ความหมาย"],[
              ["Design","Binary (Being 3, Space 5)","มองโลกแบบ 2 มิติ ตัวเองกับโลก ชัดเจน ตรงไปตรงมา"],
              ["Personality","Focused (Body 3, Personality 5, Individuality 1)","จิตใจทำงานแบบเจาะลึก มีโฟกัสชัดเจน มองรายละเอียดที่คนอื่นมองข้าม"],
          ],[1500,2800,4700]),
          p(""),
          tbl(["มิติ","Externalization","Internalization"],[
              ["Collective Role","The Abdicator — ถอยออกยกพื้นที่ให้คนที่เหมาะสม","The Fatigued — เหนื่อยเมื่อต้องแบกภาระส่วนรวม"],
              ["Individual Role","Tension — สร้างแรงตึงที่นำสู่การเปลี่ยนแปลง","The Opportunist — จับโอกาสได้เมื่อมันปรากฏ"],
              ["Tribal Role","The Magnanimous — ใจกว้าง เอื้อเฟื้อ","The Companion — เพื่อนแท้ที่คนไว้วางใจ"],
          ],[1800,3600,3600]),
          pb()]

    # COMPATIBILITY
    c += [h("ความเข้ากันได้กับคนอื่น", 1), sep(),
          p("JJ มี Single Definition — พลังงานภายในครบสมบูรณ์ในตัวเอง ไม่ 'ขาด' อะไรจากคนอื่น แต่ Open Centers คือจุดที่รับพลังงานจากผู้อื่นมาขยาย ทั้งดีและเสีย", size=22, before=80, after=60),
          h("ความเข้ากันได้ตาม Type", 2),
          tbl(["Type ของคู่","ระดับ","คำอธิบาย"],[
              ["Projector","★★★★★ ดีมาก","'เห็น' JJ ได้ชัดที่สุด ช่วยชี้ทิศทาง ไม่แย่งพลังงาน — สมดุลธรรมชาติ"],
              ["Reflector","★★★★ ดี","กระจกสะท้อนชีวิต ให้มุมมองหลากหลาย แต่ต้องการพื้นที่และเวลาตัดสินใจ"],
              ["Generator","★★★ กลาง","พลังงานสูงเท่ากัน เข้าใจกัน แต่ต้องระวังแย่งพื้นที่กัน"],
              ["Pure MG","★★★ กลาง","เข้าใจพลังงานกันดี แต่ทั้งคู่เร็ว อาจข้ามขั้นตอนสำคัญในความสัมพันธ์"],
              ["Manifestor","★★ ระวัง","เริ่มเองไม่บอก อาจ Trigger การตามไม่ทัน — ต้องสื่อสารมากพิเศษ"],
          ],[1800,1500,5700]),
          p(""),
          h("ความเข้ากันได้ตาม Profile", 2),
          tbl(["Profile คู่","ระดับ","เหตุผล"],[
              ["1/3","★★★★★","Line 3 เรียนรู้จากการลอง ตรงกับ Line 6 ของ JJ / Line 1 ให้รากฐานที่ต้องการ"],
              ["2/4","★★★★★","Line 4 เหมือนกัน เข้าใจเรื่องเครือข่าย / Line 2 ดึงดูด JJ ตามธรรมชาติ"],
              ["6/2","★★★★","Line 6 เหมือนกัน เข้าใจการเป็น Role Model"],
              ["5/1","★★★★","Line 5 นำเสนอ 'ทางออก' เสมอ / Line 1 ให้ความมั่นคง"],
              ["3/5","★★★","อาจเข้าใจกันแต่ดราม่าพอกัน"],
              ["4/1","★★","Line 4 เหมือนกัน แต่ Line 1 อาจ Rigid เกินไป"],
          ],[1500,1500,6000]),
          p(""),
          h("จุดอ่อนในความสัมพันธ์", 2),
          bullet("Heart Center เปิด → อย่าพิสูจน์คุณค่าตัวเอง อย่าทำสัญญาเกินจริงเพื่อให้คนรัก"),
          bullet("Solar Plexus เปิด → ดูดซับอารมณ์คนอื่น ต้องการพื้นที่คนเดียวเพื่อ Reset สม่ำเสมอ"),
          bullet("Transferred Motivation: Guilt → ระวังการตัดสินใจจาก 'ความรู้สึกผิด' แทน Sacral"),
          p("สูตรความรักที่ใช่: คนที่ให้พื้นที่ JJ ตอบสนองตามธรรมชาติ ไม่กดดันให้ตัดสินใจเร็ว และยังแข็งแกร่งในตัวเอง", bold=True, color="2D6B4A", size=22, before=100),
          pb()]

    # LIFE PHASE
    c += [h("ช่วงชีวิตปัจจุบัน (อายุ 35 ปี)", 1), sep(),
          tbl(["เฟส","ช่วงอายุ","ลักษณะ","คำอธิบาย"],[
              ["เฟส 1","0-30 ปี","Line 3 Energy","ลองผิดลองถูก เจ็บปวด เรียนรู้"],
              ["เฟส 2 ★ (ปัจจุบัน)","30-50 ปี","'บนหลังคา'","สังเกตชีวิต ถอยออก เลือกสรร สะสมความรู้"],
              ["เฟส 3","50+ ปี","Role Model แท้จริง","ลงจากหลังคา เป็นแบบอย่างที่คนนับถือ"],
          ],[1600,1200,1800,4400]),
          p(""),
          h("วงจรดาวเคราะห์ที่กำลังส่งผล", 2),
          tbl(["วงจร","สถานะ","ผลกระทบ"],[
              ["Saturn Return (~29-30 ปี)","ผ่านมาแล้ว","โครงสร้างชีวิตถูกทบทวน สิ่งที่ไม่ใช่ถูกสลัดทิ้ง"],
              ["Jupiter Return (~36 ปี / ~2027)","กำลังจะมาถึง ★","ช่วงขยาย โอกาสใหม่ การเติบโต — ควรเตรียมตัวตอนนี้"],
          ],[2500,1800,4700]),
          p(""),
          h("RAX The Sleeping Phoenix 1", 2),
          bullet("Gate 55 (Spirit): พลังงานจิตวิญญาณและอิสรภาพทางอารมณ์กำลังตื่นขึ้น"),
          bullet("Gate 34 + 20: พลังงาน 'ลงมือทำใน NOW' กำลังแรงขึ้น"),
          bullet("Phoenix ไม่ 'ตาย' — มันกำลังฟักตัวเพื่อการเกิดใหม่ครั้งยิ่งใหญ่"),
          p(""),
          h("สัญญาณ On Track / Off Track", 2),
          tbl(["On Track (ถูกทาง) ✓","Off Track (ผิดทาง) ✗"],[
              ["ชีวิตช้าลง แต่มีความหมายมากขึ้น","ยังวิ่งไล่ตามความสำเร็จแบบเฟส 1 ไม่หยุด"],
              ["เริ่มเลือกคนรอบข้างอย่างระวังมากขึ้น","รับทุกโอกาสเพราะกลัวพลาด (Root Center กดดัน)"],
              ["ความสนใจในเรื่องที่ลึกขึ้น","พิสูจน์ตัวเองกับคนที่ไม่เห็นคุณค่า"],
              ["ความสงบภายในมากขึ้น","ตัดสินใจด้วยหัว ไม่ใช่ Sacral"],
          ],[4500,4500]),
          pb()]

    # CAREER
    c += [h("แนวทางอาชีพ", 1), sep(),
          tbl(["Channel/Profile","พลังงาน","ประเภทงาน"],[
              ["Channel 20-34 (Charisma)","ลงมือทำ + แสดงพลังงาน","Entrepreneur, Creator, Leader"],
              ["Channel 11-56 (Curiosity)","แบ่งปันไอเดีย + เล่าเรื่อง","Speaker, Writer, Teacher"],
              ["Channel 07-31 (The Alpha)","นำคน + สร้างทิศทาง","CEO, Director, Mentor"],
              ["Profile 4/6","ผ่านเครือข่าย + Role Model","Network-based, Consultancy"],
              ["Individual 37.93%","ต้องการอิสระ ไม่ใช่ระบบ","Freelance, Solo, Startup"],
          ],[2200,2800,4000]),
          p(""),
          h("อาชีพที่แนะนำ", 2),
          h("กลุ่มที่ 1: Content & Communication (แนะนำสูงสุด)", 3),
          bullet("Content Creator / Influencer / Podcaster"),
          bullet("Public Speaker / Keynote Speaker"),
          bullet("Educator / Course Creator / Author / Blogger"),
          p("เหตุผล: Throat Center Define ด้วย 3 Channels + Channel 11-56 = พลังการสื่อสารและเล่าเรื่องระดับสูงมาก", size=20, color="555555", italic=True, indent=360),
          h("กลุ่มที่ 2: Leadership & Entrepreneurship", 3),
          bullet("ผู้ก่อตั้งธุรกิจ / Startup Founder"),
          bullet("CEO / ผู้นำองค์กร / Business Consultant"),
          bullet("Brand Builder / Creative Director"),
          p("เหตุผล: Channel 07-31 (The Alpha) = ผู้นำโดยธรรมชาติที่คนตามด้วยความเต็มใจ", size=20, color="555555", italic=True, indent=360),
          h("กลุ่มที่ 3: Coaching & Mentoring (อนาคตหลัง 50 ปี)", 3),
          bullet("Life Coach / Human Design Reader"),
          bullet("Business Mentor / Advisor"),
          bullet("Speaker ผู้เชี่ยวชาญในแวดวงของตัวเอง"),
          p("เหตุผล: Profile 6 = Role Model ที่แท้จริง คนจะมาขอคำแนะนำจากคุณเมื่อเวลาถูกต้อง", size=20, color="555555", italic=True, indent=360),
          p(""),
          h("กฎที่ต้องยึดในการเลือกงาน", 2),
          tbl(["กฎ","รายละเอียด"],[
              ["1. รอ Sacral ตอบสนอง","ท้องรู้สึก 'อืม/ใช่' → ทำ | หัวบอกดีแต่ท้องเฉย → อย่าทำ"],
              ["2. ต้องผ่านเครือข่าย (Line 4)","โอกาสงานที่ดีที่สุดมาจากคนรู้จัก ไม่ใช่ Cold Apply"],
              ["3. ต้องการอิสระ (Individual 38%)","งาน 9-5 ที่มี Micro-management จะทำให้ Frustration สูงมาก"],
              ["4. Determination: Hot","ทำงานดีในพื้นที่อุ่น อากาศแห้ง (Kitchens - Dry)"],
              ["5. Cognition: Taste","ต้องลองก่อนเสมอ เรียนรู้ดีที่สุดจากประสบการณ์จริง"],
          ],[2500,6500]),
          p(""),
          h("อาชีพที่ต้องหลีกเลี่ยง", 2),
          tbl(["งาน","เหตุผล"],[
              ["งานที่ต้องพิสูจน์ตัวเองตลอดเวลา","Heart Center เปิด → จะหมดพลังงาน"],
              ["งาน Service ที่รับอารมณ์คนอื่นเยอะ","Solar Plexus เปิด → ดูดซับอารมณ์จนป่วย"],
              ["งานที่ต้องตัดสินใจเร็วใต้ความกดดัน","Root Center เปิด → ตัดสินใจผิดพลาด"],
              ["งาน Corporate แบบ Rigid Hierarchy","Individual Energy สูง → จะขัดแย้งกับระบบ"],
              ["งานที่ทำคนเดียว 100%","Channel 07-31 ต้องการ Audience"],
          ],[3500,5500]),
          p(""),
          h("Timeline อาชีพ", 2),
          tbl(["ช่วงอายุ","เฟส","การกระทำที่เหมาะสม"],[
              ["35 ปี (ตอนนี้)","สะสมและสังเกต","เรียนรู้ ทดสอบ Sacral หาเครือข่ายที่ใช่"],
              ["36-40 ปี (Jupiter Return)","เริ่มเห็นทิศทาง","ลงทุนกับสิ่งที่ Sacral 'ใช่' สร้าง Authority ในด้านที่ถนัด"],
              ["40-50 ปี (ปลายเฟส 2)","สร้างชื่อเสียง","เริ่มถูกมองเป็นผู้เชี่ยวชาญ คนมาขอคำแนะนำมากขึ้น"],
              ["50+ ปี (เฟส 3)","Role Model เต็มตัว","RAX Sleeping Phoenix ตื่นเต็มที่ กลายเป็นผู้นำทางความคิด"],
          ],[1800,2200,5000]),
          pb()]

    # SUMMARY
    c += [h("สรุป — JJ Z Human Design", 1), sep(),
          tbl(["หัวข้อ","บทสรุป"],[
              ["ตัวตนแท้จริง","Pure Manifesting Generator — พลังงานสูง เสน่ห์ดึงดูด นักเล่าเรื่อง ผู้นำโดยธรรมชาติ"],
              ["จุดแข็งสูงสุด","การสื่อสาร (11-56) + พลังลงมือทำ (20-34) + ภาวะผู้นำ (07-31)"],
              ["พลังงานหลัก","Individual 37.93% + Mutation 50% = ออกแบบมาเพื่อนำการเปลี่ยนแปลง"],
              ["กลยุทธ์ชีวิต","ตอบสนอง → Sacral เป็นเข็มทิศ → ผ่านเครือข่าย → รอเวลาที่เหมาะ"],
              ["ช่วงชีวิตนี้","กำลังอยู่กลางเฟส 2 'บนหลังคา' → รอ Jupiter Return (~36 ปี)"],
              ["ความสัมพันธ์","ต้องการคนที่ให้พื้นที่ Sacral ทำงาน ไม่กดดัน และแข็งแกร่งในตัวเอง"],
              ["อาชีพที่ใช่","Content Creator / Speaker / Entrepreneur / Coach — ต้องผ่านเครือข่ายเสมอ"],
              ["สิ่งต้องระวัง","Heart/Solar Plexus เปิด + Transferred Motivation: Guilt"],
              ["ชะตากรรม","RAX Sleeping Phoenix 1 = ฟื้นคืนชีพซ้ำๆ นำการเปลี่ยนแปลง เป็นแสงสว่างให้ผู้อื่น"],
          ],[2500,6500]),
          p(""),
          p("สูตรชีวิตของ JJ Z", bold=True, size=26, color="1E5CBF", center=True, before=120, after=40),
          p("ตอบสนองต่อสิ่งที่ชีวิตนำเสนอ  →  ใช้ Sacral เป็นเข็มทิศ  →  สร้างแรงบันดาลใจผ่านการกระทำและการเล่าเรื่อง", bold=True, size=24, color="2D6B4A", center=True, before=40, after=40),
          p("รอจนกว่าจะถึงเวลาเป็น Role Model ที่แท้จริง — Phoenix กำลังฟักตัวอยู่", size=22, color="8B4000", center=True, italic=True, before=40, after=120),
          sep(),
          p("วิเคราะห์โดย Claude AI  |  Genetic Matrix Foundation Chart  |  System: Tropical", size=18, color="999999", center=True)]

    return "\n".join(c)


# ── DOCX FILES ───────────────────────────────────────────

CONTENT_TYPES = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="jpeg" ContentType="image/jpeg"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>'

RELS = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'

STYLES = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>'

def doc_rels():
    r = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
    for rid, _, media in IMAGES:
        r += f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/{media}"/>'
    r += '</Relationships>'
    return r

def full_doc(body):
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
  <w:body>
{body}
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1080" w:right="1080" w:bottom="1080" w:left="1080" w:header="708" w:footer="708" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>'''

def create():
    body = build_body()
    with zipfile.ZipFile(OUTPUT, 'w', zipfile.ZIP_DEFLATED) as zf:
        zf.writestr('[Content_Types].xml', CONTENT_TYPES.encode('utf-8'))
        zf.writestr('_rels/.rels', RELS.encode('utf-8'))
        zf.writestr('word/styles.xml', STYLES.encode('utf-8'))
        zf.writestr('word/_rels/document.xml.rels', doc_rels().encode('utf-8'))
        zf.writestr('word/document.xml', full_doc(body).encode('utf-8'))
        for _, src, media in IMAGES:
            path = os.path.join(UPLOAD, src)
            if os.path.exists(path):
                zf.write(path, f'word/media/{media}')
                print(f"  + {media}")
            else:
                print(f"  ! NOT FOUND: {src}")
    size = os.path.getsize(OUTPUT)
    print(f"\nDone: {OUTPUT}")
    print(f"Size: {size:,} bytes ({size//1024} KB)")

if __name__ == "__main__":
    create()
