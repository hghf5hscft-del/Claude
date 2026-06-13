#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import zipfile, io, re
from datetime import datetime

def xe(s):
    """XML escape"""
    return str(s).replace('&','&amp;').replace('<','&lt;').replace('>','&gt;').replace('"','&quot;')

def run(text, bold=False, italic=False, size=None, color=None, font=None):
    rpr = ''
    if bold: rpr += '<w:b/><w:bCs/>'
    if italic: rpr += '<w:i/><w:iCs/>'
    if size: rpr += f'<w:sz w:val="{size}"/><w:szCs w:val="{size}"/>'
    if color: rpr += f'<w:color w:val="{color}"/>'
    if font: rpr += f'<w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:cs="{font}"/>'
    rpr_xml = f'<w:rPr>{rpr}</w:rPr>' if rpr else ''
    return f'<w:r>{rpr_xml}<w:t xml:space="preserve">{xe(text)}</w:t></w:r>'

def para(content_xml, style='Normal', align=None, space_before=0, space_after=120, indent=None):
    ppr = f'<w:pStyle w:val="{style}"/>'
    if align: ppr += f'<w:jc w:val="{align}"/>'
    ppr += f'<w:spacing w:before="{space_before}" w:after="{space_after}"/>'
    if indent: ppr += f'<w:ind w:left="{indent}"/>'
    return f'<w:p><w:pPr>{ppr}</w:pPr>{content_xml}</w:p>'

def h1(text, color='1F5C2E'):
    return para(run(text, bold=True, size=28, color=color), space_before=240, space_after=120)

def h2(text, color='2E6B3E'):
    return para(run(text, bold=True, size=24, color=color), space_before=200, space_after=80)

def h3(text, color='3A7D4F'):
    return para(run(text, bold=True, size=22, color=color), space_before=160, space_after=60)

def h4(text, color='4A8F60'):
    return para(run(text, bold=True, size=20, color=color), space_before=120, space_after=40)

def p(text, bold=False, italic=False, size=20, color=None, align=None):
    return para(run(text, bold=bold, italic=italic, size=size, color=color), align=align, space_after=100)

def pq(text):
    """Quote / blockquote style"""
    return para(run(text, italic=True, size=19, color='555555'), indent=720, space_after=100)

def pbold(text):
    return p(text, bold=True)

def separator():
    return para(run('─' * 60, size=16, color='AAAAAA'), align='center', space_before=60, space_after=60)

def make_table(headers, rows, col_widths=None):
    """Create a simple table"""
    n = len(headers)
    if col_widths is None:
        w = 8000 // n
        col_widths = [w] * n

    def cell(text, bold=False, shade=None):
        shd = f'<w:shd w:val="clear" w:color="auto" w:fill="{shade}"/>' if shade else ''
        rpr = '<w:b/><w:bCs/>' if bold else ''
        run_xml = f'<w:r><w:rPr>{rpr}<w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr><w:t xml:space="preserve">{xe(text)}</w:t></w:r>'
        return f'<w:tc><w:tcPr><w:tcW w:w="{col_widths[0]}" w:type="dxa"/>{shd}<w:tcBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="AAAAAA"/><w:left w:val="single" w:sz="4" w:space="0" w:color="AAAAAA"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="AAAAAA"/><w:right w:val="single" w:sz="4" w:space="0" w:color="AAAAAA"/></w:tcBorders></w:tcPr><w:p><w:pPr><w:spacing w:before="40" w:after="40"/></w:pPr>{run_xml}</w:p></w:tc>'

    def row_xml(cells_text, is_header=False):
        shade = 'D6E8D8' if is_header else None
        cells = ''.join(cell(t, bold=is_header, shade=shade) for t in cells_text)
        return f'<w:tr>{cells}</w:tr>'

    header_row = row_xml(headers, is_header=True)
    data_rows = ''.join(row_xml(r) for r in rows)
    tbl_pr = '<w:tblPr><w:tblW w:w="8000" w:type="dxa"/><w:tblBorders><w:insideH w:val="single" w:sz="4" w:space="0" w:color="AAAAAA"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="AAAAAA"/></w:tblBorders></w:tblPr>'
    return f'<w:tbl>{tbl_pr}{header_row}{data_rows}</w:tbl>'

def bullet(text, bold_prefix=None):
    prefix_xml = run(bold_prefix, bold=True, size=20) if bold_prefix else ''
    return para(prefix_xml + run(('  ' if bold_prefix else '• ') + text, size=20), indent=360, space_after=60)

# ========================
# DOCUMENT BODY CONTENT
# ========================

def build_body():
    parts = []
    A = parts.append

    # COVER PAGE
    A(para(run('Human Design Chart — การแปลผลฉบับสมบูรณ์', bold=True, size=36, color='1F5C2E'), align='center', space_before=400, space_after=200))
    A(para(run('Jaja Z  •  เกิด 3 มีนาคม 2543', bold=True, size=26, color='2E2E2E'), align='center', space_after=80))
    A(para(run('08:29 น. (UTC+07:00)  •  ป้อมปราบศัตรูพ่าย, กรุงเทพมหานคร, ประเทศไทย', size=22, color='444444'), align='center', space_after=80))
    A(para(run('โครงสร้างหลัก  •  Variables  •  Shadow Chart  •  Quantum Data', size=22, color='555555'), align='center', space_after=80))
    A(para(run(f'วันที่จัดทำ: 13 มิถุนายน 2569', size=20, color='666666'), align='center', space_after=400))
    A(separator())

    # FOUNDATION SUMMARY TABLE
    A(h1('ตารางข้อมูลพื้นฐาน (Foundation Summary)'))
    A(make_table(
        ['รายการ', 'ค่า'],
        [
            ['Type', 'Pure Generator'],
            ['Profile', '2/4 — Hermit / Opportunist'],
            ['Definition', 'Single'],
            ['Inner Authority', 'Sacral'],
            ['Strategy', 'Respond (รอ Response)'],
            ['Signature / Not-Self Theme', 'Satisfaction / Frustration'],
            ['Incarnation Cross', 'RAX Consciousness 1'],
            ['Centers Defined', 'Sacral, Spleen'],
            ['Centers Undefined', 'Head, Ajna, Throat, G/Identity, Heart/Ego, Solar Plexus, Root'],
            ['Active Channel', '27-50 (Preservation / Custodianship)'],
        ],
        col_widths=[3000, 5000]
    ))
    A(p(''))
    A(p('Jaja เป็น Pure Generator ที่มี Channel เพียงเส้นเดียวคือ 27-50 (Preservation) เชื่อม Sacral กับ Spleen พลังงานที่มั่นคงของเธอมีจุดเดียวคือการดูแลเอาใจใส่และรักษาคุณค่า ในขณะที่ Centers ที่เหลืออีก 7 แห่งเปิดรับพลังงานจากคนรอบข้างอย่างเต็มที่ Profile 2/4 ทำให้เธอเป็นคนที่มีพรสวรรค์ที่ตนเองมักมองไม่เห็น แต่คนอื่นมองเห็นชัดเจน และ Incarnation Cross RAX Consciousness 1 บอกว่าพันธกิจชีวิตของเธอคือการตั้งคำถาม สำรวจความไม่แน่นอน และค้นหาความตระหนักรู้ผ่านความสงสัยและความสับสน', size=20))
    A(separator())

    # CHAPTER 1
    A(h1('บทที่ 1: Type และกลยุทธ์ชีวิต — Pure Generator'))
    A(pq('"Type คือ รูปแบบพลังงาน ของคุณ มี 5 ประเภทหลักใน Human Design Generator เป็น Type ที่พบมากที่สุด (ประมาณ 35-37%) Strategy คือ ท่าทาง ที่คุณถูกออกแบบมาให้ใช้ในชีวิต ผลลัพธ์จริงจะบอกว่าใช่หรือไม่ใช่"'))
    A(h2('Type: Pure Generator'))
    A(p('Pure Generator คือบุคคลที่มีพลังงาน Sacral ที่ลุกโชนอยู่ภายในร่างกายอย่างต่อเนื่อง Sacral Center ของเธอ Defined ซึ่งหมายความว่าเธอมีพลังงานที่ "สร้างขึ้นมาเองได้" — เมื่อทำในสิ่งที่ถูกต้อง พลังงานนี้จะไม่หมด เติมตัวเองได้ และทำให้รู้สึกมีชีวิตชีวา'))
    A(p('สิ่งที่ทำให้ Jaja เป็น Pure Generator (ไม่ใช่ Manifesting Generator) คือ Channel 27-50 ที่เธอมีนั้นไม่ได้เชื่อม Motor ไปยัง Throat โดยตรง — เธอไม่ได้ออกแบบมาเพื่อ "เริ่มต้นทุกอย่างด้วยตัวเอง" แต่มีพลังที่ยิ่งใหญ่ในการตอบสนองต่อชีวิตที่มาหาเธอ'))
    A(h3('ลักษณะสำคัญของ Pure Generator'))
    A(bullet('พลัง Sacral = เสียงภายใน (คำตอบสำหรับ Yes/No) เกิดขึ้นก่อนสมองคิด'))
    A(bullet('ร่างกายถูกออกแบบมาเพื่อใช้พลังงาน ไม่ใช่เก็บ — ถ้าวันไหนไม่ได้ทำอะไรที่ถูกต้อง จะรู้สึกอึดอัด ฟุ้งซ่าน หรือเครียดโดยไม่รู้สาเหตุ'))
    A(bullet('Generator สร้างทักษะผ่านการทำซ้ำ — ยิ่งทำสิ่งที่ใช่ซ้ำมากเท่าไหร่ ยิ่งเก่งขึ้น'))
    A(bullet('ร่างกายจะไม่เหนื่อยเมื่อทำในสิ่งที่ Sacral ตอบ Yes'))
    A(h2('Strategy: Respond (รอ Response)'))
    A(p('"Respond" ไม่ได้แปลว่า "นั่งรอเฉย ๆ" — มันหมายถึงการให้ชีวิตภายนอก (คน สถานการณ์ โอกาส คำถาม) นำเสนอบางอย่างมาก่อน แล้วค่อยฟังว่า Sacral ตอบสนองอย่างไร'))
    A(h3('ตัวอย่างชีวิตจริง 4 สถานการณ์'))
    A(bullet('ที่ทำงาน: รอจนมีเพื่อนพูดถึงปัญหา แล้วรู้สึกว่า Sacral ตอบรับทันที — นั่นคือ Response ที่ถูกต้อง ไม่ใช่ตื่นเช้าตัดสินใจเองทันที', 'งาน/โปรเจกต์: '))
    A(bullet('รอจนคนมาเปิดพื้นที่ให้ แล้วรู้สึกว่า Sacral บอก "ใช่" — ตอบสนองจากที่นั้น ไม่ใช่วางแผนในหัวก่อน', 'ความสัมพันธ์: '))
    A(bullet('มีคนถามว่า "อยากทำสิ่งนั้นไหม?" แล้วฟังว่า Sacral พูดว่าอะไร ก่อนสมองจะคิดทัน', 'การตัดสินใจ: '))
    A(bullet('สภาพแวดล้อมบางอย่างทำให้ Sacral "กระตุก" ขึ้นมาเอง — ตรงนั้นคือจุดเริ่มต้นที่ถูกต้อง', 'สิ่งใหม่: '))
    A(h3('Signature & Not-Self Theme'))
    A(p('Signature: Satisfaction — ความพอใจลึก ๆ ในร่างกาย รู้สึกว่า "ฉันอยู่ในที่ที่ถูกต้อง ทำสิ่งที่ถูกต้อง กับคนที่ถูกต้อง" ร่างกายไม่รู้สึกเหนื่อยแม้ทำงานหนัก เวลาผ่านไปเร็ว ทำแล้วอยากทำต่อ', bold=False))
    A(p('Not-Self Theme: Frustration — เมื่อ Jaja ใช้ชีวิตออกนอก Design ร่างกายจะส่งสัญญาณผ่าน Frustration ซึ่งอาจปรากฏเป็น: หงุดหงิดโดยไม่รู้สาเหตุ, รู้สึกว่าพยายามมากแต่ไม่ได้ผล, เบื่อหน่ายกับสิ่งที่ทำ Frustration ไม่ใช่สิ่งเลวร้าย — มันคือ Feedback ที่ร่างกายบอกว่า "มีบางอย่างไม่ตรง"'))
    A(separator())

    # CHAPTER 2
    A(h1('บทที่ 2: Inner Authority — วิธีตัดสินใจที่แท้จริง'))
    A(pq('"ใน Human Design สมองและความคิดไม่ใช่เครื่องมือตัดสินใจที่เชื่อถือได้ที่สุด Authority คือส่วนของร่างกาย/ความรู้สึกที่ออกแบบมาให้คุณ รู้ คำตอบที่ถูกต้องสำหรับตัวเอง"'))
    A(h2('Sacral Authority'))
    A(p('Sacral Authority คือเสียงหรือความรู้สึกในลำตัว (บริเวณช่องท้องถึงหน้าอก) ที่ตอบสนองต่อสิ่งต่าง ๆ โดยอัตโนมัติ ก่อนที่ความคิดจะเกิดขึ้น'))
    A(h3('เสียง Sacral คืออะไร'))
    A(bullet('"อืม" — เสียงที่ขึ้นมาเองเมื่อรู้สึก Yes'))
    A(bullet('"อ๊ะ" หรือเสียงที่ "ตัด" ลงมา — เมื่อรู้สึก No'))
    A(bullet('ความรู้สึกในร่างกาย — อุ่นขึ้น มีพลัง = Yes / เย็นลง หดหู่ = No'))
    A(h3('กฎทองของ Sacral'))
    A(bullet('เสียง Sacral เกิดขึ้นเพียงครั้งเดียว — มันไม่ซ้ำ ไม่รอ'))
    A(bullet('เสียงที่ตามมาทีหลัง (ที่คิด วิเคราะห์ เหตุผล) มักไม่ใช่ Sacral — มันคือสมอง'))
    A(bullet('ถามเป็น Yes/No เท่านั้น — "ฉันอยากทำสิ่งนั้นไหม?" ไม่ใช่ "ทำไมฉันถึงควรทำ?"'))
    A(h3('ตัวอย่างสถานการณ์ที่ Sacral ทำงาน'))
    A(bullet('มีคนเสนอโอกาสใหม่ ก่อนสมองคิด สังเกตว่าร่างกายรู้สึกอะไรทันที — ถ้ามีพลังงานผุดขึ้น = Sacral กำลังบอก Yes', 'การเลือกงาน: '))
    A(bullet('ร่างกายรู้สึกมีพลังเมื่ออยู่กับคนคนหนึ่ง หรือรู้สึกหมดแรง — นั่นคือ Sacral บอกความจริงเกี่ยวกับความสัมพันธ์', 'ความสัมพันธ์: '))
    A(bullet('เจอ content ที่น่าสนใจ ร่างกายรู้สึก "ดูด" เข้าไป ต้องการอ่านต่อ = Sacral บอกว่า "สิ่งนี้ใช่"', 'ความสนใจ: '))
    A(h3('ความแตกต่าง Sacral แท้ vs เสียงสมอง/ความกลัว'))
    A(make_table(
        ['Sacral แท้', 'เสียงสมอง/ความกลัว'],
        [
            ['เกิดขึ้นทันที ก่อนคิด', 'เกิดทีหลัง มาพร้อมเหตุผล'],
            ['รู้สึกในลำตัว', 'รู้สึกในหัว'],
            ['ไม่ซ้ำ', 'วนซ้ำในหัว'],
            ['ชัดเจน ถึงแม้จะเล็กน้อย', 'สับสน เปลี่ยนไปเรื่อย ๆ'],
            ['ไม่ต้องการเหตุผล', 'ต้องการเหตุผลยืนยัน'],
        ],
        col_widths=[4000, 4000]
    ))
    A(separator())

    # CHAPTER 3
    A(h1('บทที่ 3: Profile 2/4 — Hermit / Opportunist'))
    A(pq('"Profile คือ บทบาท ที่คุณถูกออกแบบมาเพื่อเล่นในชีวิตนี้ ตัวเลขแรก (Conscious) คือบทบาทที่คุณรู้ตัว ตัวเลขที่สอง (Unconscious) คือบทบาทที่ทำงานโดยอัตโนมัติ คนอื่นมักเห็นในตัวคุณก่อน"'))
    A(h2('6 Lines ภาพรวม'))
    A(make_table(
        ['Line', 'ชื่อ', 'ธีม'],
        [
            ['1', 'Investigator', 'สร้างรากฐานความมั่นคงจากความรู้'],
            ['2 ★', 'Hermit', 'พรสวรรค์ที่คนอื่นต้องมาเรียกออกมา'],
            ['3', 'Martyr', 'เรียนรู้ผ่านการลองผิดลองถูก'],
            ['4 ★', 'Opportunist', 'มีอิทธิพลผ่านเครือข่ายคนใกล้ชิด'],
            ['5', 'Heretic', 'แก้ปัญหาให้คนอื่นในภาพกว้าง'],
            ['6', 'Role Model', 'เป็นต้นแบบผ่านสามช่วงชีวิต'],
        ],
        col_widths=[800, 2200, 5000]
    ))
    A(p(''))
    A(h2('Line 2 — Hermit (Conscious / สิ่งที่คุณรู้ตัว)'))
    A(pbold('"คนที่ไม่เห็นพรสวรรค์ของตัวเอง แต่คนอื่นเห็นหมด"'))
    A(p('Line 2 คือ Line ที่มีพรสวรรค์ที่ดูเหมือน "ธรรมชาติ" จนตัวเองไม่รู้ว่ามัน "พิเศษ" สิ่งที่ Jaja ทำได้อย่างง่ายดายมักเป็นพรสวรรค์ที่หายากมากในสายตาคนอื่น'))
    A(bullet('ต้องการเวลาส่วนตัว — Line 2 ชาร์จพลังจากการอยู่คนเดียว ไม่ใช่เพราะ introvert แต่เพราะพรสวรรค์ "งอก" ขึ้นในความเงียบ'))
    A(bullet('ถูก "เรียก" ออกมา — มักมีคนมาชวนหรือขอความช่วยเหลือในสิ่งที่ถนัด'))
    A(bullet('ความรู้สึก "ถูกรุกราน" — เมื่อถูกเรียกออกมาจากพื้นที่ส่วนตัวบ่อยเกินไป อาจรู้สึกเหนื่อย ล้า'))
    A(bullet('ไม่เชื่อในตัวเอง — มักมองข้ามพรสวรรค์ของตัวเอง เพราะมันดูง่ายเกินไปในสายตาเธอ'))
    A(h2('Line 4 — Opportunist (Unconscious / สิ่งที่คนอื่นเห็นในตัวคุณ)'))
    A(pbold('"คนที่มีอิทธิพลผ่านความสัมพันธ์ใกล้ชิด"'))
    A(bullet('ฐานชีวิตต้องมั่นคงก่อน — Line 4 ไม่เปลี่ยนแปลงกระทันหัน จะเปลี่ยนก็ต่อเมื่อมีที่ยืนใหม่รองรับแล้ว'))
    A(bullet('อิทธิพลผ่านความสัมพันธ์ — สิ่งที่ Jaja พูดหรือทำมีผลกับคนใกล้ชิดมาก'))
    A(bullet('เครือข่ายคือทุกอย่าง — งาน ความรัก โอกาส มักมาจากคนที่รู้จักแล้ว ไม่ใช่คนแปลกหน้า'))
    A(h2('Profile 2/4 โดยรวม'))
    A(p('ความตึงเครียดสำคัญ: Line 2 ต้องการอยู่คนเดียว แต่ Line 4 ต้องการเครือข่ายและความสัมพันธ์'))
    A(p('วิธีที่ทั้งสองทำงานร่วมกัน: Jaja ชาร์จพลังในความเงียบ (Line 2) แล้วนำพรสวรรค์ที่พัฒนาขึ้นออกไปแบ่งปันกับเครือข่ายใกล้ชิด (Line 4) ไม่ใช่กับโลกกว้าง'))
    A(make_table(
        ['หัวข้อ', 'คำแนะนำ'],
        [
            ['หยุดทำ', 'ออกไปหาคนใหม่ตลอดเวลา / ละเลยเวลาส่วนตัว / พยายาม "โปรโมต" ตัวเองในที่กว้าง'],
            ['เริ่มทำ', 'ลงทุนกับความสัมพันธ์เก่า / ให้เวลาตัวเองอยู่คนเดียวโดยไม่รู้สึกผิด / เชื่อว่าโอกาสมาจากคนที่รู้จักแล้ว'],
            ['อาชีพ', 'งานที่ให้เวลาส่วนตัวพัฒนาความสามารถ + นำไปใช้กับกลุ่มคนใกล้ชิด'],
            ['ความสัมพันธ์', 'คนรักและเพื่อนสนิทจะเรียก Jaja ออกจาก Shell และเธอมีอิทธิพลลึก ๆ กับคนเหล่านั้น'],
        ],
        col_widths=[2000, 6000]
    ))
    A(separator())

    # CHAPTER 4
    A(h1('บทที่ 4: Definition — Single'))
    A(pq('"Definition คือวิธีที่ Centers ที่ Defined เชื่อมต่อกัน สำหรับ Jaja Sacral และ Spleen เชื่อมกันโดยตรง เป็นกลุ่มเดียว ทำงานสอดประสานกันตลอดเวลา"'))
    A(h2('Single Definition'))
    A(bullet('ความสมบูรณ์ในตัวเอง: ไม่ต้องการคนอื่นมา "เติมเต็ม" — เลือกความสัมพันธ์เพื่อการเติบโต ไม่ใช่เพราะขาด'))
    A(bullet('ตัดสินใจได้เร็ว: ไม่ต้องรอสภาพแวดล้อมหรือคนอื่นมาทำให้ชัดขึ้น'))
    A(bullet('พลังงานสม่ำเสมอ: ไม่ขึ้นลงตามว่าอยู่กับใคร — แต่ Centers 7 แห่งที่เปิดรับยังรับพลังจากรอบข้างมาก'))
    A(h3('สุขภาพพลังงาน'))
    A(bullet('อยู่คนเดียว: ชาร์จพลังได้ดี Centers ที่เปิดรับจะปล่อยพลังงานที่รับมาออกไป'))
    A(bullet('อยู่กับคนอื่น: ตื่นตัวและ energized มากกว่าปกติ แต่ต้องระวังว่าพลังที่รู้สึกนั้นเป็นของจริงหรือเปลือยพลังของคนรอบข้าง'))
    A(p('คำแนะนำ: ให้ตัวเองมีเวลา decompress ทุกวัน — อาบน้ำ เดินคนเดียว หรือนอนเงียบ ๆ สัก 30 นาที เพื่อให้ร่างกายปล่อยพลังงานที่รับมาจากผู้อื่นออกไป', bold=False))
    A(separator())

    # CHAPTER 5
    A(h1('บทที่ 5: Centers — ศูนย์พลัง 9 แห่ง'))
    A(pq('"Centers 9 แห่งคือเสาอากาศ 9 ต้น Centers ที่ Defined (มีสี) คือพลังงานที่คงที่ Centers ที่ Undefined (ขาว) คือตัวรับสัญญาณที่ดูดพลังงานจากคนรอบข้างมาขยายให้แรงขึ้น — ทั้งปัญญาและกับดักในเวลาเดียวกัน"'))

    A(h2('HEAD CENTER — UNDEFINED (เปิดกว้าง)'))
    A(make_table(
        ['หัวข้อ', 'รายละเอียด'],
        [
            ['ปัญญาที่ได้', 'รับรู้แรงดันความคิดของคนอื่น เข้าใจว่าใครกำลังคิดอะไรหนักหรืออยู่ภายใต้แรงดันอะไร'],
            ['กับดัก (Not-Self)', 'คิดเรื่องที่ไม่เกี่ยวกับตัวเองเลย และรู้สึกว่า "ต้องหาคำตอบ" ให้กับคำถามที่ไม่ใช่ของตัวเอง'],
            ['คำถามเช็คตัวเอง', '"ความคิดที่วนเวียนอยู่นี้ — เป็นคำถามที่ฉันอยากรู้จริง ๆ หรือมาจากคนที่เจอมา?"'],
        ],
        col_widths=[2200, 5800]
    ))
    A(p(''))
    A(h2('AJNA CENTER — UNDEFINED (เปิดกว้าง)'))
    A(make_table(
        ['หัวข้อ', 'รายละเอียด'],
        [
            ['ปัญญาที่ได้', 'สามารถเห็นหลายมุมมองพร้อมกัน ไม่ติดอยู่กับมุมมองเดียว ยืดหยุ่นทางความคิดสูง'],
            ['กับดัก (Not-Self)', 'แกล้งทำเป็นแน่ใจเพื่อให้ดูน่าเชื่อถือ ทั้งที่จริง ๆ ยังไม่แน่ใจ'],
            ['คำถามเช็คตัวเอง', '"ฉันแน่ใจเรื่องนี้จริง ๆ หรือแค่รู้สึกว่าควรต้องแน่ใจ?"'],
        ],
        col_widths=[2200, 5800]
    ))
    A(p(''))
    A(h2('THROAT CENTER — UNDEFINED (เปิดกว้าง) ★ ระดับเปราะบางสูงสุด (7/7)'))
    A(make_table(
        ['หัวข้อ', 'รายละเอียด'],
        [
            ['ปัญญาที่ได้', 'ปรับรูปแบบการสื่อสารให้เข้ากับคนตรงหน้าได้ดีมาก เป็นผู้ "แปล" ความคิดได้หลากหลายรูปแบบ'],
            ['กับดัก (Not-Self)', 'พูดเพื่อดึงความสนใจ หรือพูดในจังหวะที่ไม่ใช่ของตัวเอง (พูดก่อนหรือหลังเกินไป)'],
            ['คำถามเช็คตัวเอง', '"ฉันพูดเพราะมีอะไรอยากแชร์จริง ๆ หรือพูดเพื่อให้มีคนสังเกตฉัน?"'],
            ['สำคัญ', 'Gate 31 (North Node) และ Gate 35 (Design Earth) ต่างอยู่ใน Throat — รอจังหวะที่ถูกต้องเสมอ'],
        ],
        col_widths=[2200, 5800]
    ))
    A(p(''))
    A(h2('G / IDENTITY CENTER — UNDEFINED (เปิดกว้าง)'))
    A(make_table(
        ['หัวข้อ', 'รายละเอียด'],
        [
            ['ปัญญาที่ได้', 'กลมกลืนกับสภาพแวดล้อม ปรับตัวได้เก่ง ไม่ยึดติดกับตัวตนที่แน่นอนตายตัว'],
            ['กับดัก (Not-Self)', 'ถามตัวเองว่า "ฉันเป็นใคร? ฉันอยากทำอะไรในชีวิต?" วนซ้ำจนรู้สึกหลงทาง'],
            ['คำถามเช็คตัวเอง', '"ความสับสนนี้ — เกิดขึ้นหลังอยู่กับคนบางคนหรือสภาพแวดล้อมบางอย่างไหม?"'],
            ['คำแนะนำ', 'เลือกสภาพแวดล้อมและคนรอบข้างอย่างระมัดระวัง G Center จะ "กลายเป็น" สิ่งที่สภาพแวดล้อมสะท้อน'],
        ],
        col_widths=[2200, 5800]
    ))
    A(p(''))
    A(h2('HEART / EGO CENTER — UNDEFINED (เปิดกว้าง)'))
    A(make_table(
        ['หัวข้อ', 'รายละเอียด'],
        [
            ['ปัญญาที่ได้', 'มองเห็นว่าคนอื่นมีแรงจูงใจหรือ Ego อะไรอยู่ เข้าใจแรงจูงใจของคนได้ดี'],
            ['กับดัก (Not-Self)', 'รู้สึกว่าต้องพิสูจน์ตัวเอง สัญญาในสิ่งที่ทำได้ยาก หรือรับงานเกินกำลัง'],
            ['คำถามเช็คตัวเอง', '"ฉันสัญญาหรือรับงานนี้เพราะ Sacral บอก Yes หรือเพราะรู้สึกว่าต้องพิสูจน์ตัวเอง?"'],
        ],
        col_widths=[2200, 5800]
    ))
    A(p(''))
    A(h2('SACRAL CENTER — DEFINED ★★'))
    A(p('หน้าที่: พลังชีวิต การทำงาน การสร้างสรรค์ พลังที่สร้างขึ้นเองได้เมื่อทำในสิ่งที่ถูกต้อง'))
    A(bullet('คุณสมบัติที่มีอย่างมั่นคง: พลังงานไม่หมดเมื่อทำสิ่งที่ถูก, Authority (Sacral) สำหรับตัดสินใจ, พลังในการดูแลผ่าน Gate 27'))
    A(bullet('ผลต่อคนรอบข้าง: คนที่อยู่ใกล้ Jaja มักรู้สึกว่า "มีพลังงาน" มากกว่าปกติ — Sacral ที่ Defined ส่งพลังออกไปโดยอัตโนมัติ'))
    A(p(''))
    A(h2('SOLAR PLEXUS CENTER — UNDEFINED (เปิดกว้าง)'))
    A(make_table(
        ['หัวข้อ', 'รายละเอียด'],
        [
            ['ปัญญาที่ได้', 'รับรู้อารมณ์ของคนอื่นได้ดีมาก "รู้สึก" ได้ว่าห้องมีอารมณ์อะไร ใครกำลังเจ็บปวด — empathy สูง'],
            ['กับดัก (Not-Self)', 'หลีกเลี่ยงความขัดแย้งทุกราคา รับอารมณ์คนอื่นมาเป็นของตัวเอง'],
            ['คำถามเช็คตัวเอง', '"อารมณ์นี้เป็นของฉันจริง ๆ หรือรับมาจากคนรอบข้าง?"'],
        ],
        col_widths=[2200, 5800]
    ))
    A(p(''))
    A(h2('SPLEEN CENTER — DEFINED ★★'))
    A(p('หน้าที่: สัญชาตญาณ สุขภาพ ความรู้สึกปลอดภัย'))
    A(bullet('คงที่อย่างมั่นคง: สัญชาตญาณเรื่องสิ่งที่ดีหรือไม่ดีต่อร่างกายและชีวิต — Gate 50 (Values) ให้มาตรฐานที่แข็งแกร่ง'))
    A(bullet('ผลต่อคนรอบข้าง: คนที่อยู่ใกล้ Jaja รู้สึกปลอดภัยและสบายใจ เพราะ Spleen ที่ Defined ส่งสัญญาณความปลอดภัยออกไป'))
    A(p(''))
    A(h2('ROOT CENTER — UNDEFINED (เปิดกว้าง)'))
    A(make_table(
        ['หัวข้อ', 'รายละเอียด'],
        [
            ['ปัญญาที่ได้', 'เข้าใจความเครียดและแรงดันของคนอื่น รับรู้ว่าใครกำลังรู้สึกเร่งด่วนหรือกดดัน'],
            ['กับดัก (Not-Self)', 'รีบตัดสินใจเพื่อให้หายจากความรู้สึกกดดัน ทั้งที่แรงดันนั้นมาจากคนอื่น'],
            ['คำถามเช็คตัวเอง', '"ฉันรีบเพราะอะไรเร่งด่วนจริง ๆ หรือรีบเพราะรู้สึกว่าควรรีบ?"'],
        ],
        col_widths=[2200, 5800]
    ))
    A(separator())

    # CHAPTER 6
    A(h1('บทที่ 6: Channels — Channel 27-50 Preservation'))
    A(pq('"Channel คือเส้นที่เชื่อม Centers สองจุดด้วย Gates สองบาน Chart ของ Jaja มี Channel เพียงเส้นเดียว แต่นั่นทำให้พลังงานนั้นเป็นแกนกลางที่ชัดเจนอย่างยิ่ง"'))
    A(make_table(
        ['รายการ', 'ค่า'],
        [
            ['Channel', '27-50: Preservation / Custodianship'],
            ['เชื่อมระหว่าง', 'Sacral ↔ Spleen'],
            ['Circuit', 'Tribal — Defense Stream'],
            ['Gate 27 (Conscious / Personality Jupiter)', 'Nourishing / Caring — การหล่อเลี้ยง Gate 27, Line 2'],
            ['Gate 50 (Unconscious / Design Venus)', 'Values / Leadership — คุณค่าและมาตรฐาน Gate 50, Line 6'],
        ],
        col_widths=[2800, 5200]
    ))
    A(p(''))
    A(h2('Gate 27 — Nourishing (Caring)'))
    A(p('พลังงานในการดูแล เอาใจใส่ และหล่อเลี้ยงผู้อื่น — รู้ว่าใครต้องการอะไร และมีแรงที่จะให้สิ่งนั้น คนอื่นรู้สึกได้โดยไม่ต้องพูด'))
    A(h2('Gate 50 — Values (Leadership)'))
    A(p('พลังงานในการรักษา "สิ่งดีงาม" ที่มีอยู่ — คุณค่า มาตรฐาน กฎกติกา มีมาตรฐานที่ชัดเจนในชีวิต และรู้สึกไม่สบายใจเมื่อเห็นสิ่งที่ละเมิดคุณค่าเหล่านั้น'))
    A(h2('Channel 27-50 โดยรวม'))
    A(bullet('คนอื่นสัมผัสอย่างไร: รู้สึกว่า Jaja เอาใจใส่อย่างแท้จริง และรู้สึกปลอดภัย เพราะทั้ง Gate 27 (ความห่วงใย) + Gate 50 (มาตรฐานที่ชัดเจน) สร้างพื้นที่ที่ปลอดภัยและมีหลักการ'))
    A(bullet('วิธีใช้ให้ดีที่สุด: เลือกว่าจะดูแล "ใคร" อย่างจริงจัง — พลังนี้ออกแบบมาสำหรับกลุ่มเล็ก ๆ ที่ใกล้ชิด ไม่ใช่ทุกคน'))
    A(bullet('เงา (Shadow): ดูแลคนอื่นจนลืมดูแลตัวเอง หรือดูแลคนผิดคนจนพลังงานหมด — Sacral จะเตือนผ่าน Frustration'))
    A(separator())

    # CHAPTER 7
    A(h1('บทที่ 7: Incarnation Cross — RAX Consciousness 1'))
    A(pq('"Incarnation Cross คือ ชื่อ ของจุดประสงค์ชีวิต มาจาก Gates 4 ดวงที่สำคัญที่สุด สำหรับ Jaja เป็น Right Angle Cross (Personal) — เน้นการเดินทางชีวิตส่วนตัวมากกว่าการรับใช้โลกโดยตรง"'))
    A(make_table(
        ['ดาว', 'Gate', 'ชื่อ Gate', 'ความหมาย'],
        [
            ['Conscious Sun ☉', '63.2', 'After Completion — Doubts', 'ตั้งคำถามกับสิ่งที่ "สมบูรณ์แล้ว" / Structuring'],
            ['Conscious Earth ⊕', '64.2', 'Before Completion — Confusion', 'ถือความไม่แน่นอนก่อนความชัดเจนจะเกิด / Qualification'],
            ['Unconscious Sun ☉', '5.4', 'Fixed Rhythms — The Hunter', 'พลังแห่งความสม่ำเสมอ จังหวะที่คงที่'],
            ['Unconscious Earth ⊕', '35.4', 'Change — Hunger', 'แรงผลักในการสัมผัสประสบการณ์ใหม่ ๆ'],
        ],
        col_widths=[2000, 1000, 3000, 2000]
    ))
    A(p(''))
    A(h2('ความหมายของ RAX Consciousness 1'))
    A(p('"Consciousness" ในที่นี้หมายถึงกระบวนการสร้างความตระหนักรู้ผ่านความสงสัยและความสับสน Gate 63 (Doubt) และ Gate 64 (Confusion) อยู่ใน Head Center — ทั้งสองเป็น Gates แห่งการตั้งคำถาม Jaja ถูกออกแบบมาเพื่อตั้งคำถามกับสิ่งที่คนส่วนใหญ่ยอมรับว่า "ชัดเจนแล้ว"'))
    A(h2('พันธกิจและจุดประสงค์'))
    A(bullet('ตั้งคำถามที่คนอื่นไม่กล้าถาม — นี่คือพลังพิเศษของ Conscious Sun Gate 63'))
    A(bullet('ยืนอยู่ในความไม่แน่นอนโดยไม่รีบหาคำตอบปลอม — พลังของ Conscious Earth Gate 64'))
    A(bullet('สร้างสิ่งสำคัญผ่านการทำซ้ำอย่างสม่ำเสมอ — Unconscious Sun Gate 5 (Fixed Rhythms)'))
    A(bullet('เปิดรับประสบการณ์ใหม่ ๆ โดยไม่ยึดติดกับ "แบบเดิม" — Unconscious Earth Gate 35 (Change)'))
    A(h2('ความท้าทาย'))
    A(p('สังคมมักสอนให้ "มีคำตอบ" และ "มั่นใจ" สำหรับคนที่ Conscious Sun อยู่ใน Gate 63 (Doubt) สิ่งนี้คือแรงกดดันที่ยากมาก Jaja อาจรู้สึกว่าตัวเองผิดปกติที่ "สงสัยตลอด" แต่ความสงสัยนั้นคือแกนกลางของ Cross ของเธอ — Cross มักชัดขึ้นเรื่อย ๆ ตามอายุ'))
    A(separator())

    # CHAPTER 8
    A(h1('บทที่ 8: Variables — ธรรมชาติกาย-ใจในระดับลึก'))
    A(p('Variable Type Code: PRR DRR', bold=True))
    A(make_table(
        ['ตำแหน่ง', 'หน้าที่', 'ค่า'],
        [
            ['Design Brain (บนซ้าย)', 'วิธีประมวลผลพลังงาน', 'Passive (→)'],
            ['Design Determination', 'วิธีร่างกายย่อย/ประมวล', 'Nervous'],
            ['Design Cognition', 'วิธีที่ร่างกาย "รู้"', 'Feeling'],
            ['Design Environment', 'สภาพแวดล้อมที่ร่างกายเติบโต', 'Shores — Artificial / Observer'],
            ['Personality (บนขวา)', 'รูปแบบ Personality', 'Receptive (→)'],
            ['Personality Motivation', 'แรงผลักดันแท้', 'Desire'],
            ['Transferred Motivation', 'เมื่อ Desire ถูกเบี่ยงเบน', 'Innocence'],
            ['Personality Sense', 'วิธีที่จิตใจรับรู้', 'Acceptance'],
            ['Personality Trajectory', 'ทิศทางจิตใจ', 'Follower'],
            ['Personality View', 'มุมมองแท้', 'Possibility (Peripheral)'],
            ['Transferred View', 'เมื่อ Possibility ถูกเบี่ยงเบน', 'Probability'],
            ['Design Dependent Variable', 'แกน Design', '4th Line'],
            ['Personality Dependent Variable', 'แกน Personality', '2nd Line'],
        ],
        col_widths=[2500, 2500, 3000]
    ))
    A(p(''))
    A(h2('Determination: Nervous'))
    A(p('ร่างกายทำงานได้ดีที่สุดในสภาพแวดล้อมที่มีสิ่งกระตุ้นเปลี่ยนแปลง ไม่ใช่ความเงียบนิ่งสม่ำเสมอ อาหารที่มีรสชาติหลากหลาย งานที่มีหลายโปรเจกต์ และ routine ที่เปลี่ยนได้บ้างจะทำให้ร่างกายทำงานได้ดีกว่า'))
    A(h2('Environment: Shores — Artificial / Observer'))
    A(p('ร่างกายเติบโตได้ดีในสภาพแวดล้อมที่เป็น "ขอบระหว่างสองโลก" — ไม่ใช่กลางทะเลหรือกลางป่า แต่เป็นพื้นที่ที่มีทั้งธรรมชาติและโครงสร้างของมนุษย์ เช่น ริมแม่น้ำในเมือง คาเฟ่ริมคลอง ห้องทำงานที่มีต้นไม้ สไตล์ Observer หมายถึงพื้นที่ที่ได้ "สังเกต" สิ่งที่เกิดขึ้น ไม่ใช่ต้องมีส่วนร่วมตลอดเวลา'))
    A(h2('Motivation: Desire → Transferred: Innocence'))
    A(p('Desire คือแรงผลักดันที่แท้จริง — ขับเคลื่อนด้วยความต้องการที่แท้จริงของตัวเอง เมื่อ Desire ถูกเบี่ยงเบน จะเข้าสู่ Innocence — ทำสิ่งต่าง ๆ ด้วย "ความบริสุทธิ์ใจ" แบบไร้เดียงสา โดยไม่ได้เชื่อมกับความต้องการแท้จริง สัญญาณ: ทำแล้วรู้สึกว่า "ฉันทำเพราะมันถูกต้อง" ไม่ใช่ "ฉันทำเพราะอยากทำ"'))
    A(h2('View: Possibility → Transferred: Probability'))
    A(p('Possibility คือจิตใจที่มองเห็นศักยภาพในทุกสถานการณ์ เมื่อถูกเบี่ยงเบน จะเริ่มคำนวณโอกาส คิดว่า "มันคงไม่ได้ผลหรอก" แทนที่จะมองว่า "สิ่งนี้อาจเป็นไปได้"'))
    A(h2('สูตรรวม Variables'))
    A(p('รับพลังจากสิ่งที่มาหา (Passive/Receptive) → อยู่ในพื้นที่ขอบระหว่างสองโลก (Shores-Artificial/Observer) → ขับเคลื่อนด้วยสิ่งที่ตัวเองต้องการจริง ๆ (Desire/Possibility) → ฟื้นฟูผ่านความรู้สึกในร่างกายและการสังเกต (Feeling/Peripheral)', bold=False))
    A(separator())

    # CHAPTER 9
    A(h1('บทที่ 9: Shadow Chart — แผนที่ "เสียงในหัว"'))
    A(h2('ระดับความเปราะบางต่อ Conditioning ตาม Center'))
    A(make_table(
        ['Center', 'ระดับ', 'Gates Active ใน Chart', 'เสียง Not-Self ที่พบบ่อย'],
        [
            ['Throat', '7 ★★★', 'Gate 31, 35', 'พูดเพื่อดึงความสนใจ, พูดในจังหวะที่ไม่ใช่ของตัวเอง'],
            ['Head', '6 ★★', 'Gate 63, 64', 'คิดเรื่องที่ไม่เกี่ยวกับตัวเอง, ต้องหาคำตอบทุกคำถาม'],
            ['Root', '5 ★★', 'Gate 19, 41', 'รีบตัดสินใจเพื่อหายจากแรงดัน, เร่งโดยไม่มีเหตุผล'],
            ['Ajna', '4 ★', 'Gate 24', 'แกล้งทำเป็นแน่ใจ, rationalize สิ่งที่ไม่ใช่ของจริง'],
            ['G/Identity', '3 ★', 'Gate 13', '"ฉันเป็นใคร?" วนซ้ำ, ตัวตนเปลี่ยนตามสภาพแวดล้อม'],
            ['Solar Plexus', '2', 'Gate 37', 'หลีกเลี่ยงความขัดแย้ง, รับอารมณ์คนอื่นมาเป็นของตัวเอง'],
            ['Heart/Ego', '1', 'Gate 51', 'พิสูจน์ตัวเอง, สัญญาเกินกำลัง'],
        ],
        col_widths=[1600, 800, 1400, 4200]
    ))
    A(p(''))
    A(h2('Transferred Variables'))
    A(make_table(
        ['Variable', 'แท้จริง', 'เมื่อถูก Transfer', 'สัญญาณที่เกิดขึ้น'],
        [
            ['Motivation', 'Desire', 'Innocence', 'ทำเพราะ "ควรทำ" ไม่ใช่ "อยากทำ" เหนื่อยแต่ยังทำต่อ'],
            ['View', 'Possibility', 'Probability', '"คงไม่ได้ผลหรอก" / คำนวณความเสี่ยงแทนมองโอกาส'],
        ],
        col_widths=[1400, 1400, 1400, 3800]
    ))
    A(p(''))
    A(h2('วิธีรับมือกับทุกเสียง — กระบวนการเดียว'))
    A(bullet('หยุด — ไม่ตอบสนองทันที Sacral ของคุณรู้คำตอบแล้วตั้งแต่วินาทีแรก', '1. '))
    A(bullet('ถามว่า "ใหม่หรือเก่า" — ความรู้สึกแว่บแรก = มักจริง / ความคิดที่ตามมาวนซ้ำ = มักเสียงรบกวน', '2. '))
    A(bullet('ตั้งชื่อให้มัน — "นี่คือเสียง Heart ที่อยากพิสูจน์ตัวเอง" พลังของมันจะลดลงทันที', '3. '))
    A(bullet('ปล่อยผ่าน — ไม่ต้องสู้หรือกำจัด แค่รับรู้แล้วกลับมาที่ Sacral', '4. '))
    A(separator())

    # CHAPTER 10
    A(h1('บทที่ 10: Quantum Data — สถิติเชิงลึก'))
    A(h2('Lines Distribution'))
    A(make_table(
        ['Line', 'Total', '%', 'ชื่อ', 'นัยสำคัญ'],
        [
            ['2 ★', '6', '23.08%', 'Natural Projection', 'สูงสุด — พรสวรรค์ที่มองไม่เห็นในตัวเอง ธีมที่ทอซ้ำมากที่สุดใน Chart'],
            ['5', '5', '19.23%', 'Universalizing Projection', 'สูง — พลังในการแก้ปัญหาให้คนอื่นในภาพกว้าง'],
            ['1', '5', '19.23%', 'Introspective', 'สูง — รากฐานความรู้และความมั่นคงที่ลึก'],
            ['6', '4', '15.38%', 'Role Model', 'ปานกลาง — ต้นแบบที่ค่อย ๆ ชัดขึ้นตามอายุ'],
            ['4', '3', '11.54%', 'Fixed Externalizing', 'ปานกลาง — อิทธิพลผ่านเครือข่าย'],
            ['3', '3', '11.54%', 'Adapting', 'ต่ำสุด — ไม่ได้ถูกออกแบบให้เรียนรู้ผ่านความล้มเหลวซ้ำ ๆ'],
        ],
        col_widths=[600, 600, 1000, 2200, 3600]
    ))
    A(p(''))
    A(h2('Energy Families / Circuits'))
    A(make_table(
        ['Circuit', 'Total', '%', 'ความหมาย'],
        [
            ['Abstract - Sharing ★★', '11', '42.31%', 'สูงสุด — สัมผัสประสบการณ์แล้วแบ่งปันเป็นเรื่องราว'],
            ['Logical - Sharing', '7', '26.92%', 'รอง — แบ่งปันผ่านระบบและตรรกะ'],
            ['Individual - Empowering', '3', '11.54%', 'ปัจเจก — พลังส่วนตัวที่ไม่ต้องการการอนุมัติ'],
            ['Tribal - Support', '2', '7.69%', 'ดูแลกลุ่มคนใกล้ชิด'],
            ['Defense - Support', '2', '7.69%', 'ปกป้องและรักษาคุณค่า (Channel 27-50)'],
            ['Centering - Self-Empowering', '1', '3.85%', 'พลังงานตัวเอง'],
            ['Integration - Survival', '0', '0.00%', 'ไม่มีเลย — ไม่ได้ถูกออกแบบมาเพื่อ "ทำเพื่อตัวเอง" เป็นหลัก'],
        ],
        col_widths=[2500, 600, 900, 4000]
    ))
    A(p(''))
    A(h2('Incarnation Quarters'))
    A(make_table(
        ['Quarter', 'Total', '%', 'ดาว', 'ธีม'],
        [
            ['Mutation ★★', '11', '42.31%', 'Pluto/Saturn', 'เปลี่ยนผ่าน กลายพันธุ์ สิ่งใหม่ถือกำเนิด'],
            ['Initiation ★★', '10', '38.46%', 'Mars', 'จุดประกาย รวมกลุ่มผู้มีไฟเดียวกัน'],
            ['Civilization', '3', '11.54%', 'Venus', 'สร้างระบบ วัฒนธรรม โครงสร้าง'],
            ['Duality', '2', '7.69%', 'Jupiter', 'เติบโตผ่านความสัมพันธ์ พันธสัญญา'],
        ],
        col_widths=[1800, 600, 900, 1000, 3700]
    ))
    A(p(''))
    A(p('Mutation + Initiation รวมกัน = 80.77% ของพลังงานทั้งหมด — ชีวิตของ Jaja ถูกออกแบบมาเพื่อการเปลี่ยนแปลงและจุดประกายสิ่งใหม่ การเปลี่ยนแปลงในชีวิตเธอไม่ใช่ปัญหา มันคือ Design', bold=False))
    A(separator())

    # CHAPTER 11
    A(h1('บทที่ 11: Story Line — จาก South Node สู่ North Node'))
    A(make_table(
        ['Node', 'Gate.Line', 'Keyword หลัก', 'Keyword รอง', 'ธีม'],
        [
            ['Design South Node ☋', '41.3', 'Contraction', 'Efficiency', 'ในช่วงที่ต้องลด การรักษาตัวเองก่อนถือเป็นสิ่งที่ถูกต้อง'],
            ['Personality South Node ☋', '41.2', 'Contraction', 'Caution', 'มนุษยธรรมที่ถูกกรองด้วยความเป็นจริง'],
            ['Design North Node ☊', '31.3', 'The Logical Leader', 'Selectivity', 'ความสามารถในการประเมินและเลือกอิทธิพลที่ถูกต้องอย่างรอบคอบ'],
            ['Personality North Node ☊', '31.2', 'The Logical Leader', 'Arrogance', 'การกระทำอย่างอิสระโดยไม่ต้องรอการนำทาง'],
        ],
        col_widths=[2000, 800, 1800, 1000, 2400]
    ))
    A(p(''))
    A(pq('"Jaja เกิดมาพร้อมกับพลังงานแห่งการลดและรักษาตัวเอง (Gate 41 — Contraction: Efficiency + Caution) และชีวิตกำลังพาเธอไปสู่การเป็นเสียงที่คนเลือกฟัง — ผู้นำที่กล้าหาญ ไม่รอฉันทามติ แต่เลือกอิทธิพลรอบข้างอย่างชาญฉลาด (Gate 31 — The Logical Leader: Arrogance + Selectivity)"'))
    A(p('หมายเหตุสำคัญ: ทั้งสี่ Node เป็น Retrograde — การเดินทางนี้ไม่ได้เป็นเส้นตรง จะมีการถอยกลับ วนซ้ำ และทบทวนก่อนก้าวไปข้างหน้าได้จริง'))
    A(separator())

    # CHAPTER 12 — Role Data
    A(h1('บทที่ 12: Role Data — บทบาท 3 มิติ'))
    A(pq('"Role Data แสดงบทบาทที่คุณแสดงออกตามบริบท แต่ละบทบาทมีสองด้าน — Externalization คือสิ่งที่คนอื่นเห็น และ Internalization คือสิ่งที่เกิดขึ้นภายในโดยที่คุณอาจไม่รู้สึกตัวเอง"'))
    A(make_table(
        ['บริบท', 'Externalization (คนอื่นเห็น)', 'Internalization (ภายในตัวเอง)'],
        [
            ['Collective Role', 'The Democrat', 'The Bigot'],
            ['Individual Role', 'Beauty', 'The Hermit'],
            ['Tribal Role', 'The Giver', 'The Loner'],
        ],
        col_widths=[2000, 3000, 3000]
    ))
    A(p(''))
    A(h2('Collective: The Democrat ↔ The Bigot'))
    A(p('ในบริบทส่วนรวม คนอื่นมองเห็น Jaja เป็น "The Democrat" — ยุติธรรม ฟังทุกเสียง ให้คุณค่ากับความหลากหลาย แต่ภายในตัวเอง (โดยไม่รู้ตัว) คือ "The Bigot" — ความยึดมั่นในมุมมองของตัวเองที่แน่วแน่ Gate 50 (Values) ทำให้มีมาตรฐานที่ไม่ยอมง่าย ๆ บางครั้ง "การรับฟัง" ภายนอกอาจเป็นแค่ผิวนอก ในขณะที่การตัดสินใจภายในมาจาก The Bigot แล้วตั้งแต่ต้น'))
    A(h2('Individual: Beauty ↔ The Hermit'))
    A(p('ในบริบท Individual คนอื่นมองเห็น Jaja เป็น "Beauty" — มีความสวยงามทั้งในรูปลักษณ์และพลังงาน ดึงดูดสายตาและความสนใจโดยธรรมชาติ ภายในตัวเอง คือ "The Hermit" — ความต้องการอยู่คนเดียว โดดเดี่ยว สอดคล้องกับ Profile Line 2 อย่างลึกซึ้ง ยิ่งเธอให้เวลาตัวเอง (The Hermit) ยิ่งพลังงาน Beauty ออกมาได้อย่างสวยงามและแท้จริง'))
    A(h2('Tribal: The Giver ↔ The Loner'))
    A(p('ในบริบท Tribal คนอื่นมองเห็น Jaja เป็น "The Giver" — ผู้ให้ ผู้ดูแล สอดคล้องกับ Channel 27-50 โดยตรง ภายในตัวเอง คือ "The Loner" — แม้อยู่ในกลุ่มใกล้ชิด ก็ยังมีพื้นที่ภายในที่ไม่มีใครเข้าถึง The Giver ทำงานได้ดีที่สุดเมื่อเธอเลือกดูแล "จากความอยาก" (Sacral Response) ไม่ใช่จากแรงกดดัน'))
    A(separator())

    # CHAPTER 13 — Keywords
    A(h1('บทที่ 13: Keywords ทั้ง 26 ตำแหน่ง'))
    A(h2('ธีมที่ถูกย้ำซ้ำมากที่สุด'))
    A(make_table(
        ['คำหลัก', 'Gate', 'ครั้ง', 'ความหมาย'],
        [
            ['Contraction ★★★', '41', '5', 'Gate 41.1, 41.2, 41.3, 41.4, 41.5 — แกนกลางของ Chart ทั้งหมด'],
            ['Listener ★★', '13', '3', 'Gate 13.1 (Empathy), 13.5 (Saviour), 13.6 (Optimist)'],
            ['The Logical Leader', '31', '2', 'Gate 31.2 (Arrogance), 31.3 (Selectivity) — ทิศทาง North Node'],
            ['Fixed Rhythms', '5', '2', 'Gate 5.4 (Hunter), 5.2 (Inner Peace) — Unconscious Sun + Pluto'],
            ['Focus', '9', '2', 'Gate 9.3 (Last Straw), 9.6 (Gratitude)'],
            ['Rationalization', '24', '2', 'Gate 24.5 (Confession), 24.6 (Gift Horse) — กับดักจิตใจ'],
        ],
        col_widths=[2200, 800, 700, 4300]
    ))
    A(p(''))
    A(h2('Keywords ฝั่ง Design (Unconscious — แดง)'))
    A(make_table(
        ['ดาว', 'Gate.Line', 'Keyword หลัก', 'Keyword รอง'],
        [
            ['☉ Sun', '5.4', 'Fixed Rhythms', 'The Hunter'],
            ['⊕ Earth', '35.4', 'Change', 'Hunger'],
            ['☊ North Node', '31.3', 'The Logical Leader', 'Selectivity'],
            ['☋ South Node', '41.3', 'Contraction', 'Efficiency'],
            ['☽ Moon', '9.3', 'Focus', 'The Last Straw'],
            ['☿ Mercury', '14.1', 'Power Skills', 'Money Isn\'t Everything'],
            ['♀ Venus', '50.6', 'Values', 'Leadership'],
            ['♂ Mars', '19.1', 'Wanting', 'Interdependence'],
            ['♃ Jupiter', '42.5', 'Endings', 'Self-actualisation'],
            ['♄ Saturn', '24.5', 'Rationalization', 'Confession'],
            ['♅ Uranus', '13.1', 'Listener', 'Empathy'],
            ['♆ Neptune', '41.1', 'Contraction', 'Reasonableness'],
            ['♇ Pluto', '9.6', 'Focus', 'Gratitude'],
        ],
        col_widths=[1000, 1000, 2500, 3500]
    ))
    A(p(''))
    A(h2('Keywords ฝั่ง Personality (Conscious — ดำ)'))
    A(make_table(
        ['ดาว', 'Gate.Line', 'Keyword หลัก', 'Keyword รอง'],
        [
            ['☉ Sun', '63.2', 'Doubts', 'Structuring'],
            ['⊕ Earth', '64.2', 'Confusion', 'Qualification'],
            ['☊ North Node', '31.2', 'The Logical Leader', 'Arrogance'],
            ['☋ South Node', '41.2', 'Contraction', 'Caution'],
            ['☽ Moon', '41.5', 'Contraction', 'Authorisation'],
            ['☿ Mercury', '37.5', 'Friendship', 'Love'],
            ['♀ Venus', '13.5', 'Listener', 'The Saviour'],
            ['♂ Mars', '51.1', 'Shock', 'Reference'],
            ['♃ Jupiter', '27.2', 'Caring', 'Self-sufficiency'],
            ['♄ Saturn', '24.6', 'Rationalization', 'The Gift Horse'],
            ['♅ Uranus', '13.6', 'Listener', 'The Optimist'],
            ['♆ Neptune', '41.4', 'Contraction', 'Correction'],
            ['♇ Pluto', '5.2', 'Fixed Rhythms', 'Inner Peace'],
        ],
        col_widths=[1000, 1000, 2500, 3500]
    ))
    A(separator())

    # CHAPTER 14 — Resonance Square
    A(h1('บทที่ 14: Resonance Square — แผนที่ระดับลึกของร่างกายและจิตใจ'))
    A(pq('"Resonance Square คือการจัดวาง 18 ตำแหน่งพิเศษในรูปแบบตาราง 3×3 สองชุด ชุดบน (จิตใจ) และชุดล่าง (ร่างกาย) เพื่อดูความสัมพันธ์เชิงพลังงานที่ลึกที่สุด"'))
    A(h2('ตาราง Resonance Square บน (จิตใจ / Rave Mind)'))
    A(make_table(
        ['', 'ซ้าย', 'กลาง', 'ขวา'],
        [
            ['แถว 1', 'The Focus\n41.5.3', 'Communication\n37.5.3', 'The Sidetrack\n13.6.3'],
            ['แถว 2', 'The Standard\n13.5.1', 'Mutation ★\n51.1.1', 'Misinformation\n41.4.4'],
            ['แถว 3', 'The Constraint\n24.6.2', 'Rules\n27.2.1', 'Truth\n5.2.4'],
        ],
        col_widths=[800, 2400, 2400, 2400]
    ))
    A(p(''))
    A(h2('ตาราง Resonance Square ล่าง (ร่างกาย / Body)'))
    A(make_table(
        ['', 'ซ้าย', 'กลาง', 'ขวา'],
        [
            ['แถว 1', 'Movement\n13.1.3', 'Voice\n14.1.4', 'Gravity\n9.3.6'],
            ['แถว 2', 'Hydration\n41.1.3', 'Temperature\n19.1.6', 'Complexion\n50.6.4'],
            ['แถว 3', 'Inhalation\n9.6.1', 'Aura\n42.5.6', 'Exhaust\n24.5.1'],
        ],
        col_widths=[800, 2400, 2400, 2400]
    ))
    A(p(''))
    A(h2('จุดสำคัญใน Resonance Square'))
    A(bullet('Gate 41 ปรากฏทั้งบนและล่าง: The Focus + Misinformation (จิตใจ) และ Hydration (ร่างกาย) — "จินตนาการและความฝัน" คือเชื้อเพลิงที่หล่อเลี้ยงทั้งร่างกายและจิตใจของ Jaja'))
    A(bullet('Gate 13 ปรากฏทั้งบนและล่าง: Sidetrack + Standard (จิตใจ) และ Movement (ร่างกาย) — ร่างกายของเธอ "เคลื่อนไหว" ตามการฟัง ไม่ใช่ตามความทะเยอทะยาน'))
    A(bullet('"Mutation" (51.1.1) อยู่ตรงกลางตารางบน: การเปลี่ยนแปลงที่ยิ่งใหญ่ในชีวิต Jaja มักมาจาก "ความช็อก" หรือประสบการณ์กระทบกระเทือน ไม่ใช่การค่อย ๆ สะสม'))
    A(bullet('"Truth" (5.2 — Fixed Rhythms / Inner Peace) อยู่มุมขวาล่างตารางบน: ความจริงของ Jaja อยู่ที่ "ความสม่ำเสมอที่นำไปสู่ Inner Peace" ไม่ใช่การตอบคำถามใหญ่'))
    A(bullet('"Rules" (27.2 — Caring / Self-sufficiency) อยู่กลางแถวล่างของตารางบน: กฎที่แท้จริงของจิตใจเธอคือการดูแลตัวเองและคนที่รักอย่างพอดี'))
    A(separator())

    # CONCLUSION
    A(h1('บทสรุป — ข้อความถึง Jaja โดยตรง'))
    A(h2('แก่นของ Design นี้'))
    A(p('Jaja คือ Pure Generator ที่มีพลังงาน Sacral อันทรงพลังสำหรับการดูแลและรักษาคุณค่า (Channel 27-50) แต่ใช้ชีวิตส่วนใหญ่อยู่ใน "โลกที่เปิดกว้าง" รับพลังงาน ความคิด อารมณ์ และตัวตนจากคนรอบข้างอยู่ตลอดเวลา งานที่สำคัญที่สุดในชีวิตของเธอไม่ใช่การหา "ตัวเอง" แต่การฝึกแยกแยะว่า "อันนี้คือฉัน" กับ "อันนี้คือพลังงานที่รับมาจากคนอื่น"'))
    A(h2('จุดแข็งที่โดดเด่น 3 อย่าง'))
    A(bullet('พลังการดูแลที่แท้จริง (Channel 27-50): มีพลังในการดูแลคนที่รักอย่างสม่ำเสมอและมีมาตรฐาน คนที่อยู่ในวงของเธอจะรู้สึกปลอดภัยและมีคุณค่า', '1. '))
    A(bullet('พรสวรรค์ที่คนอื่นเห็นก่อน (Line 2 สูงสุด 23%): สิ่งที่ทำได้อย่างเป็นธรรมชาติที่สุด มักคือสิ่งที่หายากและมีคุณค่าในสายตาคนอื่น — การฟัง การตั้งคำถาม การถือความไม่แน่นอน', '2. '))
    A(bullet('เส้นทางสู่อิทธิพล (Story Line Gate 41 → Gate 31): กำลังเดินทางจากพลังงานแห่งความฝัน ไปสู่การเป็นเสียงที่คนเลือกฟัง ผ่านการสะสมและแบ่งปันประสบการณ์ (Abstract Circuit 42%)', '3. '))
    A(h2('ความท้าทายหลักและวิธีรับมือ'))
    A(p('ความท้าทายใหญ่ที่สุด: "การแยกแยะตัวเอง" เมื่อ 7 ใน 9 Centers เปิดรับ โลกภายในของ Jaja จะเต็มไปด้วยพลังงานของคนอื่น'))
    A(bullet('Sacral คือ Anchor — เมื่อสับสนว่าตัวเองคิดหรือรู้สึกอะไร กลับมาที่ Sacral เสมอ'))
    A(bullet('เวลาคนเดียวคือยาวิเศษ — ทุกวัน ให้ตัวเองมีเวลา "ปล่อย" พลังงานที่รับมาออกไป'))
    A(bullet('เลือกสภาพแวดล้อมอย่างจริงจัง — G Center เปิด สภาพแวดล้อมและคนรอบข้างจะกำหนดทิศทางชีวิตได้มาก'))
    A(h2('จุดเริ่มต้นที่แนะนำ'))
    A(p('เริ่มจากสิ่งเดียวเท่านั้น: ฝึกฟัง Sacral ก่อนตอบทุกคำถาม'))
    A(p('เมื่อมีคนถาม "อยากทำสิ่งนี้ไหม?" — แทนที่จะตอบทันที ให้หยุด 3 วินาที สูดหายใจ แล้วสังเกตว่าร่างกายรู้สึกอะไรก่อนสมองจะคิด ทำแค่นี้อย่างเดียวเป็นเวลา 1 เดือน แล้วสังเกตว่าชีวิตเปลี่ยนไปอย่างไร'))
    A(h2('ข้อความให้กำลังใจ'))
    A(pq('"Human Design ไม่ใช่โชคชะตาที่ตายตัว — มันคือแผนที่ ทุกครั้งที่รู้สึก Frustration ไม่ใช่สัญญาณว่าล้มเหลว — มันคือ Feedback ที่ร่างกายส่งมาบอกว่า มีบางอย่างไม่ตรง กลับมาฟัง Sacral และทุกครั้งที่รู้สึก Satisfaction แม้เพียงเล็กน้อย — นั่นคือร่างกายที่กำลังบอกว่า ใช่ นี่แหละคือตัวเธอ เชื่อสิ่งนั้น"'))
    A(separator())
    A(p('รายงานนี้จัดทำจาก Chart ของ Jaja Z เกิด 3 มีนาคม 2543 เวลา 08:29 น. ป้อมปราบศัตรูพ่าย กรุงเทพมหานคร | ระบบ Tropical | Genetic Matrix', size=16, color='888888'))

    return ''.join(parts)


# ========================
# DOCX BUILDER
# ========================

CONTENT_TYPES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>'''

RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>'''

DOC_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
</Relationships>'''

STYLES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:pPr><w:spacing w:after="100"/></w:pPr>
    <w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/><w:lang w:val="th-TH" w:eastAsia="th-TH" w:bidi="th-TH"/></w:rPr>
  </w:style>
</w:styles>'''

SETTINGS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:defaultTabStop w:val="720"/>
</w:settings>'''

CORE = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
  xmlns:dc="http://purl.org/dc/elements/1.1/">
  <dc:title>Human Design Report — Jaja Z</dc:title>
  <dc:creator>HDrepTH</dc:creator>
  <cp:lastModifiedBy>HDrepTH</cp:lastModifiedBy>
</cp:coreProperties>'''

APP = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
  <Application>HDrepTH Report Generator</Application>
</Properties>'''

def make_document(body_content):
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    {body_content}
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1080" w:bottom="1440" w:left="1080"/>
    </w:sectPr>
  </w:body>
</w:document>'''

# Build
output_path = '/home/user/Claude/HD_Report_Jaja_Z.docx'
body = build_body()
doc_xml = make_document(body)

with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
    zf.writestr('[Content_Types].xml', CONTENT_TYPES)
    zf.writestr('_rels/.rels', RELS)
    zf.writestr('word/document.xml', doc_xml.encode('utf-8'))
    zf.writestr('word/_rels/document.xml.rels', DOC_RELS)
    zf.writestr('word/styles.xml', STYLES)
    zf.writestr('word/settings.xml', SETTINGS)
    zf.writestr('docProps/core.xml', CORE)
    zf.writestr('docProps/app.xml', APP)

print(f'Done: {output_path}')
import os
print(f'Size: {os.path.getsize(output_path):,} bytes')
