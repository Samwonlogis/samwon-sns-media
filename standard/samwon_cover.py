"""삼원종합물류(주) 표지(메인페이지) 표준 — 대표 지시 2026-11-02

모든 Word·Excel·PowerPoint 산출물의 첫 장은 아래 4요소로 구성한다.
  ① 영문 로고(SAMWON)  ② 작업물 내용(제목·부제)  ③ 날짜  ④ 결재방(기안 · 팀장 · 최종결재자)

사용법 (samwon_kit 대신 이 모듈을 import — kit의 모든 함수를 그대로 쓰고 표지만 교체됨)
    from samwon_cover import *
    doc = docx_new("구매요청서", "구매·지출 승인 절차", team="총무팀", doc_no="SW-GA-2611-01",
                   drafter="총무 대리", leader="총무팀장", approver="대표이사")
    wb  = xlsx_new("차고지 이용현황 대장", team="총무팀", drafter="총무 대리", leader="총무팀장")
    prs = pptx_new("화주 모집 카드뉴스", team="홍보팀", drafter="홍보 과장", leader="홍보팀장")
"""
import datetime

import samwon_kit as _k
from samwon_kit import *  # noqa: F401,F403  (본문 작성 함수 재사용)
from samwon_kit import (ADDR, BLUE, COMPANY, CONTACT, DEEP, FONT, GRAY, INK, LOGO, NAVY, SKY, SYMBOL, WARN, _borders,
                        _font, _ratio, _shade)

TODAY_KO = datetime.date.today().strftime("%Y. %m. %d.")
APPROVAL = ("기안", "팀장", "최종결재자")


def _names(drafter, leader, approver):
    return [drafter or "", leader or "", approver or "대표이사"]


# ───────────────────────────── Word ─────────────────────────────
from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt


def _row_h(row, cm):
    trPr = row._tr.get_or_add_trPr(); h = OxmlElement("w:trHeight")
    h.set(qn("w:val"), str(int(cm * 567))); h.set(qn("w:hRule"), "exact"); trPr.append(h)


def _no_borders(table):
    tblPr = table._element.tblPr; b = OxmlElement("w:tblBorders")
    for e in ("top", "left", "bottom", "right", "insideH", "insideV"):
        x = OxmlElement(f"w:{e}"); x.set(qn("w:val"), "nil"); b.append(x)
    tblPr.append(b)


def _approval_docx(cell, names):
    """결재방: 왼쪽 세로 '결재' + 기안·팀장·최종결재자 (직위 / 서명 / 날짜)"""
    t = cell.add_table(rows=3, cols=4); _borders(t, "64748B", 6); t.alignment = WD_TABLE_ALIGNMENT.RIGHT
    a = t.cell(0, 0).merge(t.cell(2, 0)); a.width = Cm(0.9); _shade(a, SKY)
    a.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER; a.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    _font(a.paragraphs[0].add_run("결\n\n재"), 9, True, NAVY)
    for j, k in enumerate(APPROVAL, 1):
        h = t.cell(0, j); h.width = Cm(2.4); _shade(h, SKY)
        h.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER; _font(h.paragraphs[0].add_run(k), 9, True, NAVY)
        s = t.cell(1, j); s.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        s.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.BOTTOM; _font(s.paragraphs[0].add_run(names[j - 1]), 7.5, False, GRAY)
        d = t.cell(2, j); d.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER; _font(d.paragraphs[0].add_run("    .    ."), 8, False, GRAY)
    _row_h(t.rows[0], 0.6); _row_h(t.rows[1], 1.6); _row_h(t.rows[2], 0.55)
    return t


def docx_new(title, subtitle="", team="", doc_no="", version="제정 v1.0", landscape=False, kind="",
             drafter="", leader="", approver="대표이사", date=None):
    doc = Document()
    st = doc.styles["Normal"]; st.font.name = FONT; st.font.size = Pt(10.5)
    st.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    st.paragraph_format.space_after = Pt(4); st.paragraph_format.line_spacing = 1.35
    sec = doc.sections[0]
    if landscape:
        sec.orientation = WD_ORIENT.LANDSCAPE; sec.page_width, sec.page_height = Cm(29.7), Cm(21.0)
    else:
        sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
    sec.left_margin = sec.right_margin = Cm(2.0); sec.top_margin = Cm(1.6); sec.bottom_margin = Cm(1.6)
    sec.different_first_page_header_footer = True
    hp = sec.header.paragraphs[0]; hp.alignment = WD_ALIGN_PARAGRAPH.LEFT
    hp.add_run().add_picture(LOGO, height=Cm(0.55)); _font(hp.add_run(f"    {title}"), 8, color=GRAY)
    fp = sec.footer.paragraphs[0]; fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _font(fp.add_run(f"{COMPANY}  ·  {doc_no + '  ·  ' if doc_no else ''}"), 8, color=GRAY)
    _k._field(fp, "PAGE"); _font(fp.add_run(" / "), 8, color=GRAY); _k._field(fp, "NUMPAGES")

    usable = (29.7 if landscape else 21.0) - 4.0
    # ① 로고(좌) + ④ 결재방(우)
    top = doc.add_table(rows=1, cols=2); _no_borders(top)
    l, r = top.rows[0].cells; l.width = Cm(usable - 9.3); r.width = Cm(9.3)
    l.paragraphs[0].add_run().add_picture(LOGO, height=Cm(1.5))
    r.paragraphs[0].paragraph_format.space_after = Pt(0)
    _approval_docx(r, _names(drafter, leader, approver))
    for _ in range(5 if not landscape else 2): doc.add_paragraph()
    # ② 작업물 내용
    if kind:
        p = doc.add_paragraph(); _font(p.add_run(kind), 12, True, BLUE)
    p = doc.add_paragraph(); p.paragraph_format.space_after = Pt(6); _font(p.add_run(title), 30, True, NAVY)
    if subtitle:
        p = doc.add_paragraph(); _font(p.add_run(subtitle), 14, False, "374151")
    p = doc.add_paragraph(); pPr = p._p.get_or_add_pPr(); bd = OxmlElement("w:pBdr"); bt = OxmlElement("w:bottom")
    bt.set(qn("w:val"), "single"); bt.set(qn("w:sz"), "24"); bt.set(qn("w:color"), BLUE); bd.append(bt); pPr.append(bd)
    # ③ 날짜
    p = doc.add_paragraph(); p.paragraph_format.space_before = Pt(10)
    _font(p.add_run("작성일  "), 12, True, NAVY); _font(p.add_run(date or TODAY_KO), 16, True, "111827")
    for _ in range(2 if not landscape else 1): doc.add_paragraph()
    info = [("법인명", COMPANY), ("주관 부서", team or "-"), ("문서번호", doc_no or "-"), ("제정·개정", version)]
    t = doc.add_table(rows=len(info), cols=2); t.alignment = WD_TABLE_ALIGNMENT.LEFT; _borders(t, "D1D9E6")
    for i, (k, v) in enumerate(info):
        c0, c1 = t.rows[i].cells; c0.width = Cm(3.5); c1.width = Cm(12.5); _shade(c0, SKY)
        _font(c0.paragraphs[0].add_run(k), 10, True, NAVY); _font(c1.paragraphs[0].add_run(v), 10)
    for _ in range(2): doc.add_paragraph()
    p = doc.add_paragraph(); _font(p.add_run(_k.REVIEW_NOTE), 8, color=GRAY)
    p = doc.add_paragraph(); _font(p.add_run(f"{COMPANY}  |  {ADDR}  |  {CONTACT}"), 8, color=GRAY)
    p = doc.add_paragraph(); p.add_run().add_break(WD_BREAK.PAGE)
    return doc


# ───────────────────────────── Excel ─────────────────────────────
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Border, PatternFill, Side

XF = _k.XF
_T = Side(style="thin", color="64748B")
_BOX = Border(left=_T, right=_T, top=_T, bottom=_T)


def xlsx_new(title, subtitle="", team="", doc_no="", version="제정 v1.0", how_to=None,
             drafter="", leader="", approver="대표이사", date=None):
    wb = Workbook(); ws = wb.active; ws.title = "표지"; ws.sheet_view.showGridLines = False
    for col, w in zip("ABCDEFGHI", (3, 16, 16, 16, 5, 13, 13, 13, 3)):
        ws.column_dimensions[col].width = w
    img = XLImage(LOGO); img.height = 56; img.width = int(56 * _ratio()); ws.add_image(img, "B2")
    # 결재방 E2:H4
    ws.merge_cells("E2:E4"); c = ws["E2"]; c.value = "결\n재"
    c.font = XF(9, True, NAVY); c.fill = PatternFill("solid", fgColor=SKY)
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    names = _names(drafter, leader, approver)
    for j, k in enumerate(APPROVAL):
        col = 6 + j
        h = ws.cell(row=2, column=col, value=k); h.font = XF(9, True, NAVY); h.fill = PatternFill("solid", fgColor=SKY)
        h.alignment = Alignment(horizontal="center", vertical="center")
        s = ws.cell(row=3, column=col, value=names[j]); s.font = XF(8, False, GRAY)
        s.alignment = Alignment(horizontal="center", vertical="bottom")
        d = ws.cell(row=4, column=col, value="   .   ."); d.font = XF(8, False, GRAY); d.alignment = Alignment(horizontal="center")
    for r in (2, 3, 4):
        for col in range(5, 9): ws.cell(row=r, column=col).border = _BOX
    ws.row_dimensions[2].height = 18; ws.row_dimensions[3].height = 44; ws.row_dimensions[4].height = 16
    ws["B7"] = COMPANY; ws["B7"].font = XF(12, True, BLUE)
    ws["B9"] = title; ws["B9"].font = XF(24, True, NAVY)
    if subtitle:
        ws["B11"] = subtitle; ws["B11"].font = XF(13, False, "374151")
    for col in range(2, 9):
        ws.cell(row=12, column=col).border = Border(bottom=Side(style="thick", color=BLUE))
    ws["B13"] = "작성일"; ws["B13"].font = XF(11, True, NAVY)
    ws["C13"] = date or TODAY_KO; ws["C13"].font = XF(14, True, "111827")
    info = [("법인명", COMPANY), ("주관 부서", team or "-"), ("문서번호", doc_no or "-"), ("제정·개정", version)]
    for i, (k, v) in enumerate(info):
        r = 15 + i
        ws.cell(row=r, column=2, value=k).font = XF(10, True, NAVY); ws.cell(row=r, column=2).fill = PatternFill("solid", fgColor=SKY)
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=8)
        ws.cell(row=r, column=3, value=v).font = XF(10)
        for col in range(2, 9): ws.cell(row=r, column=col).border = _k.BOX
    r = 20
    if how_to:
        ws.cell(row=r, column=2, value="사용 방법").font = XF(11, True, BLUE); r += 1
        for line in how_to:
            ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=8)
            c = ws.cell(row=r, column=2, value=f"• {line}"); c.font = XF(10); c.alignment = Alignment(wrap_text=True, vertical="top")
            ws.row_dimensions[r].height = 30 if len(line) > 60 else 18; r += 1
    r += 1
    ws.merge_cells(start_row=r, start_column=2, end_row=r + 1, end_column=8)
    c = ws.cell(row=r, column=2, value=_k.REVIEW_NOTE); c.font = XF(8, False, GRAY); c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.cell(row=r + 3, column=2, value=f"{COMPANY} | {ADDR} | {CONTACT}").font = XF(8, False, GRAY)
    ws.page_setup.paperSize = ws.PAPERSIZE_A4; ws.print_options.horizontalCentered = True
    ws.page_setup.fitToWidth = 1; ws.sheet_properties.pageSetUpPr.fitToPage = True
    return wb


# ───────────────────────────── PowerPoint ─────────────────────────────
from pptx import Presentation
from pptx.dml.color import RGBColor as PRGB
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Inches, Pt as PPt


def _cell_txt(cell, text, size, bold=False, color="111827", fill=None, anchor=MSO_ANCHOR.MIDDLE):
    cell.text = ""; tf = cell.text_frame; p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run(); r.text = text; r.font.size = PPt(size); r.font.bold = bold; r.font.name = FONT
    r.font.color.rgb = PRGB.from_string(color); cell.vertical_anchor = anchor
    cell.margin_left = cell.margin_right = Inches(0.03); cell.margin_top = cell.margin_bottom = Inches(0.02)
    if fill:
        cell.fill.solid(); cell.fill.fore_color.rgb = PRGB.from_string(fill)
    else:
        cell.fill.solid(); cell.fill.fore_color.rgb = PRGB.from_string("F8FAFC")
    # 칸 테두리(회색 실선)
    from pptx.oxml.ns import qn as _qn
    from lxml import etree
    tcPr = cell._tc.get_or_add_tcPr()
    for tag in ("a:lnL", "a:lnR", "a:lnT", "a:lnB"):
        ln = etree.SubElement(tcPr, _qn(tag), w="9525")
        sf = etree.SubElement(ln, _qn("a:solidFill")); etree.SubElement(sf, _qn("a:srgbClr"), val="64748B")


def pptx_new(title, subtitle="", team="", doc_no="", wide=True, drafter="", leader="", approver="대표이사", date=None):
    prs = Presentation()
    if wide:
        prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    s = prs.slides.add_slide(prs.slide_layouts[6]); W, H = prs.slide_width, prs.slide_height
    bar = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, H - Inches(0.35), W, Inches(0.35))
    bar.fill.solid(); bar.fill.fore_color.rgb = PRGB.from_string(NAVY); bar.line.fill.background()
    s.shapes.add_picture(LOGO, Inches(0.7), Inches(0.6), height=Inches(0.75))
    # 결재방(우상단)
    names = _names(drafter, leader, approver)
    tw = Inches(3.9); tbl = s.shapes.add_table(3, 4, W - tw - Inches(0.6), Inches(0.5), tw, Inches(1.35)).table
    tbl.columns[0].width = Inches(0.45)
    for j in range(1, 4): tbl.columns[j].width = Inches(1.15)
    tbl.rows[0].height = Inches(0.3); tbl.rows[1].height = Inches(0.8); tbl.rows[2].height = Inches(0.25)
    m = tbl.cell(0, 0); m.merge(tbl.cell(2, 0)); _cell_txt(m, "결재", 10, True, NAVY, SKY)
    for j, k in enumerate(APPROVAL, 1):
        _cell_txt(tbl.cell(0, j), k, 10, True, NAVY, SKY)
        _cell_txt(tbl.cell(1, j), names[j - 1], 8, False, GRAY, anchor=MSO_ANCHOR.BOTTOM)
        _cell_txt(tbl.cell(2, j), "   .   .", 8, False, GRAY)
    tb = s.shapes.add_textbox(Inches(0.8), Inches(2.5), W - Inches(1.6), Inches(2.4)).text_frame; tb.word_wrap = True
    _k._ptxt(tb, COMPANY, 18, True, BLUE); _k._ptxt(tb, title, 40, True, NAVY)
    if subtitle: _k._ptxt(tb, subtitle, 18, False, "374151")
    db = s.shapes.add_textbox(Inches(0.8), Inches(5.0), Inches(6), Inches(0.6)).text_frame
    _k._ptxt(db, f"작성일  {date or TODAY_KO}", 16, True, "111827")
    ib = s.shapes.add_textbox(Inches(0.8), H - Inches(1.4), W - Inches(1.6), Inches(0.9)).text_frame
    _k._ptxt(ib, f"{team}  ·  {doc_no or ''}", 12, False, GRAY)
    _k._ptxt(ib, f"{ADDR}  |  {CONTACT}", 10, False, GRAY)
    return prs
