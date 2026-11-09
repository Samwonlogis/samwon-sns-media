"""삼원종합물류(주) 표지(메인페이지) 표준 v3 — 대표 지시 2026-11-02 · 디자인 개편 2026-11-09

모든 Word·Excel·PowerPoint 산출물의 첫 장은 아래 4요소로 구성한다.
  ① 영문 로고(SAMWON)  ② 작업물 내용(구분·제목·부제)  ③ 작성일  ④ 결재방(기안 · 팀장 · 최종결재자)

디자인(모던 클린): 흰 바탕 + 상단 가는 블루 선 + 하단 진한 네이비 띠(회사 정보) · 좌측 정렬 대제목 ·
짧은 블루 강조 막대 · 4칸 문서 정보 줄 · 우하단 결재방. 표지는 독립 구역(section)이라 본문 쪽 번호는 1부터 센다.

사용법 (samwon_kit 대신 이 모듈을 import — 본문 함수는 kit 것을 그대로 쓴다)
    from samwon_cover import *
    doc = docx_new("구매요청서", "구매·지출 승인 절차", team="총무팀", doc_no="SW-GA-2611-01",
                   drafter="총무 대리", leader="총무팀장", approver="대표이사", kind="규정·서식")
    wb  = xlsx_new("차고지 이용현황 대장", team="총무팀", drafter="총무 대리", leader="총무팀장", how_to=["…"])
    prs = pptx_new("화주 모집 카드뉴스", team="홍보팀", drafter="홍보 과장", leader="홍보팀장")
"""
import datetime
import math
import os

import samwon_kit as _k
from samwon_kit import *  # noqa: F401,F403  (본문 작성 함수 재사용)
from samwon_kit import (ADDR, BLUE, COMPANY, CONTACT, DEEP, FONT, GRAY, INK, LINE, LOGO, NAVY, PALE, SKY, SYMBOL,
                        WARN, _borders, _cell_border, _cell_margins, _fixed_widths, _font, _no_borders, _ratio,
                        _row_props, _shade, disp_len)

TODAY_KO = datetime.date.today().strftime("%Y. %m. %d.")
APPROVAL = ("기안", "팀장", "최종결재자")


def _names(drafter, leader, approver):
    return [drafter or "", leader or "", approver or "대표이사"]


def _meta(team, doc_no, date, version):
    return [("주관 부서", team or "-"), ("문서번호", doc_no or "-"), ("작성일", date or TODAY_KO), ("제정·개정", version or "제정 v1.0")]


# ───────────────────────────── Word ─────────────────────────────
from docx import Document
from docx.enum.section import WD_ORIENT, WD_SECTION
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Cm, Emu, Pt

_CM = 360000


def _exact(p, cm, before=0.0):
    f = p.paragraph_format
    f.space_before = Cm(before) if before else Pt(0); f.space_after = Pt(0)
    f.line_spacing_rule = WD_LINE_SPACING.EXACTLY; f.line_spacing = Cm(cm)
    f.widow_control = False
    return p


def _gap(container, cm):
    p = container.add_paragraph(); _exact(p, cm)
    _font(p.add_run(""), 2)
    return p


def _anchor_bg(run, path, w_cm, h_cm):
    """쪽 전체 배경 그림(글 뒤, 쪽 기준 0,0)을 run에 붙인다."""
    run.add_picture(path, width=Cm(w_cm), height=Cm(h_cm))
    inline = run._element.findall(".//" + qn("wp:inline"))[-1]
    graphic = inline.find(qn("a:graphic"))
    docPr = inline.find(qn("wp:docPr"))
    cx, cy = int(w_cm * _CM), int(h_cm * _CM)
    anchor = parse_xml(
        f'<wp:anchor {nsdecls("wp", "a", "pic", "r")} distT="0" distB="0" distL="0" distR="0" simplePos="0" '
        f'relativeHeight="0" behindDoc="1" locked="1" layoutInCell="1" allowOverlap="1">'
        f'<wp:simplePos x="0" y="0"/>'
        f'<wp:positionH relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionH>'
        f'<wp:positionV relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionV>'
        f'<wp:extent cx="{cx}" cy="{cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>'
        f'<wp:docPr id="{docPr.get("id")}" name="cover-background"/><wp:cNvGraphicFramePr/></wp:anchor>')
    anchor.append(graphic)
    inline.getparent().replace(inline, anchor)


def _approval_docx(cell, names, size=8.5):
    """결재방: 왼쪽 '결재' + 기안·팀장·최종결재자 (직위 / 서명 / 날짜). 전체 8.3cm × 2.6cm"""
    t = cell.add_table(rows=3, cols=4); t.alignment = WD_TABLE_ALIGNMENT.RIGHT
    _borders(t, LINE, 4, outer=NAVY, outer_sz=8); _cell_margins(t, 0.05, 0.02)
    _fixed_widths(t, [0.8, 2.5, 2.5, 2.5])
    a = t.cell(0, 0).merge(t.cell(2, 0)); _shade(a, NAVY)
    a.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    pa = a.paragraphs[0]; pa.alignment = WD_ALIGN_PARAGRAPH.CENTER; _exact(pa, 0.42)
    _font(pa.add_run("결"), size, True, "FFFFFF")
    pb = a.add_paragraph(); pb.alignment = WD_ALIGN_PARAGRAPH.CENTER; _exact(pb, 0.42); _font(pb.add_run("재"), size, True, "FFFFFF")
    for j, k in enumerate(APPROVAL, 1):
        h = t.cell(0, j); _shade(h, SKY); h.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        ph = h.paragraphs[0]; ph.alignment = WD_ALIGN_PARAGRAPH.CENTER; _exact(ph, 0.42); _font(ph.add_run(k), size, True, NAVY)
        s = t.cell(1, j); s.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.BOTTOM
        ps = s.paragraphs[0]; ps.alignment = WD_ALIGN_PARAGRAPH.CENTER; _exact(ps, 0.36); _font(ps.add_run(names[j - 1]), 7, False, GRAY)
        d = t.cell(2, j); d.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        pd = d.paragraphs[0]; pd.alignment = WD_ALIGN_PARAGRAPH.CENTER; _exact(pd, 0.36); _font(pd.add_run("     .     ."), 7, False, GRAY)
    _row_props(t.rows[0], True, False, 0.55, "exact"); _row_props(t.rows[1], True, False, 1.5, "exact")
    _row_props(t.rows[2], True, False, 0.5, "exact")
    return t


def _lines(text, size_pt, width_cm, factor=1.03):
    if not text:
        return 0
    per = width_cm / (size_pt * 0.0353 * factor)
    return max(1, math.ceil(disp_len(text) / per))


def _cover_docx(doc, title, subtitle, kind, team, doc_no, version, date, names, landscape):
    sec = doc.sections[0]
    pw, ph = (29.7, 21.0) if landscape else (21.0, 29.7)
    sec.top_margin = Cm(1.9); sec.bottom_margin = Cm(2.75 if not landscape else 2.55)
    sec.left_margin = sec.right_margin = Cm(2.0)
    W = pw - 4.0; H = ph - sec.top_margin.cm - sec.bottom_margin.cm
    bg = _k._asset("cover_bg_landscape.png" if landscape else "cover_bg_portrait.png")
    used = 0.0
    # ① 로고 + 배경
    p = doc.paragraphs[0] if doc.paragraphs else doc.add_paragraph(); _exact(p, 1.35); used += 1.35
    p.add_run().add_picture(LOGO, height=Cm(1.05))
    if os.path.exists(bg):
        _anchor_bg(p.add_run(), bg, pw, ph)
    top_gap = 4.9 if not landscape else 2.3
    _gap(doc, top_gap); used += top_gap
    # ② 구분 라벨 · 강조 막대 · 제목 · 부제
    p = doc.add_paragraph(); _exact(p, 0.62); used += 0.62
    r = _font(p.add_run(kind or "업무 문서"), 11, True, BLUE)
    sp = OxmlElement("w:spacing"); sp.set(qn("w:val"), "30"); r._element.get_or_add_rPr().append(sp)
    _gap(doc, 0.22); used += 0.22
    bar = doc.add_table(rows=1, cols=1); _no_borders(bar); _cell_margins(bar, 0, 0); _fixed_widths(bar, [1.6])
    bar.alignment = WD_TABLE_ALIGNMENT.LEFT; _shade(bar.rows[0].cells[0], BLUE)
    _row_props(bar.rows[0], True, False, 0.13, "exact"); _exact(bar.rows[0].cells[0].paragraphs[0], 0.1)
    used += 0.13
    _gap(doc, 0.42); used += 0.42
    tsize = 30 if disp_len(title) <= 34 else 26
    tl = min(3, _lines(title, tsize, W)); lh = tsize * 0.0353 * 1.38
    p = doc.add_paragraph(); _exact(p, lh); used += lh * tl
    _font(p.add_run(title), tsize, True, NAVY)
    if subtitle:
        _gap(doc, 0.32); used += 0.32
        sl = min(2, _lines(subtitle, 13, W)); p = doc.add_paragraph(); _exact(p, 0.68); used += 0.68 * sl
        _font(p.add_run(subtitle), 13, False, GRAY)
    # ③ 문서 정보 줄(아래쪽)  ④ 결재방
    bottom = 1.4 + 0.75 + 2.65 + 0.12
    flex = max(0.6, H - used - bottom - 0.45)
    _gap(doc, flex)
    meta = _meta(team, doc_no, date, version)
    mt = doc.add_table(rows=2, cols=4); mt.alignment = WD_TABLE_ALIGNMENT.LEFT
    _no_borders(mt); _cell_margins(mt, 0.0, 0.0); _fixed_widths(mt, [W / 4] * 4)
    for j, (k, v) in enumerate(meta):
        c0, c1 = mt.cell(0, j), mt.cell(1, j)
        _cell_border(c0, top=("single", 12, NAVY))
        c0.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.BOTTOM
        p0 = c0.paragraphs[0]; _exact(p0, 0.42); _font(p0.add_run(k), 8, False, GRAY)
        p1 = c1.paragraphs[0]; _exact(p1, 0.55); _font(p1.add_run(v), 10.5, True, INK)
    _row_props(mt.rows[0], True, False, 0.7, "exact"); _row_props(mt.rows[1], True, False, 0.7, "exact")
    _gap(doc, 0.75)
    bt = doc.add_table(rows=1, cols=2); _no_borders(bt); _cell_margins(bt, 0, 0)
    _fixed_widths(bt, [W - 8.5, 8.5]); bt.alignment = WD_TABLE_ALIGNMENT.LEFT
    _row_props(bt.rows[0], True, False, 2.65, "exact")
    l, rcell = bt.rows[0].cells
    l.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.BOTTOM
    pl = l.paragraphs[0]; pl.paragraph_format.space_after = Pt(0); pl.paragraph_format.line_spacing = 1.25
    pl.paragraph_format.right_indent = Cm(0.8)
    _font(pl.add_run(_k.REVIEW_NOTE), 7.5, False, GRAY)
    rp = rcell.paragraphs[0]; _exact(rp, 0.05)
    _approval_docx(rcell, names)
    # 표지 구역 끝 → 본문 구역(쪽 번호 1부터)
    doc.add_section(WD_SECTION.NEW_PAGE)
    for p in reversed(doc.paragraphs):
        if p._p.find(qn("w:pPr")) is not None and p._p.pPr.find(qn("w:sectPr")) is not None:
            _exact(p, 0.05); break
    return doc


def docx_new(title, subtitle="", team="", doc_no="", version="제정 v1.0", landscape=False, kind="",
             drafter="", leader="", approver="대표이사", date=None):
    doc = Document()
    _k.setup_styles(doc); _k.page_setup(doc.sections[0], landscape)
    _cover_docx(doc, title, subtitle, kind, team, doc_no, version, date, _names(drafter, leader, approver), landscape)
    # 표지 구역은 머리글·바닥글 없음 / 본문 구역: 표준 머리글·바닥글, 쪽 번호 1부터
    body = doc.sections[-1]
    _k.page_setup(body, landscape)
    _k.body_header_footer(body, title, doc_no, restart=True)
    doc._samwon_cover = True
    return doc


# ───────────────────────────── Excel ─────────────────────────────
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Border, PatternFill, Side
from openpyxl.worksheet.page import PageMargins
from openpyxl import Workbook

XF = _k.XF
_LN = Side(style="thin", color=LINE)
_NV = Side(style="thin", color=NAVY)


def _fill(c, color):
    c.fill = PatternFill("solid", fgColor=color)


def xlsx_new(title, subtitle="", team="", doc_no="", version="제정 v1.0", how_to=None,
             drafter="", leader="", approver="대표이사", date=None, kind=""):
    wb = Workbook(); ws = wb.active; ws.title = "표지"; ws.sheet_view.showGridLines = False
    widths = {"A": 2.5, "B": 11, "C": 11, "D": 11, "E": 11, "F": 4.5, "G": 12, "H": 12, "I": 12, "J": 2.5}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w
    H = ws.row_dimensions
    _fill_row = lambda r, color, c1=1, c2=10: [_fill(ws.cell(row=r, column=c), color) for c in range(c1, c2 + 1)]
    H[1].height = 5; _fill_row(1, BLUE)                       # 상단 가는 블루 선
    H[2].height = 16
    H[3].height = 40
    img = XLImage(LOGO); img.height = 44; img.width = int(44 * _ratio()); ws.add_image(img, "B3")
    for r in range(4, 9): H[r].height = 24                    # 여백
    r = 9; H[r].height = 20
    c = ws.cell(row=r, column=2, value=kind or "업무 대장·서식"); c.font = XF(11, True, BLUE)
    H[10].height = 6
    H[11].height = 4; _fill(ws["B11"], BLUE)                   # 강조 막대
    H[12].height = 10
    ws.merge_cells("B13:I13")
    tl = math.ceil(disp_len(title) / 26) or 1
    c = ws["B13"]; c.value = title; c.font = XF(26 if disp_len(title) <= 30 else 22, True, NAVY)
    c.alignment = Alignment(wrap_text=True, vertical="top"); H[13].height = 40 * min(tl, 3)
    r = 14
    if subtitle:
        ws.merge_cells("B14:I14"); c = ws["B14"]; c.value = subtitle; c.font = XF(12, False, GRAY)
        c.alignment = Alignment(wrap_text=True, vertical="top"); H[14].height = 22 * min(2, math.ceil(disp_len(subtitle) / 48) or 1)
    for r in range(15, 19): H[r].height = 22
    # 문서 정보 줄
    meta = _meta(team, doc_no, date, version)
    cols = [(2, 3), (4, 5), (6, 7), (8, 9)]
    for (a, b), (k, v) in zip(cols, meta):
        if b > a:
            ws.merge_cells(start_row=19, start_column=a, end_row=19, end_column=b)
            ws.merge_cells(start_row=20, start_column=a, end_row=20, end_column=b)
        ws.cell(row=19, column=a, value=k).font = XF(8, False, GRAY)
        ws.cell(row=20, column=a, value=v).font = XF(10.5, True, INK)
        for cc in range(a, b + 1):
            ws.cell(row=19, column=cc).border = Border(top=Side(style="medium", color=NAVY))
        ws.cell(row=19, column=a).alignment = Alignment(vertical="bottom")
        ws.cell(row=20, column=a).alignment = Alignment(vertical="center")
    H[19].height = 20; H[20].height = 22
    r = 21; H[r].height = 14
    r = 22
    if how_to:
        how_to = list(how_to) + ["행이 더 필요하면 표의 마지막 행을 복사해 바로 아래에 붙여 넣습니다(수식·목록 선택·서식이 그대로 이어집니다)."]
        ws.cell(row=r, column=2, value="사용 방법").font = XF(11, True, NAVY); H[r].height = 20; r += 1
        for line in how_to:
            ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=9)
            c = ws.cell(row=r, column=2, value=f"•  {line}"); c.font = XF(10, False, INK)
            c.alignment = Alignment(wrap_text=True, vertical="top", indent=0)
            H[r].height = 16 * max(1, math.ceil(disp_len(line) / 52)) + 4; r += 1
        H[r].height = 12; r += 1
    # 결재방(우측) + 검토 문구(좌측)
    a0 = r
    ws.merge_cells(start_row=a0, start_column=6, end_row=a0 + 2, end_column=6)
    c = ws.cell(row=a0, column=6, value="결\n재"); c.font = XF(9, True, "FFFFFF"); _fill(c, NAVY)
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    names = _names(drafter, leader, approver)
    for j, k in enumerate(APPROVAL):
        col = 7 + j
        h = ws.cell(row=a0, column=col, value=k); h.font = XF(9, True, NAVY); _fill(h, SKY)
        h.alignment = Alignment(horizontal="center", vertical="center")
        s = ws.cell(row=a0 + 1, column=col, value=names[j]); s.font = XF(7.5, False, GRAY)
        s.alignment = Alignment(horizontal="center", vertical="bottom")
        d = ws.cell(row=a0 + 2, column=col, value="    .    ."); d.font = XF(7.5, False, GRAY)
        d.alignment = Alignment(horizontal="center", vertical="center")
    for rr in range(a0, a0 + 3):
        for col in range(6, 10):
            ws.cell(row=rr, column=col).border = Border(left=_LN, right=_LN, top=_LN, bottom=_LN)
    H[a0].height = 18; H[a0 + 1].height = 44; H[a0 + 2].height = 16
    ws.merge_cells(start_row=a0, start_column=2, end_row=a0 + 2, end_column=5)
    c = ws.cell(row=a0, column=2, value=_k.REVIEW_NOTE); c.font = XF(7.5, False, GRAY)
    c.alignment = Alignment(wrap_text=True, vertical="bottom")
    r = a0 + 3; H[r].height = 14; r += 1
    # 하단 네이비 띠
    H[r].height = 4; _fill_row(r, BLUE); r += 1
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=9)
    c = ws.cell(row=r, column=2, value=f"{COMPANY}\n{ADDR}  ·  {CONTACT}")
    c.font = XF(8.5, False, "FFFFFF"); c.alignment = Alignment(wrap_text=True, vertical="center")
    _fill_row(r, DEEP); H[r].height = 34
    ws.print_area = f"A1:J{r}"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4; ws.page_setup.orientation = "portrait"
    ws.sheet_properties.pageSetUpPr.fitToPage = True; ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 1
    ws.page_margins = PageMargins(left=0.4, right=0.4, top=0.4, bottom=0.4, header=0, footer=0)
    ws.print_options.horizontalCentered = True
    return wb


# ───────────────────────────── PowerPoint ─────────────────────────────
from pptx import Presentation
from pptx.dml.color import RGBColor as PRGB
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt as PPt


def _cell_txt(cell, text, size, bold=False, color=INK, fill="FFFFFF", anchor=MSO_ANCHOR.MIDDLE, line=LINE):
    cell.text = ""; tf = cell.text_frame; p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run(); r.text = text; r.font.size = PPt(size); r.font.bold = bold; r.font.name = FONT
    r.font.color.rgb = PRGB.from_string(color); cell.vertical_anchor = anchor
    cell.margin_left = cell.margin_right = Inches(0.03); cell.margin_top = cell.margin_bottom = Inches(0.02)
    cell.fill.solid(); cell.fill.fore_color.rgb = PRGB.from_string(fill)
    from lxml import etree
    from pptx.oxml.ns import qn as _qn
    tcPr = cell._tc.get_or_add_tcPr()
    for tag in ("a:lnL", "a:lnR", "a:lnT", "a:lnB"):
        ln = etree.SubElement(tcPr, _qn(tag), w="9525")
        sf = etree.SubElement(ln, _qn("a:solidFill")); etree.SubElement(sf, _qn("a:srgbClr"), val=line)


def pptx_new(title, subtitle="", team="", doc_no="", wide=True, drafter="", leader="", approver="대표이사", date=None,
             kind="", version="제정 v1.0"):
    prs = Presentation()
    if wide:
        prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    s = prs.slides.add_slide(prs.slide_layouts[6]); W, H = prs.slide_width, prs.slide_height
    bg = _k._asset("cover_bg_slide.png")
    if wide and os.path.exists(bg):
        s.shapes.add_picture(bg, 0, 0, width=W, height=H)
    else:
        _k._rect(s, 0, H - Inches(0.57), W, Inches(0.57), DEEP)
        _k._rect(s, 0, H - Inches(0.63), W, Inches(0.06), BLUE)
        _k._rect(s, 0, 0, W, Inches(0.05), BLUE)
    s.shapes.add_picture(LOGO, Inches(0.7), Inches(0.55), height=Inches(0.55))
    lb = s.shapes.add_textbox(Inches(0.7), Inches(1.95), Inches(8), Inches(0.4)).text_frame
    _k._ptxt(lb, kind or "업무 자료", 14, True, BLUE)
    _k._rect(s, Inches(0.72), Inches(2.45), Inches(0.75), Inches(0.07), BLUE)
    tb = s.shapes.add_textbox(Inches(0.7), Inches(2.65), Inches(9.2), Inches(2.0)).text_frame; tb.word_wrap = True
    _k._ptxt(tb, title, 40 if disp_len(title) <= 22 else 32, True, NAVY, space_after=6)
    if subtitle:
        _k._ptxt(tb, subtitle, 18, False, GRAY)
    # 문서 정보 줄
    meta = _meta(team, doc_no, date, version)
    x0, y0, cw = Inches(0.7), Inches(5.35), Inches(1.95)
    for j, (k, v) in enumerate(meta):
        x = x0 + cw * j
        ln = s.shapes.add_connector(1, x, y0, x + cw - Inches(0.15), y0)
        ln.line.color.rgb = PRGB.from_string(NAVY); ln.line.width = PPt(1.5)
        bx = s.shapes.add_textbox(x, y0 + Inches(0.05), cw - Inches(0.15), Inches(0.7)).text_frame; bx.word_wrap = True
        _k._ptxt(bx, k, 9, False, GRAY); _k._ptxt(bx, v, 12, True, INK)
    # 결재방(우하단)
    names = _names(drafter, leader, approver)
    tw = Inches(4.0); tbl = s.shapes.add_table(3, 4, W - tw - Inches(0.7), Inches(5.0), tw, Inches(1.3)).table
    tbl.columns[0].width = Inches(0.4)
    for j in range(1, 4): tbl.columns[j].width = Inches(1.2)
    tbl.rows[0].height = Inches(0.3); tbl.rows[1].height = Inches(0.75); tbl.rows[2].height = Inches(0.25)
    m = tbl.cell(0, 0); m.merge(tbl.cell(2, 0)); _cell_txt(m, "결재", 10, True, "FFFFFF", NAVY)
    for j, k in enumerate(APPROVAL, 1):
        _cell_txt(tbl.cell(0, j), k, 10, True, NAVY, SKY)
        _cell_txt(tbl.cell(1, j), names[j - 1], 8, False, GRAY, anchor=MSO_ANCHOR.BOTTOM)
        _cell_txt(tbl.cell(2, j), "    .    .", 8, False, GRAY)
    prs._samwon_title = title
    return prs
