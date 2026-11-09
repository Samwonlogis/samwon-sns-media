"""삼원종합물류(주) 문서 키트 v2 (samwon-standard) — 본문 서식 함수. 표지는 samwon_cover.py가 결재방 포함 표준판으로 덮어쓴다.
※ 반드시 `from samwon_cover import *`로 쓴다(이 모듈의 docx_new 등은 구판 표지).

사용 예)
    from samwon_kit import *
    doc = docx_new("구매요청서", "구매·지출 승인 절차", team="총무팀", doc_no="SW-GA-2610-01")
    h1(doc, "1. 목적"); para(doc, "본 서식은 ..."); table(doc, [["항목","내용"],["품명",""]], widths=[4,12])
    docx_save(doc, "out/2610총무_구매요청서.docx")

    wb = xlsx_new("차고지 이용현황 대장", team="총무팀")
    ws = sheet(wb, "대장", ["No","차량번호","입고일"], widths=[6,14,12])
    xlsx_save(wb, "out/2610총무_차고지이용현황대장.xlsx")

    prs = pptx_new("화주 모집 카드뉴스", team="홍보팀")
    slide_title_body(prs, "제목", ["본문 1", "본문 2"]); prs.save("...pptx")
"""
import datetime
import os

KIT = os.path.dirname(os.path.abspath(__file__))
# 로고는 영문(SAMWON)만 사용(국문 로고 금지). 같은 폴더 → ../assets 순서로 찾는다.
LOGO = next((p for p in (os.path.join(KIT, "logo_EN.png"), os.path.join(KIT, "..", "assets", "logo_EN.png")) if os.path.exists(p)),
            os.path.join(KIT, "logo_EN.png"))
LOGO_EN = LOGO


def _ratio():
    from PIL import Image
    im = Image.open(LOGO)
    return im.width / im.height
COMPANY = "삼원종합물류(주)"
COMPANY_EN = "SAMWON LOGISTICS CO., LTD."
ADDR = "서울특별시 서초구 효령로 328 아트리트21 6층 (06721)"
CONTACT = "TEL 02-598-3001 · FAX 02-597-2564 · www.3won.kr"
# 브랜드 3색 — 영문 로고(SAMWON) 원본 색 (대표 지시 2026-11-09)
BLUE = "00489C"     # 삼원 블루: 워드마크 — 강조·구분선·번호
SYMBOL = "08448C"   # 심볼 블루: 물결 심볼 — 보조
NAVY = "282B6E"     # 첨탑 네이비: 제목·표 머리
DEEP = "141850"     # 진한 네이비: 어두운 바탕
SKY, GRAY, INK = "EAF0FA", "6E7891", "0E1530"
WARN = "D63A40"     # 경고 빨강: 위험·미준수 표시에만
FONT = "맑은 고딕"
TODAY = datetime.date.today().strftime("%Y. %m. %d.")
REVIEW_NOTE = ("본 문서는 노무·법무·세무·지식재산(저작권) 관점의 전문가 검토 기준으로 작성·점검되었습니다. "
               "법령 해석이 필요한 조항은 시행 전 회사 자문 노무사·변호사·세무사의 최종 확인을 받으십시오.")

# ───────────────────────────── Word ─────────────────────────────
from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor


def _rgb(h):
    return RGBColor.from_string(h)


def _font(run, size=None, bold=None, color=None):
    run.font.name = FONT
    rpr = run._element.get_or_add_rPr()
    rf = rpr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts"); rpr.append(rf)
    for k in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rf.set(qn(k), FONT)
    if size: run.font.size = Pt(size)
    if bold is not None: run.font.bold = bold
    if color: run.font.color.rgb = _rgb(color)
    return run


def _shade(cell, fill):
    tcPr = cell._element.get_or_add_tcPr()
    shd = OxmlElement("w:shd"); shd.set(qn("w:val"), "clear"); shd.set(qn("w:color"), "auto"); shd.set(qn("w:fill"), fill)
    tcPr.append(shd)


def _borders(table, color="9AA6B8", sz=4):
    tbl = table._element; tblPr = tbl.tblPr
    b = OxmlElement("w:tblBorders")
    for e in ("top", "left", "bottom", "right", "insideH", "insideV"):
        x = OxmlElement(f"w:{e}"); x.set(qn("w:val"), "single"); x.set(qn("w:sz"), str(sz)); x.set(qn("w:color"), color); b.append(x)
    tblPr.append(b)


def _field(par, code):
    r = par.add_run(); _font(r, 8, color=GRAY)
    f1 = OxmlElement("w:fldChar"); f1.set(qn("w:fldCharType"), "begin")
    it = OxmlElement("w:instrText"); it.set(qn("xml:space"), "preserve"); it.text = code
    f2 = OxmlElement("w:fldChar"); f2.set(qn("w:fldCharType"), "end")
    r._element.append(f1); r._element.append(it); r._element.append(f2)


def docx_new(title, subtitle="", team="", doc_no="", version="제정 v1.0", landscape=False, kind="서식"):
    """표지 1쪽 + 머리글(로고·제목)/바닥글(쪽 번호)이 있는 문서. kind: 서식/규정/매뉴얼/대본 등 표지 상단 라벨."""
    doc = Document()
    st = doc.styles["Normal"]; st.font.name = FONT; st.font.size = Pt(10.5)
    st.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    st.paragraph_format.space_after = Pt(4); st.paragraph_format.line_spacing = 1.35
    sec = doc.sections[0]
    if landscape:
        sec.orientation = WD_ORIENT.LANDSCAPE; sec.page_width, sec.page_height = Cm(29.7), Cm(21.0)
    else:
        sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
    sec.left_margin = sec.right_margin = Cm(2.0); sec.top_margin = Cm(2.0); sec.bottom_margin = Cm(1.8)
    sec.different_first_page_header_footer = True
    # 머리글/바닥글 (표지 제외)
    hp = sec.header.paragraphs[0]; hp.alignment = WD_ALIGN_PARAGRAPH.LEFT
    hp.add_run().add_picture(LOGO, height=Cm(0.55))
    _font(hp.add_run(f"    {title}"), 8, color=GRAY)
    fp = sec.footer.paragraphs[0]; fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _font(fp.add_run(f"{COMPANY}  ·  {doc_no + '  ·  ' if doc_no else ''}"), 8, color=GRAY)
    _field(fp, "PAGE"); _font(fp.add_run(" / "), 8, color=GRAY); _field(fp, "NUMPAGES")

    # ── 표지 ──
    p = doc.add_paragraph(); p.paragraph_format.space_after = Pt(0)
    p.add_run().add_picture(LOGO, height=Cm(1.6))
    for _ in range(5): doc.add_paragraph()
    p = doc.add_paragraph(); _font(p.add_run(kind.upper() if kind.isascii() else kind), 12, True, BLUE)
    p = doc.add_paragraph(); p.paragraph_format.space_after = Pt(6)
    _font(p.add_run(title), 30, True, NAVY)
    if subtitle:
        p = doc.add_paragraph(); _font(p.add_run(subtitle), 14, False, "374151")
    # 구분선
    p = doc.add_paragraph(); pPr = p._p.get_or_add_pPr(); bd = OxmlElement("w:pBdr"); bt = OxmlElement("w:bottom")
    bt.set(qn("w:val"), "single"); bt.set(qn("w:sz"), "24"); bt.set(qn("w:color"), BLUE); bd.append(bt); pPr.append(bd)
    for _ in range(3): doc.add_paragraph()
    info = [("법인명", COMPANY), ("문서번호", doc_no or "-"), ("주관 부서", team or "-"), ("제정·개정", f"{version}  ({TODAY})")]
    t = doc.add_table(rows=len(info), cols=2); t.alignment = WD_TABLE_ALIGNMENT.LEFT; _borders(t, "D1D9E6")
    for i, (k, v) in enumerate(info):
        c0, c1 = t.rows[i].cells; c0.width = Cm(3.5); c1.width = Cm(12.5); _shade(c0, SKY)
        _font(c0.paragraphs[0].add_run(k), 10, True, NAVY); _font(c1.paragraphs[0].add_run(v), 10)
    doc.add_paragraph()
    # 결재란
    t = doc.add_table(rows=2, cols=4); _borders(t, "D1D9E6"); t.alignment = WD_TABLE_ALIGNMENT.RIGHT
    for j, k in enumerate(["작성", "검토", "승인", "대표이사"]):
        _shade(t.rows[0].cells[j], SKY); _font(t.rows[0].cells[j].paragraphs[0].add_run(k), 9, True, NAVY)
        t.rows[0].cells[j].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        t.rows[1].cells[j].paragraphs[0].add_run("\n\n")
    for _ in range(2): doc.add_paragraph()
    p = doc.add_paragraph(); _font(p.add_run(REVIEW_NOTE), 8, color=GRAY)
    p = doc.add_paragraph(); _font(p.add_run(f"{COMPANY}  |  {ADDR}  |  {CONTACT}"), 8, color=GRAY)
    p = doc.add_paragraph(); p.add_run().add_break(WD_BREAK.PAGE)
    return doc


def h1(doc, text):
    p = doc.add_paragraph(); p.paragraph_format.space_before = Pt(14); p.paragraph_format.space_after = Pt(6)
    _font(p.add_run(text), 15, True, NAVY)
    pPr = p._p.get_or_add_pPr(); bd = OxmlElement("w:pBdr"); bt = OxmlElement("w:bottom")
    bt.set(qn("w:val"), "single"); bt.set(qn("w:sz"), "8"); bt.set(qn("w:color"), BLUE); bd.append(bt); pPr.append(bd)
    return p


def h2(doc, text):
    p = doc.add_paragraph(); p.paragraph_format.space_before = Pt(10); p.paragraph_format.space_after = Pt(3)
    _font(p.add_run(text), 12, True, BLUE); return p


def para(doc, text, size=10.5, bold=False, color=None, align=None, indent=0):
    """**굵게** 표기 지원"""
    p = doc.add_paragraph()
    if indent: p.paragraph_format.left_indent = Cm(indent)
    if align == "center": p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if align == "right": p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    parts = text.split("**")
    for i, s in enumerate(parts):
        if s: _font(p.add_run(s), size, bold or (i % 2 == 1), color)
    return p


def bullets(doc, items, mark="•", indent=0.5, size=10.5):
    for it in items:
        para(doc, f"{mark} {it}", size=size, indent=indent)


def checklist(doc, items, size=10.5):
    for it in items:
        para(doc, f"☐  {it}", size=size, indent=0.4)


def table(doc, rows, widths=None, header=True, size=9.5, align_center_cols=(), fill_rows=0):
    """rows: 2차원 리스트(첫 행 머리글). widths: cm 리스트. fill_rows: 빈 기입 행 추가 수"""
    rows = [list(map(lambda x: "" if x is None else str(x), r)) for r in rows]
    ncol = max(len(r) for r in rows)
    rows += [[""] * ncol for _ in range(fill_rows)]
    t = doc.add_table(rows=len(rows), cols=ncol); _borders(t)
    for i, r in enumerate(rows):
        for j in range(ncol):
            c = t.rows[i].cells[j]; v = r[j] if j < len(r) else ""
            if widths: c.width = Cm(widths[j])
            pp = c.paragraphs[0]; pp.paragraph_format.space_after = Pt(0)
            if header and i == 0:
                _shade(c, NAVY); _font(pp.add_run(v), size, True, "FFFFFF"); pp.alignment = WD_ALIGN_PARAGRAPH.CENTER
            else:
                segs = v.split("**")
                for k, s in enumerate(segs):
                    if s: _font(pp.add_run(s), size, k % 2 == 1)
                if j in align_center_cols: pp.alignment = WD_ALIGN_PARAGRAPH.CENTER
            if fill_rows and i >= len(rows) - fill_rows:
                tr = t.rows[i]._tr; trPr = tr.get_or_add_trPr(); h = OxmlElement("w:trHeight"); h.set(qn("w:val"), "420"); trPr.append(h)
    doc.add_paragraph()
    return t


def form_table(doc, pairs, cols=2, widths=None, size=9.5):
    """서식용 '항목 | 기입란' 표. pairs: [(항목, 기본값)]. cols=2면 한 행에 2쌍."""
    n = (len(pairs) + cols - 1) // cols
    t = doc.add_table(rows=n, cols=cols * 2); _borders(t)
    for idx, (k, v) in enumerate(pairs):
        r, c = divmod(idx, cols)
        kc, vc = t.rows[r].cells[c * 2], t.rows[r].cells[c * 2 + 1]
        _shade(kc, SKY); _font(kc.paragraphs[0].add_run(k), size, True, NAVY); _font(vc.paragraphs[0].add_run(v or ""), size)
        if widths:
            kc.width = Cm(widths[0]); vc.width = Cm(widths[1])
        tr = t.rows[r]._tr; trPr = tr.get_or_add_trPr(); h = OxmlElement("w:trHeight"); h.set(qn("w:val"), "460"); trPr.append(h)
    doc.add_paragraph(); return t


def sign_block(doc, lines=("작성일:      년     월     일", "작성자:                    (서명 또는 인)")):
    for l in lines:
        para(doc, l, align="right")


def page_break(doc):
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)


def docx_save(doc, path):
    os.makedirs(os.path.dirname(path) or ".", exist_ok=True); doc.save(path); return path


# ───────────────────────────── Excel ─────────────────────────────
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

XF = lambda size=10, bold=False, color="111827": Font(name=FONT, size=size, bold=bold, color=color)
THIN = Side(style="thin", color="B8C2D3")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def xlsx_new(title, subtitle="", team="", doc_no="", version="제정 v1.0", how_to=None):
    """첫 시트 '표지': 로고·법인명·제목·문서정보·사용 방법. how_to: 사용 방법 문장 리스트"""
    wb = Workbook(); ws = wb.active; ws.title = "표지"
    ws.sheet_view.showGridLines = False
    for col, w in zip("ABCDEFGH", (3, 16, 16, 16, 16, 16, 16, 3)):
        ws.column_dimensions[col].width = w
    img = XLImage(LOGO); img.height = 62; img.width = int(62 * _ratio()); ws.add_image(img, "B2")
    ws["B7"] = "삼원종합물류(주)"; ws["B7"].font = XF(12, True, BLUE)
    ws["B9"] = title; ws["B9"].font = XF(24, True, NAVY)
    if subtitle:
        ws["B11"] = subtitle; ws["B11"].font = XF(13, False, "374151")
    for c in range(2, 8):
        ws.cell(row=12, column=c).border = Border(bottom=Side(style="thick", color=BLUE))
    info = [("법인명", COMPANY), ("문서번호", doc_no or "-"), ("주관 부서", team or "-"), ("제정·개정", f"{version} ({TODAY})")]
    for i, (k, v) in enumerate(info):
        r = 14 + i
        ws.cell(row=r, column=2, value=k).font = XF(10, True, NAVY); ws.cell(row=r, column=2).fill = PatternFill("solid", fgColor=SKY)
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=7)
        ws.cell(row=r, column=3, value=v).font = XF(10)
        for c in range(2, 8): ws.cell(row=r, column=c).border = BOX
    r = 19
    if how_to:
        ws.cell(row=r, column=2, value="사용 방법").font = XF(11, True, BLUE); r += 1
        for line in how_to:
            ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
            c = ws.cell(row=r, column=2, value=f"• {line}"); c.font = XF(10); c.alignment = Alignment(wrap_text=True, vertical="top")
            ws.row_dimensions[r].height = 30 if len(line) > 60 else 18; r += 1
    r += 1
    ws.merge_cells(start_row=r, start_column=2, end_row=r + 1, end_column=7)
    c = ws.cell(row=r, column=2, value=REVIEW_NOTE); c.font = XF(8, False, GRAY); c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.cell(row=r + 3, column=2, value=f"{COMPANY} | {ADDR} | {CONTACT}").font = XF(8, False, GRAY)
    ws.page_setup.paperSize = ws.PAPERSIZE_A4; ws.print_options.horizontalCentered = True
    ws.page_setup.fitToWidth = 1; ws.sheet_properties.pageSetUpPr.fitToPage = True
    return wb


def sheet(wb, name, headers, widths=None, rows=None, title=None, note=None, blank_rows=30, landscape=True):
    """머리글 서식·틀 고정·인쇄 설정이 된 표 시트. title이 있으면 1행 제목, note가 있으면 2행 안내."""
    ws = wb.create_sheet(name); ws.sheet_view.showGridLines = False
    hr = 1
    if title:
        ws.cell(row=1, column=1, value=f"{COMPANY}  |  {title}").font = XF(13, True, NAVY); hr = 3 if note else 2
        ws.row_dimensions[1].height = 24
    if note:
        ws.cell(row=2, column=1, value=note).font = XF(9, False, GRAY)
    for j, h in enumerate(headers, 1):
        c = ws.cell(row=hr, column=j, value=h); c.font = XF(10, True, "FFFFFF"); c.fill = PatternFill("solid", fgColor=NAVY)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True); c.border = BOX
        ws.column_dimensions[get_column_letter(j)].width = (widths[j - 1] if widths else 14)
    ws.row_dimensions[hr].height = 32
    data = rows or []
    total = max(len(data), 0) + blank_rows
    for i in range(total):
        r = hr + 1 + i; vals = data[i] if i < len(data) else []
        for j in range(1, len(headers) + 1):
            c = ws.cell(row=r, column=j, value=(vals[j - 1] if j - 1 < len(vals) else None))
            c.font = XF(10); c.border = BOX; c.alignment = Alignment(vertical="center", wrap_text=True)
            if i % 2 == 1: c.fill = PatternFill("solid", fgColor="F7F9FC")
    ws.freeze_panes = ws.cell(row=hr + 1, column=1)
    ws.auto_filter.ref = f"A{hr}:{get_column_letter(len(headers))}{hr + total}"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = "landscape" if landscape else "portrait"
    ws.sheet_properties.pageSetUpPr.fitToPage = True; ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0
    ws.print_title_rows = f"{hr}:{hr}"
    ws.oddFooter.center.text = f"{COMPANY} · &P / &N"
    ws._samwon_header_row = hr
    return ws


def dropdown(ws, col_letter, options, first_row=None, last_row=500):
    """목록 선택(데이터 유효성) 추가"""
    fr = first_row or (getattr(ws, "_samwon_header_row", 1) + 1)
    dv = DataValidation(type="list", formula1='"' + ",".join(options) + '"', allow_blank=True)
    ws.add_data_validation(dv); dv.add(f"{col_letter}{fr}:{col_letter}{last_row}")


def xlsx_save(wb, path):
    os.makedirs(os.path.dirname(path) or ".", exist_ok=True); wb.active = 0; wb.save(path); return path


# ─────────────────────────── PowerPoint ───────────────────────────
from pptx import Presentation
from pptx.dml.color import RGBColor as PRGB
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Emu, Inches, Pt as PPt


def _ptxt(tf, text, size, bold=False, color="111827", align=None):
    from pptx.enum.text import PP_ALIGN
    p = tf.paragraphs[0] if not tf.paragraphs[0].text and len(tf.paragraphs) == 1 else tf.add_paragraph()
    r = p.add_run(); r.text = text; r.font.size = PPt(size); r.font.bold = bold; r.font.name = FONT
    r.font.color.rgb = PRGB.from_string(color)
    if align == "center": p.alignment = PP_ALIGN.CENTER
    return p


def pptx_new(title, subtitle="", team="", doc_no="", wide=True):
    prs = Presentation()
    if wide:
        prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    s = prs.slides.add_slide(prs.slide_layouts[6])
    W, H = prs.slide_width, prs.slide_height
    bar = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, H - Inches(0.35), W, Inches(0.35)); bar.fill.solid(); bar.fill.fore_color.rgb = PRGB.from_string(NAVY); bar.line.fill.background()
    s.shapes.add_picture(LOGO, Inches(0.8), Inches(0.7), height=Inches(0.75))
    tb = s.shapes.add_textbox(Inches(0.8), Inches(2.6), W - Inches(1.6), Inches(2.4)).text_frame; tb.word_wrap = True
    _ptxt(tb, COMPANY, 18, True, BLUE); _ptxt(tb, title, 40, True, NAVY)
    if subtitle: _ptxt(tb, subtitle, 18, False, "374151")
    ib = s.shapes.add_textbox(Inches(0.8), H - Inches(1.6), W - Inches(1.6), Inches(1.0)).text_frame
    _ptxt(ib, f"{team}  ·  {doc_no or ''}  ·  {TODAY}", 12, False, GRAY)
    _ptxt(ib, f"{ADDR}  |  {CONTACT}", 10, False, GRAY)
    return prs


def slide_title_body(prs, title, lines, note=None):
    s = prs.slides.add_slide(prs.slide_layouts[6]); W = prs.slide_width
    s.shapes.add_picture(LOGO, W - Inches(2.4), Inches(0.35), height=Inches(0.4))
    tb = s.shapes.add_textbox(Inches(0.7), Inches(0.5), W - Inches(3.4), Inches(1)).text_frame
    _ptxt(tb, title, 28, True, NAVY)
    bb = s.shapes.add_textbox(Inches(0.7), Inches(1.6), W - Inches(1.4), Inches(5)).text_frame; bb.word_wrap = True
    for l in lines: _ptxt(bb, l, 18)
    if note:
        nb = s.shapes.add_textbox(Inches(0.7), prs.slide_height - Inches(0.8), W - Inches(1.4), Inches(0.5)).text_frame
        _ptxt(nb, note, 10, False, GRAY)
    return s


def slide_image(prs, image_path, caption=None):
    """이미지(카드뉴스 등)를 가운데 배치한 슬라이드"""
    from PIL import Image
    s = prs.slides.add_slide(prs.slide_layouts[6]); W, H = prs.slide_width, prs.slide_height
    im = Image.open(image_path); ih = H - Inches(0.9); iw = int(ih * im.width / im.height)
    s.shapes.add_picture(image_path, int((W - iw) / 2), Inches(0.3), height=ih)
    if caption:
        tb = s.shapes.add_textbox(Inches(0.5), H - Inches(0.55), W - Inches(1), Inches(0.4)).text_frame
        _ptxt(tb, caption, 11, False, GRAY, align="center")
    return s
