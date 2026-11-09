"""삼원종합물류(주) 문서 키트 v3 (samwon-standard) — 본문 서식 엔진.

표지는 samwon_cover.py가 담당한다. 반드시 `from samwon_cover import *`로 쓴다.
서식 기준은 '2611총무_문서서식표준규정'(대표 지시 2026-11-09)을 따른다.

  ▸ 글꼴 맑은 고딕 단일 · 본문 10.5pt/줄간격 1.3/양쪽 정렬 · 표 9.5pt
  ▸ 제목은 다음 내용과 같은 쪽에 붙고(keep with next), 표 행은 쪽 사이에서 쪼개지지 않으며(cantSplit),
    15행 이하 표는 통째로 한 쪽에 놓인다. 긴 표는 머리글 행이 매 쪽 반복된다.
  ▸ 표 칸 너비는 내용 길이로 자동 배분하고, 낱말이 중간에서 잘리지 않도록 최소 너비를 보장한다.
  ▸ 엑셀은 빈 기입 행을 '인쇄 쪽 끝'까지만 채워 빈 종이가 나오지 않게 한다.

사용 예)
    from samwon_cover import *
    doc = docx_new("구매요청서", "구매·지출 승인 절차", team="총무팀", doc_no="SW-GA-2611-01", drafter="총무 대리", leader="총무팀장")
    h1(doc, "1. 목적"); para(doc, "본 서식은 ..."); table(doc, [["항목","내용"],["품명",""]])
    docx_save(doc, "out/2611총무_구매요청서.docx")
"""
import datetime
import math
import os
import re

KIT = os.path.dirname(os.path.abspath(__file__))


def _asset(name):
    for p in (os.path.join(KIT, name), os.path.join(KIT, "..", "assets", name)):
        if os.path.exists(p):
            return p
    return os.path.join(KIT, name)


# 로고는 영문(SAMWON)만 사용(국문 로고 금지)
LOGO = _asset("logo_EN.png")
LOGO_EN = LOGO
LOGO_WHITE = _asset("logo_EN_white.png")


def _ratio(path=None):
    from PIL import Image
    im = Image.open(path or LOGO)
    return im.width / im.height


COMPANY = "삼원종합물류(주)"
COMPANY_EN = "SAMWON LOGISTICS CO., LTD."
ADDR = "서울특별시 서초구 효령로 328 아트리트21 6층 (06721)"
CONTACT = "TEL 02-598-3001 · FAX 02-597-2564 · www.3won.kr"
# 브랜드 3색 — 영문 로고(SAMWON) 원본 색 (대표 지시 2026-11-09)
BLUE = "00489C"     # 삼원 블루: 강조·구분선·번호
SYMBOL = "08448C"   # 심볼 블루: 보조
NAVY = "282B6E"     # 첨탑 네이비: 제목·표 머리
DEEP = "141850"     # 진한 네이비: 어두운 바탕
SKY, PALE, LINE, GRAY, INK = "EAF0FA", "F5F8FD", "D5DEEE", "6E7891", "0E1530"
WARN = "D63A40"     # 경고 빨강: 위험·미준수 표시에만
FONT = "맑은 고딕"
TODAY = datetime.date.today().strftime("%Y. %m. %d.")
REVIEW_NOTE = ("본 문서는 노무·법무·세무·지식재산(저작권) 관점의 전문가 검토 기준으로 작성·점검되었습니다. "
               "법령 해석이 필요한 조항은 시행 전 회사 자문 노무사·변호사·세무사의 최종 확인을 받으십시오.")

# ───────────────────────── 서식 기준값 (규정 제2장) ─────────────────────────
# level 0 = 표준, 1·2 = 쪽 낭비를 줄이기 위한 단계적 압축(글자 크기는 줄이지 않고 간격만 줄인다)
_LEVELS = {
    0: dict(body=10.5, table=9.5, line=1.30, after=4, h1_before=16, h1_after=6, h2_before=10, h2_after=3, cell_v=0.06),
    1: dict(body=10.5, table=9.5, line=1.22, after=3, h1_before=12, h1_after=5, h2_before=8, h2_after=2, cell_v=0.04),
    2: dict(body=10.5, table=9.0, line=1.15, after=2, h1_before=10, h1_after=4, h2_before=6, h2_after=2, cell_v=0.03),
}
S = dict(_LEVELS[0])


def compact(level=0):
    """간격 압축 단계 설정(0 표준·1·2). samwon_md가 쪽 낭비를 발견하면 자동으로 올린다."""
    S.clear(); S.update(_LEVELS[max(0, min(2, int(level)))])
    return S


# ───────────────────────────── 공통 유틸 ─────────────────────────────
_WIDE = re.compile(r"[ᄀ-ᇿ　-鿿가-힣豈-﫿＀-￯]")


def disp_len(s):
    """표시 폭(한글·전각=1, 영문·숫자=0.55, 그 밖=0.8)"""
    n = 0.0
    for ch in str(s):
        if _WIDE.match(ch):
            n += 1
        elif ch.isascii():
            n += 0.55 if not ch.isspace() else 0.3
        else:
            n += 0.8
    return n


def _longest_token(s):
    toks = re.split(r"[\s/]+", str(s).replace("**", ""))
    return max((disp_len(t) for t in toks if t), default=0)


_NUM = re.compile(r"^[\s₩$+\-]*[\d,]+(\.\d+)?\s*(원|%|건|명|대|개|일|시간|회|점|km|kg|톤)?$")
_DATE = re.compile(r"^(20\d\d|\d{2})[.\-/]\s*\d{1,2}([.\-/]\s*\d{1,2})?\.?$|^\d{1,2}/\d{1,2}$|^\d{1,2}:\d{2}$")
_CENTER_HEAD = re.compile(r"^(No\.?|번호|순번|연번|NO|구분|상태|여부|등급|판정|확인|체크|선택|서명|인|결과|단위|수량|차수|요일)$")
_RIGHT_HEAD = re.compile(r"(금액|원\)|비용|단가|합계|잔액|급여|수당|\(원\)|율\(?%?\)?$|점수$)")


def col_kinds(rows, header=True):
    """열별 정렬 판단: 'center' / 'right' / 'left'"""
    if not rows:
        return []
    ncol = max(len(r) for r in rows)
    heads = rows[0] if header else [""] * ncol
    body = rows[1:] if header else rows
    kinds = []
    for j in range(ncol):
        h = str(heads[j] if j < len(heads) else "").replace("**", "").strip()
        vals = [str(r[j]).replace("**", "").strip() for r in body if j < len(r) and str(r[j]).strip()]
        if _CENTER_HEAD.match(h):
            kinds.append("center"); continue
        if vals and all(_NUM.match(v) for v in vals) and any(c.isdigit() for v in vals for c in v):
            kinds.append("right" if any(len(v) > 3 or "," in v for v in vals) or _RIGHT_HEAD.search(h) else "center"); continue
        if _RIGHT_HEAD.search(h) and (not vals or all(_NUM.match(v) or v in ("-", "") for v in vals)):
            kinds.append("right"); continue
        if vals and all(_DATE.match(v) or v in ("-", "□", "☐", "○", "●", "O", "X", "✓") for v in vals):
            kinds.append("center"); continue
        if vals and max(disp_len(v) for v in vals) <= 4 and disp_len(h) <= 6:
            kinds.append("center"); continue
        kinds.append("left")
    return kinds


def auto_widths(rows, usable_cm, size=9.5, pad_cm=0.32, max_chars=34):
    """내용 길이로 칸 너비(cm) 배분. 낱말이 잘리지 않을 최소 너비를 먼저 보장하고 남는 폭을 긴 칸에 나눈다."""
    ncol = max(len(r) for r in rows)
    ch = size * 0.0353 * 1.02           # 한글 1자 폭(cm)
    need_min, want = [], []
    for j in range(ncol):
        cells = [str(r[j]) if j < len(r) else "" for r in rows]
        tok = max((_longest_token(c) for c in cells), default=1)
        tok = max(tok, min(disp_len(cells[0].replace("**", "")) if cells else 0, 6))  # 머리글은 최대 6자까지 한 줄
        full = max((disp_len(c.replace("**", "")) for c in cells), default=1)
        need_min.append(max(1.0, min(tok, 14) * ch + pad_cm))
        want.append(max(need_min[-1], min(full, max_chars) * ch + pad_cm))
    if sum(need_min) >= usable_cm:
        k = usable_cm / sum(need_min)
        return [round(w * k, 2) for w in need_min]
    if sum(want) <= usable_cm:
        extra = usable_cm - sum(want)
        tot = sum(want)
        return [round(w + extra * w / tot, 2) for w in want]
    spare = usable_cm - sum(need_min)
    grow = [w - m for w, m in zip(want, need_min)]
    g = sum(grow) or 1
    return [round(m + spare * gr / g, 2) for m, gr in zip(need_min, grow)]


# ───────────────────────────── Word ─────────────────────────────
from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_TAB_ALIGNMENT, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Emu, Pt, RGBColor

_ALIGN = {"left": WD_ALIGN_PARAGRAPH.LEFT, "center": WD_ALIGN_PARAGRAPH.CENTER,
          "right": WD_ALIGN_PARAGRAPH.RIGHT, "justify": WD_ALIGN_PARAGRAPH.JUSTIFY}


def _rgb(h):
    return RGBColor.from_string(h)


def _font(run, size=None, bold=None, color=None):
    run.font.name = FONT
    rpr = run._element.get_or_add_rPr()
    rf = rpr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts"); rpr.insert(0, rf)
    for k in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rf.set(qn(k), FONT)
    if size: run.font.size = Pt(size)
    if bold is not None: run.font.bold = bold
    if color: run.font.color.rgb = _rgb(color)
    return run


def _shade(cell, fill):
    tcPr = cell._element.get_or_add_tcPr()
    for old in tcPr.findall(qn("w:shd")):
        tcPr.remove(old)
    shd = OxmlElement("w:shd"); shd.set(qn("w:val"), "clear"); shd.set(qn("w:color"), "auto"); shd.set(qn("w:fill"), fill)
    tcPr.append(shd)


def _set_tbl_prop(table, tag, attrs=None, children=None):
    tblPr = table._element.tblPr
    for old in tblPr.findall(qn(tag)):
        tblPr.remove(old)
    el = OxmlElement(tag)
    for k, v in (attrs or {}).items():
        el.set(qn(k), str(v))
    for c in children or []:
        el.append(c)
    tblPr.append(el)
    return el


def _border_el(tag, val="single", sz=4, color=LINE):
    x = OxmlElement(tag); x.set(qn("w:val"), val)
    if val != "nil":
        x.set(qn("w:sz"), str(sz)); x.set(qn("w:space"), "0"); x.set(qn("w:color"), color)
    return x


def _borders(table, color=LINE, sz=4, outer=None, outer_sz=8, vertical=True):
    """표 테두리: 안쪽 옅은 선, 위·아래 바깥선은 outer 색(기본 네이비)."""
    outer = outer or color
    kids = [_border_el("w:top", sz=outer_sz if outer != color else sz, color=outer),
            _border_el("w:left", "single" if vertical else "nil", sz, color),
            _border_el("w:bottom", sz=outer_sz if outer != color else sz, color=outer),
            _border_el("w:right", "single" if vertical else "nil", sz, color),
            _border_el("w:insideH", sz=sz, color=color),
            _border_el("w:insideV", "single" if vertical else "nil", sz, color)]
    _set_tbl_prop(table, "w:tblBorders", children=kids)


def _no_borders(table):
    _set_tbl_prop(table, "w:tblBorders", children=[_border_el(f"w:{e}", "nil") for e in
                                                  ("top", "left", "bottom", "right", "insideH", "insideV")])


def _cell_border(cell, **sides):
    """cell 단위 테두리. 예) _cell_border(c, bottom=("single", 12, NAVY), left=("nil",))"""
    tcPr = cell._element.get_or_add_tcPr()
    tb = tcPr.find(qn("w:tcBorders"))
    if tb is None:
        tb = OxmlElement("w:tcBorders"); tcPr.append(tb)
    for side, spec in sides.items():
        for old in tb.findall(qn(f"w:{side}")):
            tb.remove(old)
        tb.append(_border_el(f"w:{side}", *spec))


def _cell_margins(table, lr_cm=0.16, tb_cm=0.06):
    kids = []
    for side, v in (("top", tb_cm), ("left", lr_cm), ("bottom", tb_cm), ("right", lr_cm)):
        x = OxmlElement(f"w:{side}"); x.set(qn("w:w"), str(int(v * 567))); x.set(qn("w:type"), "dxa"); kids.append(x)
    _set_tbl_prop(table, "w:tblCellMar", children=kids)


def _fixed_widths(table, widths_cm):
    """고정 레이아웃 + gridCol + 각 셀 너비(Word·LibreOffice 모두 같은 너비로 보이게)."""
    _set_tbl_prop(table, "w:tblLayout", {"w:type": "fixed"})
    _set_tbl_prop(table, "w:tblW", {"w:w": int(sum(widths_cm) * 567), "w:type": "dxa"})
    grid = table._element.tblGrid
    gcs = grid.findall(qn("w:gridCol"))
    for gc, w in zip(gcs, widths_cm):
        gc.set(qn("w:w"), str(int(w * 567)))
    for row in table.rows:
        for j, c in enumerate(row.cells):
            if j < len(widths_cm):
                c.width = Cm(widths_cm[j])


def _row_props(row, cant_split=True, header=False, height_cm=None, rule="atLeast"):
    trPr = row._tr.get_or_add_trPr()
    if cant_split and trPr.find(qn("w:cantSplit")) is None:
        trPr.append(OxmlElement("w:cantSplit"))
    if header and trPr.find(qn("w:tblHeader")) is None:
        trPr.append(OxmlElement("w:tblHeader"))
    if height_cm:
        for old in trPr.findall(qn("w:trHeight")):
            trPr.remove(old)
        h = OxmlElement("w:trHeight"); h.set(qn("w:val"), str(int(height_cm * 567))); h.set(qn("w:hRule"), rule)
        trPr.append(h)


def _row_h(row, cm, rule="exact"):
    _row_props(row, cant_split=False, height_cm=cm, rule=rule)


def _pfmt(p, before=0, after=None, line=None, keep_next=False, align=None, keep_lines=True):
    f = p.paragraph_format
    f.space_before = Pt(before)
    f.space_after = Pt(S["after"] if after is None else after)
    f.line_spacing = S["line"] if line is None else line
    f.widow_control = True
    f.keep_together = keep_lines
    if keep_next:
        f.keep_with_next = True
    if align:
        p.alignment = _ALIGN[align]
    return p


def _bottom_rule(p, color=BLUE, sz=8, space=2):
    pPr = p._p.get_or_add_pPr(); bd = OxmlElement("w:pBdr")
    bt = _border_el("w:bottom", "single", sz, color); bt.set(qn("w:space"), str(space)); bd.append(bt); pPr.append(bd)


def _rich(p, text, size, bold=False, color=None):
    """**굵게** 표기 지원"""
    for i, s in enumerate(str(text).split("**")):
        if s:
            _font(p.add_run(s), size, bold or (i % 2 == 1), color)
    return p


def _field(par, code, size=8, color=GRAY):
    r = par.add_run(); _font(r, size, color=color)
    f1 = OxmlElement("w:fldChar"); f1.set(qn("w:fldCharType"), "begin")
    it = OxmlElement("w:instrText"); it.set(qn("xml:space"), "preserve"); it.text = code
    f2 = OxmlElement("w:fldChar"); f2.set(qn("w:fldCharType"), "end")
    r._element.append(f1); r._element.append(it); r._element.append(f2)


def usable_cm(doc):
    sec = doc.sections[-1]
    return round((sec.page_width - sec.left_margin - sec.right_margin) / 360000, 2)


def setup_styles(doc):
    st = doc.styles["Normal"]; st.font.name = FONT; st.font.size = Pt(S["body"])
    rpr = st.element.get_or_add_rPr(); rf = rpr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts"); rpr.insert(0, rf)
    for k in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rf.set(qn(k), FONT)
    st.font.color.rgb = _rgb(INK)
    pf = st.paragraph_format; pf.space_after = Pt(S["after"]); pf.line_spacing = S["line"]; pf.widow_control = True
    return doc


def page_setup(sec, landscape=False):
    """A4 · 여백 상 2.2 / 하 2.0 / 좌우 2.0cm · 머리글 1.0 / 바닥글 0.9cm"""
    if landscape:
        sec.orientation = WD_ORIENT.LANDSCAPE; sec.page_width, sec.page_height = Cm(29.7), Cm(21.0)
    else:
        sec.orientation = WD_ORIENT.PORTRAIT; sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
    sec.left_margin = sec.right_margin = Cm(2.0); sec.top_margin = Cm(2.2); sec.bottom_margin = Cm(2.0)
    sec.header_distance = Cm(1.0); sec.footer_distance = Cm(0.9)
    return sec


def body_header_footer(sec, title, doc_no="", restart=False):
    """본문 머리글(로고 ··· 문서 제목, 아래 옅은 선)·바닥글(회사·문서번호 ··· 쪽/전체)."""
    sec.different_first_page_header_footer = False
    sec.header.is_linked_to_previous = False; sec.footer.is_linked_to_previous = False
    W = round((sec.page_width - sec.left_margin - sec.right_margin) / 360000, 2)
    # 머리글: 로고(좌) ··· 문서 제목(우), 아래 옅은 선
    hp = sec.header.paragraphs[0]; hp.paragraph_format.space_after = Pt(0)
    hp.paragraph_format.tab_stops.add_tab_stop(Cm(W), WD_TAB_ALIGNMENT.RIGHT)
    hp.add_run().add_picture(LOGO, height=Cm(0.42))
    _font(hp.add_run("\t" + title), 8, color=GRAY)
    _bottom_rule(hp, LINE, 4, 4)
    # 바닥글: 회사·문서번호(좌) ··· 쪽(우)
    fp = sec.footer.paragraphs[0]; fp.paragraph_format.space_after = Pt(0)
    fp.paragraph_format.tab_stops.add_tab_stop(Cm(W), WD_TAB_ALIGNMENT.RIGHT)
    _font(fp.add_run(COMPANY + (f"  ·  {doc_no}" if doc_no else "")), 8, color=GRAY)
    _font(fp.add_run("\t"), 8, color=GRAY)
    _font(fp.add_run("- "), 8, color=GRAY); _field(fp, "PAGE", 8, NAVY); _font(fp.add_run(" -"), 8, color=GRAY)
    if restart:
        pg = OxmlElement("w:pgNumType"); pg.set(qn("w:start"), "1"); sec._sectPr.append(pg)
    return sec


def setup_document(doc, title, doc_no="", landscape=False):
    """(표지 없는 문서용) 글꼴·용지·머리글/바닥글"""
    setup_styles(doc); page_setup(doc.sections[0], landscape)
    body_header_footer(doc.sections[0], title, doc_no)
    return doc


def docx_new(title, subtitle="", team="", doc_no="", version="제정 v1.0", landscape=False, kind="서식", **_):
    """표지 없는 기본 문서(표지 표준판은 samwon_cover.docx_new)."""
    doc = Document(); setup_document(doc, title, doc_no, landscape)
    return doc


def h1(doc, text):
    p = doc.add_paragraph(); _pfmt(p, S["h1_before"], S["h1_after"], 1.15, keep_next=True)
    _font(p.add_run(str(text).replace("**", "")), 14, True, NAVY)
    _bottom_rule(p, BLUE, 8, 3)
    return p


def h2(doc, text):
    p = doc.add_paragraph(); _pfmt(p, S["h2_before"], S["h2_after"], 1.15, keep_next=True)
    _font(p.add_run(str(text).replace("**", "")), 12, True, BLUE)
    return p


def h3(doc, text):
    p = doc.add_paragraph(); _pfmt(p, 6, 2, 1.15, keep_next=True)
    _font(p.add_run(str(text).replace("**", "")), 10.5, True, NAVY)
    return p


def para(doc, text, size=None, bold=False, color=None, align="justify", indent=0, hanging=0, keep_next=False, after=None):
    """본문 문단(양쪽 정렬). **굵게** 지원. hanging: 내어쓰기(cm)"""
    p = doc.add_paragraph(); _pfmt(p, 0, after, keep_next=keep_next, align=align or "justify")
    if indent or hanging:
        p.paragraph_format.left_indent = Cm(indent + hanging)
        if hanging:
            p.paragraph_format.first_line_indent = Cm(-hanging)
    _rich(p, text, size or S["body"], bold, color)
    return p


def bullets(doc, items, mark="•", indent=0.2, size=None):
    for it in items:
        p = doc.add_paragraph(); _pfmt(p, 0, max(1, S["after"] - 2), align="justify")
        p.paragraph_format.left_indent = Cm(indent + 0.45); p.paragraph_format.first_line_indent = Cm(-0.45)
        _font(p.add_run(f"{mark} "), size or S["body"], False, BLUE)
        _rich(p, it, size or S["body"])


def numbered(doc, text, size=None):
    """'1. ', '①', '가.' 같은 번호 문단 — 번호 뒤 내어쓰기"""
    m = re.match(r"^(\(?\d{1,2}[.)]|[①-⑳]|[가-하]\.|[a-z][.)])\s*", text)
    pre = m.group(1) if m else ""
    rest = text[m.end():] if m else text
    p = doc.add_paragraph(); _pfmt(p, 0, max(1, S["after"] - 1), align="justify")
    hang = 0.25 + 0.21 * disp_len(pre)
    p.paragraph_format.left_indent = Cm(0.2 + hang); p.paragraph_format.first_line_indent = Cm(-hang)
    if pre:
        _font(p.add_run(pre + " "), size or S["body"], True, NAVY)
    _rich(p, rest, size or S["body"])
    return p


def checklist(doc, items, size=None):
    for it in items:
        p = doc.add_paragraph(); _pfmt(p, 0, max(1, S["after"] - 2), align="justify")
        p.paragraph_format.left_indent = Cm(0.75); p.paragraph_format.first_line_indent = Cm(-0.55)
        _font(p.add_run("☐ "), size or S["body"], False, NAVY)
        _rich(p, it, size or S["body"])


def spacer(doc, pt=6):
    """표 뒤 등 짧은 간격(빈 문단 대신 사용 — 쪽 낭비 방지)"""
    p = doc.add_paragraph(); f = p.paragraph_format
    f.space_before = Pt(0); f.space_after = Pt(0); f.line_spacing = Pt(pt); f.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    r = p.add_run(""); _font(r, 2)
    return p


def _cell_par(c, text, size, bold=False, color=None, align="left", keep_next=False):
    pp = c.paragraphs[0]
    _pfmt(pp, 0, 0, 1.15, keep_next=keep_next, align=align)
    _rich(pp, text, size, bold, color)
    return pp


def table(doc, rows, widths=None, header=True, size=None, align_center_cols=(), fill_rows=0, keep=None,
          align=None, zebra=True, first_col_bold=False):
    """표준 표.
    rows: 2차원 리스트(첫 행 머리글) · widths: cm 리스트(없으면 자동) · fill_rows: 빈 기입 행 수
    keep: True면 표 전체를 한 쪽에(기본: 15행 이하 자동) · align: 열별 'left|center|right' 리스트(없으면 자동)"""
    size = size or S["table"]
    rows = [["" if x is None else str(x) for x in r] for r in rows]
    ncol = max(len(r) for r in rows)
    rows = [r + [""] * (ncol - len(r)) for r in rows]
    kinds = list(align) if align else col_kinds(rows, header)
    for j in align_center_cols:
        if j < ncol: kinds[j] = "center"
    W = usable_cm(doc)
    if widths:
        widths = [float(w) for w in widths][:ncol] + [2.0] * max(0, ncol - len(widths))
        if sum(widths) > W + 0.05 or sum(widths) < W * 0.6:
            widths = [round(w * W / sum(widths), 2) for w in widths]
    else:
        widths = auto_widths(rows, W, size)
    total = rows + [[""] * ncol for _ in range(fill_rows)]
    t = doc.add_table(rows=len(total), cols=ncol); t.alignment = WD_TABLE_ALIGNMENT.CENTER
    _borders(t, LINE, 4, outer=NAVY, outer_sz=10)
    _cell_margins(t, 0.16, S["cell_v"])
    _fixed_widths(t, widths)
    keep = (len(total) <= 15) if keep is None else keep
    for i, r in enumerate(total):
        row = t.rows[i]
        is_head = header and i == 0
        _row_props(row, True, is_head, 0.62 if i >= len(rows) else (0.7 if is_head else None))
        kn = keep and i < len(total) - 1
        for j in range(ncol):
            c = row.cells[j]; c.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            if is_head:
                _shade(c, NAVY); _cell_par(c, r[j].replace("**", ""), size, True, "FFFFFF", "center", kn)
            else:
                if zebra and i % 2 == 0 and i < len(rows):
                    _shade(c, PALE)
                b = first_col_bold and j == 0
                _cell_par(c, r[j], size, b, NAVY if b else None, kinds[j] if j < len(kinds) else "left", kn)
    spacer(doc, 8)
    return t


def form_table(doc, pairs, cols=2, widths=None, size=None):
    """서식용 '항목 | 기입란' 표. pairs: [(항목, 기본값)]. cols=2면 한 행에 2쌍."""
    size = size or S["table"]
    n = (len(pairs) + cols - 1) // cols
    W = usable_cm(doc)
    lab = max(2.4, min(4.2, max(disp_len(k) for k, _ in pairs) * size * 0.036 + 0.45))
    if widths:
        lab, val = widths[0], widths[1]
    val = (W - lab * cols) / cols
    t = doc.add_table(rows=n, cols=cols * 2); t.alignment = WD_TABLE_ALIGNMENT.CENTER
    _borders(t, LINE, 4, outer=NAVY, outer_sz=10); _cell_margins(t, 0.16, S["cell_v"])
    _fixed_widths(t, [lab, val] * cols)
    keep = n <= 15
    for idx in range(n * cols):
        r, c = divmod(idx, cols)
        k, v = pairs[idx] if idx < len(pairs) else ("", "")
        kc, vc = t.rows[r].cells[c * 2], t.rows[r].cells[c * 2 + 1]
        kn = keep and r < n - 1
        _shade(kc, SKY); _cell_par(kc, k, size, True, NAVY, "left", kn)
        _cell_par(vc, v or "", size, False, None, "left", kn)
        for cc in (kc, vc): cc.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    for row in t.rows:
        _row_props(row, True, False, 0.78)
    spacer(doc, 8)
    return t


def note_box(doc, lines, size=None, color=NAVY):
    """안내 상자(왼쪽 블루 세로선 + 옅은 바탕). 한 쪽에 통째로 놓인다."""
    size = size or (S["body"] - 0.5)
    t = doc.add_table(rows=1, cols=1); t.alignment = WD_TABLE_ALIGNMENT.CENTER
    _no_borders(t); _cell_margins(t, 0.3, 0.14); _fixed_widths(t, [usable_cm(doc)])
    c = t.rows[0].cells[0]; _shade(c, SKY)
    _cell_border(c, left=("single", 24, BLUE))
    _row_props(t.rows[0], True)
    for n, line in enumerate(lines):
        pp = c.paragraphs[0] if n == 0 else c.add_paragraph()
        _pfmt(pp, 0, 2, 1.25, keep_next=n < len(lines) - 1, align="justify")
        _rich(pp, line, size, False, color)
    spacer(doc, 8)
    return t


def sign_block(doc, lines=("작성일:      년     월     일", "작성자:                    (서명 또는 인)")):
    for i, l in enumerate(lines):
        para(doc, l, align="right", keep_next=i < len(lines) - 1)


def page_break(doc):
    p = doc.add_paragraph(); _pfmt(p, 0, 0, 1.0)
    p.add_run().add_break(WD_BREAK.PAGE)
    return p


def docx_save(doc, path):
    os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
    # 마지막 문단이 빈 간격 문단이면 지워 빈 쪽이 생기지 않게 한다
    body = doc.element.body
    paras = [e for e in body if e.tag == qn("w:p")]
    while paras and paras[-1].getnext() is not None and paras[-1].getnext().tag == qn("w:sectPr") \
            and not "".join(paras[-1].itertext()).strip() and not paras[-1].findall(".//" + qn("w:drawing")) \
            and not paras[-1].findall(".//" + qn("w:br")):
        prev = paras[-1].getprevious()
        if prev is None or prev.tag != qn("w:p"):
            break   # 표 바로 뒤 문단은 Word 규칙상 남겨 둔다(최소 높이로)
        body.remove(paras[-1]); paras.pop()
    doc.save(path)
    return path


# ───────────────────────────── Excel ─────────────────────────────
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.page import PageMargins

XF = lambda size=10, bold=False, color=INK: Font(name=FONT, size=size, bold=bold, color=color)
THIN = Side(style="thin", color=LINE)
HEAD_LINE = Side(style="medium", color=NAVY)
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
_CM_PER_CHAR = 0.19     # 엑셀 열 너비 1 = 약 0.19cm(맑은 고딕 10pt)
_ROW_PT = 20            # 기본 행 높이(pt)


def _xl_kind(h, vals):
    k = col_kinds([[h]] + [[v] for v in vals], True)
    return k[0] if k else "left"


def _xl_lines(text, width_chars):
    if text in (None, ""):
        return 1
    n = 0
    for part in str(text).split("\n"):
        n += max(1, math.ceil(disp_len(part) * 1.75 / max(4, width_chars - 1)))
    return n


def xl_page(ws, landscape=None, total_width_chars=None, title=""):
    """인쇄 설정: A4, 가로 1쪽 맞춤(세로는 자동), 여백, 머리글 반복은 sheet()가 지정."""
    if landscape is None:
        landscape = (total_width_chars or 0) > 92
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = "landscape" if landscape else "portrait"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0
    ws.page_margins = PageMargins(left=0.47, right=0.47, top=0.6, bottom=0.6, header=0.25, footer=0.25)
    ws.print_options.horizontalCentered = True
    ws.oddFooter.left.text = f"{COMPANY}" + (f" · {title}" if title else "")
    ws.oddFooter.left.size = 8; ws.oddFooter.left.font = FONT
    ws.oddFooter.right.text = "&P / &N"; ws.oddFooter.right.size = 8; ws.oddFooter.right.font = FONT
    return landscape


def rows_per_page_fill(row_heights_pt, first_page_extra_pt, header_pt, total_width_chars, landscape):
    """현재 행 높이 목록으로 마지막 쪽에 남는 높이(pt)를 계산 — 빈 기입 행을 '쪽 끝까지만' 채우기 위함."""
    page_w_cm, page_h_cm = (29.7, 21.0) if landscape else (21.0, 29.7)
    avail_w = page_w_cm - 2.4
    avail_h_pt = (page_h_cm - 3.05 - 0.4) / 0.03528
    scale = min(1.0, avail_w / max(1.0, total_width_chars * _CM_PER_CHAR))
    cap = avail_h_pt / scale          # 실제 행 높이 기준 한 쪽 용량(pt)
    used = first_page_extra_pt + header_pt
    for h in row_heights_pt:
        if used + h > cap:
            used = header_pt + h      # 다음 쪽(머리글 반복)
        else:
            used += h
    return cap - used, cap


def sheet(wb, name, headers, widths=None, rows=None, title=None, note=None, blank_rows=None, landscape=None, min_blank=8):
    """표준 표 시트(제목·안내·머리글·자동 정렬·틀 고정·필터·인쇄 설정).
    blank_rows: None이면 데이터가 끝난 쪽의 끝까지만 빈 기입 행을 채운다(최소 min_blank행).
                숫자를 주면 '최소 행 수'로 보고, 그 행이 끝나는 쪽 끝까지 채운다."""
    ws = wb.create_sheet(name); ws.sheet_view.showGridLines = False
    ncol = len(headers)
    data = [list(r) for r in (rows or [])]
    if not widths:
        widths = []
        for j, h in enumerate(headers):
            vals = [str(r[j]) for r in data if j < len(r) and r[j] not in (None, "")]
            longest = max([disp_len(h) * 1.9 + 2] + [min(disp_len(v), 30) * 1.75 + 2 for v in vals])
            widths.append(round(max(6, min(46, longest)), 1))
    widths = [float(w) for w in widths]
    for j, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(j)].width = w
    tw = sum(widths)
    land = xl_page(ws, landscape, tw, title or name)
    last_col = get_column_letter(ncol)
    hr, extra = 1, 0.0
    if title:
        ws.merge_cells(f"A1:{last_col}1")
        c = ws["A1"]; c.value = f"{title}"; c.font = XF(14, True, NAVY)
        c.alignment = Alignment(vertical="center", horizontal="left")
        ws.row_dimensions[1].height = 28; extra += 28 + 4
        for j in range(1, ncol + 1):
            ws.cell(row=1, column=j).border = Border(bottom=Side(style="medium", color=BLUE))
        hr = 2
        if note:
            ws.merge_cells(f"A2:{last_col}2")
            c = ws["A2"]; c.value = note; c.font = XF(9, False, GRAY)
            c.alignment = Alignment(wrap_text=True, vertical="center")
            ws.row_dimensions[2].height = max(18, 13 * _xl_lines(note, tw)); extra += ws.row_dimensions[2].height
            hr = 3
        else:
            ws.row_dimensions[1].height = 32                       # 제목 아래 여백은 행 높이로 확보
    elif note:
        ws.cell(row=1, column=1, value=note).font = XF(9, False, GRAY); hr = 2; extra += 18
    kinds = []
    for j, h in enumerate(headers):
        vals = [str(r[j]) for r in data if j < len(r) and r[j] not in (None, "")]
        kinds.append(_xl_kind(h, vals))
    hlines = max(_xl_lines(h, widths[j]) for j, h in enumerate(headers))
    gap = {j for j, h in enumerate(headers, 1) if h in (None, "", "None")}
    for j in gap:
        ws.column_dimensions[get_column_letter(j)].width = min(widths[j - 1], 3)
    for j, h in enumerate(headers, 1):
        if j in gap:
            continue
        c = ws.cell(row=hr, column=j, value=h); c.font = XF(10, True, "FFFFFF"); c.fill = PatternFill("solid", fgColor=NAVY)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = Border(left=Side(style="thin", color=NAVY), right=Side(style="thin", color=NAVY), top=HEAD_LINE, bottom=HEAD_LINE)
    head_pt = max(24, 15 * hlines + 8); ws.row_dimensions[hr].height = head_pt
    heights = []
    for i, vals in enumerate(data):
        lines = max([_xl_lines(vals[j] if j < len(vals) else "", widths[j]) for j in range(ncol)] + [1])
        heights.append(max(_ROW_PT, 14.5 * lines + 6))
    # 빈 기입 행: 데이터가 끝난 쪽에 남은 자리까지만(새 쪽으로 넘기지 않는다). 데이터가 없으면 첫 쪽을 채운다.
    want = 10 if blank_rows is None else int(blank_rows)
    rem, cap = rows_per_page_fill(heights, extra, head_pt, tw, land)
    room = max(0, int(rem * 0.88 // _ROW_PT))
    fill = room if not data else min(want, room)
    total = len(data) + fill
    zfill = PatternFill("solid", fgColor=PALE)
    for i in range(total):
        r = hr + 1 + i; vals = data[i] if i < len(data) else []
        for j in range(1, ncol + 1):
            v = vals[j - 1] if j - 1 < len(vals) else None
            if j in gap:
                if v not in (None, ""):
                    ws.cell(row=r, column=j, value=v)
                continue
            c = ws.cell(row=r, column=j, value=v)
            c.font = XF(10); c.border = BOX
            c.alignment = Alignment(vertical="center", wrap_text=True, horizontal=kinds[j - 1])
            if i % 2 == 1:
                c.fill = zfill
        ws.row_dimensions[r].height = heights[i] if i < len(heights) else _ROW_PT
    last = hr + total
    for j in range(1, ncol + 1):   # 표 마지막 선
        if j in gap:
            continue
        c = ws.cell(row=last, column=j); c.border = Border(left=THIN, right=THIN, top=THIN, bottom=Side(style="medium", color=NAVY))
    ws.freeze_panes = ws.cell(row=hr + 1, column=1)
    ws.auto_filter.ref = f"A{hr}:{last_col}{last}"
    ws.print_title_rows = f"{hr}:{hr}"
    ws._samwon_header_row = hr
    ws._samwon_last_row = last
    ws._samwon_kinds = kinds
    return ws


def dropdown(ws, col_letter, options, first_row=None, last_row=None):
    """목록 선택(데이터 유효성) 추가"""
    fr = first_row or (getattr(ws, "_samwon_header_row", 1) + 1)
    lr = last_row or getattr(ws, "_samwon_last_row", 500)
    dv = DataValidation(type="list", formula1='"' + ",".join(options) + '"', allow_blank=True)
    ws.add_data_validation(dv); dv.add(f"{col_letter}{fr}:{col_letter}{lr}")


def xlsx_save(wb, path):
    os.makedirs(os.path.dirname(path) or ".", exist_ok=True); wb.active = 0
    for ws in wb.worksheets:
        ws.sheet_view.zoomScale = 100
    wb.save(path); return path


# ─────────────────────────── PowerPoint ───────────────────────────
from pptx import Presentation
from pptx.dml.color import RGBColor as PRGB
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Emu as PEmu, Inches, Pt as PPt


def _ptxt(tf, text, size, bold=False, color=INK, align=None, space_after=None):
    p = tf.paragraphs[0] if not tf.paragraphs[0].text and len(tf.paragraphs) == 1 else tf.add_paragraph()
    r = p.add_run(); r.text = text; r.font.size = PPt(size); r.font.bold = bold; r.font.name = FONT
    r.font.color.rgb = PRGB.from_string(color)
    p.alignment = {"center": PP_ALIGN.CENTER, "right": PP_ALIGN.RIGHT}.get(align, PP_ALIGN.LEFT)
    if space_after is not None: p.space_after = PPt(space_after)
    return p


def _rect(slide, x, y, w, h, color):
    s = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    s.fill.solid(); s.fill.fore_color.rgb = PRGB.from_string(color); s.line.fill.background(); s.shadow.inherit = False
    return s


def pptx_new(title, subtitle="", team="", doc_no="", wide=True, **_):
    """표지 없는 기본 발표 파일(표지 표준판은 samwon_cover.pptx_new)."""
    prs = Presentation()
    if wide:
        prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    prs._samwon_title = title
    return prs


def slide_title_body(prs, title, lines, note=None):
    """본문 슬라이드: 좌상단 블루 막대 + 제목 26pt + 본문 18pt, 우상단 로고, 하단 쪽 표시"""
    s = prs.slides.add_slide(prs.slide_layouts[6]); W, H = prs.slide_width, prs.slide_height
    _rect(s, Inches(0.6), Inches(0.55), Inches(0.09), Inches(0.5), BLUE)
    s.shapes.add_picture(LOGO, W - Inches(2.0), Inches(0.45), height=Inches(0.36))
    tb = s.shapes.add_textbox(Inches(0.85), Inches(0.42), W - Inches(3.2), Inches(0.8)).text_frame
    tb.word_wrap = True; tb.vertical_anchor = MSO_ANCHOR.MIDDLE
    _ptxt(tb, title, 26, True, NAVY)
    ln = s.shapes.add_connector(1, Inches(0.6), Inches(1.32), W - Inches(0.6), Inches(1.32))
    ln.line.color.rgb = PRGB.from_string(LINE); ln.line.width = PPt(1)
    bb = s.shapes.add_textbox(Inches(0.85), Inches(1.6), W - Inches(1.7), H - Inches(2.5)).text_frame; bb.word_wrap = True
    for l in lines:
        _ptxt(bb, l, 18, space_after=8)
    nb = s.shapes.add_textbox(Inches(0.6), H - Inches(0.55), W - Inches(1.2), Inches(0.35)).text_frame
    _ptxt(nb, (note + "    ") if note else "", 10, False, GRAY)
    _ptxt(nb, f"{COMPANY}  ·  {len(prs.slides)}", 9, False, GRAY, align="right")
    return s


def slide_table(prs, title, rows, lines=None, note=None):
    """표 슬라이드: 제목 + (요약 1~2줄) + 네이비 머리글 표"""
    s = slide_title_body(prs, title, lines or [], note)
    W, H = prs.slide_width, prs.slide_height
    top = Inches(1.6 + 0.45 * len(lines or []))
    ncol = max(len(r) for r in rows); rows = [r + [""] * (ncol - len(r)) for r in rows]
    h = min(H - top - Inches(0.9), Inches(0.42) * len(rows))
    tbl = s.shapes.add_table(len(rows), ncol, Inches(0.85), top, W - Inches(1.7), h).table
    widths = auto_widths(rows, (W - Inches(1.7)) / 914400 * 2.54, 12)
    for j, w in enumerate(widths):
        tbl.columns[j].width = int(w / 2.54 * 914400)
    kinds = col_kinds(rows)
    for i, r in enumerate(rows):
        for j, v in enumerate(r):
            c = tbl.cell(i, j); c.text = ""; p = c.text_frame.paragraphs[0]
            run = p.add_run(); run.text = v.replace("**", ""); run.font.name = FONT
            run.font.size = PPt(12 if i else 12); run.font.bold = i == 0
            run.font.color.rgb = PRGB.from_string("FFFFFF" if i == 0 else INK)
            p.alignment = PP_ALIGN.CENTER if i == 0 else {"center": PP_ALIGN.CENTER, "right": PP_ALIGN.RIGHT}.get(kinds[j], PP_ALIGN.LEFT)
            c.fill.solid(); c.fill.fore_color.rgb = PRGB.from_string(NAVY if i == 0 else (PALE if i % 2 == 0 else "FFFFFF"))
            c.vertical_anchor = MSO_ANCHOR.MIDDLE
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
