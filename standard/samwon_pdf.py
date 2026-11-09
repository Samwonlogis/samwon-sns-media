#!/usr/bin/env python3
"""삼원종합물류(주) PDF 보고서 표준 (samwon-standard).

정부지원 분석·투자 브리핑·월간 보고·사내 공지 등 weasyprint로 만드는 모든 PDF에 쓴다.
브랜드 3색(영문 로고 원본 색)·한글 줄바꿈·머리글·결재 없는 보고서용 상단·검사 기능을 한곳에 모았다.

사용 예)
    from samwon_pdf import *
    body = header("국내 투자 브리핑", "추천 3건 · 시장 지표", date="2026.11.10 (화)")
    body += h2("이번 주 핵심 3줄", 1) + bullets(["…", "…", "…"])
    body += table(["사업", "대상", "판정"], [["전기화물차 보조금", "삼원종합물류", badge("추천")]])
    body += card("물류바우처", "<p>…</p>", links=[("공고 바로가기", url1), ("신청하기", url2)])
    render(body, "20261110_국내 투자 브리핑.pdf", footer="삼원종합물류(주) · 대표 보고용")

검사 출력: PAGES n / WRAP-CHECK n건 / LINKS n — WRAP이 0이 될 때까지 문장을 다듬어 다시 만든다.
"""
import base64
import html as _html
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))

# ── 브랜드 토큰 (영문 로고 SAMWON 원본 색, 대표 지시 2026-11-09) ──
C = {
    "blue": "#00489C",    # 삼원 블루 — 강조·구분선·링크·번호
    "symbol": "#08448C",  # 심볼 블루 — 보조
    "navy": "#282B6E",    # 첨탑 네이비 — 제목·표 머리
    "deep": "#141850",    # 진한 네이비 — 어두운 바탕
    "sky": "#EAF0FA", "paper": "#F5F8FD", "line": "#D5DEEE",
    "ink": "#0E1530", "gray": "#6E7891",
    "warn": "#D63A40",    # 위험·미충족·마감 임박에만
}

FONT = "'Noto Sans CJK KR', 'Noto Sans KR', 'Malgun Gothic', sans-serif"


def esc(s):
    return _html.escape(str(s if s is not None else ""))


def logo_uri(white=False):
    """영문 로고 SVG를 data URI로. 스킬 assets → 스크립트 폴더 → 공개 주소 순서."""
    name = "SAMWON_logo_EN_white.svg" if white else "SAMWON_logo_EN.svg"
    for p in (os.path.join(HERE, "..", "assets", name), os.path.join(HERE, name)):
        if os.path.exists(p):
            return "data:image/svg+xml;base64," + base64.b64encode(open(p, "rb").read()).decode()
    png = os.path.join(HERE, "logo_EN_white.png" if white else "logo_EN.png")
    if os.path.exists(png):
        return "data:image/png;base64," + base64.b64encode(open(png, "rb").read()).decode()
    return "https://raw.githubusercontent.com/jong1231/samwon-website/main/public/samwon_logo.svg"


BASE_CSS = f"""
@page {{ size: A4; margin: 15mm 15mm 16mm 15mm;
  @bottom-left {{ content: string(foot); font: 7.5pt {FONT}; color: {C['gray']}; }}
  @bottom-right {{ content: counter(page) " / " counter(pages); font: 7.5pt {FONT}; color: {C['gray']}; }} }}
* {{ box-sizing: border-box; }}
body {{ font-family: {FONT}; color: {C['ink']}; font-size: 10.5pt; line-height: 1.6; margin: 0;
       word-break: keep-all; overflow-wrap: break-word; }}
.foot {{ string-set: foot content(); display: none; }}
a {{ color: {C['blue']}; text-decoration: none; font-weight: 700; }}
.head {{ display: flex; align-items: flex-end; justify-content: space-between;
        border-bottom: 3px solid {C['blue']}; padding-bottom: 8px; margin-bottom: 10px; }}
.head img {{ height: 30px; }}
.head .r {{ text-align: right; }}
.head .t {{ font-size: 21pt; font-weight: 900; color: {C['navy']}; letter-spacing: -0.4px; line-height: 1.15; }}
.head .s {{ font-size: 9.5pt; color: {C['gray']}; margin-top: 2px; }}
h2 {{ font-size: 13.5pt; color: {C['navy']}; margin: 14px 0 6px; padding-bottom: 4px; border-bottom: 1px solid {C['line']};
     break-after: avoid; }}
h2 .n {{ display: inline-block; background: {C['blue']}; color: #fff; border-radius: 3px; font-size: 8.5pt;
        padding: 0 6px; margin-right: 6px; vertical-align: 2px; }}
h3 {{ font-size: 11.5pt; color: {C['blue']}; margin: 10px 0 4px; break-after: avoid; }}
ul.b {{ margin: 4px 0 6px 0; padding-left: 16px; }} ul.b li {{ margin: 2px 0; }}
table.t {{ width: 100%; border-collapse: collapse; table-layout: fixed; font-size: 9.6pt; line-height: 1.45; margin: 4px 0 8px; }}
.t th {{ background: {C['navy']}; color: #fff; font-weight: 700; text-align: left; padding: 5px 6px; }}
.t td {{ border-bottom: 1px solid {C['line']}; padding: 5px 6px; vertical-align: top; }}
.t tr:nth-child(even) td {{ background: {C['paper']}; }}
.t tr.grp td {{ background: {C['sky']}; color: {C['navy']}; font-weight: 900; }}
.nw {{ white-space: nowrap; }}
.badge {{ display: inline-block; font-size: 8.3pt; font-weight: 700; border-radius: 10px; padding: 0 8px; white-space: nowrap; }}
.b-ok {{ background: {C['blue']}; color: #fff; }}
.b-mid {{ background: #fff; color: {C['blue']}; border: 1px solid {C['blue']}; }}
.b-no {{ background: #E9ECF2; color: {C['gray']}; }}
.b-warn {{ background: #fff; color: {C['warn']}; border: 1px solid {C['warn']}; }}
.card {{ border: 1px solid {C['line']}; border-left: 4px solid {C['blue']}; border-radius: 6px; padding: 8px 12px;
        margin: 6px 0 8px; break-inside: avoid; }}
.card .ct {{ font-weight: 900; color: {C['navy']}; font-size: 11pt; }}
.card .links {{ margin-top: 4px; font-size: 9.6pt; }} .card .links a {{ margin-right: 12px; }}
.tip {{ background: {C['paper']}; border-left: 3px solid {C['blue']}; padding: 5px 10px; font-size: 9.6pt; margin: 6px 0; }}
.tip b {{ color: {C['blue']}; margin-right: 4px; }}
.kpis {{ display: flex; gap: 8px; margin: 6px 0 8px; }}
.kpi {{ flex: 1; border: 1px solid {C['line']}; border-radius: 6px; padding: 6px 10px; }}
.kpi .l {{ font-size: 8.4pt; color: {C['gray']}; }} .kpi .v {{ font-size: 14pt; font-weight: 900; color: {C['navy']}; }}
.src {{ font-size: 8.2pt; color: {C['gray']}; }}
.notice {{ margin-top: 12px; font-size: 8.2pt; color: {C['gray']}; border-top: 1px solid {C['line']}; padding-top: 5px; }}
"""

BADGE = {"추천": "b-ok", "충족": "b-ok", "정상": "b-ok", "조건부": "b-mid", "확인 필요": "b-mid",
         "해당 없음": "b-no", "보류": "b-no", "미충족": "b-warn", "초과": "b-warn", "마감 임박": "b-warn", "위험": "b-warn"}


def badge(text):
    return f'<span class="badge {BADGE.get(text, "b-mid")}">{esc(text)}</span>'


def header(title, subtitle="", date=""):
    sub = " · ".join(x for x in (date, subtitle) if x)
    return (f'<div class="head"><img src="{logo_uri()}"><div class="r"><div class="t">{esc(title)}</div>'
            f'<div class="s">{esc(sub)}</div></div></div>')


def h2(text, n=None):
    num = f'<span class="n">{n:02d}</span>' if isinstance(n, int) else ""
    return f"<h2>{num}{esc(text)}</h2>"


def bullets(items):
    return '<ul class="b">' + "".join(f"<li>{x}</li>" for x in items) + "</ul>"


def table(headers, rows, widths=None, nowrap_cols=()):
    """rows의 셀은 HTML 허용(badge·링크). widths: % 리스트."""
    cg = "".join(f'<col style="width:{w}%">' for w in widths) if widths else ""
    th = "".join(f"<th>{esc(h)}</th>" for h in headers)
    body = ""
    for r in rows:
        if isinstance(r, dict) and r.get("group"):
            body += f'<tr class="grp"><td colspan="{len(headers)}">{esc(r["group"])}</td></tr>'
            continue
        body += "<tr>" + "".join(f'<td class="{"nw" if j in nowrap_cols else ""}">{c}</td>' for j, c in enumerate(r)) + "</tr>"
    return f'<table class="t"><colgroup>{cg}</colgroup><thead><tr>{th}</tr></thead><tbody>{body}</tbody></table>'


def card(title, inner_html, links=()):
    lk = "".join(f'<a href="{esc(u)}">{esc(t)} ▸</a>' for t, u in links if u)
    return f'<div class="card"><div class="ct">{esc(title)}</div>{inner_html}' + (f'<div class="links">{lk}</div>' if lk else "") + "</div>"


def tip(text, label="노하우"):
    return f'<div class="tip"><b>{esc(label)}</b>{esc(text)}</div>'


def kpis(items):
    return '<div class="kpis">' + "".join(f'<div class="kpi"><div class="l">{esc(l)}</div><div class="v">{esc(v)}</div></div>' for l, v in items) + "</div>"


def check_wraps(document, short=5):
    """몇 글자 때문에 줄이 넘어간 곳(마지막 줄이 short자 이하)을 WRAP 경고로 출력한다."""
    from weasyprint.formatting_structure import boxes as B

    def text_of(box):
        if isinstance(box, B.TextBox):
            return box.text
        return "".join(text_of(ch) for ch in getattr(box, "children", []))
    found = []

    def walk(box):
        lines = [ch for ch in getattr(box, "children", []) if isinstance(ch, B.LineBox)]
        if len(lines) >= 2:
            last = text_of(lines[-1]).strip()
            if 0 < len(last.replace(" ", "")) <= short:
                found.append((text_of(lines[-2]).strip()[-20:], last))
        for ch in getattr(box, "children", []):
            if not isinstance(ch, B.LineBox):
                walk(ch)
    for p in document.pages:
        walk(p._page_box)
    for prev, last in found:
        print(f"WRAP: …{prev} ⏎ {last}")
    print(f"WRAP-CHECK: {len(found)}건")
    return len(found)


def count_links(pdf_path):
    try:
        from pypdf import PdfReader
        n = 0
        for pg in PdfReader(pdf_path).pages:
            for a in pg.get("/Annots") or []:
                o = a.get_object()
                if o.get("/A", {}).get("/URI"):
                    n += 1
        return n
    except Exception:
        return -1


def render(body_html, out_path, footer="삼원종합물류(주)", extra_css="", max_pages=None):
    from weasyprint import HTML
    doc = (f'<!doctype html><html lang="ko"><head><meta charset="utf-8"><style>{BASE_CSS}{extra_css}</style></head>'
           f'<body><div class="foot">{esc(footer)}</div>{body_html}</body></html>')
    d = HTML(string=doc, base_url=HERE).render()
    d.write_pdf(out_path)
    print(f"PAGES: {len(d.pages)}" + (f" (기준 {max_pages}장 초과)" if max_pages and len(d.pages) > max_pages else ""))
    check_wraps(d)
    print(f"LINKS: {count_links(out_path)}")
    return out_path


if __name__ == "__main__":  # 샘플 출력: python3 samwon_pdf.py sample.pdf
    out = sys.argv[1] if len(sys.argv) > 1 else "sample.pdf"
    b = header("보고서 표준 샘플", "samwon-standard", date="2026.11.09 (월)")
    b += h2("핵심 3줄", 1) + bullets(["첫째 요지입니다.", "둘째 요지입니다.", "셋째 요지입니다."])
    b += kpis([("경유 평균", "1,612원"), ("두바이유", "$64.2"), ("신규 공고", "12건"), ("마감 임박", "3건")])
    b += h2("판정 표", 2) + table(["사업", "대상 법인", "지원 내용", "판정"],
                                 [{"group": "이천"}, ["전기화물차 보급", "삼원종합물류(주)", "대당 최대 4,000만 원", badge("추천")],
                                  ["노후 경유차 조기폐차", "삼원운수(주)", "공고 확인 필요", badge("조건부")],
                                  ["스마트물류센터", "해당 법인 없음", "창고 미보유", badge("해당 없음")]],
                                 widths=[30, 22, 34, 14], nowrap_cols=(3,))
    b += card("물류바우처 지원사업", "<p>수출 중소기업 물류비 지원. 신청 자격·필요 서류는 공고 원문 기준.</p>",
              links=[("공고 바로가기", "https://www.bizinfo.go.kr"), ("신청하기", "https://www.bizinfo.go.kr")])
    b += tip("공고 마감 2주 전까지 공통 서류를 미리 발급해 둔다.")
    render(b, out, footer="삼원종합물류(주) · 샘플")
