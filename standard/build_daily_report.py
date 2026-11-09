#!/usr/bin/env python3
"""삼원의 오늘 — 운송·물류 일간 브리핑 PDF (samwon-standard, 브랜드 3색판 2026-11-09).
레이아웃·검사(BID-PAGE·WRAP-CHECK)는 2026-11-09 대표 승인판 그대로, 색만 영문 로고 원본 색으로 통일.
사용: python3 build_daily_report.py content.json logo.svg "YYYYMMDD_삼원의 오늘.pdf"
"""
import json, sys, html, base64, mimetypes
from weasyprint import HTML

def esc(s):
    return html.escape(str(s or ""))

def bid_rows(bids):
    return "".join(
        f'<tr class="bid"><td class="bt"><a href="{esc(b.get("url",""))}">{esc(b["title"])}</a>'
        f'<div class="bn">{esc(" · ".join(x for x in [b.get("no",""), b.get("note","")] if x))}</div></td>'
        f'<td class="og">{esc(b.get("org",""))}</td>'
        f'<td class="tm">{esc(b.get("term",""))}</td>'
        f'<td class="dl">{esc(b.get("deadline",""))}</td><td class="am">{esc(b.get("amount",""))}</td></tr>'
        for b in bids)

def build(content, logo_path, out_path):
    c = content
    mime = mimetypes.guess_type(logo_path)[0] or "image/png"
    if logo_path.endswith(".svg"): mime = "image/svg+xml"
    logo = base64.b64encode(open(logo_path, "rb").read()).decode()
    logo_uri = f"data:{mime};base64,{logo}"

    kpis = "".join(
        f'<div class="kpi"><div class="kl">{esc(k["label"])}</div>'
        f'<div class="kv">{esc(k["value"])}</div>'
        f'<div class="kd {esc(k.get("tone",""))}">{esc(k.get("delta",""))}</div></div>'
        for k in c.get("kpis", []))

    groups = c.get("bid_groups") or ([{"name": "", "bids": c["bids"]}] if c.get("bids") else [])
    groups = [g for g in groups if g.get("bids")]
    if groups:
        body = ""
        for g in groups:
            if g.get("name"):
                extra = f' <em>{esc(g["note"])}</em>' if g.get("note") else ""
                body += f'<tr class="grp"><td colspan="5">■ {esc(g["name"])} <span>{len(g["bids"])}건</span>{extra}</td></tr>'
            body += bid_rows(g["bids"])
        bids = (f'<table class="bids"><colgroup><col style="width:43%"><col style="width:17%"><col style="width:14%"><col style="width:13%"><col style="width:13%"></colgroup>'
                f'<thead><tr><th>공고명 · 공고번호</th><th>발주기관</th><th>계약기간</th><th>마감</th><th>금액</th></tr></thead><tbody>{body}</tbody></table>')
    else:
        bids = ""
    if c.get("bid_note"):
        bids += f'<div class="bidnote">{esc(c["bid_note"])}</div>'

    issues = "".join(
        f'<div class="issue"><div class="ih"><span class="tag">{esc(i.get("tag",""))}</span>'
        f'<span class="it">{esc(i["title"])}</span></div>'
        f'<div class="ib">{esc(i["body"])}</div>'
        + (f'<div class="ip"><b>업무 체크포인트</b> {esc(i["point"])}</div>' if i.get("point") else "")
        + f'<div class="is">출처: <a href="{esc(i.get("url",""))}">{esc(i.get("source",""))}</a></div></div>'
        for i in c.get("issues", []))

    q = c.get("quote", {})
    doc = f"""<!doctype html><html lang="ko"><head><meta charset="utf-8"><style>
@page {{ size: A4; margin: 14mm 14mm 16mm 14mm;
  @bottom-left {{ content: "삼원종합물류㈜ · 사내 배포용"; font: 7.5pt 'Noto Sans CJK KR'; color:#8D96AD; }}
  @bottom-right {{ content: counter(page) " / " counter(pages); font: 7.5pt 'Noto Sans CJK KR'; color:#8D96AD; }} }}
* {{ box-sizing: border-box; }}
body {{ font-family: 'Noto Sans CJK KR', 'Noto Sans KR', sans-serif; color:#0E1530; font-size:10.3pt; line-height:1.55; margin:0; }}
a {{ color:#00489C; text-decoration:none; }}
.head {{ display:flex; align-items:flex-end; justify-content:space-between; border-bottom:3px solid #00489C; padding-bottom:7px; }}
.head img {{ height:34px; }}
.head .r {{ text-align:right; }}
.title {{ font-size:21pt; font-weight:900; color:#282B6E; letter-spacing:-0.5px; line-height:1.1; }}
.title span {{ color:#00489C; }}
.date {{ font-size:9pt; color:#6E7891; margin-top:3px; }}
.hero {{ margin-top:8px; background:#EAF0FA; border-left:5px solid #00489C; border-radius:8px; padding:8px 14px; }}
.hero .k {{ font-size:7.8pt; font-weight:700; color:#00489C; letter-spacing:0.5px; }}
.hero .h {{ font-size:12.5pt; font-weight:700; color:#282B6E; margin-top:1px; }}
.kpis {{ display:flex; gap:8px; margin-top:8px; }}
.kpi {{ flex:1; border:1px solid #D5DEEE; border-radius:8px; padding:6px 10px; background:#fff; }}
.kl {{ font-size:8.4pt; color:#6E7891; }} .kv {{ font-size:13.5pt; font-weight:900; color:#282B6E; }}
.kd {{ font-size:8.4pt; color:#6E7891; }} .kd.down {{ color:#1E8E5A; }} .kd.up {{ color:#D63A40; }}
h2 {{ font-size:12pt; color:#282B6E; margin:12px 0 6px; padding-bottom:4px; border-bottom:1px solid #D5DEEE; }}
h2.issues {{ page-break-before: always; margin-top:0; }}
h2 .n {{ display:inline-block; background:#282B6E; color:#fff; border-radius:4px; font-size:8.5pt; padding:0 6px; margin-right:6px; }}
table.bids {{ width:100%; border-collapse:collapse; table-layout:fixed; font-size:9.2pt; line-height:1.4; }}
.bids th {{ background:#282B6E; color:#fff; font-weight:700; text-align:left; padding:4px 5px; white-space:nowrap; }}
.bids td {{ border-bottom:1px solid #E2E8F3; padding:4px 5px; vertical-align:top; }}
.bids tr:nth-child(even) td {{ background:#F5F8FD; }}
.bids tr.grp td {{ background:#EAF0FA; color:#282B6E; font-weight:900; font-size:9.8pt; padding:4px 8px; border-bottom:1.5px solid #A9BCDB; }}
.bids tr.grp td span {{ font-weight:700; font-size:8.6pt; color:#00489C; margin-left:4px; }}
.bids tr.grp td em {{ font-style:normal; font-weight:400; font-size:8.2pt; color:#6E7891; margin-left:6px; }}
.bt a {{ font-weight:700; color:#282B6E; }} .bn {{ font-size:8.1pt; color:#6E7891; }}
.og {{ font-size:8.8pt; }}
.tm {{ font-size:8.8pt; font-weight:700; color:#282B6E; }}
.dl {{ color:#D63A40; font-weight:700; white-space:nowrap; }} .am {{ font-weight:700; white-space:nowrap; }}
.bidnote {{ margin-top:5px; font-size:8.8pt; color:#4A5372; background:#F5F8FD; border:1px dashed #A9BCDB; border-radius:6px; padding:5px 9px; text-align:justify; }}
.issue {{ border:1px solid #D5DEEE; border-radius:8px; padding:9px 12px; margin-bottom:8px; page-break-inside:avoid; }}
.ih {{ margin-bottom:3px; }}
.tag {{ display:inline-block; background:#EAF0FA; color:#00489C; font-size:8.4pt; font-weight:700; border-radius:10px; padding:1px 8px; margin-right:6px; }}
.it {{ font-weight:700; font-size:11.2pt; color:#282B6E; }}
.ib {{ color:#2A3150; text-align:justify; }}
.ip {{ margin-top:5px; font-size:9.6pt; background:#F5F8FD; border-left:3px solid #00489C; border-radius:3px; padding:4px 8px; color:#282B6E; text-align:justify; }}
.ip b {{ color:#00489C; margin-right:4px; }}
.is {{ margin-top:3px; font-size:8.2pt; color:#8D96AD; }}
.quote {{ margin-top:14px; background:#141850; color:#fff; border-radius:10px; padding:13px 18px; page-break-inside:avoid; }}
.quote .qk {{ font-size:7.8pt; letter-spacing:1px; color:#A9BCDB; font-weight:700; }}
.quote .qt {{ font-size:12.5pt; font-weight:700; margin-top:3px; }}
.quote .qa {{ font-size:9.4pt; color:#D5DEEE; margin-top:3px; }}
</style></head><body>
<div class="head"><img src="{logo_uri}"><div class="r"><div class="title">삼원의 <span>오늘</span></div>
<div class="date">{esc(c["date"])} ({esc(c["weekday"])}) · 운송·물류 일간 브리핑</div></div></div>
<div class="hero"><div class="k">오늘의 헤드라인</div><div class="h">{esc(c.get("headline",""))}</div></div>
<div class="kpis">{kpis}</div>
<h2><span class="n">01</span>입찰 공고 <span style="font-size:8.6pt;color:#6E7891;font-weight:400">장기·정기 계약 우선</span></h2>{bids}
<h2 class="issues"><span class="n">02</span>운송·물류 이슈</h2>{issues}
<div class="quote"><div class="qk">오늘의 한마디</div><div class="qt">“{esc(q.get("text",""))}”</div><div class="qa">— {esc(q.get("author",""))}</div></div>
</body></html>"""
    document = HTML(string=doc).render()
    document.write_pdf(out_path)
    check_bid_page(document)
    check_wraps(document)

def check_bid_page(document):
    """입찰 섹션(01)이 1페이지 안에 모두 들어갔는지 확인."""
    def has_bid(box):
        if 'bid' in (getattr(box, 'element_tag', '') or '') : return True
        el = getattr(box, 'element', None)
        if el is not None and el.get('class') in ('bidnote',) : return True
        if el is not None and el.tag == 'tr' and el.get('class') in ('bid', 'grp'): return True
        return any(has_bid(ch) for ch in getattr(box, 'children', []) or [])
    pages = [i + 1 for i, p in enumerate(document.pages) if has_bid(p._page_box)]
    over = [p for p in pages if p != 1]
    print("BID-PAGE: OK (1페이지)" if not over else f"BID-OVERFLOW: 입찰 표가 {sorted(set(pages))}페이지에 걸침 — 우선순위 낮은 공고를 줄이고 '외 n건 생략'으로 표기")

def check_wraps(document, short=5):
    """몇 글자 때문에 줄이 넘어간 곳(마지막 줄이 short자 이하)을 찾아 WRAP 경고로 출력."""
    from weasyprint.formatting_structure import boxes as B
    def text_of(box):
        if isinstance(box, B.TextBox): return box.text
        return "".join(text_of(ch) for ch in getattr(box, "children", []))
    found = []
    def walk(box):
        lines = [ch for ch in getattr(box, "children", []) if isinstance(ch, B.LineBox)]
        if len(lines) >= 2:
            last = text_of(lines[-1]).strip()
            if 0 < len(last.replace(" ", "")) <= short:
                found.append((text_of(lines[-2]).strip()[-20:], last))
        for ch in getattr(box, "children", []):
            if not isinstance(ch, B.LineBox): walk(ch)
    for p in document.pages: walk(p._page_box)
    for prev, last in found:
        print(f"WRAP: …{prev} ⏎ {last}")
    print(f"WRAP-CHECK: {len(found)}건")

if __name__ == "__main__":
    build(json.load(open(sys.argv[1], encoding="utf-8")), sys.argv[2], sys.argv[3])
