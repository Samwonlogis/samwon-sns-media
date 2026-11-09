#!/usr/bin/env python3
"""삼원 원고(.md) → 표준 Word·Excel 변환기 (samwon-standard, 과제7 2026-11-09).

원고 형식
  @title: 제목            @subtitle: 부제        @team: 인사팀
  @doc_no: SW-HR-2611-01  @drafter: 인사 과장    @leader: 인사팀장   @kind: 규정·서식
  @file: 2611인사_인사평가제도_규정·서식        (확장자 없이; .docx 생성)
  @landscape: yes         @date: 2026. 11. 02.

본문(Word)
  # 대제목 / ## 소제목 / ### 굵은 소항목
  - 글머리표,  [ ] 체크 항목,  > 안내 상자,  1. / ① / 가. 번호 문단(내어쓰기)
  ===  쪽 나눔(앞 쪽이 절반도 차지 않으면 자동으로 무시) · === 강제  항상 쪽 나눔
  | 머리 | 머리 |  (첫 행 머리글)   표 바로 앞 줄에 <form> 이면 '항목|기입란' 서식표, <fill N> 이면 빈 기입 행 N줄 추가
  **굵게** 지원.  빈 줄은 문단 구분.

엑셀(같은 원고 안 또는 별도 원고)
  @xlsx: 2611총무_만기통합관리대장      (본문 끝의 '@@엑셀' 줄 아래 '%% 시트' 블록이 엑셀로 간다)
  @xlsx_title: 만기 통합 관리 대장      @how_to: 사용 방법 1 || 사용 방법 2
  %% 시트: 대장 | 제목 | 안내문 | blank=40 | widths=6,20,12
  | 머리 | 머리 |  (표; 데이터 행 그대로)
  @dropdown: C=적용,미적용
  @formula: H=IF(G{r}="","",G{r}-TODAY())      ({r} = 행 번호, 데이터+빈 행 전체에 채움)

  python3 samwon_md.py 원고.md [출력폴더]
"""
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from samwon_cover import *  # noqa: F401,F403
from samwon_cover import BLUE, NAVY, SKY, GRAY, _shade, _font  # noqa
try:
    import samwon_layout
except ImportError:          # 쪽 사용률 도구가 없으면 자동 보정만 생략
    samwon_layout = None

from openpyxl.styles import Alignment, PatternFill
from openpyxl.utils import get_column_letter


def _cells(line):
    s = line.strip()
    if s.startswith("|"): s = s[1:]
    if s.endswith("|"): s = s[:-1]
    return [c.strip() for c in s.split("|")]


def _is_sep(line):
    return bool(re.match(r"^\|?\s*:?-{2,}:?\s*(\|\s*:?-{2,}:?\s*)*\|?\s*$", line.strip()))


def parse(path):
    meta, body, xl = {}, [], []
    cur = body
    for raw in open(path, encoding="utf-8").read().splitlines():
        m = re.match(r"^@(\w+):\s*(.*)$", raw)
        if m and m.group(1) not in ("dropdown", "formula"):
            meta[m.group(1)] = m.group(2).strip(); continue
        if raw.strip() == "@@엑셀":
            cur = xl; continue
        cur.append(raw)
    return meta, body, xl


def _tables(lines, i):
    rows = []
    while i < len(lines) and lines[i].strip().startswith("|"):
        if not _is_sep(lines[i]): rows.append(_cells(lines[i]))
        i += 1
    return rows, i


def _emit_docx(meta, body, outdir, level=0, soft=frozenset(), markers=False):
    compact(level)
    land = meta.get("landscape") == "yes"
    doc = docx_new(meta["title"], meta.get("subtitle", ""), team=meta.get("team", ""), doc_no=meta.get("doc_no", ""),
                   kind=meta.get("kind", ""), drafter=meta.get("drafter", ""), leader=meta.get("leader", ""),
                   approver=meta.get("approver", "대표이사"), landscape=land, date=meta.get("date"),
                   version=meta.get("version", "제정 v1.0"))
    i, mode, buf, nbreak = 0, None, [], 0

    def flush():
        if buf:
            para(doc, " ".join(buf)); buf.clear()
    while i < len(body):
        ln = body[i]; s = ln.strip()
        if not s:
            flush(); i += 1; continue
        if s == "===" or s == "=== 강제":
            flush(); nbreak += 1
            if s == "=== 강제" or nbreak not in soft:
                p = page_break(doc)
                if markers:
                    _font(p.add_run(f"§BK{nbreak}§"), 2, False, "FFFFFF")
            i += 1; continue
        if s in ("<form>",) or s.startswith("<fill"):
            flush(); mode = s; i += 1; continue
        if s.startswith("|"):
            flush(); rows, i = _tables(body, i)
            if mode == "<form>":
                pairs = []
                for r in rows:
                    for k in range(0, len(r) - 1, 2): pairs.append((r[k].replace("**", ""), r[k + 1]))
                cols = max(1, max(len(r) for r in rows) // 2)
                form_table(doc, pairs, cols=min(cols, 2))
            else:
                n = int(mode.split()[1].rstrip(">")) if mode and mode.startswith("<fill") else 0
                rows = [[c.replace("**", "") if ri == 0 else c for c in r] for ri, r in enumerate(rows)]
                table(doc, rows, fill_rows=n)
            mode = None; continue
        if s.startswith("### "):
            flush(); h3(doc, s[4:])
        elif s.startswith("## "):
            flush(); h2(doc, s[3:])
        elif s.startswith("# "):
            flush(); h1(doc, s[2:])
        elif s.startswith("[ ] "):
            flush(); checklist(doc, [s[4:]])
        elif s.startswith("- ") or s.startswith("• "):
            flush(); bullets(doc, [s[2:]])
        elif re.match(r"^(\(?\d{1,2}[.)]|[①-⑳]|[가-하]\.)\s", s):
            flush(); numbered(doc, s)
        elif s.startswith("> "):
            flush(); box = []
            while i < len(body) and body[i].strip().startswith("> "):
                box.append(body[i].strip()[2:]); i += 1
            note_box(doc, box); continue
        else:
            buf.append(s)
        i += 1
    flush()
    out = os.path.join(outdir, meta["file"] + ".docx")
    docx_save(doc, out)
    compact(0)
    return out, nbreak


def build_docx(meta, body, outdir, auto=True):
    """원고 → Word. auto=True이면 LibreOffice로 쪽 사용률을 재어
    ① '===' 쪽 나눔 앞 쪽이 절반도 차지 않으면 그 나눔을 없애고 ② 마지막 쪽이 15% 미만이면 간격을 단계적으로 줄인다."""
    out, nb = _emit_docx(meta, body, outdir)
    if not auto or samwon_layout is None or samwon_layout.SOFFICE is None:
        return out
    soft = set()
    for _ in range(4):                       # 나눔 하나를 없애면 뒤 쪽 배치가 바뀌므로 안정될 때까지 반복
        if not nb or len(soft) >= nb:
            break
        _emit_docx(meta, body, outdir, markers=True, soft=frozenset(soft))
        r = samwon_layout.page_fills(out, keep_pdf=True)
        if not r:
            break
        fills, pdf, texts = r
        added = False
        for k in range(1, nb + 1):
            if k in soft:
                continue
            pg = next((idx for idx, t in enumerate(texts) if f"§BK{k}§" in t.replace(" ", "").replace("\n", "")), None)
            # 표식은 쪽 나눔 뒤(다음 쪽 맨 위)에 찍히므로 나눔 앞 쪽은 pg-1
            if pg is not None and pg - 1 >= 1 and fills[pg - 1][1] < 0.5:
                soft.add(k); added = True
                break                         # 앞에서부터 하나씩
        if not added:
            break
    best = None
    for level in (0, 1, 2):
        out, _ = _emit_docx(meta, body, outdir, level, frozenset(soft))
        fills = samwon_layout.page_fills(out)
        if not fills:
            break
        pages, lastf = len(fills), fills[-1][1]
        if best is None or pages < best[1]:
            best = (level, pages, lastf)
        if lastf >= 0.15 or pages <= 2:
            break
    if best and best[0] != level:
        out, _ = _emit_docx(meta, body, outdir, best[0], frozenset(soft))
    if best:
        print(f"LAYOUT {os.path.basename(out)}: {best[1]}쪽(표지 포함) · 간격 단계 {best[0]} · 없앤 쪽 나눔 {len(soft)}개")
    return out


def build_xlsx(meta, xl, outdir):
    how = [h.strip() for h in meta.get("how_to", "").split("||") if h.strip()]
    wb = xlsx_new(meta.get("xlsx_title", meta["title"]), meta.get("xlsx_subtitle", meta.get("subtitle", "")),
                  team=meta.get("team", ""), doc_no=meta.get("xlsx_doc_no", meta.get("doc_no", "")), how_to=how,
                  drafter=meta.get("drafter", ""), leader=meta.get("leader", ""), date=meta.get("date"),
                  kind=meta.get("xlsx_kind", "관리 대장·서식"), version=meta.get("version", "제정 v1.0"))
    i = 0
    while i < len(xl):
        s = xl[i].strip()
        if not s.startswith("%% 시트:"):
            i += 1; continue
        parts = [p.strip() for p in s[len("%% 시트:"):].split("|")]
        name = parts[0]; title = parts[1] if len(parts) > 1 and parts[1] else name
        note = parts[2] if len(parts) > 2 and parts[2] else None
        opts = dict(p.split("=", 1) for p in parts[3:] if "=" in p)
        i += 1
        while i < len(xl) and not xl[i].strip().startswith("|"): i += 1
        rows, i = _tables(xl, i)
        headers, data = [h.replace("**", "") for h in rows[0]], rows[1:]
        data = [[(float(v) if re.fullmatch(r"-?\d+(\.\d+)?", v) and not v.startswith("0") or v == "0" else (v or None))
                 for v in r] for r in data]
        widths = [float(w) for w in opts["widths"].split(",")] if "widths" in opts else None
        blank = int(opts["blank"]) if "blank" in opts else None
        ws = sheet(wb, name, headers, widths=widths, rows=data, title=title, note=note, blank_rows=blank)
        hr = ws._samwon_header_row; last = ws._samwon_last_row
        for j, h in enumerate(headers, 1):
            if re.search(r"율", h):
                for r in range(hr + 1, last + 1): ws.cell(row=r, column=j).number_format = "0.0%"
            elif re.search(r"(일|만료일|날짜)(\(자동\))?$|일자|마감일|기한\(종료", h) and "(일)" not in h and "일수" not in h:
                for r in range(hr + 1, last + 1): ws.cell(row=r, column=j).number_format = "yyyy-mm-dd"
            elif re.search(r"금액|원\)|비용|단가", h):
                for r in range(hr + 1, last + 1): ws.cell(row=r, column=j).number_format = "#,##0"
        while i < len(xl) and (xl[i].strip().startswith("@dropdown:") or xl[i].strip().startswith("@formula:") or not xl[i].strip()):
            t = xl[i].strip()
            if t.startswith("@dropdown:"):
                col, vals = t[len("@dropdown:"):].strip().split("=", 1)
                dropdown(ws, col.strip(), [v.strip() for v in vals.split(",")], last_row=last)
            elif t.startswith("@formula:"):
                col, f = t[len("@formula:"):].strip().split("=", 1)
                for r in range(hr + 1, last + 1):
                    ws[f"{col.strip()}{r}"] = "=" + f.replace("{r}", str(r))
            i += 1
    out = os.path.join(outdir, meta["xlsx"] + ".xlsx")
    xlsx_save(wb, out); return out


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    src = args[0]; outdir = args[1] if len(args) > 1 else "."
    meta, body, xl = parse(src)
    outs = []
    auto = "--no-auto" not in sys.argv
    if meta.get("file") and any(l.strip() for l in body): outs.append(build_docx(meta, body, outdir, auto))
    if meta.get("xlsx"): outs.append(build_xlsx(meta, xl, outdir))
    for o in outs: print("OUT", o)


if __name__ == "__main__":
    main()
