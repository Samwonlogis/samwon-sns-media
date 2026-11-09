#!/usr/bin/env python3
"""기존 산출물(구판 키트로 만든 Word·Excel·PowerPoint)을 현행 서식 표준(v3)으로 다시 만든다.

  python3 restyle.py <파일 또는 폴더> [출력폴더]        (출력폴더가 없으면 같은 폴더에 덮어쓴다)

Word : 표지 정보(제목·부제·구분·부서·문서번호·작성일·결재자)와 본문(제목·문단·글머리표·표·서식표·안내 상자·쪽 나눔)을
       읽어 원고(.md)로 되살린 뒤 samwon_md로 다시 조판한다(쪽 낭비 자동 보정 포함). 원고는 임시 폴더에 둔다.
Excel: 표지를 새 표준으로 바꾸고, 표 시트는 제목·머리글·데이터·수식·목록 선택을 읽어 표준 시트로 다시 만든다.
PowerPoint: 첫 장(표지)만 새 표준 표지로 바꾼다.
"""
import copy
import os
import re
import shutil
import sys
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import samwon_md  # noqa: E402
from samwon_cover import *  # noqa: F401,F403,E402
import samwon_cover as _c  # noqa: E402

import docx  # noqa: E402
from docx.oxml.ns import qn  # noqa: E402

RE_DATE = re.compile(r"(20\d\d)\s*[.\-/]\s*(\d{1,2})\s*[.\-/]\s*(\d{1,2})")


def _color(run):
    try:
        return str(run.font.color.rgb) if run.font.color is not None and run.font.color.type is not None else ""
    except Exception:
        return ""


def _size(run):
    return run.font.size.pt if run.font.size else 0


def _rich(par):
    out = []
    for r in par.runs:
        t = r.text
        if not t:
            continue
        out.append(f"**{t}**" if r.font.bold and t.strip() else t)
    s = "".join(out).replace("****", "")
    return s


def _fill(cell):
    shd = cell._element.tcPr.find(qn("w:shd")) if cell._element.tcPr is not None else None
    return (shd.get(qn("w:fill")) or "").upper() if shd is not None else ""


def _cells_text(row):
    seen, out = set(), []
    for c in row.cells:
        if id(c._element) in seen:
            continue
        seen.add(id(c._element))
        out.append(" ".join(_rich(p) for p in c.paragraphs).strip().replace("|", "/").replace("\n", " "))
    return out


def docx_to_md(path):
    d = docx.Document(path)
    body = d.element.body
    els = list(body.iterchildren())
    meta = {"file": Path(path).stem, "approver": "대표이사"}
    # 표지 끝: 첫 쪽 나눔(구판) 또는 첫 구역 나눔(신판)
    cut = 0
    for i, el in enumerate(els):
        if el.tag == qn("w:p") and (el.findall(".//" + qn("w:br")) and any(b.get(qn("w:type")) == "page" for b in el.iter(qn("w:br")))
                                    or el.find(qn("w:pPr")) is not None and el.find(qn("w:pPr")).find(qn("w:sectPr")) is not None):
            cut = i; break
    # 표지 안의 모든 문단(표 안에 든 v4 표지 포함)과 표를 차례로 읽는다
    cover_ps, cover_tbls = [], []
    for el in els[:cut]:
        if el.tag == qn("w:p"):
            cover_ps.append(el)
        elif el.tag == qn("w:tbl"):
            cover_tbls += list(el.iter(qn("w:tbl")))
            cover_ps += [p for p in el.iter(qn("w:p"))]
    for el in cover_ps:
        p = docx.text.paragraph.Paragraph(el, d)
        runs = [r for r in p.runs if r.text.strip()]
        if not runs:
            continue
        sz, col, txt = _size(runs[0]), _color(runs[0]).upper(), p.text.strip()
        m = RE_DATE.search(txt)
        if "title" not in meta and (sz >= 24 or (sz >= 15 and col == "282B6E" and runs[0].font.bold)):
            meta["title"] = txt
        elif "작성일" in txt and m:
            meta["date"] = f"{m.group(1)}. {int(m.group(2)):02d}. {int(m.group(3)):02d}."
        elif 12.5 <= sz < 15 and "title" in meta and "subtitle" not in meta:
            meta["subtitle"] = txt
        elif sz and 10.5 < sz <= 12.5 and col in ("00489C", "08448C") and "title" not in meta:
            meta["kind"] = txt
    for el in cover_tbls:
        t = docx.table.Table(el, d)
        try:
            rows_txt = [[c.replace("**", "") for c in _cells_text(r)] for r in t.rows]
        except Exception:
            continue
        if len(rows_txt) == 2 and "주관 부서" in rows_txt[0]:          # v3·v4 표지: 1행 항목명 · 2행 값
            for k, v in zip(rows_txt[0], rows_txt[1]):
                if k == "주관 부서": meta["team"] = v
                elif k == "문서번호": meta["doc_no"] = v
                elif k == "작성일": meta["date"] = v
                elif k.startswith("제정"): meta["version"] = v
            continue
        if len(t.rows) == 3 and any("기안" in c.text for c in t.rows[0].cells):   # 결재방
            names = [c.text.strip() for c in t.rows[1].cells][-3:]
            if len(names) == 3:
                meta["drafter"], meta["leader"], meta["approver"] = names[0], names[1], names[2] or "대표이사"
            continue
        for row in t.rows:                                              # 구판 표지 정보 표
            cells = _cells_text(row)
            if len(cells) >= 2:
                k, v = cells[0].replace("**", ""), cells[1].replace("**", "")
                if k == "주관 부서": meta["team"] = v
                elif k == "문서번호": meta["doc_no"] = v
                elif k.startswith("제정"): meta["version"] = re.sub(r"\s*\(.*\)$", "", v)
    meta.setdefault("title", Path(path).stem.split("_", 1)[-1])
    # 본문
    lines = []
    for el in els[cut + 1:]:
        if el.tag == qn("w:p"):
            p = docx.text.paragraph.Paragraph(el, d)
            if any(b.get(qn("w:type")) == "page" for b in el.iter(qn("w:br"))):
                lines += ["", "===", ""]; continue
            txt = p.text.strip()
            if not txt:
                continue
            runs = [r for r in p.runs if r.text.strip()]
            sz, col = (_size(runs[0]), _color(runs[0]).upper()) if runs else (0, "")
            bold_all = runs and all(r.font.bold for r in runs)
            if sz >= 14 and bold_all:
                lines += ["", "# " + txt, ""]
            elif 11.5 <= sz < 14 and bold_all and col in ("00489C", "08448C"):
                lines += ["", "## " + txt, ""]
            elif 10.6 <= sz < 11.5 and bold_all and col in ("282B6E",):
                lines += ["### " + txt]
            elif txt.startswith("•") or txt.startswith("·"):
                lines.append("- " + _rich(p).lstrip("•·  ").strip())
            elif txt.startswith("☐"):
                lines.append("[ ] " + _rich(p).lstrip("☐  ").strip())
            else:
                lines += [_rich(p).strip(), ""]
        elif el.tag == qn("w:tbl"):
            t = docx.table.Table(el, d)
            rows = [_cells_text(r) for r in t.rows]
            if not rows:
                continue
            first = t.rows[0].cells
            if len(t.rows) == 1 and len(t.columns) == 1:
                box = [ln for ln in (_rich(pp).strip() for pp in first[0].paragraphs) if ln]
                lines += [""] + ["> " + b for b in box] + [""]
                continue
            head_navy = _fill(first[0]) in ("282B6E", "0A2A5C", "141850")
            if not head_navy and len(t.columns) % 2 == 0 and _fill(first[0]) in ("EAF0FA", "F5F8FD", "EEF2F8"):
                lines += ["", "<form>"]
                for r in rows:
                    lines.append("| " + " | ".join(c.replace("**", "") if i % 2 == 0 else c for i, c in enumerate(r)) + " |")
                lines.append("")
                continue
            ncol = max(len(r) for r in rows)
            rows = [r + [""] * (ncol - len(r)) for r in rows]
            n_fill = 0
            while len(rows) > 2 and not any(c.strip() for c in rows[-1]):
                rows.pop(); n_fill += 1
            if n_fill:
                lines.append(f"<fill {n_fill}>")
            lines.append("| " + " | ".join(c.replace("**", "") for c in rows[0]) + " |")
            lines.append("|" + "---|" * ncol)
            for r in rows[1:]:
                lines.append("| " + " | ".join(r) + " |")
            lines.append("")
    # 끝 정리: 쪽 나눔 연속·끝 쪽 나눔 제거
    out, prev = [], None
    for ln in lines:
        if ln == "" and prev == "":
            continue
        out.append(ln); prev = ln
    while out and out[-1] in ("", "==="):
        out.pop()
    head = [f"@{k}: {v}" for k, v in meta.items() if v]
    return "\n".join(head) + "\n\n" + "\n".join(out) + "\n", meta


def restyle_docx(path, outdir):
    md, meta = docx_to_md(path)
    import tempfile
    mdp = os.path.join(tempfile.mkdtemp(prefix="restyle_"), Path(path).stem + ".md")
    open(mdp, "w", encoding="utf-8").write(md)
    m, body, _ = samwon_md.parse(mdp)
    return samwon_md.build_docx(m, body, outdir)


# ───────────────────────────── Excel ─────────────────────────────
import openpyxl  # noqa: E402


def restyle_xlsx(path, outdir):
    src = openpyxl.load_workbook(path)
    cov = src.worksheets[0]
    vals = {c.coordinate: c.value for row in cov.iter_rows() for c in row if c.value not in (None, "")}
    title = next((v for k, v in vals.items() if cov[k].font and cov[k].font.sz and cov[k].font.sz >= 20), Path(path).stem)
    subtitle = next((v for k, v in vals.items() if cov[k].font and cov[k].font.sz and 12.5 <= cov[k].font.sz < 20), "")
    info, how = {}, []
    for k, v in vals.items():
        if isinstance(v, str) and v in ("주관 부서", "문서번호", "제정·개정"):
            nxt = cov.cell(row=cov[k].row, column=cov[k].column + 1).value
            below = cov.cell(row=cov[k].row + 1, column=cov[k].column).value
            if (not nxt or str(nxt) in ("주관 부서", "문서번호", "작성일", "제정·개정")) and below:
                nxt = below                                  # v3·v4 표지: 항목명 아래 줄에 값
            info[v] = re.sub(r"\s*\(.*\)$", "", str(nxt or ""))
        if isinstance(v, str) and v.startswith("•"):
            how.append(v.lstrip("• ").strip())
    names = ["", "", "대표이사"]
    for row in cov.iter_rows():
        for c in row:
            if c.value == "기안":
                names = [str(cov.cell(row=c.row + 1, column=c.column + j).value or "") for j in range(3)]
    date = next((f"{m.group(1)}. {int(m.group(2)):02d}. {int(m.group(3)):02d}." for v in vals.values()
                 if isinstance(v, str) and (m := RE_DATE.search(v)) and len(v) < 16), None)
    wb = xlsx_new(str(title), str(subtitle), team=info.get("주관 부서", ""), doc_no=info.get("문서번호", ""),
                  version=info.get("제정·개정") or "제정 v1.0", how_to=how, drafter=names[0], leader=names[1],
                  approver=names[2] or "대표이사", date=date, kind="관리 대장·서식")
    for ws in src.worksheets[1:]:
        # 머리글 행 찾기(네이비 채움)
        hr = None
        for r in range(1, min(ws.max_row, 8) + 1):
            c = ws.cell(row=r, column=1)
            if c.fill and c.fill.fgColor and str(c.fill.fgColor.rgb or "")[-6:].upper() in ("282B6E", "0A2A5C", "141850") and c.value:
                hr = r; break
        if hr is None:
            # 표가 아닌 시트는 그대로 복사
            nws = wb.create_sheet(ws.title)
            for row in ws.iter_rows():
                for c in row:
                    nws.cell(row=c.row, column=c.column, value=c.value)
            continue
        ncol = max((c.column for c in ws[hr] if c.value not in (None, "")), default=1)
        headers = [("" if ws.cell(row=hr, column=j).value is None else str(ws.cell(row=hr, column=j).value)) for j in range(1, ncol + 1)]
        ttl = ws.cell(row=1, column=1).value if hr > 1 else None
        ttl = re.sub(r"^삼원종합물류\(주\)\s*\|\s*", "", str(ttl)) if ttl else None
        note = ws.cell(row=2, column=1).value if hr > 2 else None
        from openpyxl.formula.translate import Translator
        data, keep_f, blank_f = [], [], {}
        for r in range(hr + 1, ws.max_row + 1):
            row = [ws.cell(row=r, column=j).value for j in range(1, ncol + 1)]
            fcells = {j: v for j, v in enumerate(row) if isinstance(v, str) and v.startswith("=")}
            plain = [v for j, v in enumerate(row) if j not in fcells and v not in (None, "")]
            if plain:
                for j in fcells: row[j] = None
                keep_f.append((len(data), r, fcells)); data.append(row)
            else:
                for j, f in fcells.items():
                    blank_f.setdefault(j, (r, f))
        widths = [ws.column_dimensions[openpyxl.utils.get_column_letter(j)].width or 12 for j in range(1, ncol + 1)]
        nws = sheet(wb, ws.title, headers, widths=widths, rows=data, title=ttl, note=note)
        h2 = nws._samwon_header_row; last = nws._samwon_last_row
        L = openpyxl.utils.get_column_letter
        for i, r0, fcells in keep_f:
            for j, f in fcells.items():
                nws.cell(row=h2 + 1 + i, column=j + 1, value=Translator(f, origin=f"{L(j + 1)}{r0}").translate_formula(f"{L(j + 1)}{h2 + 1 + i}"))
        for j, (r0, f) in blank_f.items():
            for r in range(h2 + 1 + len(data), last + 1):
                nws.cell(row=r, column=j + 1, value=Translator(f, origin=f"{L(j + 1)}{r0}").translate_formula(f"{L(j + 1)}{r}"))
            for i, _r0, fcells in keep_f:          # 데이터 행에 수식이 없던 자동 계산 열도 채운다
                if j not in fcells and nws.cell(row=h2 + 1 + i, column=j + 1).value in (None, ""):
                    nws.cell(row=h2 + 1 + i, column=j + 1, value=Translator(f, origin=f"{L(j + 1)}{r0}").translate_formula(f"{L(j + 1)}{h2 + 1 + i}"))
        for cf in ws.conditional_formatting:      # 조건부 서식(경보 색 등) 유지 — 머리글 행 위치가 같으므로 범위 그대로
            for rule in cf.rules:
                nws.conditional_formatting.add(str(cf.sqref), rule)
        for dv in ws.data_validations.dataValidation:
            if dv.type == "list" and dv.formula1:
                opts = [o.strip() for o in dv.formula1.strip('"').split(",")]
                for rng in str(dv.sqref).split():
                    col = re.match(r"[A-Z]+", rng).group(0)
                    dropdown(nws, col, opts)
        for j in range(1, ncol + 1):   # 숫자·날짜 표시 형식 유지
            fmt = ws.cell(row=hr + 1, column=j).number_format
            if fmt and fmt != "General":
                for r in range(h2 + 1, last + 1):
                    nws.cell(row=r, column=j).number_format = fmt
    out = os.path.join(outdir, Path(path).name)
    xlsx_save(wb, out)
    return out


# ─────────────────────────── PowerPoint ───────────────────────────
from pptx import Presentation  # noqa: E402


def restyle_pptx(path, outdir):
    prs = Presentation(path)
    s0 = prs.slides[0]
    texts = []
    for sh in s0.shapes:
        if sh.has_text_frame:
            for p in sh.text_frame.paragraphs:
                if p.text.strip():
                    sz = max((r.font.size.pt for r in p.runs if r.font.size), default=0)
                    texts.append((sz, p.text.strip()))
    title = max(texts, default=(0, Path(path).stem))[1]
    sub = next((t for s, t in texts if 16 <= s < 30 and t != title and "삼원종합물류" not in t), "")
    team = next((t.split("·")[0].strip() for s, t in texts if "팀" in t and "·" in t), "")
    flat = [t for _, t in texts]
    for i, t in enumerate(flat[:-1]):                       # v3·v4 표지: '주관 부서' 다음 문단이 값
        if t == "주관 부서" and flat[i + 1].endswith("팀"):
            team = flat[i + 1]
    doc_no = next((m.group(0) for s, t in texts if (m := re.search(r"SW-[A-Z]{2}-\d{4}-\d{2}", t))), "")
    names = ["", "", "대표이사"]
    for sh in s0.shapes:
        if sh.has_table:
            tb = sh.table
            if any("기안" in tb.cell(0, j).text for j in range(len(tb.columns))):
                names = [tb.cell(1, j).text.strip() for j in range(1, 4)]
    layout = next((l for l in prs.slide_layouts if len(l.placeholders) == 0), prs.slide_layouts[-1])
    new = prs.slides.add_slide(layout)
    for ph in list(new.placeholders):
        ph._element.getparent().remove(ph._element)
    _c.draw_cover_slide(prs, new, title, sub, team=team, doc_no=doc_no, drafter=names[0], leader=names[1],
                        approver=names[2] or "대표이사", kind="발표 자료")
    lst = prs.slides._sldIdLst
    ids = list(lst)
    lst.remove(ids[-1]); lst.insert(0, ids[-1])
    old = ids[0]
    prs.part.drop_rel(old.rId); lst.remove(old)
    out = os.path.join(outdir, Path(path).name)
    prs.save(out)
    return out


def main(argv):
    args = [a for a in argv if not a.startswith("--")]
    src = Path(args[0]); outdir = args[1] if len(args) > 1 else None
    files = [f for f in (sorted(src.rglob("*")) if src.is_dir() else [src])
             if f.suffix.lower() in (".docx", ".xlsx", ".pptx") and not f.name.startswith("~$")]
    for f in files:
        od = outdir or str(f.parent)
        os.makedirs(od, exist_ok=True)
        try:
            fn = {".docx": restyle_docx, ".xlsx": restyle_xlsx, ".pptx": restyle_pptx}[f.suffix.lower()]
            print("RESTYLED", fn(str(f), od))
        except Exception as e:  # 한 파일 실패가 전체를 멈추지 않게
            print("FAILED", f, type(e).__name__, e)


if __name__ == "__main__":
    main(sys.argv[1:])
