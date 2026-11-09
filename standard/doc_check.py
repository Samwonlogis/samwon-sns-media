#!/usr/bin/env python3
"""삼원 문서 표준 자동 검사기 (samwon-standard, 과제4 2026-11-09).

사용: python3 doc_check.py <파일 또는 폴더> [...] [--json] [--report]
      --report : 정기 보고서 PDF(YYYYMMDD_보고서명.pdf) 규칙으로 파일명 검사
검사: 파일명 · 표지 4요소(영문 로고·제목·작성일·결재방) · 브랜드 3색 외 유채색 · 금지 표현 · 차량 대수 · (PDF) 쪽수·링크
      · 쪽 낭비(LibreOffice가 있으면 쪽마다 사용률을 재서 중간에 반 이상 빈 쪽·거의 빈 마지막 쪽·빈 기입 행만 있는 쪽을 찾는다)
      · 서식 기준(본문 글꼴 맑은 고딕 외 글꼴, 표 행 쪽 나눔 방지 누락)
      --no-layout : 쪽 사용률 검사 생략
판정: FAIL / WARN / PASS. FAIL이 있으면 종료 코드 1.
"""
import colorsys
import json
import re
import sys
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

TEAMS = "인사|홍보|총무|영업|운영|재경|비서|기획|경영"
NAME_DOC = re.compile(rf"^\d{{4}}({TEAMS})_.+")
NAME_REPORT = re.compile(r"^\d{8}_.+")
BRAND = ["00489C", "08448C", "282B6E", "141850", "EAF0FA", "F5F8FD", "D5DEEE", "0E1530", "6E7891", "D63A40", "FF6B70"]
BANNED = [(r"제때", "'제때'"), (r"하느냐에\s*따라", "'~하느냐에 따라'"),
          (r"업계\s*(최고|1위)|국내\s*1위", "근거 없는 최상급"), (r"알아보겠습니다|결론적으로 말하면", "AI 티 나는 표현"),
          (r"(수익|수입)\s*(을|를)?\s*보장|고수익", "수익 보장"),
          (r"최종\s*확인을\s*받으십시오|전문가\s*검토\s*기준으로\s*작성", "검토 안내 문구(대표 지시 2026-11-09 삭제)")]
FLEET = re.compile(r"(차량|화물차|보유|운영|넘버|영업용\s*번호)[^\n.]{0,10}?(\d[\d,]{1,})\s*대(?!\s*이상)")
OLD_COLORS = {"064285": "이전 로고색 #064285", "2F8CEB": "이전 포인트색 #2F8CEB", "0A2A5C": "이전 네이비 #0A2A5C"}

NEGATION = re.compile(r"하는\s*말|약속하는|쓰지\s*않|하지\s*않|않습니다|않는다|않음|금지|불가|아닙니다|없습니다|삼가|피한다|빼고|제외")


def banned_hit(pat, text):
    """금지 표현 검색. '쓰지 않는다'·'금지' 목록·따옴표 인용처럼 금지 사항을 설명하는 문맥은 제외한다."""
    for m in re.finditer(pat, text):
        before, after = text[max(0, m.start() - 3):m.start()], text[m.end():m.end() + 30]
        wide = text[max(0, m.start() - 25):m.start()]
        if any(q in before for q in "'‘\"“(·") or NEGATION.search(after) or re.search(r"금지\s*(표현|어|문구)|금지어", wide):
            continue
        return m
    return None


def _hex_ok(h):
    h = h.upper()
    if h in BRAND or h in ("FFFFFF", "000000", "AUTO"):
        return True
    try:
        r, g, b = (int(h[i:i + 2], 16) / 255 for i in (0, 2, 4))
    except ValueError:
        return True
    _, l, s = colorsys.rgb_to_hls(r, g, b)
    if s < 0.22 or l > 0.93 or l < 0.08:      # 무채색·아주 옅은/진한 색은 허용
        return True
    # 브랜드 색과 가까우면 허용
    for c in BRAND:
        cr, cg, cb = (int(c[i:i + 2], 16) / 255 for i in (0, 2, 4))
        if abs(cr - r) + abs(cg - g) + abs(cb - b) < 0.18:
            return True
    return False


def _xml_text(z, names):
    out = []
    for n in names:
        try:
            out.append(z.read(n).decode("utf-8", "ignore"))
        except KeyError:
            pass
    return "\n".join(out)


def _strip(xml):
    import html
    return html.unescape(re.sub(r"<[^>]+>", " ", xml))


def check_office(path: Path, res):
    fail, warn = res["fail"], res["warn"]
    z = zipfile.ZipFile(path)
    names = z.namelist()
    ext = path.suffix.lower()
    if ext == ".docx":
        body = _xml_text(z, ["word/document.xml"])
        xml = body + _xml_text(z, [n for n in names if n.startswith("word/header") or n.startswith("word/footer")])
        cut = [i for i in (body.find("<w:sectPr"), body.find('w:type="page"')) if i > 0]
        first = body[:min(cut)] if cut else body[:60000]
        has_logo = "<w:drawing" in first or "<pic:pic" in first
        cover_text = _strip(first)
    elif ext == ".xlsx":
        wb = _xml_text(z, ["xl/workbook.xml"])
        sheets = re.findall(r'<sheet [^>]*name="([^"]+)"', wb)
        xml = _xml_text(z, [n for n in names if n.startswith("xl/worksheets/") or n == "xl/sharedStrings.xml" or n == "xl/styles.xml"])
        cover_text = _strip(_xml_text(z, ["xl/sharedStrings.xml", "xl/worksheets/sheet1.xml"]))
        has_logo = any(n.startswith("xl/media/") for n in names) and any(n.startswith("xl/drawings/") for n in names)
        if not sheets or sheets[0] != "표지":
            fail.append(f"첫 시트가 '표지'가 아님({sheets[:1]})")
    elif ext == ".pptx":
        xml = _xml_text(z, [n for n in names if n.startswith("ppt/slides/slide")])
        # 첫 장 = presentation.xml의 첫 sldId가 가리키는 슬라이드(파일 이름 순서와 다를 수 있음)
        pres = _xml_text(z, ["ppt/presentation.xml"]); prel = _xml_text(z, ["ppt/_rels/presentation.xml.rels"])
        first_rid = (re.findall(r'<p:sldId [^>]*r:id="([^"]+)"', pres) or [""])[0]
        m = re.search(rf'Id="{first_rid}"[^>]*Target="([^"]+)"', prel) or re.search(rf'Target="([^"]+)"[^>]*Id="{first_rid}"', prel)
        first = "ppt/" + m.group(1).lstrip("/").replace("ppt/", "") if m else "ppt/slides/slide1.xml"
        s1 = _xml_text(z, [first])
        cover_text = _strip(s1)
        rels = _xml_text(z, [first.replace("slides/", "slides/_rels/") + ".rels"])
        has_logo = "image" in rels
    else:
        return
    if not has_logo:
        fail.append("표지에 로고 이미지 없음")
    for k in ("결", "기안", "팀장"):
        if k not in cover_text:
            fail.append(f"표지 결재방 누락('{k}')")
            break
    if not re.search(r"20\d\d[.\-/년]\s*\d{1,2}", cover_text):
        warn.append("표지에 작성일 표기가 보이지 않음")
    text = _strip(xml)
    res["_text"] = text
    if ext == ".docx":
        fonts = set(re.findall(r'w:eastAsia="([^"]+)"', body)) | set(re.findall(r'w:ascii="([^"]+)"', body))
        other = sorted(f for f in fonts if f not in ("맑은 고딕", "Malgun Gothic"))
        if other:
            warn.append("표준 외 글꼴: " + ", ".join(other[:4]))
        rows = body.count("<w:tr>") + body.count("<w:tr ")
        if rows >= 6 and body.count("<w:cantSplit") < rows * 0.5:
            warn.append("표 행 쪽 나눔 방지(cantSplit) 미적용 — 표 행이 쪽 사이에서 잘릴 수 있음")
    # 색 검사
    cols = set(c.upper() for c in re.findall(r'(?:w:color w:val|w:fill|srgbClr val|rgb)="(?:FF)?([0-9A-Fa-f]{6})"', xml))
    for c, why in OLD_COLORS.items():
        if c in cols:
            fail.append(f"{why} 사용")
            cols.discard(c)
    bad = sorted(c for c in cols if not _hex_ok(c))
    if bad:
        warn.append(f"브랜드 3색 외 유채색 {len(bad)}개: " + ", ".join("#" + c for c in bad[:6]))


def check_pdf(path: Path, res):
    fail, warn = res["fail"], res["warn"]
    try:
        from pypdf import PdfReader
    except ImportError:
        warn.append("pypdf 미설치 — PDF 검사 생략")
        return
    r = PdfReader(str(path))
    res["pages"] = len(r.pages)
    text = "\n".join((p.extract_text() or "") for p in r.pages)
    res["_text"] = text
    links = 0
    for pg in r.pages:
        for a in pg.get("/Annots") or []:
            o = a.get_object()
            if o.get("/A", {}).get("/URI"):
                links += 1
    res["links"] = links
    p1 = r.pages[0]
    try:
        has_img = bool(p1.images) or "/XObject" in str(p1.get("/Resources"))
    except Exception:
        has_img = True
    if not has_img and "SAMWON" not in text[:400]:
        warn.append("1쪽에 로고가 보이지 않음")
    raw = b"".join(pg.get_contents().get_data() if pg.get_contents() else b"" for pg in r.pages[:2])
    for c, why in {"0.024 0.259 0.522": "이전 로고색 #064285", "0.184 0.549 0.922": "이전 포인트색 #2F8CEB"}.items():
        if c.encode() in raw:
            fail.append(f"{why} 사용")


def check_layout(path: Path, res):
    try:
        import samwon_layout
    except ImportError:
        return
    if samwon_layout.SOFFICE is None or path.suffix.lower() not in (".docx", ".xlsx"):
        return
    found = samwon_layout.waste_findings(str(path), has_cover=True)
    if found is None:
        return
    for pno, f, why in found:
        if why.startswith("표지"):
            res["fail"].append(f"{why} — 표지는 반드시 1쪽"); continue
        res["warn"].append(f"쪽 낭비 {pno}쪽(사용률 {f * 100:.0f}%): {why}")


def check(path: Path, report=False, layout=True):
    res = {"file": path.name, "fail": [], "warn": []}
    stem = path.stem
    if report or path.suffix.lower() == ".pdf":
        if not (NAME_REPORT.match(stem) or NAME_DOC.match(stem)):
            res["warn"].append("파일명 규칙(YYYYMMDD_보고서명 또는 YYMM팀_업무내용) 아님")
    elif not NAME_DOC.match(stem):
        res["fail"].append("파일명 규칙 'YYMM팀_업무내용' 아님")
    try:
        if path.suffix.lower() in (".docx", ".xlsx", ".pptx"):
            check_office(path, res)
        elif path.suffix.lower() == ".pdf":
            check_pdf(path, res)
        if layout:
            check_layout(path, res)
    except zipfile.BadZipFile:
        res["fail"].append("파일이 손상됨(열 수 없음)")
    text = res.pop("_text", "")
    for pat, why in BANNED:
        m = banned_hit(pat, text)
        if m:
            res["fail"].append(f"금지 표현 {why}: '{m.group(0)}'")
    m = FLEET.search(text)
    if m and int(m.group(2).replace(",", "")) >= 20:
        res["warn"].append(f"차량 대수 표기 의심: '{m.group(0)}' (대외 문서면 삭제)")
    res["result"] = "FAIL" if res["fail"] else ("WARN" if res["warn"] else "PASS")
    return res


def main(argv):
    as_json = "--json" in argv
    report = "--report" in argv
    files = []
    for a in [x for x in argv if not x.startswith("--")]:
        p = Path(a)
        if p.is_dir():
            files += [f for f in sorted(p.rglob("*")) if f.suffix.lower() in (".docx", ".xlsx", ".pptx", ".pdf")
                      and not f.name.startswith("~$")]
        elif p.exists():
            files.append(p)
    layout = "--no-layout" not in argv
    results = [check(f, report, layout) for f in files]
    if as_json:
        print(json.dumps({"checked": len(results), "fail": sum(r["result"] == "FAIL" for r in results),
                          "warn": sum(r["result"] == "WARN" for r in results), "results": results}, ensure_ascii=False, indent=1))
    else:
        for r in results:
            print(f"[{r['result']}] {r['file']}")
            for x in r["fail"]:
                print(f"   ✗ {x}")
            for x in r["warn"]:
                print(f"   △ {x}")
        print(f"DOC-CHECK: {len(results)}건 · FAIL {sum(r['result'] == 'FAIL' for r in results)} · "
              f"WARN {sum(r['result'] == 'WARN' for r in results)} · PASS {sum(r['result'] == 'PASS' for r in results)}")
    return 1 if any(r["result"] == "FAIL" for r in results) else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
