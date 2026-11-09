#!/usr/bin/env python3
"""쪽 사용률 측정(samwon-standard) — 문서를 PDF로 바꿔 쪽마다 본문 영역이 얼마나 찼는지 잰다.

  fills = page_fills("문서.docx")      # [(쪽, 사용률 0~1, 글자 수), ...]  (LibreOffice 없으면 None)
  python3 samwon_layout.py <파일...>    # 쪽별 사용률 출력

사용률 = 본문 영역(머리글·바닥글 제외) 맨 위에서 마지막 잉크(글자·선·칸)까지의 높이 비율.
"""
import os
import re
import shutil
import subprocess
import sys
import tempfile

SOFFICE = shutil.which("soffice") or shutil.which("libreoffice")
PDFTOPPM = shutil.which("pdftoppm")


def to_pdf(path, outdir=None):
    if not SOFFICE:
        return None
    outdir = outdir or tempfile.mkdtemp(prefix="samwon_pdf_")
    try:
        subprocess.run([SOFFICE, "--headless", "--convert-to", "pdf", "--outdir", outdir, path],
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=180)
    except Exception:
        return None
    pdf = os.path.join(outdir, os.path.splitext(os.path.basename(path))[0] + ".pdf")
    return pdf if os.path.exists(pdf) else None


def _page_texts(pdf):
    try:
        from pypdf import PdfReader
        return [(p.extract_text() or "") for p in PdfReader(pdf).pages]
    except Exception:
        return []


def _raster_fills(pdf, top_cm=2.05, bottom_cm=1.75):
    """pdftoppm으로 저해상도 이미지를 만들어 본문 영역의 마지막 잉크 위치를 찾는다."""
    if not PDFTOPPM:
        return None
    from PIL import Image
    tmp = tempfile.mkdtemp(prefix="samwon_fill_")
    subprocess.run([PDFTOPPM, "-r", "30", "-gray", "-png", pdf, os.path.join(tmp, "p")],
                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=180)
    files = sorted((f for f in os.listdir(tmp) if f.endswith(".png")), key=lambda s: int(re.findall(r"\d+", s)[-1]))
    out = []
    for f in files:
        im = Image.open(os.path.join(tmp, f)).convert("L")
        W, H = im.size
        landscape = W > H
        cm = H / (21.0 if landscape else 29.7)
        y0, y1 = int(top_cm * cm), int(H - bottom_cm * cm)
        px = im.load()
        last = y0
        for y in range(y1 - 1, y0, -1):
            if any(px[x, y] < 200 for x in range(int(W * 0.06), int(W * 0.94), 2)):
                last = y; break
        out.append(max(0.0, min(1.0, (last - y0) / max(1, (y1 - y0)))))
    shutil.rmtree(tmp, ignore_errors=True)
    return out


def page_fills(path, keep_pdf=False):
    pdf = path if path.lower().endswith(".pdf") else to_pdf(path)
    if not pdf:
        return None
    fills = _raster_fills(pdf)
    texts = _page_texts(pdf)
    if fills is None:
        fills = [min(1.0, len(t) / 1400) for t in texts]
    res = [(i + 1, round(f, 3), len(texts[i]) if i < len(texts) else 0) for i, f in enumerate(fills)]
    if keep_pdf:
        return res, pdf, texts
    return res


def waste_findings(path, has_cover=True, kind=None):
    """쪽 낭비 판정 목록 [(쪽, 사용률, 사유)]. kind: docx/xlsx"""
    r = page_fills(path, keep_pdf=True)
    if not r:
        return None
    fills, pdf, texts = r
    kind = kind or os.path.splitext(path)[1].lower().lstrip(".")
    out = []
    n = len(fills)
    start = 1 if has_cover else 0
    if has_cover and len(texts) >= 2:
        p1, p2 = re.sub(r"\s+", "", texts[0]), re.sub(r"\s+", "", texts[1])
        if (("주관부서" in p2 and "주관부서" not in p1) or ("최종결재자" in p2 and "최종결재자" not in p1)):
            out.append((2, fills[1][1], "표지가 2쪽으로 넘침"))
    for idx in range(start, n):
        pno, f, nchar = fills[idx]
        last = idx == n - 1
        if kind == "docx":
            if last and n - start >= 2 and f < 0.15:
                out.append((pno, f, "마지막 쪽이 거의 비어 있음"))
            elif not last and f < 0.40:
                out.append((pno, f, "쪽 중간에서 끊김(아래 절반 이상 비어 있음)"))
        elif kind == "xlsx":
            body = texts[idx] if idx < len(texts) else ""
            lines = [l for l in body.splitlines() if l.strip()]
            if len(lines) <= 3 and n - start >= 2:
                out.append((pno, f, "빈 기입 행·머리글만 있는 쪽"))
    return out


if __name__ == "__main__":
    for a in sys.argv[1:]:
        r = page_fills(a)
        print(a)
        if r is None:
            print("  (LibreOffice 없음 — 측정 생략)"); continue
        for p, f, n in r:
            print(f"  {p:>2}쪽  사용률 {f * 100:5.1f}%  글자 {n}")
