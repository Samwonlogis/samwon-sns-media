#!/usr/bin/env python3
"""삼원 PDF 메일 발송기 (samwon-standard, 과제5 2026-11-09).

admin@3won.kr 명의 발송 통로(Apps Script)로 PDF 등 첨부 메일을 보낸다.
주소·키는 지시문에 적지 않는다 — 실행 때 '2611총무_삼원_기준정보' 시트 발송설정 탭에서 읽어 환경변수로 넘긴다.

  SAMWON_MAIL_URL=... SAMWON_MAIL_KEY=... python3 samwon_mail.py \
      --to "a@3won.kr,b@3won.kr" --subject "[삼원의 오늘] 2026.11.10 (화) …" \
      --headline "오늘의 헤드라인 한 줄" [--summary "핵심 3줄(줄바꿈 \\n)"] [--links-html "<ul>…</ul>"] \
      --from-name "삼원종합물류" --attach "20261110_삼원의 오늘.pdf" [--attach 양식.hwp ...] [--dry-run]

안전장치: 수신자는 @3won.kr 와 jong1231@gmail.com 만 허용(그 밖이면 발송하지 않고 종료 코드 2).
첨부 합계 20MB 초과 시 PDF만 남긴다. 실패하면 1회 재시도하고 결과를 JSON 한 줄로 출력한다(키 값은 출력하지 않는다).
"""
import argparse
import base64
import html
import json
import os
import re
import subprocess
import sys
import tempfile
import time

ALLOWED = re.compile(r"^[A-Za-z0-9._%+-]+@3won\.kr$|^jong1231@gmail\.com$")
MIME = {".pdf": "application/pdf", ".hwp": "application/x-hwp", ".hwpx": "application/haansofthwpx",
        ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        ".pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        ".zip": "application/zip", ".png": "image/png", ".jpg": "image/jpeg"}
LIMIT = 20 * 1024 * 1024


def body_html(headline, summary, links_html, from_name, org="삼원종합물류(주)"):
    lines = "<br>".join(html.escape(x) for x in (summary or "").splitlines() if x.strip())
    return (
        "<div style=\"font-family:'Malgun Gothic',sans-serif;font-size:14px;color:#0E1530;line-height:1.65;max-width:680px\">"
        "<div style='border-bottom:3px solid #00489C;padding-bottom:6px;margin-bottom:12px;color:#282B6E;font-weight:700'>"
        + ("SAMWON | 삼원종합물류(주)" if org == "삼원종합물류(주)" else html.escape(org)) + "</div>"
        f"안녕하십니까, {html.escape(org)}입니다.<br>첨부한 파일을 확인해 주십시오.<br><br>"
        + (f"<b style='color:#282B6E'>핵심</b><br>{html.escape(headline)}<br><br>" if headline else "")
        + (f"{lines}<br><br>" if lines else "")
        + (f"<b style='color:#282B6E'>바로가기</b>{links_html}<br>" if links_html else "")
        + f"<span style='color:#6E7891;font-size:12px'>{html.escape(from_name)}</span></div>")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--to", required=True)
    ap.add_argument("--subject", required=True)
    ap.add_argument("--headline", default="")
    ap.add_argument("--summary", default="")
    ap.add_argument("--links-html", default="")
    ap.add_argument("--from-name", default="삼원종합물류")
    ap.add_argument("--org", default="삼원종합물류(주)", help="머리글·인사말 기관명(개인 임대업 등은 '아트리트21 관리')")
    ap.add_argument("--attach", action="append", default=[])
    ap.add_argument("--dry-run", action="store_true")
    a = ap.parse_args()

    to = [x.strip() for x in a.to.replace(";", ",").split(",") if x.strip()]
    bad = [x for x in to if not ALLOWED.match(x)]
    if bad or not to:
        print(json.dumps({"ok": False, "error": "허용되지 않은 수신자", "bad": bad}, ensure_ascii=False))
        sys.exit(2)
    url, key = os.environ.get("SAMWON_MAIL_URL", ""), os.environ.get("SAMWON_MAIL_KEY", "")
    if not (url.startswith("https://script.google.com/") and key):
        print(json.dumps({"ok": False, "error": "SAMWON_MAIL_URL/KEY 없음(기준정보 시트 발송설정 탭에서 읽을 것)"}, ensure_ascii=False))
        sys.exit(3)

    files = [f for f in a.attach if os.path.exists(f)]
    missing = [f for f in a.attach if not os.path.exists(f)]
    if sum(os.path.getsize(f) for f in files) > LIMIT:
        files = [f for f in files if f.lower().endswith(".pdf")]
    payload = {"key": key, "to": to, "subject": a.subject,
               "text": (a.headline + "\n" + a.summary).strip() or "첨부 파일을 확인해 주십시오.",
               "html": body_html(a.headline, a.summary, a.links_html, a.from_name, a.org), "fromName": a.from_name,
               "files": [{"name": os.path.basename(f), "mimeType": MIME.get(os.path.splitext(f)[1].lower(), "application/octet-stream"),
                          "base64": base64.b64encode(open(f, "rb").read()).decode()} for f in files]}
    if a.dry_run:
        print(json.dumps({"ok": True, "dry_run": True, "to": len(to), "files": [os.path.basename(f) for f in files],
                          "missing": missing}, ensure_ascii=False))
        return
    fd, path = tempfile.mkstemp(suffix=".json")
    with os.fdopen(fd, "w") as fh:
        fh.write(json.dumps(payload, ensure_ascii=False))
    out = ""
    for i in range(2):
        r = subprocess.run(["curl", "-sSL", "-m", "240", "-H", "Content-Type: application/json", "--data-binary", f"@{path}", url],
                           capture_output=True, text=True)
        out = (r.stdout or r.stderr)[:400]
        if '"ok":true' in out.replace(" ", ""):
            break
        time.sleep(10)
    os.remove(path)
    ok = '"ok":true' in out.replace(" ", "")
    print(json.dumps({"ok": ok, "to": len(to), "files": [os.path.basename(f) for f in files], "missing": missing,
                      "response": out}, ensure_ascii=False))
    sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()
