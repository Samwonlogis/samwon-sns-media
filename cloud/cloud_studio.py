"""클라우드 SNS 제작 스튜디오 — PC가 꺼져 있어도 카드뉴스를 만들고 게시 대기열에 올린다 (2026-11-09).

클라우드(예약 작업)에서:
    SAMWON_CLOUD_KEY=<키> python3 cloud_studio.py fetch [저장소] [브랜치]
        → 암호화 묶음(bundle.bin: .env·설정·스크립트·문서·대기열)과 자산 묶음(studio.bin: 사진·폰트·로고)을
          ./samwon-sns 로 푼다. 이후 그 폴더에서 SKILL.md 절차대로 게시물을 만든다.
        (fetch는 노션에 '대기' 중인 클라우드 제작물도 작업 폴더 queue/로 함께 가져와 날짜·주제 중복을 막는다)
    cd samwon-sns && python3 scripts/cloud_studio.py push queue/<폴더> [queue/<폴더> ...]
        → 노션 등록(미리보기 포함) → 암호화 패키지(post.json·슬라이드·이력)를 노션 페이지에 첨부
          → 노션 '클라우드 제작' = '대기'. ※ 클라우드는 GitHub 쓰기가 막혀 있어 이미지 공개 업로드는 PC가 한다.

PC(cloud_sync.py가 매시 자동 호출)에서:
    python scripts/cloud_studio.py import
        → 노션의 '대기' 패키지를 queue/ 로 가져오고(이미 있으면 건너뜀) 이력을 합친 뒤 '가져옴'으로 표시한다.
          이어서 cloud_sync가 이미지 공개 업로드·게시 묶음 갱신을 하므로, 다음 클라우드 게시 회차에 승인 요청이 나간다.
    python scripts/cloud_studio.py studio [--force]
        → 자산 묶음(studio.bin)을 자산이 바뀌었을 때만 다시 올린다.

보안: 문안·토큰은 cloud_seal로 암호화한다. 노션 첨부 패키지도 암호화되어 있고, 공개 저장소에는 게시용 이미지(sns/<id>/)만 평문으로 둔다.
"""
import base64
import hashlib
import io
import json
import os
import sys
import time
import zipfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent

# ---- 묶음 구성 (PC·클라우드 공통) ----
CODE = ["publish.py", "common.py", "notion_sync.py", "hashtags.py", "refresh_tokens.py", "cloud_seal.py",
        "cloud_tasks.py", "cloud_studio.py", "render.py", "sns_queue.py", "sns_check.py"]
DOCS = ["SKILL.md", "config.json", "templates/README.md", "data/facts.md", "data/history.json"]
STUDIO_GLOBS = ["assets/*.svg", "assets/logo*.png", "assets/kakao_channel_qr.png", "assets/SAMWON_symbol*.png", "assets/RIGHTS.md",
                "assets/fonts/*.woff2", "assets/photos/library.json", "assets/photos/LICENSE.txt",
                "assets/photos/lib/*.jpg", "assets/characters/**/*"]


def _env():
    sys.path.insert(0, str(HERE))
    from common import load_env
    return load_env()


def _gh(env):
    import requests
    repo, branch = env["GITHUB_MEDIA_REPO"], env.get("GITHUB_MEDIA_BRANCH", "main")
    h = {"Authorization": f"Bearer {env['GITHUB_TOKEN']}", "Accept": "application/vnd.github+json"}
    api = f"https://api.github.com/repos/{repo}/contents"

    def get(path):
        r = requests.get(f"{api}/{path}", headers=h, params={"ref": branch}, timeout=120)
        return r.json() if r.status_code == 200 else None

    def put(path, data: bytes, msg):
        cur = get(path)
        body = {"message": msg, "content": base64.b64encode(data).decode(), "branch": branch}
        if cur and isinstance(cur, dict):
            body["sha"] = cur["sha"]
        for i in range(3):
            r = requests.put(f"{api}/{path}", headers=h, json=body, timeout=300)
            if r.status_code in (200, 201):
                return
            if r.status_code == 409 and i < 2:          # 동시 갱신 충돌 → sha 다시 받아 재시도
                cur = get(path)
                if cur:
                    body["sha"] = cur["sha"]
                time.sleep(3)
                continue
            raise RuntimeError(f"GitHub 업로드 실패 {path}: {r.status_code} {r.text[:200]}")

    def delete(path, msg):
        cur = get(path)
        if cur and isinstance(cur, dict):
            requests.delete(f"{api}/{path}", headers=h, timeout=120,
                            json={"message": msg, "sha": cur["sha"], "branch": branch})

    def raw(path) -> bytes:
        r = requests.get(f"{api}/{path}", headers={**h, "Accept": "application/vnd.github.raw"},
                         params={"ref": branch}, timeout=300)
        r.raise_for_status()
        return r.content

    return get, put, delete, raw


def _zip(files: dict) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for arc in sorted(files):
            zi = zipfile.ZipInfo(arc, date_time=(2026, 1, 1, 0, 0, 0))
            zi.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(zi, files[arc])
    return buf.getvalue()


def bundle_files(root: Path, picked) -> dict:
    """게시 묶음(bundle.bin)에 들어갈 파일. picked = [(folder, post), ...] 게시 대기 게시물."""
    files = {".env": (root / ".env").read_bytes()}
    for d in DOCS:
        if (root / d).exists():
            files[d] = (root / d).read_bytes()
    for f in (root / "references").glob("*.md"):
        if "-DESKTOP" not in f.name and "-LAPTOP" not in f.name:
            files[f"references/{f.name}"] = f.read_bytes()
    for f in CODE:
        if (root / "scripts" / f).exists():
            files[f"scripts/{f}"] = (root / "scripts" / f).read_bytes()
    for folder, _ in picked:
        files[f"queue/{folder.name}/post.json"] = (folder / "post.json").read_bytes()
    return files


def studio_files(root: Path) -> dict:
    files = {}
    for g in STUDIO_GLOBS:
        for f in root.glob(g):
            if f.is_file() and ".bak" not in f.name:
                files[f.relative_to(root).as_posix()] = f.read_bytes()
    return files


# ---------------- 클라우드 쪽 ----------------

def fetch(repo="Samwonlogis/samwon-sns-media", branch="main"):
    import urllib.request
    key = os.environ["SAMWON_CLOUD_KEY"]
    base = f"https://raw.githubusercontent.com/{repo}/{branch}/cloud"
    work = Path.cwd() / "samwon-sns"
    work.mkdir(exist_ok=True)
    sys.path.insert(0, str(HERE))
    from cloud_seal import open_sealed

    def get(name):
        with urllib.request.urlopen(f"{base}/{name}?t={int(time.time())}", timeout=300) as r:
            return r.read()

    for name in ("studio.bin", "bundle.bin"):
        zipfile.ZipFile(io.BytesIO(open_sealed(key, get(name)))).extractall(work)
    meta = json.loads(get("bundle_meta.json"))
    try:
        os.chdir(work)
        sys.path.insert(0, str(work / "scripts"))
        import cloud_studio as cs  # 묶음 안의 최신 버전
        k = cs.pull_pending(work, mark=False)
        if k:
            print(f"노션 대기 중인 클라우드 제작물 {k}건을 작업 폴더에 반영(중복 제작 방지)")
    except Exception as e:  # noqa: BLE001
        print(f"노션 대기분 확인 실패(제작 전 노션 'SNS 콘텐츠 목록'을 직접 확인할 것): {e}")
    print(f"스튜디오 준비 완료: {work} · 묶음 생성 {meta.get('made_at')} · 게시 대기 {meta.get('posts')}건")


# ---------------- 노션 수신함 (클라우드 → PC) ----------------
INBOX_PROP = "클라우드 제작"


def _notion():
    sys.path.insert(0, str(HERE))
    import notion_sync as n
    return n


def _ensure_prop(env):
    import requests
    n = _notion()
    requests.patch(f"{n.API}/databases/{env['NOTION_DATABASE_ID']}", headers=n.H(env), timeout=60,
                   json={"properties": {INBOX_PROP: {"rich_text": {}}}})


def _upload_zip(env, name: str, data: bytes) -> str:
    import requests
    n = _notion()
    r = requests.post(f"{n.API}/file_uploads", headers=n.H(env), timeout=60,
                      json={"filename": name, "content_type": "application/zip"})
    r.raise_for_status()
    fid = r.json()["id"]
    r = requests.post(f"{n.API}/file_uploads/{fid}/send", headers=n.H(env, json_ct=False), timeout=300,
                      files={"file": (name, data, "application/zip")})
    r.raise_for_status()
    return fid


def _set_inbox(env, page_id: str, value: str):
    import requests
    n = _notion()
    requests.patch(f"{n.API}/pages/{page_id}", headers=n.H(env), timeout=60,
                   json={"properties": {INBOX_PROP: {"rich_text": n._rt(value)}}}).raise_for_status()


def push(folders):
    """클라우드에서 만든 게시물을 노션에 등록하고 암호화 패키지를 첨부한다(이미지 공개 업로드·묶음 갱신은 PC가 한다)."""
    import requests
    os.environ.pop("SAMWON_CLOUD", None)
    env = _env()
    sys.path.insert(0, str(HERE))
    from cloud_seal import seal
    from common import load_post, now, save_post
    n = _notion()
    _ensure_prop(env)
    for f in folders:
        folder = Path(f) if Path(f).is_absolute() else ROOT / f
        post = load_post(folder)
        post["made_by"] = "cloud"
        save_post(folder, post)
        os.system(f'"{sys.executable}" "{HERE / "sns_queue.py"}" log "{folder}"')
        n.add(env, folder)
        page = n.find_page(env, post["id"])
        if not page:
            raise RuntimeError(f"노션 등록 확인 실패: {post['id']}")
        pkg = {f"queue/{folder.name}/{p.name}": p.read_bytes() for p in folder.iterdir() if p.is_file()}
        pkg["data/history.json"] = (ROOT / "data" / "history.json").read_bytes()
        fid = _upload_zip(env, f"{folder.name}.zip", seal(env["SAMWON_CLOUD_KEY"], _zip(pkg)))
        requests.patch(f"{n.API}/blocks/{page['id']}/children", headers=n.H(env), timeout=60, json={"children": [
            n.h2("클라우드 제작 패키지(암호화 · PC 자동 가져오기용, 삭제 금지)"),
            {"type": "file", "file": {"type": "file_upload", "file_upload": {"id": fid}}}]}).raise_for_status()
        _set_inbox(env, page["id"], f"대기 {now():%Y-%m-%d %H:%M}")
        print(f"클라우드 제작 등록: {post['id']} → 노션(검토 대기) · PC 가져오기 대기")


def pull_pending(dest: Path, mark: bool) -> int:
    """노션에서 '대기' 중인 클라우드 제작 패키지를 dest/queue 로 풀고 이력을 합친다. mark=True면 '가져옴' 표시(PC 전용)."""
    import requests
    env = _env()
    if not env.get("NOTION_TOKEN") or not env.get("SAMWON_CLOUD_KEY"):
        return 0
    sys.path.insert(0, str(HERE))
    from cloud_seal import open_sealed
    from common import log, now
    n = _notion()
    _ensure_prop(env)
    r = requests.post(f"{n.API}/databases/{env['NOTION_DATABASE_ID']}/query", headers=n.H(env), timeout=60,
                      json={"filter": {"property": INBOX_PROP, "rich_text": {"starts_with": "대기"}}})
    r.raise_for_status()
    count = 0
    hist_f = dest / "data" / "history.json"
    for page in r.json().get("results", []):
        blocks = requests.get(f"{n.API}/blocks/{page['id']}/children?page_size=100", headers=n.H(env), timeout=60).json()
        url = next((b["file"]["file"]["url"] for b in blocks.get("results", [])
                    if b.get("type") == "file" and b["file"].get("type") == "file" and ".zip" in b["file"]["file"]["url"]), None)
        if not url:
            log(f"클라우드 패키지 첨부를 찾지 못함: {page['id']}")
            continue
        try:
            z = zipfile.ZipFile(io.BytesIO(open_sealed(env["SAMWON_CLOUD_KEY"], requests.get(url, timeout=300).content)))
        except Exception as e:  # noqa: BLE001
            log(f"클라우드 패키지 열기 실패(다음 회차 재시도): {page['id']} — {e}")
            continue
        names = {Path(a).parts[1] for a in z.namelist() if a.startswith("queue/")}
        for name in names:
            out = dest / "queue" / name
            if (out / "post.json").exists():
                log(f"클라우드 패키지: 이미 대기열에 있음 → 건너뜀 {name}")
                continue
            out.mkdir(parents=True, exist_ok=True)
            for arc in z.namelist():
                if arc.startswith(f"queue/{name}/"):
                    (out / Path(arc).name).write_bytes(z.read(arc))
            log(f"클라우드 제작 게시물 가져옴: {name}")
        if "data/history.json" in z.namelist():
            mine = json.loads(hist_f.read_text(encoding="utf-8")) if hist_f.exists() else []
            merged = {h["id"]: h for h in mine + json.loads(z.read("data/history.json"))}
            hist_f.write_text(json.dumps(sorted(merged.values(), key=lambda h: h["date"])[-60:], ensure_ascii=False, indent=2), encoding="utf-8")
        if mark:
            _set_inbox(env, page["id"], f"가져옴 {now():%Y-%m-%d %H:%M}")
        count += 1
    return count


def import_inbox():
    """PC: 노션의 클라우드 제작 패키지를 queue/로 가져온다(cloud_sync가 매시 호출)."""
    return pull_pending(ROOT, mark=True)


def upload_studio(force=False):
    env = _env()
    sys.path.insert(0, str(HERE))
    from cloud_seal import seal
    from common import log
    files = studio_files(ROOT)
    digest = hashlib.sha256(b"".join(k.encode() + hashlib.sha256(v).digest() for k, v in sorted(files.items()))).hexdigest()
    state_f = ROOT / "data" / "cloud_sync_state.json"
    state = json.loads(state_f.read_text(encoding="utf-8")) if state_f.exists() else {}
    if state.get("studio") == digest and not force:
        return False
    _, put, _, _ = _gh(env)
    put("cloud/studio.bin", seal(env["SAMWON_CLOUD_KEY"], _zip(files)), "cloud: studio assets")
    put("cloud/cloud_studio.py", (HERE / "cloud_studio.py").read_bytes(), "cloud: cloud_studio.py")  # 클라우드 부트스트랩용
    put("cloud/cloud_seal.py", (HERE / "cloud_seal.py").read_bytes(), "cloud: cloud_seal.py")
    state["studio"] = digest
    state_f.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")
    log(f"클라우드 스튜디오 자산 업로드: {len(files)}개 파일")
    return True


if __name__ == "__main__":
    a = sys.argv[1:] or ["import"]
    if a[0] == "fetch":
        fetch(*a[1:3])
    elif a[0] == "push":
        folders = [x for x in a[1:] if x != "--force"]
        if "--force" not in a:   # 과제4: 자동 검사 FAIL이면 노션 등록 차단
            import subprocess
            r = subprocess.run([sys.executable, str(HERE / "sns_check.py"), *folders], cwd=str(ROOT))
            if r.returncode != 0:
                sys.exit("등록 중단: sns_check FAIL 항목을 고친 뒤 다시 push 하십시오(대표 승인 시에만 --force)")
        push(folders)
    elif a[0] == "import":
        print(f"노션 클라우드 제작물 가져오기: {import_inbox()}건")
    elif a[0] == "studio":
        print("업로드함" if upload_studio("--force" in a) else "변경 없음")
    else:
        sys.exit(__doc__)
