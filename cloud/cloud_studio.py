"""클라우드 SNS 제작 스튜디오 — PC가 꺼져 있어도 카드뉴스를 만들고 게시 대기열에 올린다 (2026-11-09).

클라우드(예약 작업)에서:
    SAMWON_CLOUD_KEY=<키> python3 cloud_studio.py fetch [저장소] [브랜치]
        → 암호화 묶음(bundle.bin: .env·설정·스크립트·문서·대기열)과 자산 묶음(studio.bin: 사진·폰트·로고)을
          ./samwon-sns 로 푼다. 이후 그 폴더에서 SKILL.md 절차대로 게시물을 만든다.
    cd samwon-sns && python3 scripts/cloud_studio.py push queue/<폴더> [queue/<폴더> ...]
        → 이미지 공개 URL 업로드 → 노션 등록 → 게시 이력 기록 → 암호화 수신함(cloud/inbox/)에 보관
          → 게시 묶음(bundle.bin) 재생성. 클라우드 게시 작업이 바로 이 게시물을 승인 요청·게시한다.

PC(cloud_sync.py가 매시 자동 호출)에서:
    python scripts/cloud_studio.py import
        → 수신함의 게시물을 queue/ 로 가져오고(이미 있으면 건너뜀) 게시 이력을 합친 뒤 수신함을 비운다.
    python scripts/cloud_studio.py studio [--force]
        → 자산 묶음(studio.bin)을 자산이 바뀌었을 때만 다시 올린다.

보안: 모든 문안·토큰은 cloud_seal로 암호화해 올린다. 공개 저장소에는 게시용 이미지(sns/<id>/)만 평문으로 둔다.
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
        "cloud_tasks.py", "cloud_studio.py", "render.py", "sns_queue.py"]
DOCS = ["SKILL.md", "config.json", "templates/README.md", "data/facts.md", "data/history.json"]
STUDIO_GLOBS = ["assets/*.svg", "assets/logo*.png", "assets/SAMWON_symbol*.png", "assets/RIGHTS.md",
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
    print(f"스튜디오 준비 완료: {work} · 묶음 생성 {meta.get('made_at')} · 게시 대기 {meta.get('posts')}건")


def push(folders):
    """클라우드에서 만든 게시물을 공개 이미지 업로드 → 노션 등록 → 이력 기록 → 수신함 보관 → 게시 묶음 재생성."""
    os.environ.pop("SAMWON_CLOUD", None)                  # 이미지 업로드를 허용(게시 모드 아님)
    env = _env()
    sys.path.insert(0, str(HERE))
    from cloud_seal import seal
    from common import iter_posts, load_post, now, save_post
    import notion_sync
    import publish
    get, put, _, _ = _gh(env)
    key = env["SAMWON_CLOUD_KEY"]
    index = (get("cloud/inbox/index.json") or {})
    idx = json.loads(base64.b64decode(index["content"])) if index.get("content") else []
    for f in folders:
        folder = Path(f) if Path(f).is_absolute() else ROOT / f
        post = load_post(folder)
        publish.host_images.no_verify = True
        publish.host_images(folder, post, env)
        post.setdefault("made_by", "cloud")
        save_post(folder, post)
        notion_sync.add(env, folder)
        os.system(f'{sys.executable} "{HERE / "sns_queue.py"}" log "{folder}"')
        pkg = {f"queue/{folder.name}/{p.name}": p.read_bytes() for p in folder.iterdir() if p.is_file()}
        pkg["data/history.json"] = (ROOT / "data" / "history.json").read_bytes()
        put(f"cloud/inbox/{folder.name}.bin", seal(key, _zip(pkg)), f"inbox: {folder.name}")
        if folder.name not in idx:
            idx.append(folder.name)
        print(f"클라우드 제작 등록: {post['id']} → 노션·수신함")
    put("cloud/inbox/index.json", json.dumps(idx, ensure_ascii=False).encode(), "inbox: index")
    # 게시 묶음 재생성(새 게시물 포함) — 클라우드 게시 작업이 바로 승인 요청을 보낼 수 있게
    from datetime import timedelta, datetime
    cutoff = now() - timedelta(days=10)
    picked = [(fd, p) for fd, p in iter_posts()
              if p.get("status") not in ("published", "cancelled", "rejected")
              and (not p.get("scheduled_at") or datetime.fromisoformat(p["scheduled_at"]) >= cutoff)]
    put("cloud/bundle.bin", seal(key, _zip(bundle_files(ROOT, picked))), f"cloud: bundle (studio) {now():%m%d %H%M}")
    put("cloud/bundle_meta.json", json.dumps({"made_at": now().isoformat(), "posts": len(picked), "by": "cloud-studio"}).encode(), "cloud: meta")
    print(f"게시 묶음 재생성: 게시 대기 {len(picked)}건")


# ---------------- PC 쪽 ----------------

def import_inbox():
    env = _env()
    if not env.get("GITHUB_TOKEN"):
        return 0
    sys.path.insert(0, str(HERE))
    from cloud_seal import open_sealed
    from common import log
    get, put, delete, raw = _gh(env)
    index = get("cloud/inbox/index.json")
    idx = json.loads(base64.b64decode(index["content"])) if index and index.get("content") else []
    if not idx:
        return 0
    done = []
    hist_f = ROOT / "data" / "history.json"
    for name in idx:
        try:
            z = zipfile.ZipFile(io.BytesIO(open_sealed(env["SAMWON_CLOUD_KEY"], raw(f"cloud/inbox/{name}.bin"))))
        except Exception as e:  # noqa: BLE001
            log(f"수신함 가져오기 실패(다음 회차 재시도): {name} — {e}")
            continue
        dest = ROOT / "queue" / name
        if dest.exists() and (dest / "post.json").exists():
            log(f"수신함: 이미 대기열에 있음 → 건너뜀 {name}")
        else:
            dest.mkdir(parents=True, exist_ok=True)
            for arc in z.namelist():
                if arc.startswith(f"queue/{name}/"):
                    (dest / Path(arc).name).write_bytes(z.read(arc))
            log(f"수신함: 클라우드 제작 게시물 가져옴 {name}")
        if "data/history.json" in z.namelist():
            mine = json.loads(hist_f.read_text(encoding="utf-8")) if hist_f.exists() else []
            theirs = json.loads(z.read("data/history.json"))
            merged = {h["id"]: h for h in mine + theirs}
            hist_f.write_text(json.dumps(sorted(merged.values(), key=lambda h: h["date"])[-60:], ensure_ascii=False, indent=2), encoding="utf-8")
        delete(f"cloud/inbox/{name}.bin", f"inbox: imported {name}")
        done.append(name)
    rest = [n for n in idx if n not in done]
    put("cloud/inbox/index.json", json.dumps(rest, ensure_ascii=False).encode(), "inbox: index")
    return len(done)


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
    put("cloud/cloud_studio.py", (HERE / "cloud_studio.py").read_bytes(), "cloud: cloud_studio.py")
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
        push(a[1:])
    elif a[0] == "import":
        print(f"수신함 가져오기: {import_inbox()}건")
    elif a[0] == "studio":
        print("업로드함" if upload_studio("--force" in a) else "변경 없음")
    else:
        sys.exit(__doc__)
