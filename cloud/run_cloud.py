"""클라우드 백업 게시 실행기 — PC가 꺼져 있어도 노션 '검토 완료'(핸드폰 승인) 게시물을 올린다.

    SAMWON_CLOUD_KEY=<키> python3 run_cloud.py [저장소] [브랜치]

암호화된 묶음(cloud/bundle.bin)을 내려받아 풀고, scripts/publish.py 를 클라우드 모드로 실행한다.
"""
import io
import json
import os
import subprocess
import sys
import time
import urllib.request
import zipfile
from pathlib import Path

repo = sys.argv[1] if len(sys.argv) > 1 else "Samwonlogis/samwon-sns-media"
branch = sys.argv[2] if len(sys.argv) > 2 else "main"
key = os.environ["SAMWON_CLOUD_KEY"]
base = f"https://raw.githubusercontent.com/{repo}/{branch}/cloud"
work = Path.cwd() / "samwon-sns"


def get(name):
    with urllib.request.urlopen(f"{base}/{name}?t={int(time.time())}", timeout=120) as r:
        return r.read()


blob = get("bundle.bin")
meta = json.loads(get("bundle_meta.json"))
print(f"묶음 생성 시각: {meta.get('made_at')} · 게시 대기 {meta.get('posts')}건")

# 묶음 안의 cloud_seal.py 로 풀어야 하므로, 먼저 동일 알고리즘으로 복호화
import hashlib, hmac  # noqa: E401,E402
k = bytes.fromhex(key.strip())
ek, mk = hashlib.sha256(k + b"enc").digest(), hashlib.sha256(k + b"mac").digest()
assert blob.startswith(b"SAMWON1"), "형식 오류"
nonce, mac, ct = blob[7:23], blob[23:55], blob[55:]
if not hmac.compare_digest(mac, hmac.new(mk, b"SAMWON1" + nonce + ct, hashlib.sha256).digest()):
    sys.exit("복호화 실패: 키가 다르거나 파일 손상")
ks = hashlib.shake_256(ek + nonce).digest(len(ct))
plain = (int.from_bytes(ct, "big") ^ int.from_bytes(ks, "big")).to_bytes(len(ct), "big")
work.mkdir(exist_ok=True)
zipfile.ZipFile(io.BytesIO(plain)).extractall(work)

try:
    import requests  # noqa: F401
except ImportError:
    subprocess.run([sys.executable, "-m", "pip", "install", "-q", "--break-system-packages", "requests"], check=False)

env = dict(os.environ, SAMWON_CLOUD="1", SAMWON_HOST="cloud", PYTHONIOENCODING="utf-8")
r = subprocess.run([sys.executable, "scripts/publish.py"], cwd=work, env=env, capture_output=True, text=True, timeout=1500)
out = (r.stdout + r.stderr).strip()
print(out[-6000:] if out else "(출력 없음)")
sys.exit(r.returncode)
