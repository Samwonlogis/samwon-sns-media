"""클라우드 백업 게시용 암호화(표준 라이브러리만 사용).
SHAKE-256 키스트림 + HMAC-SHA256(암호화 후 인증). 키는 .env SAMWON_CLOUD_KEY(64자리 16진수)."""
import hashlib
import hmac
import os

MAGIC = b"SAMWON1"


def _keys(key_hex: str):
    k = bytes.fromhex(key_hex.strip())
    return hashlib.sha256(k + b"enc").digest(), hashlib.sha256(k + b"mac").digest()


def _xor(data: bytes, ks: bytes) -> bytes:
    n = len(data)
    return (int.from_bytes(data, "big") ^ int.from_bytes(ks, "big")).to_bytes(n, "big") if n else b""


def seal(key_hex: str, data: bytes) -> bytes:
    ek, mk = _keys(key_hex)
    nonce = os.urandom(16)
    ct = _xor(data, hashlib.shake_256(ek + nonce).digest(len(data)))
    mac = hmac.new(mk, MAGIC + nonce + ct, hashlib.sha256).digest()
    return MAGIC + nonce + mac + ct


def open_sealed(key_hex: str, blob: bytes) -> bytes:
    ek, mk = _keys(key_hex)
    if not blob.startswith(MAGIC):
        raise ValueError("형식이 다른 파일입니다")
    nonce, mac, ct = blob[7:23], blob[23:55], blob[55:]
    if not hmac.compare_digest(mac, hmac.new(mk, MAGIC + nonce + ct, hashlib.sha256).digest()):
        raise ValueError("키가 다르거나 파일이 손상되었습니다")
    return _xor(ct, hashlib.shake_256(ek + nonce).digest(len(ct)))
