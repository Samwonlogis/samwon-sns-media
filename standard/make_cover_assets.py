#!/usr/bin/env python3
"""표지 배경 이미지(하단 네이비 띠 + 블루 선 + 회사 정보) 생성 — 한 번 만들어 저장소에 넣어 둔다.
   python3 make_cover_assets.py  → cover_bg_portrait.png, cover_bg_landscape.png, cover_bg_slide.png
"""
import os

from PIL import Image, ImageDraw, ImageFont

HERE = os.path.dirname(os.path.abspath(__file__))
DEEP, BLUE, WHITE, SOFT = (0x14, 0x18, 0x50), (0x00, 0x48, 0x9C), (255, 255, 255), (0xC9, 0xD3, 0xEA)
REG = ("/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc", 1)
BOLD = ("/usr/share/fonts/opentype/noto/NotoSansCJK-Bold.ttc", 1)
COMPANY = "삼원종합물류(주)"
INFO = "서울특별시 서초구 효령로 328 아트리트21 6층 (06721)   ·   TEL 02-598-3001   ·   FAX 02-597-2564   ·   www.3won.kr"


def _f(spec, px):
    return ImageFont.truetype(spec[0], px, index=spec[1])


def make(path, w_cm, h_cm, band_cm, left_cm, dpi=200):
    px = lambda cm: int(round(cm / 2.54 * dpi))
    W, H = px(w_cm), px(h_cm)
    im = Image.new("RGB", (W, H), WHITE); d = ImageDraw.Draw(im)
    band = px(band_cm); stripe = max(2, px(0.16))
    d.rectangle([0, H - band, W, H], fill=DEEP)
    d.rectangle([0, H - band - stripe, W, H - band], fill=BLUE)
    d.rectangle([0, 0, W, max(2, px(0.12))], fill=BLUE)            # 상단 가는 블루 선
    fb, fr = _f(BOLD, px(0.36)), _f(REG, px(0.27))
    x = px(left_cm); cy = H - band // 2
    d.text((x, cy - px(0.05)), COMPANY, font=fb, fill=WHITE, anchor="ls")
    d.text((x, cy + px(0.42)), INFO, font=fr, fill=SOFT, anchor="ls")
    im.save(path, optimize=True)
    return path


if __name__ == "__main__":
    print(make(os.path.join(HERE, "cover_bg_portrait.png"), 21.0, 29.7, 1.9, 2.0))
    print(make(os.path.join(HERE, "cover_bg_landscape.png"), 29.7, 21.0, 1.7, 2.0))
    print(make(os.path.join(HERE, "cover_bg_slide.png"), 33.867, 19.05, 1.45, 1.8))
