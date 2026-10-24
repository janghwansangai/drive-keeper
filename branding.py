"""드라이브 지킴이 로고: 파란 방패 + 흰 자물쇠. 그림 파일 없이 코드로 그린다 (PNG·ICO·SVG).

python branding.py  → brand_assets.py(프로그램에 넣을 PNG), build/app.ico, docs/assets/ 로고 파일을 만든다.
"""
from __future__ import annotations

import base64
import math
import struct
import sys
import zlib
from pathlib import Path

ROOT = Path(__file__).resolve().parent

TOP = (0x5B, 0x95, 0xFF)      # 방패 위쪽 색
BOTTOM = (0x1D, 0x4E, 0xD8)   # 방패 아래쪽 색
WHITE = (0xFF, 0xFF, 0xFF)

W = 0.40          # 방패 반폭
Y0, YM, YB = 0.06, 0.50, 0.95   # 방패 위, 곧은 옆면이 끝나는 곳, 아래 끝
R = 0.07          # 위 모서리 둥글기


def _in_shield(x: float, y: float) -> bool:
    if y < Y0 or y > YB:
        return False
    dx = abs(x - 0.5)
    if y <= YM:
        half = W
    else:
        half = W * math.cos((y - YM) / (YB - YM) * math.pi / 2)
    if dx > half:
        return False
    if y < Y0 + R and dx > W - R:   # 위쪽 둥근 모서리
        cx, cy = W - R, Y0 + R
        return (dx - cx) ** 2 + (y - cy) ** 2 <= R * R
    return True


def _in_lock(x: float, y: float) -> int:
    """0: 아님, 1: 흰 자물쇠, 2: 열쇠 구멍(방패색)."""
    # 몸통 (둥근 사각형)
    bx0, bx1, by0, by1, br = 0.33, 0.67, 0.47, 0.76, 0.045
    inside_body = bx0 <= x <= bx1 and by0 <= y <= by1
    if inside_body:
        cx = min(max(x, bx0 + br), bx1 - br)
        cy = min(max(y, by0 + br), by1 - br)
        inside_body = (x - cx) ** 2 + (y - cy) ** 2 <= br * br
    if inside_body:
        # 열쇠 구멍
        if (x - 0.5) ** 2 + (y - 0.585) ** 2 <= 0.038 ** 2:
            return 2
        if abs(x - 0.5) <= 0.017 and 0.585 <= y <= 0.68:
            return 2
        return 1
    # 고리 (위쪽 반원 띠 + 다리)
    cx, cy, ro, ri = 0.5, 0.40, 0.125, 0.075
    d = math.hypot(x - cx, y - cy)
    if y <= cy and ri <= d <= ro:
        return 1
    if cy < y <= by0 + 0.01 and (ri <= abs(x - cx) <= ro):
        return 1
    return 0


def render(size: int, ss: int = 4) -> list[list[tuple[int, int, int, int]]]:
    """size×size RGBA 픽셀 (ss×ss 겹쳐 그려 가장자리를 부드럽게)."""
    rows = []
    for py in range(size):
        row = []
        for px in range(size):
            acc = [0.0, 0.0, 0.0, 0.0]
            for sy in range(ss):
                for sx in range(ss):
                    x = (px + (sx + 0.5) / ss) / size
                    y = (py + (sy + 0.5) / ss) / size
                    if not _in_shield(x, y):
                        continue
                    t = (y - Y0) / (YB - Y0)
                    col = tuple(TOP[i] + (BOTTOM[i] - TOP[i]) * t for i in range(3))
                    lk = _in_lock(x, y) if size >= 20 else (1 if _in_lock(x, y) == 1 else 0)
                    if lk == 1:
                        col = WHITE
                    acc[0] += col[0]
                    acc[1] += col[1]
                    acc[2] += col[2]
                    acc[3] += 1
            n = ss * ss
            if acc[3]:
                row.append((int(acc[0] / acc[3]), int(acc[1] / acc[3]), int(acc[2] / acc[3]),
                            int(255 * acc[3] / n)))
            else:
                row.append((0, 0, 0, 0))
        rows.append(row)
    return rows


def png_bytes(size: int) -> bytes:
    px = render(size)
    raw = bytearray()
    for row in px:
        raw.append(0)
        for r, g, b, a in row:
            raw.extend((r, g, b, a))

    def chunk(kind: bytes, data: bytes) -> bytes:
        return struct.pack(">I", len(data)) + kind + data + struct.pack(">I", zlib.crc32(kind + data) & 0xFFFFFFFF)
    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", size, size, 8, 6, 0, 0, 0))
            + chunk(b"IDAT", zlib.compress(bytes(raw), 9)) + chunk(b"IEND", b""))


def ico_bytes(sizes=(16, 24, 32, 48, 64, 128, 256)) -> bytes:
    images = [png_bytes(s) for s in sizes]
    header = struct.pack("<HHH", 0, 1, len(images))
    offset = 6 + 16 * len(images)
    entries, data = b"", b""
    for s, img in zip(sizes, images):
        entries += struct.pack("<BBBBHHII", s % 256, s % 256, 0, 0, 1, 32, len(img), offset + len(data))
        data += img
    return header + entries + data


SVG = """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100" role="img" aria-label="드라이브 지킴이">
  <defs><linearGradient id="g" x1="0" y1="0" x2="0" y2="1">
    <stop offset="0" stop-color="#5B95FF"/><stop offset="1" stop-color="#1D4ED8"/></linearGradient></defs>
  <path d="M17 6 H83 Q90 6 90 13 V50 C90 75 66 88 50 95 C34 88 10 75 10 50 V13 Q10 6 17 6 Z" fill="url(#g)"/>
  <path d="M37.5 47 V40 A12.5 12.5 0 0 1 62.5 40 V47" fill="none" stroke="#fff" stroke-width="5"/>
  <rect x="33" y="47" width="34" height="29" rx="4.5" fill="#fff"/>
  <circle cx="50" cy="58.5" r="3.8" fill="#2B63E6"/><rect x="48.3" y="58.5" width="3.4" height="9.5" rx="1.2" fill="#2B63E6"/>
</svg>
"""


def main() -> None:
    sizes = (16, 20, 24, 32, 40, 48, 64, 96, 128)
    lines = ['"""자동 생성 파일 (branding.py). 드라이브 지킴이 로고 PNG를 base64로 담는다."""', "LOGO_PNG = {"]
    for s in sizes:
        lines.append(f'    {s}: "{base64.b64encode(png_bytes(s)).decode("ascii")}",')
    lines.append("}")
    (ROOT / "brand_assets.py").write_text("\n".join(lines) + "\n", encoding="utf-8")
    (ROOT / "build").mkdir(exist_ok=True)
    (ROOT / "build" / "app.ico").write_bytes(ico_bytes())
    assets = ROOT / "docs" / "assets"
    assets.mkdir(parents=True, exist_ok=True)
    (assets / "logo.svg").write_text(SVG, encoding="utf-8")
    (assets / "favicon.png").write_bytes(png_bytes(64))
    (assets / "logo-256.png").write_bytes(png_bytes(256))
    (assets / "favicon.ico").write_bytes(ico_bytes((16, 32, 48)))
    print("만듦: brand_assets.py, build/app.ico, docs/assets/")


if __name__ == "__main__":
    sys.exit(main())
