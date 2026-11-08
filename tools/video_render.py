"""사용 안내 영상 만들기: 찍어 둔 화면 + 마이크로소프트 여성 음성(SunHi) + 한글 자막 → 1080p MP4.

먼저 tools/video_capture.py, tools/video_capture_shell.py, tools/video_pages.py 로 화면을 찍어 둔다.
실행: build\\tools_env\\Scripts\\python tools\\video_render.py [--check]   (edge-tts, imageio-ffmpeg, numpy, pillow 필요)
결과: 영상/드라이브 지킴이 사용 안내.mp4, 영상/자막.srt, build/video/chapters.json
"""
from __future__ import annotations

import asyncio
import hashlib
import json
import subprocess
import sys
import wave
from dataclasses import dataclass, field
from pathlib import Path

import edge_tts
import imageio_ffmpeg
import numpy as np
from PIL import Image, ImageDraw, ImageFont

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(Path(__file__).resolve().parent))
from video_script import SCRIPT  # noqa: E402

CAPS = ROOT / "build" / "video" / "caps"
TTS = ROOT / "build" / "video" / "tts"
OUT = ROOT / "영상"
VOICE, RATE = "ko-KR-SunHiNeural", "-3%"
W, H, FPS, STAGE_H = 1920, 1080, 30, 950
SR = 48000
GAP = 0.32
ORANGE = (249, 115, 22)
NAVY = (13, 26, 69)
FONT_B, FONT_R = r"C:\Windows\Fonts\malgunbd.ttf", r"C:\Windows\Fonts\malgun.ttf"
FFMPEG = imageio_ffmpeg.get_ffmpeg_exe()

SPOKEN = [("drive-keeper", "드라이브 키퍼"), ("Ctrl + S", "컨트롤 에스"), ("Ctrl + L", "컨트롤 엘"), ("JSON", "제이슨"),
          ("Google", "구글"), ("Add users", "애드 유저스"), ("Windows", "윈도우"), ("Gmail", "지메일"), ("OAuth", "오어스"),
          ("Word", "워드"), ("Excel", "엑셀"), ("Enter", "엔터"), ("OneDrive", "원드라이브")]


def font(size: int, bold: bool = True) -> ImageFont.FreeTypeFont:
    return ImageFont.truetype(FONT_B if bold else FONT_R, size)


# ───────── 음성 ─────────
def spoken(text: str) -> str:
    for a, b in SPOKEN:
        text = text.replace(a, b)
    return text


def tts_path(text: str) -> Path:
    return TTS / (hashlib.sha1(f"{VOICE}|{RATE}|{text}".encode(), usedforsecurity=False).hexdigest()[:16] + ".wav")


async def synth_one(text: str, sem: asyncio.Semaphore) -> None:
    out = tts_path(text)
    if out.exists():
        return
    async with sem:
        mp3 = out.with_suffix(".mp3")
        for attempt in range(5):
            try:
                await edge_tts.Communicate(text, VOICE, rate=RATE).save(str(mp3))
                if mp3.stat().st_size > 500:
                    break
            except Exception as e:  # noqa: BLE001
                print("TTS 다시 시도:", text[:20], e, flush=True)
                await asyncio.sleep(2 + attempt * 2)
        else:
            raise RuntimeError("TTS 실패: " + text)
    subprocess.run([FFMPEG, "-y", "-loglevel", "error", "-i", str(mp3), "-ac", "1", "-ar", str(SR), str(out)], check=True)
    mp3.unlink()


async def synth_all(texts: list[str]) -> None:
    TTS.mkdir(parents=True, exist_ok=True)
    sem = asyncio.Semaphore(4)
    await asyncio.gather(*(synth_one(t, sem) for t in dict.fromkeys(texts)))


def read_wav(path: Path) -> np.ndarray:
    with wave.open(str(path)) as w:
        return np.frombuffer(w.readframes(w.getnframes()), dtype=np.int16)


# ───────── 그림 재료 ─────────
RECTS = json.loads((CAPS / "rects.json").read_text(encoding="utf-8"))
SHELL = json.loads((CAPS / "shell_rects.json").read_text(encoding="utf-8"))
PAGES = json.loads((CAPS / "page_rects.json").read_text(encoding="utf-8"))
MANUAL = {  # 탐색기 창 안의 자리 (창 크기를 고정해서 찍었음)
    "ex_zip": {"file": [205, 167, 598, 24]}, "ex_top": {"folder": [205, 167, 598, 24]},
    "ex_files": {"install": [205, 251, 598, 26]}, "ex_extract": {"extract": [427, 422, 99, 22]},
}
STAGE_BG = Image.new("RGB", (W, STAGE_H), (238, 242, 249))
_stage: dict[tuple, tuple] = {}


def stage(name: str, zoom: float = 1.0):
    """화면 사진을 무대 가운데에 놓는다 → (무대 그림, 왼쪽 위, 배율)."""
    key = (name, zoom)
    if key not in _stage:
        img = Image.open(CAPS / f"{name}.png").convert("RGB")
        sc = min(zoom, (W - 40) / img.width, (STAGE_H - 24) / img.height)
        if abs(sc - 1) > 1e-3:
            img = img.resize((round(img.width * sc), round(img.height * sc)), Image.LANCZOS)
        canvas = STAGE_BG.copy()
        ox, oy = (W - img.width) // 2, (STAGE_H - img.height) // 2
        if img.size != (1874, 922) or name.startswith("ex_"):
            sh = Image.new("RGBA", (img.width + 60, img.height + 60), (0, 0, 0, 0))
            ImageDraw.Draw(sh).rounded_rectangle((30, 30, img.width + 30, img.height + 30), 12, fill=(15, 30, 79, 60))
            from PIL import ImageFilter
            sh = sh.filter(ImageFilter.GaussianBlur(14))
            canvas.paste(sh, (ox - 30, oy - 22), sh)
        canvas.paste(img, (ox, oy))
        _stage[key] = (canvas, (ox, oy), sc)
    return _stage[key]


def raw_rect(name: str, at) -> list[float]:
    if isinstance(at, (list, tuple)):
        return list(at)
    if name.startswith("m_") or name.startswith("page_"):
        d = PAGES[name]
        if name.startswith("m_"):
            x, y, w, h = d["rects"][int(at)][:4]
            return [x + (1874 - d["vw"]) / 2, y + (922 - d["vh"]) / 2, w, h]
        x, y, w, h = d["rects"][at]
        return [x + (1874 - d["vw"]) / 2, y, w, h]
    for table in (MANUAL, SHELL, RECTS):
        if name in table and at in table[name]:
            return list(table[name][at])
    raise KeyError((name, at))


def stage_rect(name: str, at, zoom: float) -> list[float]:
    _, (ox, oy), sc = stage(name, zoom)
    x, y, w, h = raw_rect(name, at)
    return [ox + x * sc, oy + y * sc, w * sc, h * sc]


def rounded_panel(size, radius, fill):
    im = Image.new("RGBA", (size[0] * 2, size[1] * 2), (0, 0, 0, 0))
    ImageDraw.Draw(im).rounded_rectangle((0, 0, size[0] * 2 - 1, size[1] * 2 - 1), radius * 2, fill=fill)
    return im.resize(size, Image.LANCZOS)


def text_card(lines: list[str], sub: str | None, big: bool) -> Image.Image:
    g = np.linspace(0, 1, STAGE_H)[:, None]
    a, b = np.array([11, 22, 64]), np.array([27, 58, 154])
    arr = (a + (b - a) * g).astype(np.uint8)[:, None, :]
    img = Image.fromarray(np.repeat(arr[:, :, :], W, axis=1).reshape(STAGE_H, W, 3), "RGB")
    logo = Image.open(ROOT / "docs" / "assets" / "logo-256.png").convert("RGBA")
    size = 170 if big else 110
    logo = logo.resize((size, size), Image.LANCZOS)
    f1 = font(92 if big else 74)
    f2 = font(42, False)
    d = ImageDraw.Draw(img)
    lh = int(f1.size * 1.3)
    total = size + 36 + lh * len(lines) + (70 if sub else 0)
    y = (STAGE_H - total) // 2
    img.paste(logo, ((W - size) // 2, y), logo)
    y += size + 36
    for ln in lines:
        d.text((W / 2, y), ln, font=f1, fill=(255, 255, 255), anchor="ma")
        y += lh
    if sub:
        d.text((W / 2, y + 14), sub, font=f2, fill=(190, 208, 250), anchor="ma")
    return img


def wrap(text: str, f, width: int) -> list[str]:
    out, cur = [], ""
    for word in text.split(" "):
        t = (cur + " " + word).strip()
        if f.getlength(t) <= width or not cur:
            cur = t
        else:
            out.append(cur)
            cur = word
    return out + ([cur] if cur else [])


SUB_FONT = font(44)


def subtitle_strip(text: str, progress: float, label: str) -> Image.Image:
    im = Image.new("RGB", (W, H - STAGE_H), NAVY)
    d = ImageDraw.Draw(im)
    d.rectangle((0, 0, int(W * progress), 4), fill=ORANGE)
    lines = wrap(text, SUB_FONT, 1660)
    assert len(lines) <= 2, ("자막이 너무 깁니다", text)
    y = (H - STAGE_H) / 2 + 3 - 30 * len(lines)
    for i, ln in enumerate(lines):
        d.text((W / 2, y + i * 60), ln, font=SUB_FONT, fill=(255, 255, 255), anchor="ma")
    return im


CURSOR = None


def cursor_sprite() -> Image.Image:
    global CURSOR
    if CURSOR is None:
        k = 6
        pts = [(0, 0), (0, 33), (8, 26), (13, 38), (19, 35), (14, 24), (25, 24)]
        pts = [((x + 4) * k, (y + 4) * k) for x, y in pts]
        im = Image.new("RGBA", (40 * k, 50 * k), (0, 0, 0, 0))
        d = ImageDraw.Draw(im)
        d.polygon(pts, fill=(255, 255, 255, 255), outline=(0, 0, 0, 255), width=int(2.2 * k))
        CURSOR = im.resize((44, 55), Image.LANCZOS)
    return CURSOR


_ring: dict[tuple, Image.Image] = {}


def ring_sprite(w: int, h: int, phase: int) -> Image.Image:
    key = (w, h, phase)
    if key not in _ring:
        pad, k = 34, 2
        im = Image.new("RGBA", ((w + pad * 2) * k, (h + pad * 2) * k), (0, 0, 0, 0))
        d = ImageDraw.Draw(im)
        grow = phase / 20
        alpha = int(150 * (1 - grow))
        e = int(26 * grow * k)
        box = (pad * k - 6 * k - e, pad * k - 6 * k - e, (pad + w + 6) * k + e, (pad + h + 6) * k + e)
        if alpha > 0:
            d.rounded_rectangle(box, 14 * k + e, outline=ORANGE + (alpha,), width=6 * k)
        box = (pad * k - 6 * k, pad * k - 6 * k, (pad + w + 6) * k, (pad + h + 6) * k)
        d.rounded_rectangle(box, 14 * k, outline=ORANGE + (255,), width=5 * k)
        _ring[key] = im.resize((w + pad * 2, h + pad * 2), Image.LANCZOS)
    return _ring[key]


_pill: dict[str, Image.Image] = {}


def pill(text: str) -> Image.Image:
    if text not in _pill:
        f = font(30)
        w = int(f.getlength(text)) + 44
        im = rounded_panel((w, 56), 28, ORANGE + (255,))
        ImageDraw.Draw(im).text((w / 2, 28), text, font=f, fill=(255, 255, 255), anchor="mm")
        _pill[text] = im
    return _pill[text]


def keycaps(combo: str) -> Image.Image:
    f = font(58)
    keys = combo.split("+")
    ws = [max(120, int(f.getlength(k)) + 70) for k in keys]
    total = sum(ws) + 90 * (len(keys) - 1) + 60
    im = Image.new("RGBA", (total, 170), (0, 0, 0, 0))
    bg = rounded_panel((total, 170), 28, (13, 26, 69, 235))
    im.alpha_composite(bg)
    d = ImageDraw.Draw(im)
    x = 30
    for i, (k, kw) in enumerate(zip(keys, ws)):
        cap = rounded_panel((kw, 110), 16, (255, 255, 255, 255))
        im.alpha_composite(cap, (x, 30))
        d.rectangle((x + 14, 136, x + kw - 14, 139), fill=(180, 190, 210, 255))
        d.text((x + kw / 2, 84), k, font=f, fill=(15, 23, 42), anchor="mm")
        x += kw
        if i < len(keys) - 1:
            d.text((x + 45, 84), "+", font=f, fill=(255, 255, 255), anchor="mm")
            x += 90
    return im


def callout(text: str) -> Image.Image:
    f = font(38)
    lines = wrap(text, f, 1400)
    w = int(max(f.getlength(x) for x in lines)) + 90
    h = 54 * len(lines) + 46
    im = rounded_panel((w, h), 20, (255, 255, 255, 250))
    d = ImageDraw.Draw(im)
    d.rectangle((0, 18, 9, h - 18), fill=ORANGE + (255,))
    for i, ln in enumerate(lines):
        d.text((w / 2, 23 + 27 + 54 * i), ln, font=f, fill=(15, 23, 42), anchor="mm")
    return im


def note_badge(text: str) -> Image.Image:
    f = font(26, False)
    w = int(f.getlength(text)) + 40
    im = rounded_panel((w, 48), 24, (15, 30, 79, 215))
    ImageDraw.Draw(im).text((w / 2, 24), text, font=f, fill=(255, 255, 255), anchor="mm")
    return im


MENU_ITEM, MENU_W = 58, 360


def menu_layout(anchor: list[float], items: list[str]):
    h = sum(14 if not s else MENU_ITEM for s in items) + 20
    w = max(MENU_W, int(max(font(30, False).getlength(s) for s in items)) + 70)
    x = min(anchor[0] + 40, W - w - 30)
    y = min(anchor[1] + anchor[3] / 2, STAGE_H - h - 14)
    return x, y, w, h


def menu_sprite(items: list[str], w: int, h: int) -> Image.Image:
    f = font(30, False)
    im = rounded_panel((w, h), 14, (255, 255, 255, 255))
    d = ImageDraw.Draw(im)
    d.rounded_rectangle((0, 0, w - 1, h - 1), 14, outline=(200, 205, 215, 255), width=2)
    y = 10
    for s in items:
        if not s:
            d.line((16, y + 7, w - 16, y + 7), fill=(225, 228, 235, 255), width=2)
            y += 14
            continue
        d.text((30, y + MENU_ITEM / 2), s, font=f, fill=(20, 24, 35), anchor="lm")
        y += MENU_ITEM
    return im


def menu_item_rect(anchor, items, idx):
    x, y, w, _ = menu_layout(anchor, items)
    yy = y + 10
    for i, s in enumerate(items):
        if i == idx:
            return [x + 8, yy + 4, w - 16, MENU_ITEM - 8]
        yy += 14 if not s else MENU_ITEM
    raise IndexError(idx)


def ease(t: float) -> float:
    return t * t * (3 - 2 * t)


# ───────── 타임라인 ─────────
@dataclass
class Beat:
    say: str
    img: str = ""
    card: tuple | None = None
    start: float = 0.0
    dur: float = 0.0
    tts: float = 0.0
    zoom: float = 1.0
    at: object = None
    tag: str | None = None
    click: bool = False
    then: str | None = None
    keys: str | None = None
    call: str | None = None
    menu: list | None = None
    menu_at: object = None
    menu_pick: int | None = None
    chapter: str = ""
    extra: dict = field(default_factory=dict)


def build() -> tuple[list[Beat], list[tuple[float, str]]]:
    beats: list[Beat] = []
    for ch_title, items in SCRIPT:
        for it in items:
            b = Beat(chapter=ch_title, **it)
            beats.append(b)
    asyncio.run(synth_all([spoken(b.say) for b in beats]))
    t = 0.0
    chapters: list[tuple[float, str]] = []
    seen = set()
    for b in beats:
        if b.chapter not in seen:
            seen.add(b.chapter)
            chapters.append((t, b.chapter))
        b.tts = len(read_wav(tts_path(spoken(b.say)))) / SR
        b.start = t
        b.dur = b.tts + GAP + (0.55 if b.then else 0.0) + (0.5 if b.card else 0.0)
        t += b.dur
    return beats, chapters


def render(beats: list[Beat], out: Path) -> None:
    total = beats[-1].start + beats[-1].dur
    cmd = [FFMPEG, "-y", "-loglevel", "error", "-f", "rawvideo", "-pix_fmt", "rgb24", "-s", f"{W}x{H}", "-r", str(FPS),
           "-i", "-", "-i", str(ROOT / "build" / "video" / "voice.wav"),
           "-vf", "scale=out_color_matrix=bt709:out_range=tv,format=yuv420p",
           "-af", "loudnorm=I=-14:TP=-1.5:LRA=11,aresample=48000",
           "-c:v", "libx264", "-preset", "slow", "-crf", "14", "-profile:v", "high",
           "-colorspace", "bt709", "-color_primaries", "bt709", "-color_trc", "bt709",
           "-c:a", "aac", "-b:a", "192k", "-movflags", "+faststart", str(out)]
    ff = subprocess.Popen(cmd, stdin=subprocess.PIPE)
    cur_pos = None
    card_cache: dict = {}
    for n, b in enumerate(beats):
        frames = round((b.start + b.dur) * FPS) - round(b.start * FPS)
        strip = subtitle_strip(b.say, (b.start + b.dur / 2) / total, b.chapter)
        if b.card:
            key = json.dumps(b.card, ensure_ascii=False)
            card_cache.setdefault(key, text_card(*b.card))
            first_img = card_cache[key]
            second_img = None
            is_m = False
        else:
            first_img = stage(b.img, b.zoom)[0]
            second_img = stage(b.then, b.zoom)[0] if b.then else None
            is_m = b.img.startswith("m_")
        base_a = Image.new("RGB", (W, H))
        base_a.paste(first_img, (0, 0))
        base_a.paste(strip, (0, STAGE_H))
        base_b = None
        if second_img is not None:
            base_b = Image.new("RGB", (W, H))
            base_b.paste(second_img, (0, 0))
            base_b.paste(strip, (0, STAGE_H))
        target = None
        if b.menu_pick is not None and b.menu:
            anchor = stage_rect(b.img, b.menu_at, b.zoom)
            target = menu_item_rect(anchor, b.menu, b.menu_pick)
        elif b.at is not None:
            target = stage_rect(b.img, b.at, b.zoom)
        arrive = min(1.1, max(0.45, b.tts * 0.42))
        click_t = max(arrive + 0.12, b.tts * 0.78) if b.click else None
        start_pos = cur_pos or (W * 0.5, STAGE_H * 0.62)
        goal = None
        if target:
            goal = (target[0] + target[2] * (0.5 if target[2] < 420 else 0.22) - 2, target[1] + target[3] * 0.55)
        menu_img, menu_xy = None, None
        if b.menu:
            anchor = stage_rect(b.img, b.menu_at, b.zoom)
            x, y, w, h = menu_layout(anchor, b.menu)
            menu_img, menu_xy = menu_sprite(b.menu, int(w), int(h)), (int(x), int(y))
        overlay_keys = keycaps(b.keys) if b.keys else None
        overlay_call = callout(b.call) if b.call else None
        badge = note_badge("구글 화면 예시 그림") if is_m and b.img in GOOGLE_MOCKS else None
        for i in range(frames):
            t = i / FPS
            fr = base_a.copy()
            if base_b is not None and click_t is not None and t >= click_t + 0.1:
                a = min(1.0, (t - click_t - 0.1) / 0.2)
                fr = Image.blend(base_a, base_b, a) if a < 1 else base_b.copy()
            switched = base_b is not None and click_t is not None and t >= click_t + 0.1
            if badge is not None and not switched:
                fr.paste(badge, (W - badge.width - 26, 20), badge)
            if menu_img is not None and not switched:
                fr.paste(menu_img, menu_xy, menu_img)
            if overlay_call is not None and not switched:
                fr.paste(overlay_call, ((W - overlay_call.width) // 2, 26), overlay_call)
            if overlay_keys is not None:
                ka = min(1.0, t / 0.25)
                if ka >= 1 or i % 2 == 0:
                    fr.paste(overlay_keys, ((W - overlay_keys.width) // 2, STAGE_H - overlay_keys.height - 40), overlay_keys)
            pos = cur_pos
            if goal is not None:
                k = ease(min(1.0, t / arrive))
                pos = (start_pos[0] + (goal[0] - start_pos[0]) * k, start_pos[1] + (goal[1] - start_pos[1]) * k)
                if not switched:
                    rw, rh = int(target[2]), int(target[3])
                    phase = int((t * 30) % 21) if not is_m else None
                    if phase is not None:
                        sp = ring_sprite(rw, rh, phase)
                        fr.paste(sp, (int(target[0]) - 34, int(target[1]) - 34), sp)
                    if b.tag and t > arrive * 0.5:
                        pl = pill(b.tag)
                        px = int(min(max(20, target[0] + target[2] / 2 - pl.width / 2), W - pl.width - 20))
                        py = int(target[1] + target[3] + 20)
                        if py + pl.height > STAGE_H - 6:
                            py = int(target[1] - pl.height - 20)
                        fr.paste(pl, (px, py), pl)
            if click_t is not None and click_t <= t <= click_t + 0.4 and goal is not None:
                p = (t - click_t) / 0.4
                r = int(10 + 52 * p)
                ov = Image.new("RGBA", (r * 2 + 8, r * 2 + 8), (0, 0, 0, 0))
                ImageDraw.Draw(ov).ellipse((4, 4, r * 2 + 4, r * 2 + 4), outline=ORANGE + (int(255 * (1 - p)),), width=6)
                fr.paste(ov, (int(goal[0]) - r - 4, int(goal[1]) - r - 4), ov)
            if pos is not None and not b.card:
                c = cursor_sprite()
                fr.paste(c, (int(pos[0]) - 6, int(pos[1]) - 6), c)
            ff.stdin.write(fr.tobytes())
        if goal is not None:
            cur_pos = goal
        if b.card:
            cur_pos = None
        if n % 10 == 0:
            print(f"렌더 {n}/{len(beats)}  {b.start:.0f}/{total:.0f}초", flush=True)
    ff.stdin.close()
    ff.wait()
    if ff.returncode:
        raise RuntimeError("ffmpeg 실패")


GOOGLE_MOCKS = {"m_b1", "m_b2", "m_b3", "m_b4", "m_b5", "m_b6", "m_choose", "m_warn", "m_allow", "m_e1", "m_other",
                "m_download", "m_smart"}


def stamp(t: float) -> str:
    ms = int(round(t * 1000))
    return f"{ms // 3600000:02d}:{ms // 60000 % 60:02d}:{ms // 1000 % 60:02d},{ms % 1000:03d}"


def main() -> int:
    beats, chapters = build()
    total = beats[-1].start + beats[-1].dur
    for b in beats:   # 자막 길이·그림 자리를 미리 검사
        subtitle_strip(b.say, 0.5, b.chapter)
        if b.card:
            continue
        stage(b.img, b.zoom)
        if b.then:
            stage(b.then, b.zoom)
        if b.at is not None and b.menu_pick is None:
            stage_rect(b.img, b.at, b.zoom)
    print(f"장면 {len(beats)}개, 길이 {total / 60:.1f}분", flush=True)
    if "--check" in sys.argv:
        return 0
    OUT.mkdir(exist_ok=True)
    voice = np.zeros(int(total * SR) + SR, dtype=np.int16)
    srt = []
    for i, b in enumerate(beats, 1):
        clip = read_wav(tts_path(spoken(b.say)))
        at = int(b.start * SR)
        voice[at:at + len(clip)] = clip
        lines = wrap(b.say, SUB_FONT, 1660)
        srt.append(f"{i}\n{stamp(b.start)} --> {stamp(b.start + b.tts + 0.15)}\n" + "\n".join(lines) + "\n")
    with wave.open(str(ROOT / "build" / "video" / "voice.wav"), "wb") as w:
        w.setnchannels(1)
        w.setsampwidth(2)
        w.setframerate(SR)
        w.writeframes(voice.tobytes())
    (OUT / "자막.srt").write_text("\n".join(srt), encoding="utf-8-sig")
    (ROOT / "build" / "video" / "chapters.json").write_text(
        json.dumps({"total": total, "chapters": chapters}, ensure_ascii=False, indent=1), encoding="utf-8")
    render(beats, OUT / "드라이브 지킴이 사용 안내.mp4")
    print("끝", total, flush=True)
    return 0


if __name__ == "__main__":
    sys.exit(main())
