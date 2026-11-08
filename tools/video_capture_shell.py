"""사용 안내 영상용: 설치 과정의 실제 Windows 창을 찍는다 (탐색기·압축 풀기·설치 완료 창).

실행: .venv\\Scripts\\python tools\\video_capture_shell.py   → build/video/caps/ex_*.png, shell_rects.json
새로 연 창만 찍고 닫는다 (사용자가 열어 둔 창은 건드리지 않음). 예시 폴더 C:\\다운로드 는 끝나면 지운다.
"""
from __future__ import annotations

import ctypes
import json
import shutil
import subprocess
import sys
import time
import zipfile
from ctypes import wintypes
from pathlib import Path

from PIL import Image

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "build" / "video" / "caps"
DEMO = Path("C:/다운로드")
ZIP = ROOT / "release" / "DriveKeeper.zip"
u = ctypes.windll.user32
g = ctypes.windll.gdi32
u.SetProcessDPIAware()
WM_CLOSE, BM_CLICK = 0x10, 0xF5
ENUM = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)


def text_of(h) -> str:
    n = u.GetWindowTextLengthW(h)
    b = ctypes.create_unicode_buffer(n + 1)
    u.GetWindowTextW(h, b, n + 1)
    return b.value


def class_of(h) -> str:
    b = ctypes.create_unicode_buffer(128)
    u.GetClassNameW(h, b, 128)
    return b.value


def tops() -> dict[int, tuple[str, str]]:
    out = {}

    def cb(h, _):
        if u.IsWindowVisible(h):
            out[h] = (class_of(h), text_of(h))
        return True
    u.EnumWindows(ENUM(cb), 0)
    return out


def wait_new(before, want, sec=15.0):
    end = time.time() + sec
    while time.time() < end:
        for h, info in tops().items():
            if h not in before and want(info):
                return h
        time.sleep(0.2)
    return None


def grab(h) -> Image.Image:
    r = wintypes.RECT()
    u.GetWindowRect(h, ctypes.byref(r))
    w, ht = r.right - r.left, r.bottom - r.top
    hdc = u.GetWindowDC(h)
    mdc = g.CreateCompatibleDC(hdc)
    bmp = g.CreateCompatibleBitmap(hdc, w, ht)
    g.SelectObject(mdc, bmp)
    u.PrintWindow(h, mdc, 2)

    class BMI(ctypes.Structure):
        _fields_ = [("biSize", wintypes.DWORD), ("biWidth", wintypes.LONG), ("biHeight", wintypes.LONG),
                    ("biPlanes", wintypes.WORD), ("biBitCount", wintypes.WORD), ("biCompression", wintypes.DWORD),
                    ("biSizeImage", wintypes.DWORD), ("a", wintypes.LONG), ("b", wintypes.LONG),
                    ("c", wintypes.DWORD), ("d", wintypes.DWORD)]
    bmi = BMI()
    bmi.biSize, bmi.biWidth, bmi.biHeight, bmi.biPlanes, bmi.biBitCount = ctypes.sizeof(BMI), w, -ht, 1, 32
    buf = ctypes.create_string_buffer(w * ht * 4)
    g.GetDIBits(mdc, bmp, 0, ht, buf, ctypes.byref(bmi), 0)
    g.DeleteObject(bmp)
    g.DeleteDC(mdc)
    u.ReleaseDC(h, hdc)
    img = Image.frombuffer("RGB", (w, ht), buf.raw, "raw", "BGRX", 0, 1).copy()
    fr = wintypes.RECT()
    if ctypes.windll.dwmapi.DwmGetWindowAttribute(h, 9, ctypes.byref(fr), ctypes.sizeof(fr)) == 0:
        img = img.crop((fr.left - r.left, fr.top - r.top, fr.right - r.left, fr.bottom - r.top))
    return img


def clean_nav(img: Image.Image) -> Image.Image:
    """탐색기 왼쪽 목록에서 이 PC 사용자의 개인 폴더 이름을 지운다 (기본 폴더만 남김)."""
    white = (255, 255, 255)
    below = img.crop((0, 333, 190, 492))
    img.paste(below, (0, 301))
    img.paste(white, (0, 301 + below.height, 190, 600))
    return img


def close(h) -> None:
    u.PostMessageW(h, WM_CLOSE, 0, 0)


def explore(args: list[str], name: str, title_has: str):
    """탐색기 창을 새로 열어 찍고 닫는다. 기존 창에 탭으로 열리면 찍지 않는다."""
    before = tops()
    subprocess.Popen(["explorer.exe", *args])
    h = wait_new(before, lambda i: i[0] == "CabinetWClass")
    if not h:
        print("새 탐색기 창이 없음:", name, flush=True)
        return False
    u.MoveWindow(h, 20, 20, 1500, 860, True)
    u.keybd_event(0x12, 0, 0, 0)       # 비활성 창은 윗부분이 검게 찍혀서 앞으로 가져온다 (Alt: 앞으로 가져오기 허용)
    u.SetForegroundWindow(h)
    u.keybd_event(0x12, 0, 2, 0)
    time.sleep(2.5)
    print("창 제목:", text_of(h), flush=True)
    clean_nav(grab(h)).save(OUT / f"{name}.png")
    close(h)
    time.sleep(0.8)
    return True


def main() -> int:
    if DEMO.exists():
        print("C:/다운로드 가 이미 있어서 멈춥니다.")
        return 1
    DEMO.mkdir()
    rects: dict = {}
    try:
        zp = DEMO / "DriveKeeper.zip"
        shutil.copyfile(ZIP, zp)
        explore([f"/select,{zp}"], "ex_zip", "다운로드")

        # 실제 '압축 풀기' 창
        before = tops()
        ps = ("$s=New-Object -ComObject Shell.Application;"
              "$i=$s.Namespace('C:\\다운로드').ParseName('DriveKeeper.zip');"
              "$v=$i.Verbs()|?{$_.Name -match 'tract All|모두 압축 풀기'}|select -First 1;"
              "if($v){$v.Name;$v.DoIt()}else{$i.Verbs()|%{$_.Name}};Start-Sleep 25")
        p = subprocess.Popen(["powershell.exe", "-NoProfile", "-Command", ps])
        h = wait_new(before, lambda i: "압축" in i[1] and i[0] != "CabinetWClass", 14)
        if h:
            u.MoveWindow(h, 300, 120, 900, 680, True)
            time.sleep(1.5)
            grab(h).save(OUT / "ex_extract.png")
            print("압축 풀기 창:", class_of(h), text_of(h), flush=True)
            close(h)
        else:
            print("압축 풀기 창을 찾지 못함", flush=True)
        p.kill()
        time.sleep(1)

        out = DEMO / "DriveKeeper"
        with zipfile.ZipFile(zp) as z:
            z.extractall(out)
        explore([str(out)], "ex_top", "DriveKeeper")
        top = out / "드라이브 지킴이"
        explore([str(top)], "ex_files", "드라이브 지킴이")

        # 실제 설치 → '설치가 끝났어요' 창
        running = subprocess.run(["tasklist", "/FI", "IMAGENAME eq DriveKeeper.exe"], capture_output=True).stdout
        if (OUT / "ex_done.png").exists():
            print("설치 완료 창은 이미 찍었습니다.", flush=True)
            rects = json.loads((OUT / "shell_rects.json").read_text(encoding="utf-8"))
        elif b"DriveKeeper.exe" in running:
            print("드라이브 지킴이가 켜져 있어 설치 창은 찍지 않습니다.", flush=True)
        else:
            before = tops()
            subprocess.Popen(["explorer.exe", str(top / "설치하기.cmd")])
            h = wait_new(before, lambda i: i[1] == "드라이브 지킴이 설치", 60)
            if h:
                time.sleep(0.6)
                grab(h).save(OUT / "ex_done.png")
                kids = []

                def cb(c, _):
                    if class_of(c) == "Button":
                        r = wintypes.RECT()
                        u.GetWindowRect(c, ctypes.byref(r))
                        kids.append((c, text_of(c), r.left, r.top, r.right - r.left, r.bottom - r.top))
                    return True
                u.EnumChildWindows(h, ENUM(cb), 0)
                f = wintypes.RECT()
                ctypes.windll.dwmapi.DwmGetWindowAttribute(h, 9, ctypes.byref(f), ctypes.sizeof(f))
                rects["ex_done"] = {("yes" if "Y" in k[1] or "예" in k[1] else "no"):
                                    [k[2] - f.left, k[3] - f.top, k[4], k[5]] for k in kids}
                print("단추:", kids, flush=True)
                no = next(k[0] for k in kids if "N" in k[1] or "아니" in k[1])
                u.SendMessageW(no, BM_CLICK, 0, 0)
            else:
                print("설치 창을 찾지 못함", flush=True)
            time.sleep(2)
    finally:
        (OUT / "shell_rects.json").write_text(json.dumps(rects, ensure_ascii=False), encoding="utf-8")
        for _ in range(5):
            try:
                shutil.rmtree(DEMO)
                break
            except OSError:
                time.sleep(1.5)
    print("끝", flush=True)
    return 0


if __name__ == "__main__":
    sys.exit(main())
