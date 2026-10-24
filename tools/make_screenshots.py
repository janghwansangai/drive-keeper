"""설명서·홈페이지용 화면 사진을 만든다 (가짜 자료·가짜 계정, 개인 정보 없음).

실행: .venv\\Scripts\\python tools\\make_screenshots.py   → docs/assets/shots/*.png
"""
from __future__ import annotations

import json
import shutil
import sys
import tempfile
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "tests"))
OUT = ROOT / "docs" / "assets" / "shots"


def capture(widget, name: str) -> None:
    """창 하나를 PNG로 저장 (Windows PrintWindow)."""
    import ctypes
    import struct
    import zlib
    from ctypes import wintypes
    user32, gdi32 = ctypes.windll.user32, ctypes.windll.gdi32
    hwnd = int(widget.wm_frame(), 16)
    r = wintypes.RECT()
    user32.GetWindowRect(hwnd, ctypes.byref(r))
    w, h = r.right - r.left, r.bottom - r.top
    hdc = user32.GetWindowDC(hwnd)
    mdc = gdi32.CreateCompatibleDC(hdc)
    bmp = gdi32.CreateCompatibleBitmap(hdc, w, h)
    gdi32.SelectObject(mdc, bmp)
    user32.PrintWindow(hwnd, mdc, 2)

    class BMI(ctypes.Structure):
        _fields_ = [("biSize", wintypes.DWORD), ("biWidth", wintypes.LONG), ("biHeight", wintypes.LONG),
                    ("biPlanes", wintypes.WORD), ("biBitCount", wintypes.WORD), ("biCompression", wintypes.DWORD),
                    ("biSizeImage", wintypes.DWORD), ("a", wintypes.LONG), ("b", wintypes.LONG),
                    ("c", wintypes.DWORD), ("d", wintypes.DWORD)]
    bmi = BMI()
    bmi.biSize, bmi.biWidth, bmi.biHeight, bmi.biPlanes, bmi.biBitCount = ctypes.sizeof(BMI), w, -h, 1, 32
    buf = ctypes.create_string_buffer(w * h * 4)
    gdi32.GetDIBits(mdc, bmp, 0, h, buf, ctypes.byref(bmi), 0)
    gdi32.DeleteObject(bmp)
    gdi32.DeleteDC(mdc)
    user32.ReleaseDC(hwnd, hdc)
    # Windows 11의 보이지 않는 창 테두리(검은 띠)를 잘라 낸다
    fr = wintypes.RECT()
    x0, y0, x1, y1 = 0, 0, w, h
    if ctypes.windll.dwmapi.DwmGetWindowAttribute(hwnd, 9, ctypes.byref(fr), ctypes.sizeof(fr)) == 0:
        x0, y0 = max(0, fr.left - r.left), max(0, fr.top - r.top)
        x1, y1 = min(w, fr.right - r.left), min(h, fr.bottom - r.top)
    raw = bytearray()
    data = buf.raw
    for y in range(y0, y1):
        raw.append(0)
        row = data[y * w * 4:(y + 1) * w * 4]
        for x in range(x0, x1):
            raw += bytes((row[x * 4 + 2], row[x * 4 + 1], row[x * 4]))
    w, h = x1 - x0, y1 - y0

    def chunk(t, d):
        return struct.pack(">I", len(d)) + t + d + struct.pack(">I", zlib.crc32(t + d) & 0xFFFFFFFF)
    png = (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", w, h, 8, 2, 0, 0, 0))
           + chunk(b"IDAT", zlib.compress(bytes(raw), 9)) + chunk(b"IEND", b""))
    OUT.mkdir(parents=True, exist_ok=True)
    (OUT / f"{name}.png").write_bytes(png)
    print("찍음:", name, w, "x", h)


def main() -> int:
    import config
    base = Path(tempfile.mkdtemp(prefix="gde_shots_"))
    config.BASE_DIR = base
    config.PROJECT_DIR = base           # 개발 폴더의 진짜 열쇠를 쓰지 않게
    config.WORK_DIR, config.STATE_PATH = base / "work", base / "state.json"
    config.STAGING_DIR, config.TOKEN_PATH, config.LOG_PATH = base / "st", base / "t.json", base / "a.log"
    import accounts
    import app as app_mod
    import security
    from fake_drive import FakeDrive
    app_mod.enable_dpi_awareness()
    app_mod.open_with_default = lambda p: None
    for n in ("showinfo", "showwarning", "showerror"):
        setattr(app_mod.messagebox, n, lambda *a, **k: None)
    email = "teacher@example.com"
    d = FakeDrive(email)
    folders = ["2학기 학급 운영", "수업 자료", "업무 문서"]
    for f in folders:
        d.add_file(f, None, mime="application/vnd.google-apps.folder")
    d.add_file("공유받은 폴더", None, mime="application/vnd.google-apps.shortcut",
               extra={"shortcutDetails": {"targetId": "folder1", "targetMimeType": "application/vnd.google-apps.folder"}})
    files = [("가정통신문 (10월).hwp", 48000, "application/x-hwp"),
             ("학급 명렬표.xlsx", 15000, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
             ("현장체험학습 계획서.docx", 23000, "application/vnd.openxmlformats-officedocument.wordprocessingml.document"),
             ("수업 발표 자료.pptx", 2300000, "application/vnd.openxmlformats-officedocument.presentationml.presentation"),
             ("교육과정 안내.pdf", 830000, "application/pdf"), ("단체 사진.jpg", 3200000, "image/jpeg")]
    metas = {}
    for name, size, mime in files:
        metas[name] = d.add_file(name, b"x" * size, mime=mime)
    d.add_file("학부모 상담 기록", b"g", mime="application/vnd.google-apps.document")
    d.add_file("성적 집계", b"g", mime="application/vnd.google-apps.spreadsheet")
    for txt in (b"v2 ...", b"v3 ......", b"v4"):
        time.sleep(0.01)
        d.remote_edit(metas["가정통신문 (10월).hwp"]["id"], txt)
    for n in ("지난 학기 시간표.xlsx", "회의록 초안.docx", "행사 사진", "옛 가정통신문.hwp"):
        f = d.add_file(n, b"x" * 9000, mime="application/vnd.google-apps.folder" if n == "행사 사진" else "application/x-hwp")
        d.trash(f["id"])

    # 첫 실행 화면: 열쇠도 계정도 없음
    a = app_mod.App(drive_factory=lambda acct, cb: d, start_connect=False)
    a.geometry("1180x760+30+30")

    def pump(sec):
        end = time.time() + sec
        while time.time() < end:
            a.update()
            time.sleep(0.02)
    pump(0.8)
    a.open_setup()
    pump(0.8)
    capture(a._setup, "setup_first")
    a._setup.destroy()
    # 열쇠를 넣은 뒤
    (base / "client_secret.json").write_text(json.dumps({"installed": {
        "client_id": "123-demo.apps.googleusercontent.com", "client_secret": "demo", "project_id": "drive-keeper-12345"}}),
        encoding="utf-8")
    a.open_setup()
    pump(0.6)
    capture(a._setup, "setup_key")
    a._setup.destroy()
    a.help_menu.unpost()
    a.show_help("key")
    pump(0.6)
    helps = [w for w in a.winfo_children() if isinstance(w, app_mod.HelpDialog)]
    capture(helps[-1], "help_key")
    helps[-1].destroy()
    a.shutdown()

    # 로그인한 뒤 평소 화면
    st = accounts.AccountStore()
    st.load()
    st.add(email)
    a = app_mod.App(drive_factory=lambda acct, cb: d)
    a.geometry("1180x760+30+30")
    pump(2.5)
    a.open_item(next(m for m in a.item_meta.values() if m["name"].endswith(".hwp")))
    pump(1.2)
    a.set_status("가정통신문 (10월).hwp: 드라이브에 저장됨 14:32:10")
    sm = a.sync
    for r in sm.rows():
        sm._recs[r["key"]]["status"] = "드라이브에 저장됨 14:32:10"
    a._refresh_track_view()
    pump(0.4)
    capture(a, "main")
    # 타임머신
    a.open_time_machine()
    pump(1.2)
    capture(a._tm_dialog, "tm_trash")
    a.open_time_machine(next(m for m in a.item_meta.values() if m["name"].endswith(".hwp")))
    pump(1.0)
    capture(a._tm_dialog, "tm_versions")
    a._tm_dialog.destroy()
    # 보안 설정 (비밀번호 만들기 화면)
    a.open_security()
    pump(0.6)
    capture(a._sec_dialog, "security")
    a._sec_dialog.destroy()
    # 잠금 화면
    s = security.Security()
    s.load()
    s.set_password("1234")
    a.security.load()
    a.security_changed("")
    a.lock_now("idle")
    pump(0.6)
    capture(a, "lock")
    a.unlock()
    a.shutdown()
    shutil.rmtree(base, ignore_errors=True)
    return 0


if __name__ == "__main__":
    sys.exit(main())
