"""사용 안내 영상용 화면 캡처: 실제 프로그램을 가짜 드라이브·예시 계정으로 돌리며 화면과 단추 자리를 기록한다.

실행: .venv\\Scripts\\python tools\\video_capture.py   → build/video/caps/*.png, rects.json
화면을 1.5배로 키워 찍어서 1080p 영상에서 글자가 또렷하게 보인다. 개인 정보는 들어가지 않는다.
"""
from __future__ import annotations

import ctypes
import json
import shutil
import sys
import tempfile
import time
import tkinter as tk
from ctypes import wintypes
from pathlib import Path

from PIL import Image, ImageFilter

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "tests"))
OUT = ROOT / "build" / "video" / "caps"
SCALE = 1.5
RECTS: dict[str, dict] = {}
OLE = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"

_tk_init = tk.Tk.__init__


def _scaled_init(self, *a, **k):
    _tk_init(self, *a, **k)
    self.tk.call("tk", "scaling", SCALE * 96 / 72)


tk.Tk.__init__ = _scaled_init


def grab(widget) -> tuple[Image.Image, tuple[int, int]]:
    """창 하나를 찍는다 (다른 창에 가려져 있어도 됨). (그림, 그림의 화면 왼쪽 위 좌표)"""
    u, g = ctypes.windll.user32, ctypes.windll.gdi32
    hwnd = int(widget.wm_frame(), 16)
    r = wintypes.RECT()
    u.GetWindowRect(hwnd, ctypes.byref(r))
    w, h = r.right - r.left, r.bottom - r.top
    hdc = u.GetWindowDC(hwnd)
    mdc = g.CreateCompatibleDC(hdc)
    bmp = g.CreateCompatibleBitmap(hdc, w, h)
    g.SelectObject(mdc, bmp)
    u.PrintWindow(hwnd, mdc, 2)

    class BMI(ctypes.Structure):
        _fields_ = [("biSize", wintypes.DWORD), ("biWidth", wintypes.LONG), ("biHeight", wintypes.LONG),
                    ("biPlanes", wintypes.WORD), ("biBitCount", wintypes.WORD), ("biCompression", wintypes.DWORD),
                    ("biSizeImage", wintypes.DWORD), ("a", wintypes.LONG), ("b", wintypes.LONG),
                    ("c", wintypes.DWORD), ("d", wintypes.DWORD)]
    bmi = BMI()
    bmi.biSize, bmi.biWidth, bmi.biHeight, bmi.biPlanes, bmi.biBitCount = ctypes.sizeof(BMI), w, -h, 1, 32
    buf = ctypes.create_string_buffer(w * h * 4)
    g.GetDIBits(mdc, bmp, 0, h, buf, ctypes.byref(bmi), 0)
    g.DeleteObject(bmp)
    g.DeleteDC(mdc)
    u.ReleaseDC(hwnd, hdc)
    img = Image.frombuffer("RGB", (w, h), buf.raw, "raw", "BGRX", 0, 1).copy()
    fr = wintypes.RECT()
    x0, y0, x1, y1 = 0, 0, w, h
    if ctypes.windll.dwmapi.DwmGetWindowAttribute(hwnd, 9, ctypes.byref(fr), ctypes.sizeof(fr)) == 0:
        x0, y0 = max(0, fr.left - r.left), max(0, fr.top - r.top)
        x1, y1 = min(w, fr.right - r.left), min(h, fr.bottom - r.top)
    return img.crop((x0, y0, x1, y1)), (r.left + x0, r.top + y0)


def rect(target, origin) -> list[int]:
    """위젯 또는 (트리, 줄) 또는 (트리, 줄, 칸)의 자리를 찍은 그림 안 좌표로."""
    l, t = origin
    if isinstance(target, tuple):
        tree, iid = target[0], target[1]
        x, y, w, h = tree.bbox(iid, target[2]) if len(target) > 2 else tree.bbox(iid)
        return [tree.winfo_rootx() + x - l, tree.winfo_rooty() + y - t, w, h]
    return [target.winfo_rootx() - l, target.winfo_rooty() - t, target.winfo_width(), target.winfo_height()]


def shot(name: str, main, *dialogs, rects: dict | None = None) -> None:
    """주 창 + 그 위에 뜬 대화 상자를 한 장으로 찍고, 단추 자리를 기록한다."""
    main.update_idletasks()
    main.update()
    img, origin = grab(main)
    for d in dialogs:
        d.update_idletasks()
        d.update()
        di, (dl, dt) = grab(d)
        at = (dl - origin[0], dt - origin[1])
        sh = Image.new("L", img.size, 0)
        sh.paste(70, (at[0] + 4, at[1] + 8, at[0] + di.width + 4, at[1] + di.height + 8))
        img.paste((10, 20, 50), (0, 0), sh.filter(ImageFilter.GaussianBlur(14)))
        img.paste(di, at)
    OUT.mkdir(parents=True, exist_ok=True)
    img.save(OUT / f"{name}.png")
    RECTS[name] = {k: rect(v, origin) for k, v in (rects or {}).items()}
    print("찍음:", name, img.size, flush=True)


def walk(w):
    yield w
    for c in w.winfo_children():
        yield from walk(c)


def find(root, text: str):
    for w in walk(root):
        try:
            if w.cget("text") == text:
                return w
        except tk.TclError:
            continue
    raise LookupError(text)


def center(parent, d) -> None:
    d.update_idletasks()
    x = parent.winfo_rootx() + (parent.winfo_width() - d.winfo_reqwidth()) // 2
    y = parent.winfo_rooty() + max(10, (parent.winfo_height() - d.winfo_reqheight()) // 2)
    limit = parent.winfo_height() - 60           # 큰 창은 주 창 안에 들어오게 줄인다
    if d.winfo_reqheight() > limit:
        d.geometry(f"{d.winfo_reqwidth()}x{limit}+{x}+{parent.winfo_rooty() + 10}")
    else:
        d.geometry(f"+{x}+{y}")


def main() -> int:
    import config
    base = Path(tempfile.mkdtemp(prefix="gde_video_"))
    config.BASE_DIR = base
    config.PROJECT_DIR = base           # 개발 폴더의 진짜 열쇠를 쓰지 않게
    config.WORK_DIR, config.STATE_PATH = base / "work", base / "state.json"
    config.STAGING_DIR, config.TOKEN_PATH, config.LOG_PATH = base / "st", base / "t.json", base / "a.log"
    import accounts
    import app as app_mod
    from fake_drive import FakeDrive, write_in_place
    app_mod.enable_dpi_awareness()
    app_mod.open_with_default = lambda p: None
    for n in ("showinfo", "showwarning", "showerror"):
        setattr(app_mod.messagebox, n, lambda *a, **k: None)
    app_mod.messagebox.askyesno = lambda *a, **k: True

    # ── 처음 실행 안내 ──
    w = app_mod.DisclaimerWindow()
    for _ in range(25):
        w.update()
        time.sleep(0.02)
    shot("disc_off", w, rects={"check": w.check, "agree": w.agree_btn})
    w.check.invoke()
    w.update()
    shot("disc_on", w, rects={"check": w.check, "agree": w.agree_btn})
    w.destroy()
    del w

    email, school = "teacher@example.com", "school@example.com"
    d, d2 = FakeDrive(email), FakeDrive(school)
    fold = "application/vnd.google-apps.folder"
    folders = {f: d.add_file(f, None, mime=fold) for f in ("2학기 학급 운영", "수업 자료", "업무 문서")}
    files = [("가정통신문 (10월).hwp", 48000, "application/x-hwp"),
             ("학급 명렬표.xlsx", 15000, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
             ("현장체험학습 계획서.docx", 23000, "application/vnd.openxmlformats-officedocument.wordprocessingml.document"),
             ("수업 발표 자료.pptx", 2300000, "application/vnd.openxmlformats-officedocument.presentationml.presentation"),
             ("교육과정 안내.pdf", 830000, "application/pdf"), ("단체 사진.jpg", 3200000, "image/jpeg")]
    metas = {n: d.add_file(n, OLE + b"x" * s, mime=m) for n, s, m in files}
    for n, m in (("1단원 활동지.hwp", "application/x-hwp"), ("2단원 평가 문제.docx", files[2][2]),
                 ("수업 영상 목록.xlsx", files[1][2])):
        d.add_file(n, OLE + b"y" * 9000, parent=folders["수업 자료"]["id"], mime=m)
    hwp = metas["가정통신문 (10월).hwp"]
    for txt in (b"v2 ...", b"v3 ......", b"v4"):
        time.sleep(0.01)
        d.remote_edit(hwp["id"], OLE + txt * 8000)
    trashed = {}
    for n in ("지난 학기 시간표.xlsx", "회의록 초안.docx", "행사 사진", "옛 가정통신문.hwp"):
        f = d.add_file(n, b"x" * 9000, mime=fold if n == "행사 사진" else "application/x-hwp")
        d.trash(f["id"])
        trashed[n] = f
    for n, s in (("3학년 2반 시간표.xlsx", 12000), ("학년 협의회 자료.hwp", 52000), ("방과후 신청서.docx", 21000)):
        d2.add_file(n, b"x" * s, mime="application/octet-stream")
    d2.add_file("학년 공용 폴더", None, mime=fold)

    def pump(app, sec):
        end = time.time() + sec
        while time.time() < end:
            app.update()
            time.sleep(0.02)

    GEOM = "1872x890+8+8"
    # ── 시작 도우미 ──
    a = app_mod.App(drive_factory=lambda acct, cb: d, start_connect=False)
    a.geometry(GEOM)
    pump(a, 1.0)
    a.open_setup()
    center(a, a._setup)
    pump(a, 0.8)
    wz = a._setup
    btns = {"guide": wz.cards[1][1][0], "find": wz.cards[2][1][0], "login": wz.cards[3][1][0]}
    shot("wiz_first", a, wz, rects=btns)
    (base / "client_secret.json").write_text(json.dumps({"installed": {
        "client_id": "123-demo.apps.googleusercontent.com", "client_secret": "demo",
        "project_id": "drive-keeper-12345"}}), encoding="utf-8")
    wz.refresh()
    pump(a, 0.5)
    shot("wiz_key", a, wz, rects=btns)
    wz.destroy()
    a.shutdown()
    del a, wz, btns

    # ── 로그인한 뒤 ──
    st = accounts.AccountStore()
    st.load()
    st.add(email)
    a = app_mod.App(drive_factory=lambda acct, cb: d2 if acct.email == school else d)
    a.geometry(GEOM)
    pump(a, 3.0)
    acct = a.store.accounts[0]
    root_loc = app_mod.loc_mydrive(acct.id)
    row = {m["name"]: i for i, m in a.item_meta.items()}
    clean = find(a, "작업 폴더 정리")
    base_rects = {"nav_acct": (a.nav, f"a:{acct.id}"), "nav_my": (a.nav, f"l:{acct.id}:mydrive"), "tree": a.tree,
                  "nav": a.nav, "track": a.track_tree, "search": a.search, "lock": a.btn_lock,
                  "security": a.btn_security, "tm": a.btn_tm, "add": a.btn_add, "help": a.help_btn, "clean": clean,
                  "hwp": (a.tree, row["가정통신문 (10월).hwp"]), "folder": (a.tree, row["수업 자료"])}
    shot("main_root", a, rects=base_rects)
    file_menu_at = rect((a.tree, row["가정통신문 (10월).hwp"]), grab(a)[1])

    # 폴더 들어가기
    a.open_item(folders["수업 자료"])
    pump(a, 1.0)
    shot("main_folder", a, rects={"up": a.btn_up, "back": a.btn_back})
    a.navigate([root_loc])
    pump(a, 1.0)

    # 파일 열기 → 저장 → 자동 업로드 (실제 동작)
    a.open_item(a.item_meta[next(i for i, m in a.item_meta.items() if m["id"] == hwp["id"])])
    pump(a, 3.0)
    print("열기 상태:", a.status_var.get(), a.sync.rows(), flush=True)
    first = a.track_tree.get_children()[0]
    shot("main_open", a, rects={"status": (a.track_tree, first, "status"), "row": (a.track_tree, first)})
    sm = a.sync
    path = Path(next(iter(sm._recs.values()))["path"])
    write_in_place(path, OLE + b"z" * 50000)
    end = time.time() + 12
    while time.time() < end:
        pump(a, 0.2)
        if any("저장됨" in str(r.get("status")) for r in sm.rows()):
            break
    pump(a, 0.5)
    first = a.track_tree.get_children()[0]
    print("상태:", [r.get("status") for r in sm.rows()], flush=True)
    shot("main_saved", a, rects={"status": (a.track_tree, first, "status"), "row": (a.track_tree, first), "clean": clean})

    # 검색
    a.search.focus_force()
    pump(a, 0.2)
    a.search_var.set("통신")
    pump(a, 0.2)
    shot("search_typed", a, rects={"search": a.search})
    a._on_search(None)
    pump(a, 1.0)
    shot("search_result", a)
    a.navigate([root_loc])
    pump(a, 1.0)

    # 비밀번호 잠금
    a.open_security()
    sd = a._sec_dialog
    center(a, sd)
    pump(a, 0.8)
    save = find(sd, "저장")
    sec_rects = {"new1": sd.new1, "new2": sd.new2, "save": save}
    shot("sec_empty", a, sd, rects=sec_rects)
    sd.new1.insert(0, "1234")
    sd.new2.insert(0, "1234")
    pump(a, 0.2)
    shot("sec_filled", a, sd, rects=sec_rects)
    save.invoke()
    pump(a, 0.8)
    shot("main_secured", a, rects={"lock": a.btn_lock})
    a.lock_now("manual")
    pump(a, 0.8)
    lk = a.lock
    lk_rects = {"entry": lk.entry, "unlock": find(lk, "잠금 풀기"), "forgot": find(lk, "비밀번호를 잊었어요")}
    shot("lock_empty", a, rects=lk_rects)
    lk.entry.insert(0, "1234")
    pump(a, 0.2)
    shot("lock_typed", a, rects=lk_rects)
    lk.try_unlock()
    pump(a, 0.8)

    # 타임머신
    a.open_time_machine()
    tm = a._tm_dialog
    center(a, tm)
    pump(a, 1.6)
    tree = tm.trash_tree
    old = next(i for i in tree.get_children() if "옛 가정통신문" in tree.item(i, "text"))
    restore = find(tm, "선택한 항목 되살리기")
    shot("tm_trash", a, tm, rects={"row": (tree, old), "restore": restore})
    tree.selection_set(old)
    pump(a, 0.3)
    shot("tm_sel", a, tm, rects={"row": (tree, old), "restore": restore})
    tm.restore_selected()
    pump(a, 1.6)
    shot("tm_done", a, tm)
    tm.destroy()
    pump(a, 0.5)
    a.navigate([root_loc])
    pump(a, 1.0)
    a.open_time_machine(next(m for m in a.item_meta.values() if m["id"] == hwp["id"]))
    tm = a._tm_dialog
    center(a, tm)
    pump(a, 1.4)
    vers = tm.ver_tree.get_children()
    tm.ver_tree.selection_set(vers[1] if len(vers) > 1 else vers[0])
    pump(a, 0.3)
    shot("tm_ver", a, tm, rects={"row": (tm.ver_tree, vers[1] if len(vers) > 1 else vers[0]),
                                 "open": tm.btn_ver_open, "restore": tm.btn_ver_restore})
    tm.destroy()
    pump(a, 0.5)

    # 계정 추가
    a._on_add_ok(('{"refresh_token": "demo"}', school))
    pump(a, 2.2)
    acct2 = next(x for x in a.store.accounts if x.email == school)
    a.set_status(f"{school} 계정을 추가했습니다.")
    pump(a, 0.3)
    shot("two_accounts", a, rects={"new": (a.nav, f"a:{acct2.id}"), "mine": (a.nav, f"a:{acct.id}")})
    a.navigate([root_loc])
    pump(a, 1.0)

    # 새 버전 단추 (예시 번호)
    a.update_btn.configure(text="새 버전 2.2.0 업데이트")
    a.update_btn.pack(side="right", padx=4, pady=2)
    pump(a, 0.4)
    shot("main_update", a, rects={"update": a.update_btn, "help": a.help_btn})

    def labels(menu) -> list[str]:
        out = []
        for i in range((menu.index("end") or 0) + 1):
            out.append("" if menu.type(i) == "separator" else str(menu.entrycget(i, "label")))
        return out
    RECTS["_menus"] = {"file": labels(a.menu), "help": labels(a.help_menu), "file_at": file_menu_at}
    try:
        a.shutdown()
    except Exception:
        pass
    (OUT / "rects.json").write_text(json.dumps(RECTS, ensure_ascii=False, indent=1), encoding="utf-8")
    shutil.rmtree(base, ignore_errors=True)
    print("끝", flush=True)
    return 0


if __name__ == "__main__":
    sys.exit(main())
