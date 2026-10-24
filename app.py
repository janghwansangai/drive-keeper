"""드라이브 지킴이 (DriveKeeper) - 구글 드라이브를 비밀번호로 지키며 필요한 파일만 꺼내 쓰는 프로그램.

드라이브를 통째로 마운트하지 않는다.
- 폴더 탐색: API로 목록만 가져옴 (로컬에 파일이 생기지 않음)
- 열기: 더블클릭한 파일 하나만 작업 폴더에 내려받아 기본 프로그램(Word/Excel/한글)으로 연다
- 저장: 파일 변경을 감지해 원래 드라이브 파일에 자동 업로드 (sync.py)
- 여러 구글 계정: 왼쪽 목록에 계정별로 표시. 계정마다 로그인·작업 폴더·기록이 따로 (accounts.py)
- 비밀번호 잠금: 켤 때 / 자리 비움 / 절전 후 / 직접 잠금, 모두 끄기 (security.py)

속도: 창을 먼저 띄우고 구글 라이브러리는 뒤에서 불러온다. 한 번 본 폴더는 바로 보여주고
뒤에서 새로 고친다. 목록·다운로드·업로드는 서로 다른 스레드라 서로 기다리지 않는다.
"""
from __future__ import annotations

import ctypes
import datetime as _dt
import gc
import hashlib
import json
import logging
import logging.handlers
import os
import queue
import re
import shutil
import subprocess
import sys
import threading
import time
import tkinter as tk
import tkinter.font as tkfont
import webbrowser
from pathlib import Path
from tkinter import messagebox, ttk

import accounts
import config
import gtypes
import protect
import security
import sync
from errors import (FileLockedError, friendly_error, http_reason, http_status, is_auth_error,
                    is_network_error)

log = logging.getLogger("app")


# ---------- 시작 준비 ----------
def setup_logging() -> None:
    config.ensure_dirs()
    handler = logging.handlers.RotatingFileHandler(
        config.LOG_PATH, maxBytes=1_000_000, backupCount=3, encoding="utf-8", delay=True)
    handler.setFormatter(logging.Formatter(
        "%(asctime)s %(levelname)s [%(threadName)s] %(name)s: %(message)s"))
    root = logging.getLogger()
    root.setLevel(logging.INFO)
    root.addHandler(handler)
    logging.getLogger("googleapiclient.discovery_cache").setLevel(logging.ERROR)
    logging.getLogger("googleapiclient").setLevel(logging.WARNING)  # 재시도 경고만 남김
    # 로그인 라이브러리가 INFO로 남기는 주소에 로그인 코드가 들어 있으므로 기록하지 않는다
    logging.getLogger("google_auth_oauthlib").setLevel(logging.WARNING)
    logging.getLogger("oauthlib").setLevel(logging.WARNING)
    logging.getLogger("requests_oauthlib").setLevel(logging.WARNING)

    def thread_hook(args):
        log.error("스레드 오류", exc_info=(args.exc_type, args.exc_value, args.exc_traceback))
    threading.excepthook = thread_hook


def enable_dpi_awareness() -> None:
    # 화면 배율(125%, 150%)에서 글자가 흐리게 보이지 않게
    if os.name != "nt":
        return
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass


class SingleInstance:
    """두 개가 동시에 실행되면 같은 파일을 두 번 올리게 되므로 막는다."""

    def __init__(self, path: Path) -> None:
        self.path = path
        self.fh = None

    def acquire(self) -> bool:
        try:
            self.path.parent.mkdir(parents=True, exist_ok=True)
            self.fh = open(self.path, "a+b")
            if os.name == "nt":
                import msvcrt
                self.fh.seek(0)
                msvcrt.locking(self.fh.fileno(), msvcrt.LK_NBLCK, 1)
            else:
                import fcntl
                fcntl.flock(self.fh.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
            return True
        except OSError:
            return False


def activate_existing_window() -> bool:
    if os.name != "nt":
        return False
    try:
        user32 = ctypes.windll.user32
        hwnd = user32.FindWindowW(None, config.APP_TITLE)
        if not hwnd:
            return False
        user32.ShowWindow(hwnd, 9)  # SW_RESTORE
        user32.SetForegroundWindow(hwnd)
        return True
    except Exception:
        return False


# ---------- 표시 도우미 ----------
def human_size(value) -> str:
    try:
        n = float(value)
    except (TypeError, ValueError):
        return ""
    for unit in ("B", "KB", "MB", "GB"):
        if n < 1024:
            return f"{n:.0f} {unit}" if unit == "B" else f"{n:.1f} {unit}"
        n /= 1024
    return f"{n:.1f} TB"


def fmt_time(value: str) -> str:
    if not value:
        return ""
    try:
        dt = _dt.datetime.fromisoformat(value.replace("Z", "+00:00"))
        return dt.astimezone().strftime("%Y-%m-%d %H:%M")
    except (ValueError, OSError):
        return value[:16].replace("T", " ")


_NUM = re.compile(r"([0-9]+)")


def natural_key(name: str):
    # '2반' < '10반' 처럼 숫자를 숫자로 비교
    return [int(t) if t.isascii() and t.isdigit() else t for t in _NUM.split((name or "").casefold())]


def short_list(names: list[str], limit: int = 8) -> str:
    lines = [f" - {n}" for n in names[:limit]]
    if len(names) > limit:
        lines.append(f" 외 {len(names) - limit}개")
    return "\n".join(lines)


def open_with_default(path: Path) -> None:
    """OS 기본 프로그램으로 파일 열기."""
    if os.name == "nt":
        os.startfile(str(path))  # type: ignore[attr-defined]
    elif sys.platform == "darwin":
        subprocess.Popen(["open", str(path)])
    else:
        subprocess.Popen(["xdg-open", str(path)])


def reveal_in_explorer(path: Path) -> None:
    path = Path(path)
    if os.name == "nt":
        if path.is_file():
            subprocess.Popen(f'explorer /select,"{path}"')
        else:
            os.startfile(str(path if path.exists() else path.parent))  # type: ignore[attr-defined]
    elif sys.platform == "darwin":
        subprocess.Popen(["open", "-R", str(path)])
    else:
        subprocess.Popen(["xdg-open", str(path.parent)])


# ---------- 아이콘 (그림 파일 없이 코드로 그림) ----------
class Icons:
    PAGE_COLORS = {
        "word": "#2B579A", "excel": "#217346", "ppt": "#D24726", "hwp": "#2D8CEB",
        "pdf": "#D93025", "image": "#8E44AD", "video": "#C0392B", "audio": "#E67E22",
        "zip": "#8D6E63", "text": "#7F8C8D", "file": "#B0B7BF",
        "gdoc": "#4285F4", "gsheet": "#0F9D58", "gslides": "#F4B400", "gother": "#5F6368",
    }

    def __init__(self, master: tk.Misc, scale: float) -> None:
        self.master = master
        self.scale = scale
        self.images: dict[str, tk.PhotoImage] = {}
        self.images["folder"] = self._fit(self._folder("#F2C14E", "#C8951F", "#F8D775"))
        self.images["drive"] = self._fit(self._folder("#8AB4F8", "#4A78C2", "#AECBFA"))
        self.images["sc_folder"] = self._fit(self._arrow(self._folder("#F2C14E", "#C8951F", "#F8D775")))
        self.images["sc_file"] = self._fit(self._arrow(self._page(self.PAGE_COLORS["file"])))
        for key, color in self.PAGE_COLORS.items():
            self.images[key] = self._fit(self._page(color))
        self.size = self.images["folder"].width()

    def get(self, key: str) -> tk.PhotoImage:
        return self.images.get(key) or self.images["file"]

    def _new(self) -> tk.PhotoImage:
        return tk.PhotoImage(master=self.master, width=16, height=16)

    def _folder(self, mid: str, dark: str, light: str) -> tk.PhotoImage:
        img = self._new()
        img.put(dark, to=(1, 2, 7, 4))
        img.put(dark, to=(1, 3, 15, 14))
        img.put(mid, to=(2, 5, 14, 13))
        img.put(light, to=(2, 4, 14, 5))
        return img

    def _page(self, color: str) -> tk.PhotoImage:
        img = self._new()
        img.put("#8A8F96", to=(3, 1, 13, 15))
        img.put("#FFFFFF", to=(4, 2, 12, 14))
        img.put("#8A8F96", to=(10, 1, 13, 4))
        img.put("#FFFFFF", to=(10, 2, 12, 4))
        img.put("#C9CDD2", to=(10, 3, 12, 4))
        img.put(color, to=(4, 8, 12, 14))
        img.put("#FFFFFF", to=(5, 10, 11, 11))
        return img

    def _arrow(self, img: tk.PhotoImage) -> tk.PhotoImage:
        img.put("#1A73E8", to=(0, 9, 7, 16))
        for x, y in ((1, 14), (2, 13), (3, 12), (4, 11), (2, 11), (3, 11), (4, 12), (4, 13)):
            img.put("#FFFFFF", to=(x, y, x + 1, y + 1))
        return img

    def _fit(self, img: tk.PhotoImage) -> tk.PhotoImage:
        if self.scale >= 1.75:
            return img.zoom(2)
        if self.scale >= 1.25:
            return img.zoom(3).subsample(2)
        return img


# ---------- 백그라운드 작업자 ----------
class Workers:
    """고정 개수의 백그라운드 스레드. 결과는 화면 스레드에서 콜백으로 받는다."""

    def __init__(self, count: int, name: str, post) -> None:
        self.q: queue.Queue = queue.Queue()
        self.post = post
        for i in range(count):
            threading.Thread(target=self._run, name=f"{name}-{i}", daemon=True).start()

    def submit(self, fn, on_ok=None, on_err=None) -> None:
        self.q.put((fn, on_ok, on_err))

    def _run(self) -> None:
        while True:
            fn, on_ok, on_err = self.q.get()
            try:
                result = fn()
            except BaseException as e:  # noqa: BLE001 - 모든 오류를 화면에 전달
                if not isinstance(e, Exception):
                    log.error("작업 중단: %r", e)
                if on_err is not None:
                    self.post(on_err, e)
                else:
                    log.error("작업 오류", exc_info=e)
            else:
                if on_ok is not None:
                    self.post(on_ok, result)


class Location:
    """탐색 위치 하나 (계정의 폴더, 공유 문서함, 공유 드라이브 목록, 검색 결과)."""

    __slots__ = ("kind", "id", "name", "drive_id", "account")

    def __init__(self, kind: str, id: str, name: str, drive_id: str | None = None,
                 account: str | None = None) -> None:
        self.kind = kind
        self.id = id
        self.name = name
        self.drive_id = drive_id
        self.account = account

    @property
    def key(self) -> tuple:
        return (self.account, self.kind, self.id, self.drive_id)

    def __eq__(self, other) -> bool:
        return isinstance(other, Location) and self.key == other.key

    def __hash__(self) -> int:
        return hash(self.key)


def loc_mydrive(account: str | None) -> Location:
    return Location("folder", "root", "내 드라이브", account=account)


def loc_shared(account: str | None) -> Location:
    return Location("shared", "shared", "공유 문서함", account=account)


def loc_drives(account: str | None) -> Location:
    return Location("drives", "drives", "공유 드라이브", account=account)


def loc_none() -> Location:
    return Location("none", "none", "계정 없음")


def _default_drive_factory(account, on_browser_login):
    from drive_client import DriveClient  # 무거운 구글 라이브러리는 창이 뜬 뒤에 불러온다
    return DriveClient(token_path=account.token_path, expected_email=account.email or None,
                       on_browser_login=on_browser_login)


def _default_add_flow(on_browser_login) -> tuple[str, str]:
    from drive_client import login_new_account
    return login_new_account(on_browser_login)


def _default_revoke(creds_json: str) -> None:
    from drive_client import revoke_token
    revoke_token(creds_json)


# 실행 파일·스크립트는 더블클릭해도 실행하지 않는다 (공유받은 악성 파일 대비)
DANGEROUS_EXT = {".exe", ".msi", ".msp", ".bat", ".cmd", ".com", ".scr", ".pif", ".cpl", ".js", ".jse",
                 ".vbs", ".vbe", ".wsf", ".wsh", ".hta", ".ps1", ".psm1", ".lnk", ".reg", ".jar", ".msc",
                 ".url", ".scf", ".appref-ms", ".application", ".gadget", ".chm", ".dll", ".sys", ".inf",
                 ".iso", ".img", ".vhd", ".vhdx"}
# 매크로가 든 오피스 파일: '인터넷에서 받은 파일' 표시를 붙여 오피스가 매크로를 막게 한다
MACRO_EXT = {".docm", ".dotm", ".xlsm", ".xltm", ".xlam", ".pptm", ".potm", ".ppsm", ".ppam"}


def mark_internet_zone(path: Path) -> None:
    """Windows의 '인터넷에서 받은 파일' 표시(Zone.Identifier). 실행·매크로 때 경고/차단이 뜬다."""
    if os.name != "nt":
        return
    try:
        with open(str(path) + ":Zone.Identifier", "w", encoding="ascii") as f:
            f.write("[ZoneTransfer]\r\nZoneId=3\r\n")
    except OSError:
        pass


NATIVE_NOTICE = ("Google 문서/시트/슬라이드는 오피스 파일로 '내보낸 복사본'으로 열립니다.\n"
                 "여기서 저장한 내용은 드라이브 원본에 반영되지 않아요.\n"
                 "원본을 고치려면 브라우저에서 편집하세요. (오른쪽 클릭 → 브라우저에서 열기)")
READONLY_NOTICE = ("이 파일은 보기 권한만 있어서, 저장해도 드라이브에 반영되지 않습니다.\n"
                   "수정하려면 파일 소유자에게 편집 권한을 요청하세요.")
SLEEP_GAP_SEC = 30      # 화면 타이머가 이만큼 멈췄으면 절전(최대 절전)에서 돌아온 것으로 본다


# ---------- 잠금 화면 ----------
# ---------- 디자인: 색·아이콘·버튼 ----------
class T:
    """드라이브 지킴이 색 (로고의 파란 방패에 맞춤)."""
    BG = "#F3F6FB"
    SURFACE = "#FFFFFF"
    SIDEBAR = "#EDF2FA"
    BORDER = "#DCE3EE"
    TEXT = "#1E293B"
    MUTED = "#64748B"
    PRIMARY = "#2563EB"
    PRIMARY_DARK = "#1D4ED8"
    PRIMARY_SOFT = "#DCE8FF"
    HOVER = "#EAF1FF"
    HEADER = "#13245E"
    HEADER_HOVER = "#203A86"
    HEADER_MUTED = "#AFC0F0"
    OK = "#059669"
    WARN = "#B45309"
    DANGER = "#DC2626"
    STRIPE = "#F8FAFD"


# Windows 기본 아이콘 글꼴(Segoe Fluent Icons / MDL2)의 그림 문자. 글꼴이 없으면 오른쪽 글자로 대신함
ICONS = {
    "back": ("", "←"), "forward": ("", "→"), "up": ("", "↑"), "refresh": ("", "⟳"),
    "search": ("", "⌕"), "lock": ("", "🔒"), "history": ("", "↺"), "shield": ("", "◆"),
    "help": ("", "?"), "add": ("", "+"), "open": ("", "▶"), "folder": ("", "▣"),
    "clean": ("", "✕"), "download": ("", "↓"), "key": ("", "✱"), "check": ("", "✓"),
    "person": ("", "●"), "info": ("", "i"),
}


def icon_font_family(root: tk.Misc) -> str | None:
    for fam in ("Segoe Fluent Icons", "Segoe MDL2 Assets"):
        try:
            if tkfont.Font(root=root, family=fam).actual("family") == fam:
                return fam
        except tk.TclError:
            pass
    return None


def logo_image(master: tk.Misc, size: int) -> tk.PhotoImage | None:
    try:
        import brand_assets
        best = min(brand_assets.LOGO_PNG, key=lambda s: abs(s - size))
        return tk.PhotoImage(master=master, data=brand_assets.LOGO_PNG[best])
    except Exception:
        return None


class Tooltip:
    """마우스를 올리면 잠깐 뜨는 설명."""

    def __init__(self, widget: tk.Misc, text: str) -> None:
        self.widget, self.text, self.tip, self.job = widget, text, None, None
        widget.bind("<Enter>", self._schedule, add="+")
        widget.bind("<Leave>", self._hide, add="+")
        widget.bind("<ButtonPress>", self._hide, add="+")

    def _schedule(self, _e=None) -> None:
        self._cancel()
        self.job = self.widget.after(550, self._show)

    def _cancel(self) -> None:
        if self.job:
            try:
                self.widget.after_cancel(self.job)
            except tk.TclError:
                pass
            self.job = None

    def _show(self) -> None:
        if self.tip or not self.text:
            return
        try:
            x = self.widget.winfo_rootx() + 4
            y = self.widget.winfo_rooty() + self.widget.winfo_height() + 4
            self.tip = tk.Toplevel(self.widget)
            self.tip.wm_overrideredirect(True)
            self.tip.wm_geometry(f"+{x}+{y}")
            tk.Label(self.tip, text=self.text, bg="#1E293B", fg="#FFFFFF", padx=8, pady=4,
                     font=tkfont.nametofont("TkTooltipFont")).pack()
        except tk.TclError:
            self.tip = None

    def _hide(self, _e=None) -> None:
        self._cancel()
        if self.tip is not None:
            try:
                self.tip.destroy()
            except tk.TclError:
                pass
            self.tip = None


class IconButton(tk.Frame):
    """아이콘 + 글자 단추 (평평한 현대식). ttk 단추처럼 state(["disabled"]) 를 쓸 수 있다."""

    VARIANTS = {
        "tool": (T.SURFACE, T.HOVER, T.TEXT, T.PRIMARY),
        "side": (T.SIDEBAR, T.HOVER, T.TEXT, T.PRIMARY),
        "header": (T.HEADER, T.HEADER_HOVER, "#FFFFFF", "#FFFFFF"),
        "accent": (T.PRIMARY, T.PRIMARY_DARK, "#FFFFFF", "#FFFFFF"),
    }

    def __init__(self, master, icon: str = "", text: str = "", command=None, variant: str = "tool",
                 tooltip: str = "", app: "App | None" = None) -> None:
        bg, hover, fg, icon_fg = self.VARIANTS[variant]
        super().__init__(master, bg=bg, cursor="hand2")
        self._bg, self._hover, self._fg, self._icon_fg = bg, hover, fg, icon_fg
        self._command = command
        self._disabled = False
        root = app or self.winfo_toplevel()
        fam = getattr(root, "icon_family", None)
        glyph, fallback = ICONS.get(icon, ("", ""))
        s = getattr(root, "scale", 1.0)
        padx, pady = int(8 * s), int(5 * s)
        self.icon_label = None
        if icon:
            self.icon_label = tk.Label(self, text=glyph if fam else fallback, bg=bg, fg=icon_fg,
                                       font=(fam, 11) if fam else tkfont.nametofont("TkDefaultFont"))
            self.icon_label.pack(side="left", padx=(padx, 0 if text else padx), pady=pady)
        self.text_label = None
        if text:
            self.text_label = tk.Label(self, text=text, bg=bg, fg=fg, font=tkfont.nametofont("TkDefaultFont"))
            self.text_label.pack(side="left", padx=(int(5 * s) if icon else padx, padx), pady=pady)
        for w in (self, self.icon_label, self.text_label):
            if w is not None:
                w.bind("<Enter>", self._enter)
                w.bind("<Leave>", self._leave)
                w.bind("<ButtonRelease-1>", self._click)
        if tooltip:
            Tooltip(self, tooltip)

    def _paint(self, bg: str) -> None:
        for w in (self, self.icon_label, self.text_label):
            if w is not None:
                w.configure(bg=bg)

    def _enter(self, _e=None) -> None:
        if not self._disabled:
            self._paint(self._hover)

    def _leave(self, _e=None) -> None:
        self._paint(self._bg)

    def _click(self, e=None) -> None:
        if self._disabled or self._command is None:
            return
        x, y = self.winfo_pointerx() - self.winfo_rootx(), self.winfo_pointery() - self.winfo_rooty()
        if 0 <= x <= self.winfo_width() and 0 <= y <= self.winfo_height():
            self._command()

    def invoke(self) -> None:
        if not self._disabled and self._command is not None:
            self._command()

    def state(self, spec=None):
        if spec is not None:
            for item in spec:
                if item == "disabled":
                    self._disabled = True
                elif item == "!disabled":
                    self._disabled = False
            muted = "#94A3B8" if self._bg != T.HEADER else "#5D6FA8"
            if self.icon_label is not None:
                self.icon_label.configure(fg=muted if self._disabled else self._icon_fg)
            if self.text_label is not None:
                self.text_label.configure(fg=muted if self._disabled else self._fg)
            self.configure(cursor="arrow" if self._disabled else "hand2")
            if self._disabled:
                self._paint(self._bg)
        return ("disabled",) if self._disabled else ()

    def set_text(self, text: str) -> None:
        if self.text_label is not None:
            self.text_label.configure(text=text)


# ---------- 잠금 화면 ----------
class LockScreen(tk.Frame):
    """창 전체를 덮는 잠금 화면. 풀기 전에는 목록·버튼·단축키를 쓸 수 없다."""

    def __init__(self, app: "App", reason: str) -> None:
        super().__init__(app, bg=T.HEADER)
        self.app = app
        self.reason = reason
        s = app.scale
        self.place(x=0, y=0, relwidth=1, relheight=1)
        self.lift()
        card = tk.Frame(self, bg=T.SURFACE, padx=int(40 * s), pady=int(32 * s))
        card.place(relx=0.5, rely=0.46, anchor="center")
        self._logo = logo_image(self, int(96 * s))
        if self._logo is not None:
            tk.Label(card, image=self._logo, bg=T.SURFACE).pack(pady=(0, int(10 * s)))
        base = tkfont.nametofont("TkDefaultFont")
        big = base.copy()
        big.configure(size=base.cget("size") + 8, weight="bold")
        mid = base.copy()
        mid.configure(size=base.cget("size") + 2)
        self._fonts = (big, mid)
        tk.Label(card, text="드라이브 지킴이가 잠겨 있어요", font=big, bg=T.SURFACE, fg=T.TEXT).pack()
        minutes = int(app.security.data.get("idle_minutes") or 0)
        texts = {
            "start": "비밀번호를 넣으면 열립니다.",
            "idle": f"{minutes}분 동안 컴퓨터를 사용하지 않아 잠겼어요.",
            "sleep": "절전 모드에서 돌아와 잠겼어요.",
            "manual": "잠금 버튼으로 잠갔어요.",
            "reset": "보안 설정 파일이 없거나 바뀌었어요.\n구글 계정으로 본인 확인을 한 뒤 새 비밀번호를 정해 주세요.",
        }
        tk.Label(card, text=texts.get(reason, ""), font=mid, bg=T.SURFACE, fg=T.MUTED,
                 justify="center").pack(pady=(int(4 * s), int(18 * s)))
        self.msg = tk.StringVar(value="")
        if reason != "reset":
            self.entry = ttk.Entry(card, show="●", width=26, font=mid, justify="center")
            self.entry.pack(ipady=int(5 * s))
            self.entry.bind("<Return>", lambda e: self.try_unlock())
            ttk.Button(card, text="잠금 해제", style="Accent.TButton", command=self.try_unlock).pack(
                fill="x", pady=(int(10 * s), 0), ipady=int(3 * s))
            tk.Label(card, textvariable=self.msg, bg=T.SURFACE, fg=T.DANGER).pack(pady=(int(8 * s), 0))
            link = tk.Label(card, text="비밀번호를 잊었어요", bg=T.SURFACE, fg=T.PRIMARY, cursor="hand2",
                            font=(base.cget("family"), base.cget("size"), "underline"))
            link.pack(pady=(int(6 * s), 0))
            link.bind("<Button-1>", lambda e: self.forgot())
        else:
            self.entry = None
            ttk.Button(card, text="구글 계정으로 본인 확인", style="Accent.TButton", command=self.forgot).pack(
                fill="x", ipady=int(3 * s))
            tk.Label(card, textvariable=self.msg, bg=T.SURFACE, fg=T.DANGER).pack(pady=(int(8 * s), 0))
        tk.Label(self, text="잠겨 있는 동안에도 열어 둔 파일을 저장하면 드라이브에 계속 올라갑니다.",
                 bg=T.HEADER, fg=T.HEADER_MUTED).place(relx=0.5, rely=0.96, anchor="s")
        try:
            self.grab_set()
        except tk.TclError:
            pass
        if self.entry is not None:
            self.entry.focus_set()

    def try_unlock(self) -> None:
        if self.entry is None:
            return
        ok, wait = self.app.security.check(self.entry.get())
        self.entry.delete(0, "end")
        if ok:
            self.app.unlock()
            return
        if wait:
            self.msg.set(f"여러 번 틀려서 {wait}초 뒤에 다시 할 수 있어요.")
        else:
            self.msg.set(f"비밀번호가 틀렸어요. ({self.app.security.fails()}번)")

    def forgot(self) -> None:
        self.msg.set("브라우저에서 이 프로그램에 등록된 구글 계정으로 로그인해 주세요...")
        self.app.verify_identity_for_reset()


class SecurityDialog(tk.Toplevel):
    """보안 설정: 비밀번호 만들기·바꾸기, 잠금 조건, 모든 잠금 끄기."""

    def __init__(self, app: "App") -> None:
        super().__init__(app)
        self.app = app
        self.sec = app.security
        self.title("보안 설정")
        self.transient(app)
        self.configure(bg=T.BG)
        self.resizable(False, False)
        pad = int(10 * app.scale)
        self.body = ttk.Frame(self, padding=pad)
        self.body.pack(fill="both", expand=True)
        self.msg = tk.StringVar(value="")
        d = self.sec.data
        self.v_enabled = tk.BooleanVar(value=bool(d.get("enabled")) or not self.sec.has_password)
        self.v_start = tk.BooleanVar(value=bool(d.get("lock_on_start")))
        idle = int(d.get("idle_minutes") or 0)
        self.v_idle_on = tk.BooleanVar(value=idle > 0)
        self.v_idle = tk.StringVar(value=str(idle or 10))
        self.v_sleep = tk.BooleanVar(value=bool(d.get("lock_on_sleep")))
        if self.sec.has_password:
            self._build_verify()
        else:
            self._build_settings()
        self.protocol("WM_DELETE_WINDOW", self.destroy)
        self.bind("<Escape>", lambda e: self.destroy())
        try:
            self.grab_set()
        except tk.TclError:
            pass

    def destroy(self) -> None:
        super().destroy()
        if self.app._sec_dialog is self:
            self.app._sec_dialog = None
        # 창 객체의 순환 참조를 화면 스레드에서 바로 정리 (다른 스레드에서 정리되면 Tk 경고가 남)
        if not self.app._closing:
            try:
                self.app.after_idle(gc.collect)
            except tk.TclError:
                pass

    def _clear(self) -> None:
        for w in self.body.winfo_children():
            w.destroy()

    def _build_verify(self) -> None:
        self._clear()
        ttk.Label(self.body, text="보안 설정을 바꾸려면 현재 비밀번호를 입력하세요.").pack(anchor="w")
        row = ttk.Frame(self.body)
        row.pack(fill="x", pady=8)
        self.cur = ttk.Entry(row, show="●", width=26)
        self.cur.pack(side="left")
        self.cur.bind("<Return>", lambda e: self._verify())
        ttk.Button(row, text="확인", command=self._verify).pack(side="left", padx=6)
        ttk.Label(self.body, textvariable=self.msg, foreground="#C62828").pack(anchor="w")
        self.cur.focus_set()

    def _verify(self) -> None:
        ok, wait = self.sec.check(self.cur.get())
        if ok:
            self.msg.set("")
            self._build_settings()
        else:
            self.cur.delete(0, "end")
            self.msg.set(f"{wait}초 뒤에 다시 시도하세요." if wait else "비밀번호가 틀렸습니다.")

    def _build_settings(self) -> None:
        self._clear()
        b = self.body
        has = self.sec.has_password
        ttk.Checkbutton(b, text="잠금 사용", variable=self.v_enabled).grid(row=0, column=0, columnspan=3, sticky="w")
        ttk.Label(b, text="이럴 때 잠그기:").grid(row=1, column=0, columnspan=3, sticky="w", pady=(8, 2))
        ttk.Checkbutton(b, text="프로그램을 켤 때 (컴퓨터를 켠 뒤 처음 실행할 때 포함)",
                        variable=self.v_start).grid(row=2, column=0, columnspan=3, sticky="w", padx=(16, 0))
        ttk.Checkbutton(b, text="컴퓨터를 사용하지 않은 지", variable=self.v_idle_on).grid(
            row=3, column=0, sticky="w", padx=(16, 0))
        ttk.Spinbox(b, from_=1, to=240, width=5, textvariable=self.v_idle).grid(row=3, column=1, sticky="w")
        ttk.Label(b, text="분이 지나면").grid(row=3, column=2, sticky="w")
        ttk.Checkbutton(b, text="절전 모드에서 돌아오면", variable=self.v_sleep).grid(
            row=4, column=0, columnspan=3, sticky="w", padx=(16, 0))
        ttk.Label(b, text="언제든 위쪽 '잠금' 버튼이나 Ctrl+L로 바로 잠글 수 있습니다.",
                  foreground="#555").grid(row=5, column=0, columnspan=3, sticky="w", pady=(2, 10))
        ttk.Label(b, text="비밀번호 만들기 (4자 이상)" if not has else "새 비밀번호 (바꿀 때만 입력)").grid(
            row=6, column=0, columnspan=3, sticky="w")
        self.new1 = ttk.Entry(b, show="●", width=26)
        self.new1.grid(row=7, column=0, columnspan=3, sticky="w", pady=2)
        ttk.Label(b, text="한 번 더 입력").grid(row=8, column=0, columnspan=3, sticky="w")
        self.new2 = ttk.Entry(b, show="●", width=26)
        self.new2.grid(row=9, column=0, columnspan=3, sticky="w", pady=2)
        ttk.Label(b, textvariable=self.msg, foreground="#C62828").grid(row=10, column=0, columnspan=3, sticky="w")
        btns = ttk.Frame(b)
        btns.grid(row=11, column=0, columnspan=3, sticky="e", pady=(10, 0))
        if has:
            ttk.Button(btns, text="모든 잠금 끄기", command=self._disable_all).pack(side="left", padx=(0, 20))
        ttk.Button(btns, text="저장", command=self._save).pack(side="left")
        ttk.Button(btns, text="취소", command=self.destroy).pack(side="left", padx=(6, 0))
        self.new1.focus_set()

    def _save(self) -> None:
        p1, p2 = self.new1.get(), self.new2.get()
        try:
            minutes = int(self.v_idle.get())
        except ValueError:
            minutes = 0
        if self.v_idle_on.get() and not 1 <= minutes <= 240:
            self.msg.set("사용하지 않은 시간은 1~240분으로 정해 주세요.")
            return
        if p1 or p2 or not self.sec.has_password:
            if p1 != p2:
                self.msg.set("두 비밀번호가 다릅니다.")
                return
            if len(p1) < security.MIN_PASSWORD:
                self.msg.set(f"비밀번호는 {security.MIN_PASSWORD}자 이상이어야 합니다.")
                return
        try:
            if p1:
                self.sec.set_password(p1)
            self.sec.update(self.v_enabled.get(), self.v_start.get(),
                            minutes if self.v_idle_on.get() else 0, self.v_sleep.get())
        except Exception as e:
            log.exception("보안 설정 저장 실패")
            self.msg.set(f"저장하지 못했습니다: {e}")
            return
        self.app.security_changed("보안 설정을 저장했습니다." + ("" if self.sec.active else " (잠금 꺼짐)"))
        self.destroy()

    def _disable_all(self) -> None:
        if not messagebox.askyesno("모든 잠금 끄기",
                                   "모든 잠금을 끕니다.\n(비밀번호는 남아 있어서 나중에 다시 켤 수 있습니다)\n\n끌까요?",
                                   parent=self):
            return
        self.sec.disable_all()
        self.app.security_changed("모든 잠금을 껐습니다.")
        self.destroy()


# ---------- 도움말 ----------
def _project_id() -> str:
    try:
        secret = config.find_client_secret()
        if secret:
            return (json.loads(secret.read_text(encoding="utf-8")).get("installed") or {}).get("project_id", "")
    except Exception:
        pass
    return ""


def _console_url(page: str = "auth/audience") -> str:
    """내 Google Cloud 콘솔의 화면 주소 (내 열쇠의 프로젝트가 있으면 그 프로젝트로)."""
    pid = _project_id()
    return f"https://console.cloud.google.com/{page}" + (f"?project={pid}" if pid else "")


def open_guide(part: str = "") -> None:
    """홈페이지의 그림 따라 하기 안내를 연다 (part: install·key·connect·add)."""
    webbrowser.open(config.GUIDE_URL + (f"#{part}" if part else ""))


CONSOLE = {
    "project": "https://console.cloud.google.com/projectcreate",
    "api": "https://console.cloud.google.com/apis/library/drive.googleapis.com",
    "auth": "https://console.cloud.google.com/auth/overview",
    "audience": "https://console.cloud.google.com/auth/audience",
    "client": "https://console.cloud.google.com/auth/clients/create",
}

HELP_TOPICS: dict[str, tuple[str, list[tuple[str, str]]]] = {
    "key": ("내 구글 열쇠 만들기 (처음 한 번, 약 5분)", [
        ("h1", "내 구글 열쇠 만들기"),
        ("p", "드라이브 지킴이는 내 구글 드라이브에 들어갈 때 '나만의 열쇠'를 씁니다.\n"
              "열쇠는 구글 클라우드에서 무료로 만들 수 있고, 처음 한 번만 만들면 됩니다.\n"
              "아래 파란 단추를 차례대로 누르며 따라 하세요. 모두 같은 구글 계정(내 드라이브 계정)으로 하세요."),
        ("h2", "1. 프로젝트 만들기"),
        ("link", "프로젝트 만들기 화면 열기|" + CONSOLE["project"]),
        ("step", "• 프로젝트 이름에 drive-keeper 라고 쓰고 [만들기]를 누릅니다.\n"
                 "• 처음이면 나라 선택·약관 동의가 나옵니다. 대한민국을 고르고 동의하세요."),
        ("h2", "2. 구글 드라이브 기능 켜기"),
        ("link", "Google Drive API 화면 열기|" + CONSOLE["api"]),
        ("step", "• 화면 맨 위 프로젝트 이름이 방금 만든 drive-keeper 인지 확인합니다.\n• 파란 [사용] 단추를 누릅니다."),
        ("h2", "3. 앱 정보 쓰기"),
        ("link", "Google 인증 플랫폼 열기|" + CONSOLE["auth"]),
        ("step", "• [시작하기]를 누릅니다.\n"
                 "• 앱 이름: 드라이브 지킴이 / 사용자 지원 이메일: 내 이메일 → [다음]\n"
                 "• 대상: [외부] → [다음]\n• 연락처 정보: 내 이메일 → [다음]\n"
                 "• 동의 칸에 체크 → [계속] → [만들기]"),
        ("h2", "4. 나를 '테스트 사용자'로 넣기"),
        ("link", "대상(테스트 사용자) 화면 열기|" + CONSOLE["audience"]),
        ("step", "• 아래쪽 '테스트 사용자'의 [+ Add users]를 누릅니다.\n• 내 Gmail 주소를 쓰고 [저장]을 누릅니다.\n"
                 "• 나중에 쓸 다른 계정(예: 학교 계정)도 지금 같이 넣어 두면, 계정 추가가 바로 돼요."),
        ("h2", "5. 열쇠 파일 받기"),
        ("link", "클라이언트(열쇠) 만들기 화면 열기|" + CONSOLE["client"]),
        ("step", "• 애플리케이션 유형: [데스크톱 앱]을 고릅니다. 이름은 아무거나 → [만들기]\n"
                 "• 뜨는 창에서 [JSON 다운로드]를 꼭 누르세요. (이 창에서만 받을 수 있어요!)\n"
                 "  놓쳤다면 그 열쇠를 지우고 5번을 다시 하면 됩니다."),
        ("h2", "6. 프로그램에 열쇠 넣기"),
        ("step", "• 드라이브 지킴이 '시작 도우미'에서 [다운로드 폴더에서 찾기]를 누르면 끝!"),
        ("h2", "알아 두세요"),
        ("note", "• 열쇠 파일(client_secret_....json)은 다른 사람에게 주지 마세요."),
        ("note", "• 이 방식에서는 구글 규칙상 7일마다 한 번 다시 로그인합니다. 때가 되면 브라우저가 저절로 열려요."),
        ("note", "• 화면 모양이 조금 달라도 같은 이름의 단추를 찾아 누르면 됩니다."),
    ]),
    "account": ("다른 구글 계정 추가하기", [
        ("h1", "다른 구글 계정 추가하기"),
        ("p", "내 열쇠는 '테스트 중' 상태라서, 새 계정을 쓰려면 먼저 그 계정을 내 프로젝트의 '테스트 사용자'로 넣어야 합니다."),
        ("h2", "1단계. 추가할 계정을 '테스트 사용자'로 넣기 (계정마다 한 번)"),
        ("link", "내 프로젝트의 테스트 사용자 화면 열기|@audience"),
        ("step", "• 열쇠를 만든 내 계정으로 로그인되어 있는지 확인합니다.\n"
                 "• '테스트 사용자'의 [+ Add users] → 추가할 Gmail 주소 → [저장]"),
        ("h2", "2단계. 프로그램에서 계정 추가하기"),
        ("step", "• 아래 [계정 추가 시작] (왼쪽 목록 맨 아래 '+ 계정 추가'와 같아요)\n"
                 "• 브라우저에서 추가할 계정을 고릅니다. 없으면 '다른 계정 사용'\n"
                 "• 'Google에서 확인하지 않은 앱'이 나오면 파란 단추 말고 왼쪽 글자 [계속]을 누릅니다.\n"
                 "  ([계속]이 없으면 [고급] → '...(으)로 이동'. 내가 만든 앱이라 나오는 안내예요)\n"
                 "• 권한 화면에서 체크 칸이 있으면 체크하고 [계속] → '로그인이 완료되었습니다'가 뜨면 브라우저를 닫고 돌아오세요."),
        ("h2", "잘 안 될 때"),
        ("note", "• '액세스 차단됨' / 'access_denied' → 1단계를 안 했거나 주소에 오타가 있어요."),
        ("note", "• 학교 계정에서 '관리자가 차단함' → 학교 관리자에게 허용을 요청하거나 개인 계정을 쓰세요."),
        ("note", "• 5분 안에 로그인을 마치지 않으면 시간이 초과돼요 → 다시 누르세요."),
        ("h2", "계정 빼기"),
        ("p", "왼쪽 목록의 계정 이름에서 오른쪽 클릭 → '계정 제거'. 이 PC의 로그인 정보만 지워지고, 드라이브의 파일은 그대로예요."),
    ]),
    "lock": ("비밀번호 잠금 사용법", [
        ("h1", "비밀번호 잠금"),
        ("step", "① 위쪽 [보안 설정]을 누릅니다.\n② 비밀번호(4자 이상)를 두 번 씁니다.\n"
                 "③ 잠글 때를 고릅니다: 프로그램을 켤 때 / 컴퓨터를 N분 동안 안 쓰면 / 절전 모드에서 돌아오면\n"
                 "④ [저장]. 이제 위쪽 [잠금] 단추나 Ctrl+L로 언제든 바로 잠글 수 있어요."),
        ("note", "• 잠겨 있어도 열어 둔 파일을 저장하면 드라이브에 계속 올라가요."),
        ("note", "• 모두 끄려면 [보안 설정] → [모든 잠금 끄기]"),
        ("note", "• 비밀번호를 잊으면 잠금 화면의 '비밀번호를 잊었어요' → 등록된 구글 계정으로 로그인하면 새로 정할 수 있어요."),
    ]),
    "timemachine": ("타임머신 사용법", [
        ("h1", "타임머신: 지운 파일·예전 내용 되돌리기"),
        ("h2", "실수로 지운 파일 되살리기"),
        ("step", "① 위쪽 [타임머신] → '지운 파일 되살리기'\n② 파일을 고르고 [선택한 항목 되살리기] → 원래 폴더로 돌아가요."),
        ("note", "• 구글 드라이브는 지운 파일을 휴지통에 30일 동안 보관해요. 휴지통을 비우면 되살릴 수 없어요."),
        ("h2", "잘못 저장한 파일을 예전 내용으로"),
        ("step", "① 파일을 오른쪽 클릭 → '이전 버전 보기 (타임머신)'\n② 버전을 고르고 [이 버전 열어보기]로 확인\n"
                 "③ 맞으면 [이 버전으로 되돌리기]. 지금 내용도 기록에 남아요."),
        ("note", "• 일반 파일(한글·Word·Excel 등)의 예전 버전은 30일 동안, 최대 100개까지 보관돼요."),
        ("note", "• Google 문서·시트·슬라이드는 브라우저의 '파일 → 버전 기록'에서 되돌려요."),
    ]),
}


class HelpDialog(tk.Toplevel):
    def __init__(self, app: "App", topic: str) -> None:
        super().__init__(app)
        self.app = app
        title, parts = HELP_TOPICS[topic]
        self.title(title)
        self.transient(app)
        self.configure(bg=T.BG)
        s = app.scale
        self.geometry(f"{int(780 * s)}x{int(640 * s)}")
        frame = ttk.Frame(self, padding=int(12 * s))
        frame.pack(fill="both", expand=True)
        btns = ttk.Frame(frame)
        btns.pack(side="bottom", fill="x", pady=(int(10 * s), 0))
        ttk.Button(btns, text="닫기", command=self.destroy).pack(side="right")
        if topic == "account":
            ttk.Button(btns, text="계정 추가 시작", style="Accent.TButton", command=self._start_add).pack(
                side="right", padx=6)
            ttk.Button(btns, text="그림으로 보기", command=lambda: open_guide("add")).pack(side="right")
        if topic == "key":
            ttk.Button(btns, text="시작 도우미 열기", style="Accent.TButton",
                       command=lambda: (self.destroy(), app.open_setup())).pack(side="right", padx=6)
            ttk.Button(btns, text="그림으로 따라 하기", command=lambda: open_guide("key")).pack(side="right")
        base = tkfont.nametofont("TkDefaultFont")
        f_text = base.copy()
        f_text.configure(size=base.cget("size") + 1)
        f_h1 = base.copy()
        f_h1.configure(size=base.cget("size") + 7, weight="bold")
        f_h2 = base.copy()
        f_h2.configure(size=base.cget("size") + 2, weight="bold")
        self._fonts = (f_text, f_h1, f_h2)
        text = tk.Text(frame, wrap="word", font=f_text, relief="flat", padx=int(16 * s), pady=int(12 * s),
                       background=T.SURFACE, foreground=T.TEXT, cursor="arrow", highlightthickness=1,
                       highlightbackground=T.BORDER, spacing1=int(2 * s))
        sb = ttk.Scrollbar(frame, orient="vertical", command=text.yview)
        text.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        text.pack(side="left", fill="both", expand=True)
        text.tag_configure("h1", font=f_h1, foreground=T.TEXT, spacing3=int(8 * s))
        text.tag_configure("h2", font=f_h2, foreground=T.PRIMARY, spacing1=int(16 * s), spacing3=int(4 * s))
        text.tag_configure("step", lmargin1=int(14 * s), lmargin2=int(28 * s), spacing1=int(3 * s))
        text.tag_configure("note", lmargin1=int(14 * s), lmargin2=int(26 * s), foreground=T.MUTED, spacing1=int(2 * s))
        text.tag_configure("p", spacing1=int(2 * s), spacing3=int(4 * s), foreground="#334155")
        for tag, line in parts:
            if tag == "link":
                label, url = line.split("|", 1)
                if url == "@audience":
                    url = _console_url("auth/audience")
                text.insert("end", "   ")
                btn = ttk.Button(text, text="↗ " + label, style="Link.TButton",
                                 command=lambda u=url: webbrowser.open(u))
                text.window_create("end", window=btn, padx=2, pady=4)
                text.insert("end", "\n")
            else:
                text.insert("end", line + "\n", tag)
        text.configure(state="disabled")
        self.text = text
        self.bind("<Escape>", lambda e: self.destroy())

    def _start_add(self) -> None:
        self.destroy()
        self.app.add_account()


class SetupWizard(tk.Toplevel):
    """처음 쓰는 사람을 위한 시작 도우미: ① 내 열쇠 만들기 ② 열쇠 불러오기 ③ 로그인."""

    def __init__(self, app: "App") -> None:
        super().__init__(app)
        self.app = app
        self.title("드라이브 지킴이 시작 도우미")
        self.transient(app)
        self.configure(bg=T.BG)
        self.resizable(False, False)
        s = app.scale
        self.s = s
        head = tk.Frame(self, bg=T.HEADER, padx=int(24 * s), pady=int(18 * s))
        head.pack(fill="x")
        self._logo = logo_image(self, int(48 * s))
        if self._logo is not None:
            tk.Label(head, image=self._logo, bg=T.HEADER).pack(side="left", padx=(0, int(14 * s)))
        base = tkfont.nametofont("TkDefaultFont")
        big = base.copy()
        big.configure(size=base.cget("size") + 7, weight="bold")
        mid = base.copy()
        mid.configure(size=base.cget("size") + 2, weight="bold")
        self._fonts = (big, mid)
        tt = tk.Frame(head, bg=T.HEADER)
        tt.pack(side="left")
        tk.Label(tt, text="드라이브 지킴이를 시작해요", font=big, bg=T.HEADER, fg="#FFFFFF").pack(anchor="w")
        tk.Label(tt, text="아래 세 단계를 차례대로 하면 됩니다. 처음 한 번만 하면 돼요.", bg=T.HEADER,
                 fg=T.HEADER_MUTED).pack(anchor="w")
        body = tk.Frame(self, bg=T.BG, padx=int(24 * s), pady=int(18 * s))
        body.pack(fill="both", expand=True)
        self.status = {}
        self.cards = {}
        self._card(body, 1, "내 구글 열쇠 만들기", "구글 클라우드에서 나만의 열쇠 파일을 만들어요 (처음 한 번, 약 5분).\n"
                   "[그림으로 따라 하기]를 누르면 한 단계씩 그림으로 보여 줘요.",
                   [("그림으로 따라 하기", lambda: open_guide("key"), True),
                    ("글로 보기", lambda: app.show_help("key"), False)])
        self._card(body, 2, "열쇠 파일 불러오기", "받은 client_secret_....json 파일을 프로그램에 넣어요.\n"
                   "[다운로드 폴더에서 찾기]를 누르면 저절로 찾아 줘요.",
                   [("다운로드 폴더에서 찾기", self.find_key, True), ("직접 고르기...", self.pick_key, False)])
        self._card(body, 3, "내 구글 계정으로 로그인", "브라우저가 열리면 열쇠를 만든 구글 계정을 고르세요.\n"
                   "'확인하지 않은 앱'이 나오면 파란 단추 말고 왼쪽 글자 [계속]을 눌러요.",
                   [("로그인하기", self.login, True)])
        foot = tk.Frame(self, bg=T.BG, padx=int(24 * s), pady=int(0 * s))
        foot.pack(fill="x", pady=(0, int(16 * s)))
        ttk.Button(foot, text="나중에 하기", command=self.destroy).pack(side="right")
        self.refresh()
        self.bind("<Escape>", lambda e: self.destroy())

    def _card(self, parent, num: int, title: str, desc: str, actions) -> None:
        s = self.s
        card = tk.Frame(parent, bg=T.SURFACE, highlightthickness=1, highlightbackground=T.BORDER,
                        padx=int(18 * s), pady=int(14 * s))
        card.pack(fill="x", pady=(0, int(12 * s)))
        badge = tk.Label(card, text=str(num), width=2, bg=T.PRIMARY, fg="#FFFFFF", font=self._fonts[1])
        badge.grid(row=0, column=0, rowspan=2, sticky="n", padx=(0, int(14 * s)))
        tk.Label(card, text=title, font=self._fonts[1], bg=T.SURFACE, fg=T.TEXT).grid(row=0, column=1, sticky="w")
        tk.Label(card, text=desc, bg=T.SURFACE, fg=T.MUTED, justify="left").grid(row=1, column=1, sticky="w")
        st = tk.StringVar(value="")
        tk.Label(card, textvariable=st, bg=T.SURFACE, fg=T.OK).grid(row=2, column=1, sticky="w", pady=(int(4 * s), 0))
        btns = tk.Frame(card, bg=T.SURFACE)
        btns.grid(row=0, column=2, rowspan=3, sticky="e", padx=(int(18 * s), 0))
        made = []
        for label, cmd, primary in actions:
            b = ttk.Button(btns, text=label, command=cmd, style="Accent.TButton" if primary else "TButton")
            b.pack(fill="x", pady=2)
            made.append(b)
        card.grid_columnconfigure(1, weight=1)
        self.status[num] = st
        self.cards[num] = (badge, made)

    def refresh(self) -> None:
        has_key = config.find_client_secret() is not None
        if has_key:
            pid = _project_id()
            self.status[2].set("✓ 열쇠가 들어 있어요" + (f" (프로젝트: {pid})" if pid else ""))
            self.cards[2][0].configure(bg=T.OK, text="✓")
            self.cards[1][0].configure(bg=T.OK, text="✓")
        for b in self.cards[3][1]:
            b.state(["!disabled"] if has_key else ["disabled"])
        if not has_key:
            self.status[3].set("")
        if self.app.store.accounts:
            self.status[3].set("✓ 로그인했어요: " + ", ".join(a.label for a in self.app.store.accounts))
            self.cards[3][0].configure(bg=T.OK, text="✓")

    def _use_key(self, path: Path) -> None:
        try:
            pid = config.import_client_secret(path)
        except ValueError as e:
            messagebox.showerror("열쇠 파일이 아니에요", str(e), parent=self)
            return
        log.info("열쇠 파일을 불러옴 (프로젝트 %s)", pid or "?")
        self.refresh()

    def find_key(self) -> None:
        found = config.find_downloaded_secrets()
        if not found:
            messagebox.showinfo("열쇠 파일을 찾지 못했어요",
                                "다운로드 폴더와 바탕화면에 client_secret_....json 파일이 없어요.\n\n"
                                "1단계 안내의 5번(열쇠 파일 받기)을 하고 다시 눌러 주세요.\n"
                                "다른 곳에 저장했다면 [직접 고르기...]를 누르세요.", parent=self)
            return
        path = found[0]
        if len(found) > 1 and not messagebox.askyesno(
                "열쇠 파일", f"가장 최근에 받은 열쇠 파일을 쓸까요?\n\n{path.name}", parent=self):
            self.pick_key()
            return
        self._use_key(path)

    def pick_key(self) -> None:
        from tkinter import filedialog
        name = filedialog.askopenfilename(
            parent=self, title="구글에서 받은 열쇠 파일(client_secret_....json) 고르기",
            initialdir=str(config.downloads_dir()), filetypes=[("구글 열쇠 파일", "*.json"), ("모든 파일", "*.*")])
        if name:
            self._use_key(Path(name))

    def login(self) -> None:
        if config.find_client_secret() is None:
            return
        self.app.add_account()


class AboutDialog(tk.Toplevel):
    def __init__(self, app: "App") -> None:
        super().__init__(app)
        self.title("드라이브 지킴이 정보")
        self.transient(app)
        self.configure(bg=T.SURFACE)
        self.resizable(False, False)
        s = app.scale
        box = tk.Frame(self, bg=T.SURFACE, padx=int(32 * s), pady=int(24 * s))
        box.pack()
        self._logo = logo_image(self, int(96 * s))
        if self._logo is not None:
            tk.Label(box, image=self._logo, bg=T.SURFACE).pack()
        base = tkfont.nametofont("TkDefaultFont")
        big = base.copy()
        big.configure(size=base.cget("size") + 6, weight="bold")
        self._font = big
        tk.Label(box, text="드라이브 지킴이", font=big, bg=T.SURFACE, fg=T.TEXT).pack(pady=(int(8 * s), 0))
        tk.Label(box, text=f"버전 {config.VERSION}", bg=T.SURFACE, fg=T.MUTED).pack()
        tk.Label(box, text="구글 드라이브를 비밀번호로 지키며 필요한 파일만 꺼내 쓰는 프로그램",
                 bg=T.SURFACE, fg=T.TEXT).pack(pady=(int(10 * s), int(4 * s)))
        tk.Label(box, text=f"데이터 폴더: {config.BASE_DIR}", bg=T.SURFACE, fg=T.MUTED).pack()
        row = tk.Frame(box, bg=T.SURFACE)
        row.pack(pady=(int(14 * s), 0))
        ttk.Button(row, text="홈페이지", command=lambda: webbrowser.open(config.SITE_URL)).pack(side="left", padx=3)
        ttk.Button(row, text="업데이트 확인", command=lambda: app.check_update(manual=True)).pack(side="left", padx=3)
        ttk.Button(row, text="닫기", command=self.destroy).pack(side="left", padx=3)
        self.bind("<Escape>", lambda e: self.destroy())


# ---------- 타임머신 ----------
class TimeMachineDialog(tk.Toplevel):
    """지운 파일 되살리기(드라이브 휴지통)와 이전 버전으로 되돌리기(드라이브 버전 기록)."""

    def __init__(self, app: "App", account_id: str, target: dict | None = None) -> None:
        super().__init__(app)
        self.app = app
        self.account_id = account_id
        self.target: dict | None = None
        self.trash_items: list[dict] = []
        self.revs: list[dict] = []
        acct = app._acct(account_id)
        self.title("타임머신 - 지운 파일·이전 버전 되돌리기")
        self.transient(app)
        self.configure(bg=T.BG)
        s = app.scale
        self.geometry(f"{int(900 * s)}x{int(560 * s)}")
        pad = int(8 * s)
        top = ttk.Frame(self, padding=(pad, pad, pad, 0))
        top.pack(fill="x")
        ttk.Label(top, text=f"계정: {acct.label if acct else ''}").pack(side="left")
        ttk.Button(top, text="사용법", command=lambda: app.show_help("timemachine")).pack(side="right")
        self.nb = ttk.Notebook(self)
        self.nb.pack(fill="both", expand=True, padx=pad, pady=pad)
        self.trash_tab = ttk.Frame(self.nb, padding=pad)
        self.ver_tab = ttk.Frame(self.nb, padding=pad)
        self.nb.add(self.trash_tab, text="  지운 파일 되살리기  ")
        self.nb.add(self.ver_tab, text="  이전 버전으로 되돌리기  ")
        self._build_trash()
        self._build_versions()
        self.bind("<Escape>", lambda e: self.destroy())
        self.load_trash()
        if target is not None:
            self.show_versions(target)

    @property
    def drive(self):
        return self.app.drives.get(self.account_id)

    def destroy(self) -> None:
        super().destroy()
        if self.app._tm_dialog is self:
            self.app._tm_dialog = None
        if not self.app._closing:
            try:
                self.app.after_idle(gc.collect)
            except tk.TclError:
                pass

    def _alive(self) -> bool:
        try:
            return bool(self.winfo_exists())
        except tk.TclError:
            return False

    # ----- 지운 파일 -----
    def _build_trash(self) -> None:
        t = self.trash_tab
        ttk.Label(t, text="구글 드라이브는 지운 파일을 휴지통에 30일 동안 보관합니다. "
                          "30일이 지나거나 휴지통을 비우면 되살릴 수 없습니다.\n"
                          "최근에 고친 파일부터 보입니다. 찾는 파일이 안 보이면 이름 일부를 넣고 [찾기]를 누르세요.",
                  foreground="#444").pack(anchor="w")
        row = ttk.Frame(t)
        row.pack(fill="x", pady=(6, 4))
        ttk.Label(row, text="이름으로 찾기").pack(side="left")
        self.filter_var = tk.StringVar()
        ent = ttk.Entry(row, textvariable=self.filter_var, width=30)
        ent.pack(side="left", padx=4)
        ent.bind("<Return>", lambda e: self.search_trash())
        ttk.Button(row, text="찾기", command=self.search_trash).pack(side="left", padx=2)
        ttk.Button(row, text="처음부터 다시", command=self._reset_search).pack(side="left", padx=2)
        btns = ttk.Frame(t)
        btns.pack(side="bottom", fill="x", pady=(6, 0))
        self.trash_status = tk.StringVar(value="")
        ttk.Label(btns, textvariable=self.trash_status).pack(side="left")
        ttk.Button(btns, text="브라우저에서 휴지통 열기",
                   command=lambda: webbrowser.open("https://drive.google.com/drive/trash")).pack(side="right")
        ttk.Button(btns, text="선택한 항목 되살리기", command=self.restore_selected).pack(side="right", padx=4)
        self.btn_more = ttk.Button(btns, text="더 불러오기", command=self._fetch_trash_page)
        self.btn_more.pack(side="right", padx=4)
        self.btn_more.state(["disabled"])
        s = self.app.scale
        self.trash_tree = ttk.Treeview(t, columns=("kind", "size", "when"), selectmode="extended")
        self.trash_tree.heading("#0", text="이름")
        self.trash_tree.heading("kind", text="유형")
        self.trash_tree.heading("size", text="크기")
        self.trash_tree.heading("when", text="수정한 날짜")
        self.trash_tree.column("#0", width=int(420 * s))
        self.trash_tree.column("kind", width=int(150 * s))
        self.trash_tree.column("size", width=int(90 * s), anchor="e")
        self.trash_tree.column("when", width=int(150 * s))
        sb = ttk.Scrollbar(t, orient="vertical", command=self.trash_tree.yview)
        self.trash_tree.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self.trash_tree.pack(fill="both", expand=True)
        self.trash_tree.bind("<Double-1>", lambda e: self.restore_selected())

    def search_trash(self) -> None:
        self.load_trash(self.filter_var.get().strip())

    def _reset_search(self) -> None:
        self.filter_var.set("")
        self.load_trash()

    def load_trash(self, query: str = "") -> None:
        """처음부터 다시 읽는다 (query가 있으면 그 이름이 든 것만 드라이브에서 찾음)."""
        self._trash_gen = getattr(self, "_trash_gen", 0) + 1
        self._trash_query = query
        self._trash_next: str | None = None
        self._trash_pages = 0
        self._trash_more = True
        self.trash_items = []
        self._render_trash()
        if self.drive is None:
            self.trash_status.set("아직 이 계정에 연결되지 않았습니다. 잠시 후 '처음부터 다시'를 눌러 주세요.")
            return
        self._fetch_trash_page()

    def _fetch_trash_page(self) -> None:
        drive = self.drive
        if drive is None or not self._trash_more:
            return
        gen, token, query = self._trash_gen, self._trash_next, self._trash_query
        self.btn_more.state(["disabled"])
        self.trash_status.set("찾는 중..." if query else "휴지통을 읽는 중...")
        self.app.io.submit(lambda: drive.list_trash_page(token, query or None),
                           lambda r: self._on_trash_page(gen, r), self._on_trash_error)

    def _on_trash_page(self, gen: int, result) -> None:
        if not self._alive() or gen != self._trash_gen:
            return  # 그 사이 새로 찾기 시작함 → 옛 결과는 버림
        items, nxt = result
        seen = {m["id"] for m in self.trash_items}
        self.trash_items.extend(m for m in items if m["id"] not in seen)
        self._trash_next = nxt
        self._trash_more = bool(nxt)
        self._trash_pages += 1
        self._render_trash()
        # 지운 폴더 안의 파일만 잔뜩 나온 쪽이면 몇 쪽 더 읽어 둔다 (화면이 비어 보이지 않게)
        if self._trash_more and len(self.trash_items) < 40 and self._trash_pages < 8:
            self._fetch_trash_page()
            return
        self.btn_more.state(["!disabled"] if self._trash_more else ["disabled"])
        n = len(self.trash_items)
        if self._trash_query:
            msg = f"'{self._trash_query}' 찾음: {n}개" if n else f"'{self._trash_query}'이(가) 든 지운 파일이 없습니다."
        else:
            msg = f"{n}개 표시 (최근에 고친 순서)" if n else "휴지통이 비어 있습니다."
        if self._trash_more:
            msg += " · [더 불러오기]로 더 볼 수 있음"
        self.trash_status.set(msg)

    def _on_trash_error(self, e: BaseException) -> None:
        if self._alive():
            self.btn_more.state(["!disabled"] if getattr(self, "_trash_more", False) else ["disabled"])
            self.trash_status.set("휴지통을 읽지 못했습니다: " + friendly_error(e).splitlines()[0])

    def _render_trash(self) -> None:
        tree = self.trash_tree
        tree.delete(*tree.get_children())
        for m in self.trash_items:
            when = m.get("trashedTime") or m.get("modifiedTime") or ""
            size = "" if gtypes.is_folderish(m) or gtypes.is_google_native(m) else human_size(m.get("size"))
            tree.insert("", "end", iid=m["id"], text="  " + (m.get("name") or ""),
                        image=self.app.icons.get(gtypes.icon_key(m)),
                        values=(gtypes.kind_label(m), size, fmt_time(when)))

    def restore_selected(self) -> None:
        ids = list(self.trash_tree.selection())
        drive = self.drive
        if not ids or drive is None:
            self.trash_status.set("되살릴 항목을 고르세요.")
            return
        names = [next((m.get("name") for m in self.trash_items if m["id"] == i), i) for i in ids]
        if not messagebox.askyesno("지운 파일 되살리기",
                                   "다음 항목을 휴지통에서 꺼내 원래 폴더로 되돌립니다:\n\n" + short_list(names)
                                   + "\n\n진행할까요?", parent=self):
            return
        self.trash_status.set("되살리는 중...")

        def work():
            done, failed = [], []
            for fid, name in zip(ids, names):
                try:
                    drive.untrash(fid)
                    done.append(fid)
                except Exception as e:
                    failed.append(f"{name}: {friendly_error(e).splitlines()[0]}")
            return done, failed
        self.app.io.submit(work, self._on_restored, self._on_trash_error)

    def _on_restored(self, result) -> None:
        done, failed = result
        log.info("타임머신: %d개 되살림, %d개 실패", len(done), len(failed))
        self.app.after_restore(self.account_id)
        if not self._alive():
            return
        self.trash_items = [m for m in self.trash_items if m["id"] not in done]
        self._render_trash()
        msg = f"{len(done)}개를 되살렸습니다 (원래 폴더로 돌아갔습니다)."
        self.trash_status.set(msg + (f" {len(failed)}개 실패" if failed else ""))
        if failed:
            messagebox.showwarning("일부 실패", "되살리지 못한 항목:\n\n" + short_list(failed), parent=self)

    # ----- 이전 버전 -----
    def _build_versions(self) -> None:
        v = self.ver_tab
        self.ver_title = tk.StringVar(value="파일 목록에서 파일을 오른쪽 클릭 → '이전 버전 보기 (타임머신)'을 고르세요.")
        ttk.Label(v, textvariable=self.ver_title, font=tkfont.nametofont("TkHeadingFont")).pack(anchor="w")
        ttk.Label(v, text="구글 드라이브는 일반 파일(한글·Word·Excel 등)의 이전 버전을 30일 동안, 최대 100개까지 보관합니다.\n"
                          "되돌려도 지금 버전은 기록에 남아서 언제든 다시 돌아올 수 있습니다.",
                  foreground="#444").pack(anchor="w", pady=(2, 6))
        btns = ttk.Frame(v)
        btns.pack(side="bottom", fill="x", pady=(6, 0))
        self.ver_status = tk.StringVar(value="")
        ttk.Label(btns, textvariable=self.ver_status).pack(side="left")
        self.btn_ver_browser = ttk.Button(btns, text="브라우저에서 열기", command=self._open_target_browser)
        self.btn_ver_browser.pack(side="right")
        self.btn_ver_restore = ttk.Button(btns, text="이 버전으로 되돌리기", command=self.restore_version)
        self.btn_ver_restore.pack(side="right", padx=4)
        self.btn_ver_open = ttk.Button(btns, text="이 버전 열어보기", command=self.open_version)
        self.btn_ver_open.pack(side="right")
        s = self.app.scale
        self.ver_tree = ttk.Treeview(v, columns=("who", "size", "note"), selectmode="browse")
        self.ver_tree.heading("#0", text="저장한 때")
        self.ver_tree.heading("who", text="저장한 사람")
        self.ver_tree.heading("size", text="크기")
        self.ver_tree.heading("note", text="")
        self.ver_tree.column("#0", width=int(200 * s))
        self.ver_tree.column("who", width=int(220 * s))
        self.ver_tree.column("size", width=int(100 * s), anchor="e")
        self.ver_tree.column("note", width=int(200 * s))
        sb = ttk.Scrollbar(v, orient="vertical", command=self.ver_tree.yview)
        self.ver_tree.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self.ver_tree.pack(fill="both", expand=True)
        self.ver_tree.bind("<Double-1>", lambda e: self.open_version())
        self._set_version_buttons(False)

    def _set_version_buttons(self, enabled: bool) -> None:
        for b in (self.btn_ver_open, self.btn_ver_restore):
            b.state(["!disabled"] if enabled else ["disabled"])
        self.btn_ver_browser.state(["!disabled"] if self.target else ["disabled"])

    def show_versions(self, meta: dict) -> None:
        self.nb.select(self.ver_tab)
        self.ver_tree.delete(*self.ver_tree.get_children())
        self.revs = []
        drive = self.drive
        if gtypes.is_shortcut(meta):
            target = (meta.get("shortcutDetails") or {}).get("targetId")
            if target and drive is not None:
                self.ver_status.set("바로가기 대상을 찾는 중...")
                self.app.io.submit(lambda: drive.get_meta(target), self.show_versions, self._on_ver_error)
            return
        self.target = meta
        name = meta.get("name") or ""
        self.ver_title.set(f"'{name}'의 이전 버전")
        self._set_version_buttons(False)
        if gtypes.is_folder(meta):
            self.ver_status.set("폴더는 이전 버전이 없습니다. 폴더를 되살리려면 '지운 파일 되살리기' 탭을 쓰세요.")
            return
        if gtypes.is_google_native(meta):
            self.ver_status.set("Google 문서·시트·슬라이드는 [브라우저에서 열기] → '파일 → 버전 기록'에서 되돌립니다.")
            return
        if drive is None:
            self.ver_status.set("아직 이 계정에 연결되지 않았습니다.")
            return
        self.ver_status.set("버전 기록을 읽는 중...")
        fid = meta["id"]
        self.app.io.submit(lambda: drive.list_revisions(fid), lambda revs: self._on_revs(fid, revs),
                           self._on_ver_error)

    def _on_revs(self, file_id: str, revs: list[dict]) -> None:
        if not self._alive() or not self.target or self.target.get("id") != file_id:
            return
        self.revs = revs
        tree = self.ver_tree
        tree.delete(*tree.get_children())
        for i, r in enumerate(reversed(revs)):
            user = r.get("lastModifyingUser") or {}
            note = "지금 버전" if i == 0 else ("계속 보관" if r.get("keepForever") else "")
            tree.insert("", "end", iid=r["id"], text="  " + fmt_time(r.get("modifiedTime", "")),
                        values=(user.get("displayName") or user.get("emailAddress") or "", human_size(r.get("size")),
                                note))
        self.ver_status.set(f"버전 {len(revs)}개" if revs else "저장된 이전 버전이 없습니다.")
        self._set_version_buttons(bool(revs))

    def _on_ver_error(self, e: BaseException) -> None:
        if self._alive():
            self.ver_status.set("버전 기록을 읽지 못했습니다: " + friendly_error(e).splitlines()[0])

    def _selected_rev(self) -> dict | None:
        sel = self.ver_tree.selection()
        if not sel:
            self.ver_status.set("버전을 고르세요.")
            return None
        return next((r for r in self.revs if r["id"] == sel[0]), None)

    def _view_path(self, rev: dict) -> Path:
        name = self.target.get("name") or "파일"
        stem, dot, ext = name.rpartition(".")
        if not dot or not stem:
            stem, ext = name, ""
        else:
            ext = "." + ext
        when = fmt_time(rev.get("modifiedTime", "")).replace(":", "")
        folder = config.BASE_DIR / "versions" / self.account_id / sync.short_dir_name(self.target["id"])
        fname = sync.safe_filename(f"{stem} (이전 버전 {when}){ext}",
                                   config.MAX_LOCAL_PATH - len(str(folder)) - 1)
        return folder / fname

    def open_version(self) -> None:
        rev = self._selected_rev()
        drive = self.drive
        if rev is None or drive is None or self.target is None:
            return
        path = self._view_path(rev)
        fid = self.target["id"]
        self.ver_status.set("이전 버전을 받는 중...")

        def work():
            drive.download_revision(fid, rev, path)
            return path
        self.app.io.submit(work, self._on_version_downloaded, self._on_ver_error)

    def _on_version_downloaded(self, path: Path) -> None:
        if path.suffix.lower() in DANGEROUS_EXT:
            if self._alive():
                self.ver_status.set("실행 파일은 열지 않습니다.")
            self.app._reveal(path)
            return
        try:
            open_with_default(path)
        except OSError as e:
            self.app._dialog("error", "열 수 없음", f"{path.name}\n\n{friendly_error(e)}")
            return
        if self._alive():
            self.ver_status.set("열었습니다. 이 사본에서 저장한 내용은 드라이브에 올라가지 않습니다 (보기용).")

    def restore_version(self) -> None:
        rev = self._selected_rev()
        drive = self.drive
        if rev is None or drive is None or self.target is None:
            return
        if self.revs and rev["id"] == self.revs[-1]["id"]:
            self.ver_status.set("이미 지금 버전입니다.")
            return
        target = self.target
        name = target.get("name") or ""
        in_use = self.app.local_copy_in_use(self.account_id, target["id"])
        warn = ("\n\n⚠ 이 파일이 Word·한글 등에서 열려 있습니다. 닫은 뒤 되돌리는 것이 안전합니다.\n"
                "(열린 채로 저장하면 '충돌 사본'으로 따로 저장됩니다)") if in_use else ""
        when = fmt_time(rev.get("modifiedTime", ""))
        if not messagebox.askyesno("이 버전으로 되돌리기",
                                   f"'{name}'을(를) {when}에 저장한 내용으로 되돌립니다.\n"
                                   "지금 내용도 버전 기록에 남아서 다시 돌아올 수 있습니다." + warn + "\n\n진행할까요?",
                                   icon="warning" if in_use else "question", parent=self):
            return
        self.ver_status.set("되돌리는 중...")
        tmp = config.BASE_DIR / "versions" / "restore" / f"{target['id']}-{rev['id']}"

        def work():
            try:
                drive.download_revision(target["id"], rev, tmp)
                if tmp.stat().st_size <= config.MEMORY_UPLOAD_LIMIT:
                    source = tmp.read_bytes()
                else:
                    source = tmp
                return drive.update_content(target["id"], source, target.get("mimeType") or "application/octet-stream")
            finally:
                try:
                    tmp.unlink()
                except OSError:
                    pass
        self.app.io.submit(work, lambda new: self._on_version_restored(target, when), self._on_ver_error)

    def _on_version_restored(self, target: dict, when: str) -> None:
        log.info("타임머신: 이전 버전으로 되돌림 (%s)", when)
        self.app.after_restore(self.account_id, target["id"])
        self.app.set_status(f"'{target.get('name')}'을(를) {when} 버전으로 되돌렸습니다.")
        if self._alive():
            self.ver_status.set(f"{when} 버전으로 되돌렸습니다.")
            self.show_versions(target)

    def _open_target_browser(self) -> None:
        if self.target:
            self.app.open_in_browser(self.target)


class App(tk.Tk):
    CACHE_MAX = 300

    def __init__(self, drive_factory=None, start_connect: bool = True, add_account_flow=None,
                 revoke=None) -> None:
        super().__init__()
        self.title(config.APP_TITLE)
        config.ensure_dirs()
        self.drive_factory = drive_factory or _default_drive_factory
        self.add_account_flow = add_account_flow or _default_add_flow
        self.revoke = revoke or _default_revoke
        self.ui_queue: queue.Queue = queue.Queue()
        self.store = accounts.AccountStore()
        self.store.load()
        self.drives: dict[str, object] = {}
        self.syncs: dict[str, sync.SyncManager] = {}
        self.connecting: set[str] = set()
        self.auth_trouble: set[str] = set()
        self.cache: dict[tuple, list[dict]] = {}
        self.want: dict[tuple, int] = {}
        self.seq = 0
        first = self.store.accounts[0].id if self.store.accounts else None
        self.stack: list[Location] = [loc_mydrive(first) if first else loc_none()]
        self.back_hist: list[list[Location]] = []
        self.fwd_hist: list[list[Location]] = []
        self.sort_col = "name"
        self.sort_rev = False
        self.item_meta: dict[str, dict] = {}
        self.shown_sig = None
        self.downloading: set[tuple] = set()
        self.noticed: set[str] = set()
        self.track_keys: dict[str, tuple[str, str]] = {}   # 화면 iid -> (계정 id, sync key)
        self.track_status: dict[str, str] = {}
        self._track_refresh_pending = False
        self._closing = False
        self._adding = False
        self._typeahead = ("", 0.0)
        self._pending_open: tuple[str, dict] | None = None
        self.lock: LockScreen | None = None
        self._after_unlock: list = []
        self._last_wall = time.time()
        self._last_activity = time.monotonic()
        self._sec_dialog: SecurityDialog | None = None
        self._tm_dialog: TimeMachineDialog | None = None

        self.security = security.Security()
        self.security.load(lock_marker_seen=self._any_token_lock_marker())
        protect.LOCK_MARKER = self.security.active or self.security.reset_required

        self._load_ui_settings()
        self._setup_style()
        self._build_ui()
        self.report_callback_exception = self._on_tk_error

        self.browse = Workers(2, "browse", self.post)
        self.io = Workers(3, "io", self.post)
        self.io.submit(self._clean_version_copies)
        for acct in self.store.accounts:
            self._start_sync(acct)

        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self.protocol("WM_SAVE_YOURSELF", self._on_session_end)   # Windows 종료·로그오프
        self._pump_id = self.after(30, self._pump)
        self._load_root_caches()
        self._rebuild_nav()
        self._after_nav(fetch=False)
        if self.stack[0].kind != "none" and self.stack[0].key not in self.cache:
            self._show_placeholder("Google 로그인 확인 중...")
        self._refresh_track_view()
        self._update_lock_button()

        if self.security.reset_required:
            self.lock_now("reset")
        elif self.security.active and self.security.data.get("lock_on_start"):
            self.lock_now("start")
        self.after(2000, self._security_tick)
        self.after(5000, self._health_tick)
        if os.name != "nt":
            for seq in ("<Any-KeyPress>", "<Any-ButtonPress>", "<Motion>"):
                self.bind_all(seq, self._touch, add="+")
        if start_connect:
            if self.store.accounts:
                self.after(1, self.connect_all)
            else:
                self.after(300, self._first_run)

    # ======================= 계정별 도우미 =======================
    @property
    def cur_acct(self) -> str | None:
        return self.stack[0].account

    @property
    def drive(self):
        return self.drives.get(self.cur_acct) if self.cur_acct else None

    @drive.setter
    def drive(self, value) -> None:
        if self.cur_acct:
            if value is None:
                self.drives.pop(self.cur_acct, None)
            else:
                self.drives[self.cur_acct] = value

    @property
    def sync(self) -> sync.SyncManager | None:
        return self.syncs.get(self.cur_acct) if self.cur_acct else None

    def _acct(self, account_id: str | None) -> accounts.Account | None:
        return self.store.get(account_id)

    def _acct_label(self, acct: accounts.Account) -> str:
        label = acct.label
        if acct.id in self.auth_trouble:
            label += " (로그인 필요)"
        return label

    def _any_token_lock_marker(self) -> bool:
        for acct in self.store.accounts:
            try:
                if acct.token_path.exists() and protect.read_token(acct.token_path)[1]:
                    return True
            except Exception:
                pass
        return False

    def _start_sync(self, acct: accounts.Account) -> None:
        if acct.id in self.syncs:
            return
        acct.ensure_dirs()
        aid = acct.id
        sm = sync.SyncManager(
            acct.work_dir, acct.state_path, acct.staging_dir,
            get_drive=lambda: self.drives.get(aid),
            notify=lambda ev, payload: self.post(self._on_sync_event, (aid, ev, payload)))
        self.syncs[aid] = sm
        sm.start()

    # ======================= 화면 구성 =======================
    def _setup_style(self) -> None:
        try:
            self.scale = max(1.0, float(self.winfo_fpixels("1i")) / 96.0)
        except tk.TclError:
            self.scale = 1.0
        s = self.scale
        if os.name == "nt":
            try:
                if tkfont.Font(root=self, family="Malgun Gothic").actual("family") == "Malgun Gothic":
                    for nm in ("TkDefaultFont", "TkTextFont", "TkMenuFont", "TkHeadingFont",
                               "TkTooltipFont", "TkCaptionFont", "TkSmallCaptionFont", "TkIconFont"):
                        try:
                            tkfont.nametofont(nm).configure(family="Malgun Gothic")
                        except tk.TclError:
                            pass
            except tk.TclError:
                pass
        self.icon_family = icon_font_family(self)
        self.icons = Icons(self, self.scale)
        base = tkfont.nametofont("TkDefaultFont")
        self.f_bold = base.copy()
        self.f_bold.configure(weight="bold")
        self.f_title = base.copy()
        self.f_title.configure(size=base.cget("size") + 4, weight="bold")
        self.configure(bg=T.BG)
        style = ttk.Style(self)
        try:
            style.theme_use("clam")       # 색을 마음대로 정할 수 있는 테마
        except tk.TclError:
            pass
        line = base.metrics("linespace")
        row = max(line + int(10 * s), self.icons.size + 6)
        style.configure(".", background=T.BG, foreground=T.TEXT, font=base, bordercolor=T.BORDER,
                        lightcolor=T.BORDER, darkcolor=T.BORDER, troughcolor=T.BG, focuscolor=T.PRIMARY)
        style.configure("TFrame", background=T.BG)
        style.configure("Card.TFrame", background=T.SURFACE)
        style.configure("Toolbar.TFrame", background=T.SURFACE)
        style.configure("TLabel", background=T.BG, foreground=T.TEXT)
        style.configure("Card.TLabel", background=T.SURFACE, foreground=T.TEXT)
        style.configure("CardTitle.TLabel", background=T.SURFACE, foreground=T.TEXT, font=self.f_bold)
        style.configure("CardHint.TLabel", background=T.SURFACE, foreground=T.MUTED)
        style.configure("Crumb.TLabel", background=T.SURFACE, foreground=T.MUTED)
        style.configure("Status.TLabel", background=T.SURFACE, foreground=T.MUTED,
                        padding=(int(10 * s), int(4 * s)))
        style.configure("TButton", background=T.SURFACE, foreground=T.TEXT, bordercolor=T.BORDER,
                        padding=(int(12 * s), int(5 * s)), relief="flat")
        style.map("TButton", background=[("disabled", T.BG), ("pressed", T.PRIMARY_SOFT), ("active", T.HOVER)],
                  foreground=[("disabled", "#94A3B8")], bordercolor=[("active", "#B9CCF5")])
        style.configure("Accent.TButton", background=T.PRIMARY, foreground="#FFFFFF", bordercolor=T.PRIMARY,
                        font=self.f_bold)
        style.map("Accent.TButton", background=[("disabled", "#A9BFF3"), ("pressed", T.PRIMARY_DARK),
                                                ("active", T.PRIMARY_DARK)],
                  foreground=[("disabled", "#EEF3FF")], bordercolor=[("active", T.PRIMARY_DARK)])
        style.configure("Link.TButton", background=T.PRIMARY_SOFT, foreground=T.PRIMARY_DARK, bordercolor=T.PRIMARY_SOFT,
                        font=self.f_bold, padding=(int(10 * s), int(3 * s)))
        style.map("Link.TButton", background=[("active", "#C9DBFF")])
        style.configure("Crumb.TButton", background=T.SURFACE, foreground=T.TEXT, bordercolor=T.SURFACE,
                        padding=(int(6 * s), int(2 * s)))
        style.map("Crumb.TButton", background=[("active", T.HOVER)], bordercolor=[("active", T.HOVER)])
        style.configure("TEntry", fieldbackground=T.SURFACE, bordercolor=T.BORDER, padding=int(4 * s))
        style.map("TEntry", bordercolor=[("focus", T.PRIMARY)], lightcolor=[("focus", T.PRIMARY)])
        style.configure("TCheckbutton", background=T.BG)
        style.configure("TSpinbox", fieldbackground=T.SURFACE, arrowsize=int(12 * s))
        style.configure("Treeview", background=T.SURFACE, fieldbackground=T.SURFACE, foreground=T.TEXT,
                        rowheight=row, bordercolor=T.BORDER, borderwidth=0)
        style.map("Treeview", background=[("selected", T.PRIMARY_SOFT)], foreground=[("selected", "#0F2A6B")])
        style.configure("Treeview.Heading", background="#F1F5FB", foreground=T.MUTED, font=self.f_bold,
                        bordercolor=T.BORDER, relief="flat", padding=(int(6 * s), int(5 * s)))
        style.map("Treeview.Heading", background=[("active", T.HOVER)])
        style.configure("Nav.Treeview", background=T.SIDEBAR, fieldbackground=T.SIDEBAR,
                        rowheight=row + int(4 * s), bordercolor=T.SIDEBAR)
        style.map("Nav.Treeview", background=[("selected", T.PRIMARY_SOFT)], foreground=[("selected", "#0F2A6B")])
        style.configure("TNotebook", background=T.BG, bordercolor=T.BORDER)
        style.configure("TNotebook.Tab", background="#E8EEF8", padding=(int(14 * s), int(6 * s)), bordercolor=T.BORDER)
        style.map("TNotebook.Tab", background=[("selected", T.SURFACE)], foreground=[("selected", T.PRIMARY_DARK)])
        style.configure("TPanedwindow", background=T.BG)
        style.configure("Sash", sashthickness=int(6 * s), gripcount=0, background=T.BG)
        style.configure("TScrollbar", background="#E2E8F0", troughcolor=T.SURFACE, bordercolor=T.SURFACE,
                        arrowcolor=T.MUTED, gripcount=0)
        style.map("TScrollbar", background=[("active", "#CBD5E1")])
        style.configure("TMenubutton", background=T.SURFACE)
        self.option_add("*Menu.background", T.SURFACE)
        self.option_add("*Menu.activeBackground", T.PRIMARY_SOFT)
        self.option_add("*Menu.activeForeground", "#0F2A6B")
        self.option_add("*Menu.relief", "flat")
        self.option_add("*TCombobox*Listbox.background", T.SURFACE)
        try:
            imgs = [logo_image(self, n) for n in (64, 32, 16)]
            imgs = [i for i in imgs if i is not None]
            if imgs:
                self._app_icons = imgs
                self.iconphoto(True, *imgs)
        except tk.TclError:
            pass

    def _build_ui(self) -> None:
        s = self.scale
        pad = int(8 * s)

        # ── 맨 위: 로고·이름 / 계정·잠금·보안·도움말 ──
        head = tk.Frame(self, bg=T.HEADER, padx=pad, pady=int(6 * s))
        head.pack(fill="x")
        self._head_logo = logo_image(self, int(28 * s))
        if self._head_logo is not None:
            tk.Label(head, image=self._head_logo, bg=T.HEADER).pack(side="left", padx=(int(4 * s), int(8 * s)))
        tk.Label(head, text="드라이브 지킴이", font=self.f_title, bg=T.HEADER, fg="#FFFFFF").pack(side="left")
        tk.Label(head, text=f"v{config.VERSION}", bg=T.HEADER, fg=T.HEADER_MUTED).pack(side="left", padx=(int(6 * s), 0),
                                                                                     pady=(int(4 * s), 0))
        self.help_btn = IconButton(head, "help", "도움말", self._popup_help, variant="header", app=self)
        self.help_btn.pack(side="right")
        self.btn_security = IconButton(head, "shield", "보안 설정", self.open_security, variant="header", app=self)
        self.btn_security.pack(side="right")
        self.btn_lock = IconButton(head, "lock", "잠금", lambda: self.lock_now("manual"), variant="header", app=self,
                                   tooltip="지금 바로 잠그기 (Ctrl+L)")
        self.btn_lock.pack(side="right")
        self.account_var = tk.StringVar(value="")
        acc = tk.Frame(head, bg=T.HEADER)
        acc.pack(side="right", padx=(0, int(12 * s)))
        if self.icon_family:
            tk.Label(acc, text=ICONS["person"][0], font=(self.icon_family, 10), bg=T.HEADER,
                     fg=T.HEADER_MUTED).pack(side="left", padx=(0, int(4 * s)))
        tk.Label(acc, textvariable=self.account_var, bg=T.HEADER, fg="#E3EAFF").pack(side="left")
        self.help_menu = tk.Menu(self, tearoff=False)
        self.help_menu.add_command(label="그림으로 따라 하기 (홈페이지)", command=open_guide)
        self.help_menu.add_command(label="시작 도우미 (열쇠 넣기·로그인)", command=self.open_setup)
        self.help_menu.add_command(label="내 구글 열쇠 만들기 안내", command=lambda: self.show_help("key"))
        self.help_menu.add_command(label="다른 계정 추가 방법", command=lambda: self.show_help("account"))
        self.help_menu.add_command(label="비밀번호 잠금 사용법", command=lambda: self.show_help("lock"))
        self.help_menu.add_command(label="타임머신 사용법", command=lambda: self.show_help("timemachine"))
        self.help_menu.add_separator()
        self.help_menu.add_command(label="작업 폴더 열기", command=self.reveal_selected_track)
        self.help_menu.add_command(label="기록 파일(app.log) 위치 열기", command=lambda: self._reveal(config.LOG_PATH))
        self.help_menu.add_separator()
        self.help_menu.add_command(label="업데이트 확인", command=lambda: self.check_update(manual=True))
        self.help_menu.add_command(label="홈페이지 열기", command=lambda: webbrowser.open(config.SITE_URL))
        self.help_menu.add_command(label="드라이브 지킴이 정보", command=lambda: AboutDialog(self))

        # ── 도구 줄: 이동 / 경로 / 검색 / 타임머신 ──
        tool = ttk.Frame(self, style="Toolbar.TFrame", padding=(pad, int(5 * s)))
        tool.pack(fill="x")
        tk.Frame(self, bg=T.BORDER, height=1).pack(fill="x")
        self.btn_back = IconButton(tool, "back", "", self.go_back, tooltip="뒤로 (Alt+←)", app=self)
        self.btn_back.pack(side="left")
        self.btn_fwd = IconButton(tool, "forward", "", self.go_forward, tooltip="앞으로 (Alt+→)", app=self)
        self.btn_fwd.pack(side="left")
        self.btn_up = IconButton(tool, "up", "", self.go_up, tooltip="위 폴더로 (Backspace)", app=self)
        self.btn_up.pack(side="left")
        IconButton(tool, "refresh", "", self.refresh, tooltip="새로고침 (F5)", app=self).pack(side="left")
        tk.Frame(tool, bg=T.BORDER, width=1, height=int(22 * s)).pack(side="left", padx=int(8 * s))
        self.btn_tm = IconButton(tool, "history", "타임머신", lambda: self.open_time_machine(), app=self,
                                 tooltip="지운 파일 되살리기 · 예전 버전으로 되돌리기")
        self.btn_tm.pack(side="right")
        sbox = tk.Frame(tool, bg=T.SURFACE, highlightthickness=1, highlightbackground=T.BORDER,
                        highlightcolor=T.PRIMARY)
        sbox.pack(side="right", padx=(0, int(10 * s)))
        if self.icon_family:
            tk.Label(sbox, text=ICONS["search"][0], font=(self.icon_family, 10), bg=T.SURFACE,
                     fg=T.MUTED).pack(side="left", padx=(int(8 * s), 0))
        self.search_var = tk.StringVar()
        self.search = tk.Entry(sbox, textvariable=self.search_var, width=26, relief="flat", bg=T.SURFACE,
                               fg=T.TEXT, insertbackground=T.TEXT, font=tkfont.nametofont("TkDefaultFont"))
        self.search.pack(side="left", padx=int(6 * s), pady=int(5 * s))
        self._search_hint = "드라이브에서 찾기 (Ctrl+F)"
        self._search_placeholder(True)
        self.search.bind("<FocusIn>", lambda e: self._search_placeholder(False))
        self.search.bind("<FocusOut>", lambda e: self._search_placeholder(not self.search_var.get().strip()))
        self.search.bind("<Return>", self._on_search)
        self.search.bind("<Escape>", lambda e: (self.search_var.set(""), self.tree.focus_set()))
        self.crumbs = ttk.Frame(tool, style="Toolbar.TFrame")
        self.crumbs.pack(side="left", fill="x", expand=True)

        # ── 가운데: 왼쪽 계정 목록 + 파일 목록 / 아래: 편집 중인 파일 ──
        vpane = ttk.Panedwindow(self, orient="vertical")
        vpane.pack(fill="both", expand=True)
        hpane = ttk.Panedwindow(vpane, orient="horizontal")
        vpane.add(hpane, weight=4)

        navf = tk.Frame(hpane, bg=T.SIDEBAR)
        tk.Label(navf, text="내 드라이브 계정", bg=T.SIDEBAR, fg=T.MUTED, font=self.f_bold, anchor="w").pack(
            fill="x", padx=int(14 * s), pady=(int(12 * s), int(4 * s)))
        self.nav = ttk.Treeview(navf, show="tree", selectmode="browse", style="Nav.Treeview")
        self.nav.column("#0", width=int(220 * s))
        self.nav.pack(fill="both", expand=True, padx=(int(6 * s), 0))
        self.nav.bind("<ButtonRelease-1>", self._on_nav_click)
        self.nav.bind("<Return>", self._on_nav_click)
        self.nav.bind("<Button-3>", self._on_nav_menu)
        hpane.add(navf, weight=0)
        self.nav_menu = tk.Menu(self, tearoff=False)

        mainf = tk.Frame(hpane, bg=T.SURFACE, highlightthickness=1, highlightbackground=T.BORDER)
        self.tree = ttk.Treeview(mainf, columns=("modified", "kind", "size"), selectmode="browse")
        self.tree.heading("#0", text="이름", command=lambda: self._sort_by("name"))
        self.tree.heading("modified", text="수정한 날짜", command=lambda: self._sort_by("modified"))
        self.tree.heading("kind", text="유형", command=lambda: self._sort_by("kind"))
        self.tree.heading("size", text="크기", command=lambda: self._sort_by("size"))
        widths = self.ui_settings.get("columns") or {}
        self.tree.column("#0", width=int(widths.get("#0", 420 * s)), minwidth=int(120 * s))
        self.tree.column("modified", width=int(widths.get("modified", 140 * s)), minwidth=int(60 * s))
        self.tree.column("kind", width=int(widths.get("kind", 150 * s)), minwidth=int(60 * s))
        self.tree.column("size", width=int(widths.get("size", 90 * s)), minwidth=int(50 * s), anchor="e")
        self.tree.tag_configure("odd", background=T.STRIPE)
        ysb = ttk.Scrollbar(mainf, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=ysb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        ysb.pack(side="right", fill="y")
        hpane.add(mainf, weight=1)

        self.tree.bind("<Double-1>", self._on_double_click)
        self.tree.bind("<Return>", lambda e: self.open_selected())
        self.tree.bind("<BackSpace>", lambda e: self.go_up())
        self.tree.bind("<Button-3>", self._on_tree_menu)
        self.tree.bind("<Key>", self._on_typeahead)

        self.menu = tk.Menu(self, tearoff=False)
        self.menu.add_command(label="열기", command=self.open_selected)
        self.menu.add_command(label="브라우저에서 열기", command=self.open_selected_in_browser)
        self.menu.add_command(label="이전 버전 보기 (타임머신)", command=self.versions_of_selected)
        self.menu.add_separator()
        self.menu.add_command(label="새로고침", command=self.refresh)

        bottom = tk.Frame(vpane, bg=T.BG, padx=pad, pady=int(6 * s))
        vpane.add(bottom, weight=1)
        card = tk.Frame(bottom, bg=T.SURFACE, highlightthickness=1, highlightbackground=T.BORDER)
        card.pack(fill="both", expand=True)
        ch = tk.Frame(card, bg=T.SURFACE, padx=int(12 * s), pady=int(8 * s))
        ch.pack(fill="x")
        tk.Label(ch, text="편집 중인 파일", font=self.f_bold, bg=T.SURFACE, fg=T.TEXT).pack(side="left")
        tk.Label(ch, text="  Word·Excel·한글에서 저장(Ctrl+S)하면 자동으로 드라이브에 올라가요", bg=T.SURFACE,
                 fg=T.MUTED).pack(side="left")
        IconButton(ch, "clean", "작업 폴더 정리", self.cleanup_work, app=self,
                   tooltip="다 쓴 파일의 PC 사본 지우기 (열려 있거나 올리기 전인 파일은 남겨 둠)").pack(side="right")
        IconButton(ch, "folder", "파일 위치", self.reveal_selected_track, app=self).pack(side="right")
        IconButton(ch, "open", "파일 열기", self.open_selected_track, app=self).pack(side="right")
        tbody = tk.Frame(card, bg=T.SURFACE)
        tbody.pack(fill="both", expand=True)
        self.track_tree = ttk.Treeview(tbody, columns=("account", "status"), height=5, selectmode="browse")
        self.track_tree.heading("#0", text="파일")
        self.track_tree.heading("account", text="계정")
        self.track_tree.heading("status", text="상태")
        self.track_tree.column("#0", width=int(330 * s))
        self.track_tree.column("account", width=int(160 * s))
        self.track_tree.column("status", width=int(400 * s))
        tsb = ttk.Scrollbar(tbody, orient="vertical", command=self.track_tree.yview)
        self.track_tree.configure(yscrollcommand=tsb.set)
        self.track_tree.pack(side="left", fill="both", expand=True)
        tsb.pack(side="left", fill="y")
        self.track_tree.bind("<Double-1>", lambda e: self.open_selected_track())
        self.track_tree.bind("<Return>", lambda e: self.open_selected_track())
        self.track_tree.bind("<Button-3>", self._on_track_menu)
        self.track_menu = tk.Menu(self, tearoff=False)
        self.track_menu.add_command(label="파일 열기", command=self.open_selected_track)
        self.track_menu.add_command(label="파일 위치 열기", command=self.reveal_selected_track)
        self.track_menu.add_command(label="이전 버전 보기 (타임머신)", command=self.versions_of_selected_track)
        self.track_menu.add_separator()
        self.track_menu.add_command(label="그래도 업로드 (보류 풀기)", command=self.force_selected_track)

        # ── 맨 아래: 상태 표시줄 ──
        bar = ttk.Frame(self, style="Toolbar.TFrame")
        bar.pack(fill="x", side="bottom", before=vpane)
        tk.Frame(self, bg=T.BORDER, height=1).pack(fill="x", side="bottom", before=vpane)
        self.login_btn = ttk.Button(bar, text="다시 로그인", style="Accent.TButton", command=self.on_login_button)
        self.resume_btn = ttk.Button(bar, text="업로드 다시 시작", style="Accent.TButton", command=self.resume_uploads)
        self.update_btn = ttk.Button(bar, text="새 버전 받기", style="Accent.TButton", command=self._open_update)
        self.status_var = tk.StringVar(value="")
        ttk.Label(bar, textvariable=self.status_var, anchor="w", style="Status.TLabel").pack(
            side="left", fill="x", expand=True)
        self._login_btn_shown = False
        self._resume_btn_shown = False
        self._update_info: dict | None = None

        self.bind_all("<F5>", lambda e: self._when_unlocked(self.refresh))
        self.bind_all("<Alt-Left>", lambda e: self._when_unlocked(self.go_back))
        self.bind_all("<Alt-Right>", lambda e: self._when_unlocked(self.go_forward))
        self.bind_all("<Alt-Up>", lambda e: self._when_unlocked(self.go_up))
        self.bind_all("<Control-f>", lambda e: self._when_unlocked(self._focus_search))
        self.bind_all("<Control-F>", lambda e: self._when_unlocked(self._focus_search))
        self.bind_all("<Control-l>", lambda e: self.lock_now("manual"))
        self.bind_all("<Control-L>", lambda e: self.lock_now("manual"))
        self.tree.focus_set()

    def _search_placeholder(self, show: bool) -> None:
        if show:
            if not self.search_var.get():
                self.search_var.set(self._search_hint)
                self.search.configure(fg=T.MUTED)
        elif self.search_var.get() == self._search_hint:
            self.search_var.set("")
            self.search.configure(fg=T.TEXT)

    def _popup_help(self) -> None:
        b = self.help_btn
        try:
            self.help_menu.tk_popup(b.winfo_rootx(), b.winfo_rooty() + b.winfo_height())
        finally:
            self.help_menu.grab_release()

    # ======================= 시작 도우미·업데이트 =======================
    def open_setup(self) -> None:
        if self.lock is not None:
            return
        w = getattr(self, "_setup", None)
        try:
            if w is not None and w.winfo_exists():
                w.refresh()
                w.lift()
                return
        except tk.TclError:
            pass
        self._setup = SetupWizard(self)

    def check_update(self, manual: bool = False) -> None:
        if not manual:
            last = float(self.ui_settings.get("last_update_check") or 0)
            if time.time() - last < 20 * 3600:
                return
        self.ui_settings["last_update_check"] = time.time()
        import updater

        def done(info):
            if info and updater.is_newer(info["version"]):
                self._update_info = info
                self.update_btn.configure(text=f"새 버전 {info['version']} 받기")
                self.update_btn.pack(side="right", padx=4, pady=2)
                self.set_status(f"새 버전 {info['version']}이 나왔어요. 오른쪽 아래 단추로 받을 수 있어요.")
                if manual and messagebox.askyesno(
                        "새 버전", f"새 버전 {info['version']}이 있어요. (지금 {config.VERSION})\n\n다운로드 페이지를 열까요?",
                        parent=self):
                    self._open_update()
            elif manual:
                messagebox.showinfo("업데이트 확인", f"지금 최신 버전({config.VERSION})을 쓰고 있어요.", parent=self)

        def fail(e):
            log.info("업데이트 확인 실패: %r", e)
            if manual:
                messagebox.showwarning("업데이트 확인", "인터넷 연결을 확인한 뒤 다시 해 주세요.", parent=self)
        self.io.submit(updater.check_latest, done, fail)

    def _open_update(self) -> None:
        info = self._update_info or {}
        webbrowser.open(info.get("url") or config.RELEASES_URL)

    def _when_unlocked(self, fn) -> None:
        if self.lock is None:
            fn()

    def set_status(self, text: str) -> None:
        self.status_var.set(text)

    def _set_login_button(self, show: bool, text: str = "다시 로그인") -> None:
        if show:
            self.login_btn.configure(text=text)
            if not self._login_btn_shown:
                self.login_btn.pack(side="right", padx=4, pady=2)
                self._login_btn_shown = True
        elif self._login_btn_shown:
            self.login_btn.pack_forget()
            self._login_btn_shown = False

    def _show_resume_button(self, show: bool) -> None:
        if show and not self._resume_btn_shown:
            self.resume_btn.pack(side="right", padx=4, pady=2)
            self._resume_btn_shown = True
        elif not show and self._resume_btn_shown:
            self.resume_btn.pack_forget()
            self._resume_btn_shown = False

    def _focus_search(self) -> None:
        self.search.focus_set()
        self.search.select_range(0, "end")

    def _dialog(self, kind: str, title: str, message: str) -> None:
        """안내 창. 잠겨 있으면 (파일 이름 등이 보이지 않게) 잠금을 푼 뒤에 띄운다."""
        if self.lock is not None:
            self._after_unlock.append(lambda: self._dialog(kind, title, message))
            return
        fn = {"info": messagebox.showinfo, "warning": messagebox.showwarning,
              "error": messagebox.showerror}[kind]
        fn(title, message, parent=self)

    # ======================= 스레드 → 화면 =======================
    def post(self, callback, arg=None) -> None:
        self.ui_queue.put((callback, arg))

    def _pump(self) -> None:
        deadline = time.perf_counter() + 0.04
        try:
            while time.perf_counter() < deadline:
                try:
                    callback, arg = self.ui_queue.get_nowait()
                except queue.Empty:
                    break
                try:
                    callback(arg)
                except Exception as e:  # 하나가 실패해도 화면 갱신은 계속
                    self._on_tk_error(type(e), e, e.__traceback__)
        finally:
            if not self._closing:
                self._pump_id = self.after(15 if not self.ui_queue.empty() else 30, self._pump)

    def _on_tk_error(self, exc, val, tb) -> None:
        log.error("화면 처리 오류", exc_info=(exc, val, tb))
        try:
            self.set_status(f"오류: {friendly_error(val)}  (기록: {config.LOG_PATH})")
        except Exception:
            pass

    # ======================= 잠금 =======================
    def _touch(self, _event=None) -> None:
        self._last_activity = time.monotonic()

    def _update_lock_button(self) -> None:
        self.btn_lock.state(["!disabled"] if self.security.active else ["disabled"])

    def lock_now(self, reason: str = "manual") -> None:
        if self.lock is not None or self._closing:
            return
        if reason == "manual" and not self.security.active:
            if not self.security.has_password:
                self._dialog("info", "잠금", "먼저 '보안 설정'에서 비밀번호를 만들어 주세요.")
                self.open_security()
            else:
                self._dialog("info", "잠금", "잠금이 꺼져 있습니다. '보안 설정'에서 '잠금 사용'을 켜 주세요.")
            return
        for dlg in (self._sec_dialog, self._tm_dialog):   # 파일 이름이 보이는 창은 닫는다
            if dlg is not None:
                try:
                    dlg.destroy()
                except tk.TclError:
                    pass
        self._sec_dialog = None
        self._tm_dialog = None
        try:
            self.menu.unpost()
            self.nav_menu.unpost()
            self.track_menu.unpost()
        except tk.TclError:
            pass
        log.info("잠금: %s", reason)
        self.lock = LockScreen(self, reason)

    def unlock(self) -> None:
        if self.lock is None:
            return
        try:
            self.lock.grab_release()
        except tk.TclError:
            pass
        self.lock.destroy()
        self.lock = None
        gc.collect()   # 잠금 화면 객체를 화면 스레드에서 정리
        self._last_wall = time.time()
        self._last_activity = time.monotonic()
        log.info("잠금 해제")
        self.tree.focus_set()
        pending, self._after_unlock = self._after_unlock, []
        for fn in pending:
            try:
                fn()
            except Exception as e:
                self._on_tk_error(type(e), e, e.__traceback__)
        if not self.store.accounts and not self._adding:
            self._first_run()

    def _security_tick(self) -> None:
        try:
            now = time.time()
            gap = now - self._last_wall
            self._last_wall = now
            sec = self.security
            if sec.active and self.lock is None:
                if sec.data.get("lock_on_sleep") and gap > SLEEP_GAP_SEC:
                    self.lock_now("sleep")
                else:
                    minutes = int(sec.data.get("idle_minutes") or 0)
                    if minutes > 0:
                        idle = security.system_idle_seconds()
                        if idle is None:
                            idle = time.monotonic() - self._last_activity
                        if idle >= minutes * 60:
                            self.lock_now("idle")
        except Exception as e:
            self._on_tk_error(type(e), e, e.__traceback__)
        finally:
            if not self._closing:
                self.after(2000, self._security_tick)

    def open_security(self) -> None:
        if self.lock is not None:
            return
        if self._sec_dialog is not None and self._sec_dialog.winfo_exists():
            self._sec_dialog.lift()
            return
        self._sec_dialog = SecurityDialog(self)

    def security_changed(self, message: str) -> None:
        protect.LOCK_MARKER = self.security.active
        self._update_lock_button()
        self.set_status(message)
        log.info("보안 설정 변경: active=%s", self.security.active)
        # 토큰 파일 안의 잠금 표시를 맞춘다 (security.json을 지워 잠금을 푸는 것 방지)
        paths = [a.token_path for a in self.store.accounts]

        def work():
            for p in paths:
                if p.exists():
                    creds, _ = protect.read_token(p)
                    protect.write_token(p, creds)
        self.io.submit(work, on_err=lambda e: log.warning("토큰 잠금 표시 갱신 실패: %r", e))

    def verify_identity_for_reset(self) -> None:
        """비밀번호를 잊었을 때: 등록된 구글 계정으로 로그인하면 새 비밀번호를 정하게 한다."""
        def run():
            try:
                creds_json, email = self.add_account_flow(lambda: None)
            except BaseException as e:  # noqa: BLE001
                self.post(self._on_identity_failed, e)
            else:
                self.post(self._on_identity_ok, (creds_json, email))
        threading.Thread(target=run, name="identity", daemon=True).start()

    def _on_identity_failed(self, e: BaseException) -> None:
        if self.lock is not None:
            self.lock.msg.set("본인 확인 실패: " + friendly_error(e).splitlines()[0])

    def _on_identity_ok(self, arg) -> None:
        creds_json, email = arg
        acct = self.store.find_email(email)
        if self.store.accounts and acct is None:
            if self.lock is not None:
                self.lock.msg.set(f"{email} 은(는) 이 프로그램에 등록된 계정이 아닙니다.")
            return
        if acct is not None:
            protect.write_token(acct.token_path, creds_json)   # 새로 받은 로그인으로 바꿔 둔다
        log.info("본인 확인 성공 → 비밀번호 다시 만들기")
        self.security.reset_required = False
        self.security.data["enabled"] = False
        self.security.data.pop("hash", None)
        self.security.data.pop("salt", None)
        self.security.data["fails"] = 0
        self.security.data["wait_until"] = 0
        try:
            self.security.save()
        except Exception:
            log.exception("보안 설정 저장 실패")
        self.unlock()
        self.security_changed("본인 확인이 끝났습니다. 새 비밀번호를 정해 주세요.")
        if acct is not None and acct.id in self.drives:
            self.drives.pop(acct.id, None)
            self.connect(acct.id)
        self.open_security()

    # ======================= 계정·로그인 =======================
    def _first_run(self) -> None:
        if self.store.accounts or self._closing:
            return
        if self.lock is not None:
            return  # 잠금을 풀면 다시 부름
        self._show_placeholder("시작 도우미를 따라 내 열쇠를 넣고 구글 계정으로 로그인하세요.")
        self.open_setup()

    def connect_all(self) -> None:
        for acct in self.store.accounts:
            self.connect(acct.id)

    def connect(self, account_id: str | None = None) -> None:
        account_id = account_id or self.cur_acct
        acct = self._acct(account_id)
        if acct is None or account_id in self.connecting or self._closing:
            return
        self.connecting.add(account_id)
        if account_id == self.cur_acct:
            self._set_login_button(False)
            self.set_status("Google 로그인 확인 중...")

        def run():
            try:
                drive = self.drive_factory(acct, lambda: self.post(self._on_browser_login, account_id))
            except BaseException as e:  # noqa: BLE001
                self.post(self._on_connect_failed, (account_id, e))
                return
            self.post(self._on_connected, (account_id, drive))   # 목록을 바로 요청하도록 먼저 알림
            try:
                self.post(self._on_email, (account_id, drive.about_email()))
            except Exception as e:
                log.warning("계정 정보를 가져오지 못함: %r", e)
        threading.Thread(target=run, name=f"login-{account_id}", daemon=True).start()

    def _on_browser_login(self, account_id: str | None) -> None:
        acct = self._acct(account_id)
        who = f"{acct.email} " if acct and acct.email else ""
        self.set_status(f"브라우저에서 {who}구글 로그인을 완료해 주세요 (5분 안에). "
                        "'확인되지 않은 앱' 경고가 뜨면 고급 → 이동을 누르세요.")
        if account_id is None or account_id == self.cur_acct:
            self._show_placeholder("브라우저에서 구글 로그인을 기다리는 중...")

    def _on_connected(self, arg) -> None:
        account_id, drive = arg
        self.connecting.discard(account_id)
        if self._acct(account_id) is None:
            return  # 그 사이 계정을 제거함
        self.drives[account_id] = drive
        self.auth_trouble.discard(account_id)
        self._update_nav_labels()
        log.info("연결됨: 계정 %s", account_id)
        if account_id == self.cur_acct:
            self._set_login_button(False)
            self.set_status("연결됨")
            self._after_nav(fetch=True)
        pending = self._pending_open
        if pending is not None and pending[0] == account_id:
            self._pending_open = None
            self.open_item(pending[1], account_id)

    def _on_email(self, arg) -> None:
        account_id, email = arg
        acct = self._acct(account_id)
        if acct is None or not email:
            return
        if acct.email != email:
            self.store.set_email(account_id, email)
        log.info("계정 %s: %s", account_id, email)
        self._update_nav_labels()
        self._update_account_label()
        root = self.cache.get(loc_mydrive(account_id).key)
        if root is not None:
            self._save_root_cache(account_id, root)  # 계정 정보와 함께 다시 저장

    def _on_connect_failed(self, arg) -> None:
        account_id, e = arg
        self.connecting.discard(account_id)
        log.error("연결 실패 (계정 %s)", account_id, exc_info=e)
        if is_auth_error(e):
            self.auth_trouble.add(account_id)
        self._update_nav_labels()
        msg = friendly_error(e)
        if account_id == self.cur_acct:
            self.set_status("연결 실패: " + msg.splitlines()[0])
            if self.stack[-1].key not in self.cache:
                self._show_placeholder("연결하지 못했습니다. 아래 '다시 연결'을 눌러 주세요.")
            self._set_login_button(True, "다시 로그인" if is_auth_error(e) else "다시 연결")
        if isinstance(e, Exception) and not is_network_error(e):
            acct = self._acct(account_id)
            self._dialog("error", "연결 실패", (f"{acct.label}\n\n" if acct else "") + msg)

    def on_login_button(self) -> None:
        account_id = self.cur_acct
        if account_id is None:
            self.add_account()
            return
        drive = self.drives.get(account_id)
        if drive is None:
            self.connect(account_id)
            return
        self._set_login_button(False)
        self.set_status("브라우저에서 다시 로그인해 주세요...")

        def run():
            try:
                drive.relogin()
                email = drive.about_email()
            except BaseException as e:  # noqa: BLE001
                self.post(self._on_relogin_failed, (account_id, e))
            else:
                self.post(self._on_relogin_ok, (account_id, email))
        threading.Thread(target=run, name="relogin", daemon=True).start()

    def _on_relogin_ok(self, arg) -> None:
        account_id, email = arg
        self.auth_trouble.discard(account_id)
        self._on_email((account_id, email))
        self.set_status("다시 로그인했습니다. 대기 중인 업로드를 이어서 진행합니다.")
        self.refresh()

    def _on_relogin_failed(self, arg) -> None:
        account_id, e = arg
        self.set_status("로그인 실패: " + friendly_error(e).splitlines()[0])
        if account_id == self.cur_acct:
            self._set_login_button(True)

    def _note_error(self, e: BaseException, account_id: str | None = None) -> str:
        """오류를 기록하고 안내 문구를 돌려준다. 로그인 문제면 '다시 로그인' 버튼을 띄운다."""
        account_id = account_id or self.cur_acct
        if is_auth_error(e) and account_id:
            self.auth_trouble.add(account_id)
            self._update_nav_labels()
            if account_id == self.cur_acct:
                self._set_login_button(True)
        log.warning("작업 오류: %r", e)
        return friendly_error(e)

    def add_account(self) -> None:
        if self.lock is not None:
            return
        if self._adding:
            self.set_status("계정 추가가 진행 중입니다. 브라우저를 확인해 주세요.")
            return
        self._adding = True
        self.set_status("브라우저에서 추가할 구글 계정을 골라 로그인해 주세요 (5분 안에)...")

        def run():
            try:
                creds_json, email = self.add_account_flow(lambda: self.post(self._on_browser_login, None))
            except BaseException as e:  # noqa: BLE001
                self.post(self._on_add_failed, e)
            else:
                self.post(self._on_add_ok, (creds_json, email))
        threading.Thread(target=run, name="add-account", daemon=True).start()

    def _on_add_ok(self, arg) -> None:
        creds_json, email = arg
        self._adding = False
        acct, created = self.store.add(email)
        acct.ensure_dirs()
        protect.write_token(acct.token_path, creds_json)
        if created:
            self._start_sync(acct)
        self.drives.pop(acct.id, None)          # 새 로그인으로 다시 연결
        self.auth_trouble.discard(acct.id)
        log.info("계정 %s: %s", "추가" if created else "다시 로그인", email)
        self._rebuild_nav()
        self._refresh_track_view()
        self.navigate([loc_mydrive(acct.id)])
        self.connect(acct.id)
        self.set_status(f"{email} 계정을 {'추가했습니다' if created else '다시 로그인했습니다'}.")
        w = getattr(self, "_setup", None)
        try:
            if w is not None and w.winfo_exists():
                w.destroy()
        except tk.TclError:
            pass

    def _on_add_failed(self, e: BaseException) -> None:
        self._adding = False
        log.warning("계정 추가 실패: %r", e)
        msg = friendly_error(e)
        self.set_status("계정을 추가하지 못했습니다: " + msg.splitlines()[0])
        if not self.store.accounts:
            self._show_placeholder("로그인하지 못했습니다. 왼쪽의 '+ 계정 추가'를 눌러 다시 시도하세요.")
        if isinstance(e, Exception) and not is_network_error(e):
            if "테스트 사용자" in msg and self.lock is None:
                if messagebox.askyesno("계정 추가 실패", msg + "\n\n단계별 안내를 볼까요?", parent=self):
                    self.show_help("account")
                return
            self._dialog("error", "계정 추가 실패", msg)

    def remove_account(self, account_id: str) -> None:
        acct = self._acct(account_id)
        sm = self.syncs.get(account_id)
        if acct is None:
            return
        unsynced = sm.unsynced_names() if sm else []
        if unsynced:
            self._dialog("warning", "계정 제거",
                         "아직 드라이브에 올라가지 않은 변경이 있어 지금은 제거할 수 없습니다:\n\n"
                         + short_list(unsynced) + "\n\n업로드가 끝난 뒤 다시 시도해 주세요.")
            return
        in_use = sm.in_use_names() if sm else []
        if in_use:
            self._dialog("warning", "계정 제거", "이 계정의 파일이 아직 열려 있습니다:\n\n" + short_list(in_use)
                         + "\n\n닫은 뒤 다시 시도해 주세요.")
            return
        if not messagebox.askyesno(
                "계정 제거",
                f"{acct.label} 계정을 이 프로그램에서 제거할까요?\n\n"
                "- 이 PC에 저장된 로그인 정보와 작업 폴더의 사본이 지워집니다.\n"
                "- 구글 드라이브의 파일은 그대로 있습니다.", parent=self):
            return
        self.set_status(f"{acct.label} 계정을 제거하는 중...")

        def work():
            removed, kept = sm.cleanup() if sm else (0, [])
            if kept:
                return kept
            if sm:
                sm.stop()
            try:
                creds, _ = protect.read_token(acct.token_path)
                self.revoke(creds)
            except Exception as e:
                log.warning("토큰 취소 생략: %r", e)
            for p in (acct.token_path, acct.state_path, acct.state_path.with_name("state.bak.json"),
                      acct.root_cache_path, acct.dir / "account.json"):
                try:
                    p.unlink()
                except OSError:
                    pass
            for d in (acct.staging_dir, acct.work_dir, acct.dir):
                try:
                    shutil.rmtree(d) if d == acct.staging_dir else d.rmdir()
                except OSError:
                    pass  # 비어 있지 않으면 남겨 둔다 (사용자 파일 보호)
            return []
        self.io.submit(work, lambda kept: self._on_account_removed(account_id, kept),
                       lambda e: self.set_status("계정 제거 실패: " + friendly_error(e)))

    def _on_account_removed(self, account_id: str, kept: list[str]) -> None:
        if kept:
            self._dialog("warning", "계정 제거", "열려 있거나 업로드 전인 파일이 있어 제거하지 않았습니다:\n\n"
                         + short_list(kept))
            return
        acct = self.store.remove(account_id)
        self.drives.pop(account_id, None)
        self.syncs.pop(account_id, None)
        self.auth_trouble.discard(account_id)
        for key in [k for k in self.cache if k[0] == account_id]:
            self.cache.pop(key, None)
        self.back_hist = [st for st in self.back_hist if st[0].account != account_id]
        self.fwd_hist = [st for st in self.fwd_hist if st[0].account != account_id]
        self._rebuild_nav()
        self._refresh_track_view()
        if self.cur_acct == account_id or self.cur_acct is None:
            first = self.store.accounts[0].id if self.store.accounts else None
            self.stack = [loc_mydrive(first) if first else loc_none()]
            self._after_nav(fetch=True)
        self.set_status(f"{acct.label if acct else ''} 계정을 제거했습니다.")

    # ======================= 탐색 =======================
    def navigate(self, stack: list[Location], record: bool = True) -> None:
        if record and stack != self.stack:
            self.back_hist.append(self.stack)
            del self.back_hist[:-100]
            self.fwd_hist.clear()
        self.stack = list(stack)
        self._after_nav(fetch=True)

    def go_back(self) -> None:
        if self.back_hist:
            self.fwd_hist.append(self.stack)
            self.stack = self.back_hist.pop()
            self._after_nav(fetch=True)

    def go_forward(self) -> None:
        if self.fwd_hist:
            self.back_hist.append(self.stack)
            self.stack = self.fwd_hist.pop()
            self._after_nav(fetch=True)

    def go_up(self) -> None:
        if len(self.stack) > 1:
            self.navigate(self.stack[:-1])

    def refresh(self) -> None:
        if self.cur_acct is None:
            return
        if self.drive is None:
            self.connect(self.cur_acct)
            return
        self._fetch(self.stack[-1])

    def _after_nav(self, fetch: bool) -> None:
        loc = self.stack[-1]
        self._update_crumbs()
        self.btn_back.state(["!disabled"] if self.back_hist else ["disabled"])
        self.btn_fwd.state(["!disabled"] if self.fwd_hist else ["disabled"])
        self.btn_up.state(["!disabled"] if len(self.stack) > 1 else ["disabled"])
        self._sync_nav_selection()
        self._update_account_label()
        if loc.kind == "none":
            self._show_placeholder("왼쪽의 '+ 계정 추가'를 눌러 구글 계정으로 로그인하세요.")
            self._set_login_button(False)
            return
        self._set_login_button(loc.account in self.auth_trouble)
        items = self.cache.get(loc.key)
        drive = self.drives.get(loc.account)
        if items is not None:
            self._render(items, keep_view=False)
            self.set_status(f"{len(items)}개 항목 · 새로 고치는 중..." if drive else f"{len(items)}개 항목 (지난번 목록)")
        elif drive is not None:
            self._show_placeholder("불러오는 중...")
            self.set_status("불러오는 중...")
        else:
            self._show_placeholder("Google 로그인 확인 중...")
        if fetch and drive is not None:
            self._fetch(loc)

    def _fetch(self, loc: Location) -> None:
        drive = self.drives.get(loc.account)
        if drive is None or loc.kind == "none":
            return
        self.seq += 1
        seq = self.seq
        self.want[loc.key] = seq
        self.browse.submit(lambda: self._list_for(drive, loc),
                           lambda items: self._on_listing(loc, seq, items),
                           lambda e: self._on_listing_error(loc, seq, e))

    @staticmethod
    def _list_for(drive, loc: Location) -> list[dict]:
        if loc.kind == "folder":
            return drive.list_children(loc.id, loc.drive_id)
        if loc.kind == "shared":
            return drive.list_shared_with_me()
        if loc.kind == "drives":
            return drive.list_drives()
        if loc.kind == "search":
            return drive.search(loc.id)
        raise ValueError(loc.kind)

    def _on_listing(self, loc: Location, seq: int, items: list[dict]) -> None:
        if self.want.get(loc.key) != seq:
            return  # 더 나중에 보낸 요청이 있음 → 늦게 도착한 옛 결과는 버림
        if self._acct(loc.account) is None:
            return
        self.cache.pop(loc.key, None)
        self.cache[loc.key] = items
        if loc.kind == "folder" and loc.id == "root":
            self._save_root_cache(loc.account, items)
        while len(self.cache) > self.CACHE_MAX:
            self.cache.pop(next(iter(self.cache)))
        if self.stack[-1].key != loc.key:
            return  # 이미 다른 폴더로 이동함
        self._render(items, keep_view=True)
        self.set_status(f"{len(items)}개 항목")

    def _on_listing_error(self, loc: Location, seq: int, e: BaseException) -> None:
        if self.want.get(loc.key) != seq or self.stack[-1].key != loc.key:
            return
        msg = self._note_error(e, loc.account)
        if loc.key in self.cache:
            self.set_status(f"새로 고치지 못했습니다 (이전 목록 표시 중): {msg}")
        else:
            self._show_placeholder(f"불러오지 못했습니다: {msg}  (F5로 다시 시도)")
            self.set_status(f"불러오지 못했습니다: {msg}")

    def _show_placeholder(self, text: str) -> None:
        self.tree.delete(*self.tree.get_children())
        self.item_meta = {}
        self.shown_sig = None
        self.tree.insert("", "end", iid="__ph__", text="  " + text)

    def _sorted(self, items: list[dict]) -> list[dict]:
        col, rev = self.sort_col, self.sort_rev
        name_key = lambda m: natural_key(m.get("name", ""))  # noqa: E731
        keys = {
            "name": name_key,
            "modified": lambda m: m.get("modifiedTime", ""),
            "kind": lambda m: (gtypes.kind_label(m), name_key(m)),
            "size": lambda m: int(m.get("size") or 0),
        }
        folders = [m for m in items if gtypes.is_folderish(m)]
        files = [m for m in items if not gtypes.is_folderish(m)]
        folders.sort(key=name_key if col == "size" else keys[col], reverse=rev)
        files.sort(key=keys[col], reverse=rev)
        return folders + files

    def _render(self, items: list[dict], keep_view: bool) -> None:
        loc = self.stack[-1]
        sig = (loc.key, self.sort_col, self.sort_rev, tuple(
            (m.get("id"), m.get("name"), m.get("modifiedTime"), m.get("size")) for m in items))
        if sig == self.shown_sig:
            return  # 바뀐 게 없으면 다시 그리지 않음 (깜박임·스크롤 위치 유지)
        tree = self.tree
        sel = tree.selection()
        sel_id = sel[0] if sel else None
        top = tree.yview()[0] if keep_view else 0.0
        tree.delete(*tree.get_children())
        meta_map: dict[str, dict] = {}
        for m in self._sorted(items):
            iid = m.get("id")
            if not iid or iid in meta_map:
                continue
            meta_map[iid] = m
            size = "" if gtypes.is_folderish(m) or gtypes.is_google_native(m) else human_size(m.get("size"))
            tree.insert("", "end", iid=iid, text="  " + (m.get("name") or ""),
                        image=self.icons.get(gtypes.icon_key(m)), tags=("odd",) if len(meta_map) % 2 == 0 else (),
                        values=(fmt_time(m.get("modifiedTime", "")), gtypes.kind_label(m), size))
        self.item_meta = meta_map
        self.shown_sig = sig
        if not meta_map:
            tree.insert("", "end", iid="__ph__", text="  (비어 있음)")
        if keep_view and sel_id in meta_map:
            tree.selection_set(sel_id)
            tree.focus(sel_id)
        if keep_view:
            tree.yview_moveto(top)
        self._update_headings()

    def _sort_by(self, col: str) -> None:
        if self.sort_col == col:
            self.sort_rev = not self.sort_rev
        else:
            self.sort_col, self.sort_rev = col, False
        items = self.cache.get(self.stack[-1].key)
        if items is not None:
            self._render(items, keep_view=True)
        self._update_headings()

    def _update_headings(self) -> None:
        names = {"name": ("#0", "이름"), "modified": ("modified", "수정한 날짜"),
                 "kind": ("kind", "유형"), "size": ("size", "크기")}
        for col, (cid, label) in names.items():
            arrow = (" ▼" if self.sort_rev else " ▲") if col == self.sort_col else ""
            self.tree.heading(cid, text=label + arrow)

    def _update_crumbs(self) -> None:
        for w in self.crumbs.winfo_children():
            w.destroy()
        stack = self.stack
        acct = self._acct(self.cur_acct)
        if acct is not None and len(self.store.accounts) > 1:
            ttk.Button(self.crumbs, text=acct.short, style="Crumb.TButton",
                       command=lambda a=acct.id: self.navigate([loc_mydrive(a)])).pack(side="left", padx=1)
            ttk.Label(self.crumbs, text="›", style="Crumb.TLabel").pack(side="left")
        start = max(0, len(stack) - 5)
        if start > 0:
            ttk.Label(self.crumbs, text="… ›", style="Crumb.TLabel").pack(side="left")
        for i in range(start, len(stack)):
            if i > start:
                ttk.Label(self.crumbs, text="›", style="Crumb.TLabel").pack(side="left")
            name = stack[i].name
            if len(name) > 30:
                name = name[:28] + "…"
            ttk.Button(self.crumbs, text=name, style="Crumb.TButton",
                       command=lambda i=i: self.navigate(self.stack[:i + 1])).pack(side="left", padx=1)

    def _update_account_label(self) -> None:
        acct = self._acct(self.cur_acct)
        self.account_var.set(acct.email if acct else "")

    # ----- 왼쪽 계정·위치 목록 -----
    def _rebuild_nav(self) -> None:
        nav = self.nav
        nav.delete(*nav.get_children())
        for acct in self.store.accounts:
            aid = f"a:{acct.id}"
            nav.insert("", "end", iid=aid, text=" " + self._acct_label(acct), image=self.icons.get("drive"),
                       open=True)
            nav.insert(aid, "end", iid=f"l:{acct.id}:mydrive", text=" 내 드라이브", image=self.icons.get("folder"))
            nav.insert(aid, "end", iid=f"l:{acct.id}:shared", text=" 공유 문서함", image=self.icons.get("sc_folder"))
            nav.insert(aid, "end", iid=f"l:{acct.id}:drives", text=" 공유 드라이브", image=self.icons.get("drive"))
        nav.insert("", "end", iid="add", text=" + 계정 추가")
        self._sync_nav_selection()
        self._update_track_columns()

    def _update_nav_labels(self) -> None:
        for acct in self.store.accounts:
            iid = f"a:{acct.id}"
            if self.nav.exists(iid):
                self.nav.item(iid, text=" " + self._acct_label(acct))
        self._update_crumbs()

    def _sync_nav_selection(self) -> None:
        first = self.stack[0]
        target = None
        if first.account:
            what = {"root": "mydrive", "shared": "shared", "drives": "drives"}.get(first.id)
            if what and first.kind in ("folder", "shared", "drives"):
                target = f"l:{first.account}:{what}"
        if target and self.nav.exists(target):
            if self.nav.selection() != (target,):
                self.nav.selection_set(target)
        elif self.nav.selection():
            self.nav.selection_remove(*self.nav.selection())

    def _on_nav_click(self, event=None) -> None:
        if event is not None and event.type == tk.EventType.ButtonRelease:
            if self.nav.identify_element(event.x, event.y) == "Treeitem.indicator":
                return  # 펼치기 표시만 누름
            row = self.nav.identify_row(event.y)
        else:
            row = self.nav.focus()
        if not row:
            return
        if row == "add":
            self.add_account()
        elif row.startswith("a:"):
            self.navigate([loc_mydrive(row[2:])])
        elif row.startswith("l:"):
            _, aid, what = row.split(":", 2)
            maker = {"mydrive": loc_mydrive, "shared": loc_shared, "drives": loc_drives}[what]
            self.navigate([maker(aid)])  # 같은 곳을 다시 누르면 맨 위로 가서 새로 고침

    def _on_nav_menu(self, event) -> None:
        row = self.nav.identify_row(event.y)
        if not row or not row.startswith(("a:", "l:")):
            return
        aid = row[2:] if row.startswith("a:") else row.split(":", 2)[1]
        acct = self._acct(aid)
        if acct is None:
            return
        m = self.nav_menu
        m.delete(0, "end")
        m.add_command(label=f"{acct.label}", state="disabled")
        m.add_separator()
        m.add_command(label="다시 로그인", command=lambda: self._relogin_account(aid))
        m.add_command(label="지운 파일 되살리기 (타임머신)", command=lambda: self.open_time_machine(account_id=aid))
        m.add_command(label="이 계정의 작업 폴더 열기", command=lambda: self._reveal(acct.work_dir))
        m.add_separator()
        m.add_command(label="계정 제거...", command=lambda: self.remove_account(aid))
        try:
            m.tk_popup(event.x_root, event.y_root)
        finally:
            m.grab_release()

    def _relogin_account(self, account_id: str) -> None:
        self.navigate([loc_mydrive(account_id)])
        self.on_login_button()

    def _on_search(self, _event=None) -> None:
        text = self.search_var.get().strip()
        if text == getattr(self, "_search_hint", None):
            text = ""
        if not text or self.cur_acct is None:
            return
        self.navigate([Location("search", text, f"'{text}' 검색 결과", account=self.cur_acct)])
        self.tree.focus_set()

    def _on_typeahead(self, event) -> None:
        ch = event.char
        if not ch or not ch.isprintable() or event.state & 0x4:  # Ctrl 조합은 무시
            return
        prefix, last = self._typeahead
        now = time.monotonic()
        prefix = (prefix + ch) if now - last < 1.0 else ch
        self._typeahead = (prefix, now)
        p = prefix.casefold()
        for iid in self.tree.get_children():
            meta = self.item_meta.get(iid)
            if meta and (meta.get("name") or "").casefold().startswith(p):
                self.tree.selection_set(iid)
                self.tree.focus(iid)
                self.tree.see(iid)
                break

    # ======================= 열기 =======================
    def _on_double_click(self, event) -> str | None:
        if self.tree.identify_region(event.x, event.y) in ("heading", "separator"):
            return None
        row = self.tree.identify_row(event.y)
        if row:
            self.tree.focus(row)
            self.open_selected()
        return "break"

    def _selected_meta(self) -> dict | None:
        iid = self.tree.focus() or (self.tree.selection() or [None])[0]
        return self.item_meta.get(iid) if iid else None

    def open_selected(self) -> None:
        meta = self._selected_meta()
        if meta and self.lock is None:
            self.open_item(meta)

    def open_selected_in_browser(self) -> None:
        meta = self._selected_meta()
        if meta:
            self.open_in_browser(meta)

    def _on_tree_menu(self, event) -> None:
        row = self.tree.identify_row(event.y)
        if row and row in self.item_meta:
            self.tree.selection_set(row)
            self.tree.focus(row)
        try:
            self.menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.menu.grab_release()

    def open_item(self, meta: dict, account_id: str | None = None) -> None:
        account_id = account_id or self.stack[-1].account
        if account_id is None:
            return
        if gtypes.is_folder(meta):
            self.navigate(self.stack + [Location("folder", meta["id"], meta.get("name") or "",
                                                 meta.get("driveId"), account=account_id)])
        elif gtypes.is_shortcut(meta):
            target = (meta.get("shortcutDetails") or {}).get("targetId")
            if not target:
                self.set_status("바로가기 대상을 찾을 수 없습니다.")
                return
            drive = self.drives.get(account_id)
            if drive is None:
                self._pending_open = (account_id, meta)
                self.set_status(f"연결되면 엽니다: {meta.get('name')}")
                return
            self.set_status(f"바로가기 여는 중: {meta.get('name')}")
            self.io.submit(lambda: drive.get_meta(target),
                           lambda t: self._open_shortcut_target(meta, t, account_id),
                           lambda e: self._show_open_error(meta, e, account_id))
        elif gtypes.is_google_native(meta) and not gtypes.is_exportable(meta):
            self.open_in_browser(meta)
        else:
            self.open_file(meta, account_id)

    def _open_shortcut_target(self, shortcut: dict, target: dict, account_id: str) -> None:
        if gtypes.is_folder(target):
            self.navigate(self.stack + [Location("folder", target["id"], shortcut.get("name") or "",
                                                 target.get("driveId"), account=account_id)])
        elif gtypes.is_google_native(target) and not gtypes.is_exportable(target):
            self.open_in_browser(target)
        else:
            self.open_file(target, account_id)

    def open_in_browser(self, meta: dict) -> None:
        url = meta.get("webViewLink") or f"https://drive.google.com/open?id={meta['id']}"
        self.set_status(f"브라우저에서 엽니다: {meta.get('name')}")
        webbrowser.open(url)

    def open_file(self, meta: dict, account_id: str | None = None) -> None:
        account_id = account_id or self.stack[-1].account
        file_id = meta["id"]
        drive = self.drives.get(account_id)
        if drive is None:
            self._pending_open = (account_id, meta)
            self.set_status(f"연결되면 엽니다: {meta.get('name')}")
            self.connect(account_id)
            return
        key = (account_id, file_id)
        if key in self.downloading:
            self.set_status(f"이미 여는 중입니다: {meta.get('name')}")
            return
        self.downloading.add(key)
        self.set_status(f"여는 중: {meta.get('name')}")
        self.io.submit(lambda: self._prepare_local(account_id, drive, meta),
                       lambda r: self._on_prepared(key, r),
                       lambda e: self._on_prepare_failed(key, meta, e))

    def _prepare_local(self, account_id: str, drive, meta: dict) -> dict:
        """(작업 스레드) 열 로컬 파일을 준비한다. 최신이면 바로, 아니면 내려받는다."""
        sm = self.syncs[account_id]
        acct = self._acct(account_id)
        if acct is None:
            raise RuntimeError("계정이 제거되었습니다.")
        file_id = meta["id"]
        rec = sm.get_by_file_id(file_id)
        path: Path | None = None
        if rec and os.path.isfile(rec["path"]):
            path = Path(rec["path"])
            state = sm.local_state(rec["key"])
            if state != "clean":
                # 아직 안 올린 변경이 있음 → 덮어쓰지 않고 그대로 연다 (곧 업로드됨)
                return {"path": path, "kind": rec["kind"], "fresh": False, "name": rec["name"]}
            if rec["kind"] == sync.KIND_NATIVE:
                changed = (meta.get("modifiedTime") or "") > (rec.get("remote_modified") or "")
            else:
                # 드라이브 내용이 이 PC 사본과 다르면 다시 받는다 (타임머신으로 예전 버전을 되돌린 경우 포함)
                listed = meta.get("md5Checksum")
                changed = bool(listed) and listed != rec["synced_md5"]
            if not changed:
                return {"path": path, "kind": rec["kind"], "fresh": False, "name": rec["name"]}

        # 목록에 필요한 정보(권한·검사값)가 다 있으면 정보 요청 없이 바로 내려받는다 (0.3초 절약)
        fresh = meta if "capabilities" in meta and meta.get("mimeType") else drive.get_meta(file_id)
        if gtypes.is_google_native(fresh):
            kind = sync.KIND_NATIVE
        elif (fresh.get("capabilities") or {}).get("canEdit") is False:
            kind = sync.KIND_READONLY
        else:
            kind = sync.KIND_EDIT
        if path is None:
            folder = acct.work_dir / sync.short_dir_name(file_id)
            name = sync.safe_filename(gtypes.local_name(fresh),
                                      config.MAX_LOCAL_PATH - len(str(folder)) - 1)
            path = folder / name

        last = [-1.0]
        name = fresh.get("name") or path.name

        def progress(p: float) -> None:
            if p - last[0] >= 0.05:
                last[0] = p
                self.post(self._on_progress, (name, p))

        try:
            md5 = drive.download(fresh, path, progress)
        except FileLockedError:
            if path.exists():
                return {"path": path, "kind": rec["kind"] if rec else kind, "fresh": False, "name": name,
                        "note": "Word/Excel/한글에서 열려 있는 파일이라 드라이브의 최신 내용으로 바꾸지 못했습니다.\n"
                                "열려 있는 창을 그대로 쓰거나, 닫은 뒤 다시 열어 주세요."}
            raise
        ext = path.suffix.lower()
        if ext in DANGEROUS_EXT or ext in MACRO_EXT:
            mark_internet_zone(path)
        if fresh is meta and meta.get("md5Checksum") and meta["md5Checksum"] != md5:
            fresh = drive.get_meta(file_id)  # 목록을 받은 뒤 드라이브에서 바뀐 파일 → 최신 정보로 기록
        sm.add(path, fresh, md5, kind)
        return {"path": path, "kind": kind, "fresh": True, "name": name}

    def _on_progress(self, arg) -> None:
        name, p = arg
        self.set_status(f"내려받는 중 {int(p * 100)}%: {name}")

    def _on_prepared(self, key: tuple, r: dict) -> None:
        self.downloading.discard(key)
        if self.lock is not None:
            # 여는 사이 잠김 → 잠금을 푼 뒤에 연다 (잠긴 화면 뒤로 문서가 뜨지 않게)
            self._after_unlock.append(lambda: self._on_prepared(key, r))
            return
        path = Path(r["path"])
        if path.suffix.lower() in DANGEROUS_EXT:
            self.set_status(f"실행 파일이라 열지 않았습니다: {path.name}")
            if messagebox.askyesno(
                    "보안 경고",
                    f"{path.name}\n\n실행 파일·스크립트는 바이러스일 수 있어 바로 실행하지 않습니다.\n"
                    "믿을 수 있는 파일인지 확인한 뒤 직접 실행하세요.\n\n파일 위치를 열까요?",
                    icon="warning", parent=self):
                self._reveal(path)
            return
        try:
            open_with_default(path)
        except OSError as e:
            log.error("파일 열기 실패: %s", path, exc_info=e)
            if messagebox.askyesno(
                    "열 수 없음", f"{path.name}\n\n{friendly_error(e)}\n\n작업 폴더에서 파일 위치를 열까요?",
                    parent=self):
                self._reveal(path)
            return
        self.set_status(f"열었습니다: {r['name']}")
        file_id = key[1]
        if r.get("note"):
            self._dialog("warning", "안내", r["note"])
        elif r.get("fresh") and file_id not in self.noticed:
            if r["kind"] == sync.KIND_NATIVE:
                self.noticed.add(file_id)
                self._dialog("info", "안내", NATIVE_NOTICE)
            elif r["kind"] == sync.KIND_READONLY:
                self.noticed.add(file_id)
                self._dialog("info", "안내", READONLY_NOTICE)

    def _on_prepare_failed(self, key: tuple, meta: dict, e: BaseException) -> None:
        self.downloading.discard(key)
        self._show_open_error(meta, e, key[0])

    def _show_open_error(self, meta: dict, e: BaseException, account_id: str | None = None) -> None:
        msg = self._note_error(e, account_id)
        name = meta.get("name") or ""
        self.set_status(f"열지 못했습니다: {name} - {msg.splitlines()[0]}")
        if self.lock is not None:
            self._after_unlock.append(lambda: self._dialog("error", "열 수 없음", f"{name}\n\n{msg}"))
            return
        if http_status(e) == 403 and http_reason(e) in (
                "exportSizeLimitExceeded", "cannotDownloadFile", "fileNotDownloadable"):
            if messagebox.askyesno("열 수 없음", f"{name}\n\n{msg}\n\n브라우저에서 열까요?", parent=self):
                self.open_in_browser(meta)
            return
        messagebox.showerror("열 수 없음", f"{name}\n\n{msg}", parent=self)

    # ======================= 타임머신·도움말 =======================
    def open_time_machine(self, target: dict | None = None, account_id: str | None = None) -> None:
        if self.lock is not None:
            return
        account_id = account_id or self.cur_acct
        if account_id is None:
            self._dialog("info", "타임머신", "먼저 구글 계정을 추가해 주세요.")
            return
        dlg = self._tm_dialog
        if dlg is not None and dlg._alive() and dlg.account_id == account_id:
            dlg.lift()
            if target is not None:
                dlg.show_versions(target)
            return
        if dlg is not None and dlg._alive():
            dlg.destroy()
        self._tm_dialog = TimeMachineDialog(self, account_id, target)

    def versions_of_selected(self) -> None:
        meta = self._selected_meta()
        if meta:
            self.open_time_machine(meta, self.stack[-1].account)

    def versions_of_selected_track(self) -> None:
        sel = self._selected_track()
        if not sel:
            return
        account_id, rec = sel
        if not rec.get("file_id"):
            self.set_status("아직 드라이브에 올라가지 않은 새 파일입니다.")
            return
        meta = {"id": rec["file_id"], "name": rec["name"], "mimeType": rec.get("mime") or ""}
        if rec["kind"] == sync.KIND_NATIVE:
            meta["mimeType"] = "application/vnd.google-apps.document"
        self.open_time_machine(meta, account_id)

    def show_help(self, topic: str) -> None:
        HelpDialog(self, topic)

    def after_restore(self, account_id: str, file_id: str | None = None) -> None:
        """타임머신으로 되살리거나 되돌린 뒤: 목록을 새로 고치고, 이 PC 사본을 낡은 것으로 표시."""
        if file_id and account_id in self.syncs:
            self.syncs[account_id].remote_changed(file_id)
        for key in [k for k in self.cache if k[0] == account_id and k != self.stack[-1].key]:
            self.cache.pop(key, None)
        if self.stack[-1].account == account_id:
            self._fetch(self.stack[-1])

    def local_copy_in_use(self, account_id: str, file_id: str) -> bool:
        sm = self.syncs.get(account_id)
        rec = sm.get_by_file_id(file_id) if sm else None
        if not rec or not os.path.exists(rec["path"]):
            return False
        return sync.is_in_use(rec["path"]) or sync.title_mentions(sync.visible_window_titles(),
                                                                  os.path.basename(rec["path"]))

    @staticmethod
    def _clean_version_copies() -> None:
        """'이전 버전 열어보기'로 받은 보기용 사본 중 하루가 지난 것을 지운다."""
        root = config.BASE_DIR / "versions"
        if not root.exists():
            return
        cutoff = time.time() - 86400
        for p in root.rglob("*"):
            try:
                if p.is_file() and p.stat().st_mtime < cutoff and not sync.is_in_use(p):
                    p.unlink()
            except OSError:
                pass

    def _reveal(self, path: Path) -> None:
        try:
            reveal_in_explorer(path)
        except OSError as e:
            self._dialog("error", "열 수 없음", friendly_error(e))

    # ======================= 편집 중인 파일 =======================
    def _on_sync_event(self, arg) -> None:
        account_id, event, payload = arg
        if self._closing:
            return
        if event == "changed":
            if not self._track_refresh_pending:
                self._track_refresh_pending = True
                self.after(40, self._refresh_track_view)
        elif event == "new_file":
            self.set_status(f"새 파일을 드라이브에 올렸습니다 ({payload['where']}): {payload['name']}")
            if account_id == self.cur_acct:
                self._fetch(self.stack[-1])
        elif event == "conflict":
            self._refresh_track_view()
            if account_id == self.cur_acct:
                self._fetch(self.stack[-1])  # 사본이 생긴 폴더일 수 있음
            self._dialog(
                "warning", "충돌 사본으로 저장했습니다",
                f"{payload['reason']} 원본을 덮어쓰지 않고 같은 폴더에 사본으로 저장했습니다.\n\n"
                f"사본 이름: {payload['name']}\n\n이 파일에서 이후에 저장하는 내용은 사본에 반영됩니다.")
        elif event == "suspicious":
            self.set_status(f"업로드 보류 (손상·암호화 의심): {payload['name']}")
            self._dialog(
                "warning", "업로드 보류 - 드라이브 원본 보호",
                f"'{payload['name']}' 파일이 손상됐거나 암호화된 것으로 보여 드라이브에 올리지 않았습니다.\n"
                "(드라이브의 원본은 그대로 안전합니다)\n\n"
                "파일을 열어 내용을 확인한 뒤 다시 저장하면 올라갑니다.\n"
                "정상인데도 보류됐다면 아래 목록에서 오른쪽 클릭 → '그래도 업로드'를 누르세요.")
        elif event == "paused":
            self._show_resume_button(True)
            self.set_status("보안: 여러 파일이 한꺼번에 손상·암호화되어 모든 업로드를 멈췄습니다.")
            self._dialog(
                "error", "모든 업로드를 멈췄습니다 (랜섬웨어 의심)",
                "여러 파일이 짧은 시간에 한꺼번에 손상되거나 암호화되었습니다:\n\n"
                + short_list(payload.get("names") or []) +
                "\n\n랜섬웨어(파일을 암호화하는 악성코드)일 수 있어 드라이브 원본을 지키려고 모든 업로드를 멈췄습니다.\n"
                "1) 인터넷을 끊고 백신으로 검사하세요.\n"
                "2) 문제가 없으면 아래 '업로드 다시 시작'을 누르세요.")

    def resume_uploads(self) -> None:
        if not messagebox.askyesno("업로드 다시 시작",
                                   "백신 검사 등으로 문제가 없음을 확인했나요?\n업로드를 다시 시작합니다.", parent=self):
            return
        for sm in self.syncs.values():
            sm.resume()
        self._show_resume_button(False)
        self.set_status("업로드를 다시 시작했습니다.")

    def _update_track_columns(self) -> None:
        cols = ("account", "status") if len(self.store.accounts) > 1 else ("status",)
        try:
            self.track_tree.configure(displaycolumns=cols)
        except tk.TclError:
            pass

    def _refresh_track_view(self) -> None:
        self._track_refresh_pending = False
        if self._closing:
            return
        tv = self.track_tree
        rows = []
        for acct in self.store.accounts:
            sm = self.syncs.get(acct.id)
            if sm is not None:
                for r in sm.rows():
                    rows.append((acct, r))
        wanted = []
        for acct, r in rows:
            iid = "t" + hashlib.sha1((acct.id + "|" + r["key"]).encode("utf-8")).hexdigest()[:16]
            wanted.append(iid)
            self.track_keys[iid] = (acct.id, r["key"])
            old = self.track_status.get(iid)
            values = (acct.short, r["status"])
            if tv.exists(iid):
                if old != r["status"] or tv.item(iid, "text") != r["name"]:
                    tv.item(iid, text=r["name"], values=values)
            else:
                tv.insert("", "end", iid=iid, text=r["name"], values=values,
                          image=self.icons.get(gtypes.icon_key({"name": r["name"]})))
            if old is not None and old != r["status"]:
                self.set_status(f"{r['name']}: {r['status']}")
            self.track_status[iid] = r["status"]
        for iid in tv.get_children():
            if iid not in wanted:
                tv.delete(iid)
                self.track_keys.pop(iid, None)
                self.track_status.pop(iid, None)
        for index, iid in enumerate(wanted):
            if tv.index(iid) != index:
                tv.move(iid, "", index)
        self._show_resume_button(any(sm.paused for sm in self.syncs.values()))

    def _selected_track(self) -> tuple[str, dict] | None:
        sel = self.track_tree.selection() or ((self.track_tree.focus(),) if self.track_tree.focus() else ())
        if not sel:
            return None
        ref = self.track_keys.get(sel[0])
        if not ref:
            return None
        sm = self.syncs.get(ref[0])
        rec = sm.get(ref[1]) if sm else None
        return (ref[0], rec) if rec else None

    def _on_track_menu(self, event) -> None:
        row = self.track_tree.identify_row(event.y)
        if not row:
            return
        self.track_tree.selection_set(row)
        self.track_tree.focus(row)
        try:
            self.track_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.track_menu.grab_release()

    def open_selected_track(self) -> None:
        sel = self._selected_track()
        if not sel or self.lock is not None:
            return
        path = Path(sel[1]["path"])
        if not path.exists():
            self.set_status("로컬 파일이 없습니다. 위 목록에서 다시 열어 주세요.")
            return
        if path.suffix.lower() in DANGEROUS_EXT:
            self.set_status(f"실행 파일이라 열지 않았습니다: {path.name}")
            return
        try:
            open_with_default(path)
        except OSError as e:
            self._dialog("error", "열 수 없음", f"{path.name}\n\n{friendly_error(e)}")

    def reveal_selected_track(self) -> None:
        sel = self._selected_track()
        if sel:
            self._reveal(Path(sel[1]["path"]))
        else:
            acct = self._acct(self.cur_acct)
            self._reveal(acct.work_dir if acct else config.WORK_DIR)

    def force_selected_track(self) -> None:
        sel = self._selected_track()
        if not sel:
            return
        account_id, rec = sel
        if rec["status"] != sync.STATUS_HELD:
            self.set_status("보류된 파일이 아닙니다.")
            return
        if not messagebox.askyesno(
                "그래도 업로드",
                f"{rec['name']}\n\n파일을 직접 열어 내용이 정상인지 확인했나요?\n"
                "업로드하면 드라이브의 원본이 이 내용으로 바뀝니다. (이전 버전은 드라이브 '버전 관리'에 남음)",
                icon="warning", parent=self):
            return
        self.syncs[account_id].force_upload(rec["key"])
        self.set_status(f"업로드합니다: {rec['name']}")

    def cleanup_work(self) -> None:
        self.set_status("작업 폴더 정리 중...")
        managers = list(self.syncs.values())

        def work():
            removed, kept = 0, []
            for sm in managers:
                r, k = sm.cleanup()
                removed += r
                kept += k
            return removed, kept
        self.io.submit(work, self._on_cleanup_done,
                       lambda e: self.set_status("정리 실패: " + friendly_error(e)))

    def _on_cleanup_done(self, result) -> None:
        removed, kept = result
        msg = f"정리 완료: {removed}개 삭제"
        if kept:
            msg += f", {len(kept)}개는 열려 있거나 업로드 전이라 남겨둠"
        self.set_status(msg)
        self._refresh_track_view()

    def _health_tick(self) -> None:
        try:
            for sm in list(self.syncs.values()):
                sm.ensure_running()
        except Exception:
            log.exception("상태 점검 오류")
        finally:
            if not self._closing:
                self.after(5000, self._health_tick)

    # ======================= 시작 목록 캐시 =======================
    def _load_root_caches(self) -> None:
        for acct in self.store.accounts:
            try:
                data = json.loads(acct.root_cache_path.read_text(encoding="utf-8"))
                items = data["items"]
                if isinstance(items, list) and all(isinstance(m, dict) and "id" in m for m in items):
                    self.cache[loc_mydrive(acct.id).key] = items
            except Exception:
                pass

    def _save_root_cache(self, account_id: str, items: list[dict]) -> None:
        acct = self._acct(account_id)
        if acct is None:
            return
        text = json.dumps({"email": acct.email, "items": items}, ensure_ascii=False)
        path = acct.root_cache_path
        self.io.submit(lambda: config.atomic_write_text(path, text),
                       on_err=lambda e: log.warning("목록 캐시 저장 실패: %r", e))

    # ======================= 설정·종료 =======================
    def _load_ui_settings(self) -> None:
        self.ui_settings: dict = {}
        try:
            self.ui_settings = json.loads((config.BASE_DIR / "ui.json").read_text(encoding="utf-8"))
        except Exception:
            self.ui_settings = {}
        geo = self.ui_settings.get("geometry")
        m = re.fullmatch(r"(\d+)x(\d+)([+-]-?\d+)([+-]-?\d+)", geo or "")
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        if m:
            w, h, x, y = int(m.group(1)), int(m.group(2)), int(m.group(3)), int(m.group(4))
            w, h = min(max(w, 600), sw), min(max(h, 400), sh)
            # 모니터 구성이 바뀌어 창이 화면 밖에 뜨지 않게
            vx, vy = self.winfo_vrootx(), self.winfo_vrooty()
            vw = max(self.winfo_vrootwidth(), sw)
            vh = max(self.winfo_vrootheight(), sh)
            if vx - w + 80 <= x <= vx + vw - 80 and vy <= y <= vy + vh - 80:
                self.geometry(f"{w}x{h}{x:+d}{y:+d}")
            else:
                self.geometry(f"{w}x{h}")
            if self.ui_settings.get("zoomed"):
                try:
                    self.state("zoomed")
                except tk.TclError:
                    pass
        else:
            w, h = min(int(1100 * self._guess_scale()), sw - 40), min(int(720 * self._guess_scale()), sh - 80)
            self.geometry(f"{w}x{h}")

    def _guess_scale(self) -> float:
        try:
            return max(1.0, float(self.winfo_fpixels("1i")) / 96.0)
        except tk.TclError:
            return 1.0

    def _save_ui_settings(self) -> None:
        try:
            zoomed = self.state() == "zoomed"
            data = {
                # 최대화 상태면 이전의 보통 크기를 그대로 둔다
                "geometry": self.ui_settings.get("geometry") if zoomed else self.geometry(),
                "zoomed": zoomed,
                "columns": {c: self.tree.column(c, "width") for c in ("#0", "modified", "kind", "size")},
                "last_update_check": self.ui_settings.get("last_update_check", 0),
            }
            config.atomic_write_text(config.BASE_DIR / "ui.json", json.dumps(data))
        except Exception:
            log.exception("화면 설정 저장 실패")

    def on_close(self) -> None:
        if self._closing:
            return
        unsynced = [n for sm in self.syncs.values() for n in sm.unsynced_names()]
        if unsynced:
            if not messagebox.askyesno(
                    "종료 확인",
                    "아직 드라이브에 올라가지 않은 변경이 있습니다:\n\n" + short_list(unsynced) +
                    "\n\n지금 닫아도 변경 내용은 작업 폴더에 보관되고,\n"
                    "다음에 프로그램을 켜면 자동으로 업로드됩니다.\n\n종료할까요?",
                    icon="warning", parent=self):
                return
        else:
            in_use = [n for sm in self.syncs.values() for n in sm.in_use_names()]
            if in_use and not messagebox.askyesno(
                    "종료 확인",
                    "Word·Excel·한글에서 아직 열려 있는 파일이 있습니다:\n\n" + short_list(in_use) +
                    "\n\n프로그램을 닫은 뒤에 저장한 내용은\n다음에 프로그램을 켤 때 업로드됩니다.\n\n종료할까요?",
                    parent=self):
                return
        self.shutdown()

    def _on_session_end(self) -> None:
        """Windows를 끄거나 로그오프할 때: 묻지 않고 기록을 저장한 뒤 닫는다."""
        log.info("Windows 종료/로그오프 감지 → 기록 저장 후 종료")
        self.shutdown()

    def shutdown(self) -> None:
        """기록을 저장하고 창을 닫는다 (확인 없이)."""
        if self._closing:
            return
        self._closing = True
        self._save_ui_settings()
        for sm in list(self.syncs.values()):
            try:
                sm.stop()
            except Exception:
                log.exception("종료 처리 오류")
        # 예약된 화면 작업을 모두 취소해야 창을 없앤 뒤 'invalid command' 오류가 안 남
        try:
            for after_id in self.tk.splitlist(self.tk.call("after", "info")):
                self.after_cancel(after_id)
        except tk.TclError:
            pass
        log.info("종료")
        self.destroy()


def main() -> None:
    enable_dpi_awareness()
    try:
        setup_logging()
    except Exception:
        pass
    log.info("시작 (Python %s, exe=%s)", sys.version.split()[0], getattr(sys, "frozen", False))
    lock = SingleInstance(config.LOCK_PATH)
    acquired = lock.acquire()
    if acquired:
        try:
            moved = config.migrate_legacy_base()
            if moved:
                log.info("예전 위치(%s)의 기록을 %s 로 옮겼습니다", moved, config.BASE_DIR)
        except Exception:
            log.exception("예전 기록 옮기기 실패")
    if not acquired:
        if not activate_existing_window():
            try:
                r = tk.Tk()
                r.withdraw()
                messagebox.showinfo(config.APP_TITLE, "프로그램이 이미 실행 중입니다.", parent=r)
                r.destroy()
            except Exception:
                pass
        return
    try:
        app = App()
    except Exception as e:
        log.exception("시작 실패")
        try:
            r = tk.Tk()
            r.withdraw()
            messagebox.showerror(config.APP_TITLE,
                                 f"프로그램을 시작하지 못했습니다.\n\n{e}\n\n기록: {config.LOG_PATH}", parent=r)
            r.destroy()
        except Exception:
            pass
        return
    app.after(4000, app.check_update)
    app.mainloop()
    logging.shutdown()


if __name__ == "__main__":
    main()
