"""편집 중인 파일 추적: 저장 감지 → 드라이브 업로드.

신뢰성 설계
- 감지 두 겹: watchdog(즉시) + 주기적 상태 확인(watchdog가 놓친 저장도 잡음)
- 업로드할 내용은 한 번 읽어 둔 '스냅샷'. 읽는 도중 파일이 바뀌면 버리고 다시 읽음
- 읽을 때 Windows 공유 모드(읽기/쓰기/삭제 허용)로 열어 Word·한글의 저장을 방해하지 않음
- 실패하면 지수 백오프로 자동 재시도. 기록은 state.json에 남겨서 프로그램을 껐다 켜도 이어서 업로드
- 올리기 직전 드라이브 쪽 MD5를 확인. 다른 곳에서 바뀌었거나 지워졌으면 덮어쓰지 않고 '충돌 사본'으로 올림
- '다른 이름으로 저장'으로 같은 폴더에 생긴 새 파일은 원래 파일이 있던 드라이브 폴더에 새 파일로 올림
  (2초 이상 그대로 있는 파일만. 임시·잠금·백업 파일은 제외)

정전·오류 대비
- 기록(state.json)을 두 벌(state.json, state.bak.json) + 작업 폴더마다 한 벌(.gde-meta.json) 남긴다.
  어느 하나가 깨지거나 지워져도 나머지로 복원해 '올리지 못한 변경'을 잃지 않는다.
- 업로드 스레드·파일 감시가 멈추면 스스로 다시 시작한다 (ensure_running).

악성코드(랜섬웨어) 대비
- 문서가 갑자기 손상되거나 암호화되면(형식이 깨짐, 내용이 0바이트) 올리지 않고 '보류'한다. 드라이브 원본 보호.
- 10분 안에 3개 이상 보류되면 랜섬웨어로 보고 모든 업로드를 멈춘다 (paused). 사용자가 확인 후 다시 시작.
"""
from __future__ import annotations

import ctypes
import hashlib
import io
import json
import logging
import mimetypes
import os
import re
import shutil
import threading
import time
import uuid
import zipfile
from pathlib import Path
from typing import Callable

import config
from errors import AuthError, IntegrityError, friendly_error, is_not_found, is_transient

log = logging.getLogger(__name__)

PART_SUFFIX = ".gde-part"   # drive_client.PART_SUFFIX 와 같아야 함
DEL_SUFFIX = ".gde-del"

KIND_EDIT = "edit"          # 저장하면 원본에 업로드
KIND_NATIVE = "native"      # Google 문서를 내보낸 복사본 (업로드 안 함)
KIND_READONLY = "readonly"  # 보기 전용 권한 (업로드 안 함)

STATUS_IDLE = "열림 - 저장하면 자동 업로드"
STATUS_NATIVE = "복사본 (드라이브에 반영 안 됨)"
STATUS_READONLY = "보기 전용 (드라이브에 반영 안 됨)"
STATUS_HELD = "업로드 보류: 파일이 손상됐거나 암호화된 것 같습니다 (드라이브 원본 보호)"

SIDECAR = ".gde-meta.json"      # 작업 폴더마다 남기는 기록 사본 (정전·기록 손상 대비)
EMPTY_MD5 = hashlib.md5(b"").hexdigest()
ZIP_TEST_LIMIT = 64 * 1024 * 1024
_OLE = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"


# ---------- 파일 형식 검사 (손상·암호화 감지) ----------
def content_format(head: bytes) -> str | None:
    """파일 앞부분으로 본 형식. docx/xlsx/pptx/hwpx = zip, hwp/doc/xls = ole."""
    if head.startswith(b"PK\x03\x04"):
        return "zip"
    if head.startswith(_OLE):
        return "ole"
    if head.startswith(b"HWP Document File"):
        return "hwp3"
    if head.startswith(b"%PDF-"):
        return "pdf"
    if head.startswith(b"\x89PNG\r\n\x1a\n"):
        return "png"
    if head.startswith(b"\xff\xd8\xff"):
        return "jpg"
    return None


def _zip_ok(data: bytes) -> bool:
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as z:
            return z.testzip() is None
    except Exception:
        return False


def file_format(path) -> str | None:
    """내려받은 원본의 형식. 원본부터 온전하지 않으면 None (그 파일은 검사하지 않음)."""
    try:
        with open_shared_read(path) as f:
            head = f.read(32)
        fmt = content_format(head)
        if fmt == "zip" and os.path.getsize(path) <= ZIP_TEST_LIMIT:
            with open_shared_read(path) as f:
                if not _zip_ok(f.read()):
                    return None
        return fmt
    except OSError:
        return None


# ---------- 파일 도우미 ----------
def norm(path) -> str:
    # 감지된 경로와 저장된 경로를 같은 형태로 맞춘다 (바로가기/심볼릭 링크 해제, Windows 대소문자 무시)
    return os.path.normcase(os.path.realpath(str(path)))


def is_office_temp(name: str) -> bool:
    # Word/Excel/한글이 저장할 때 만드는 잠금·임시 파일과 이 프로그램의 임시 파일은 무시
    low = name.lower()
    return (name.startswith("~$") or name.startswith(".~") or low.endswith(".tmp")
            or low.endswith(PART_SUFFIX) or low.endswith(DEL_SUFFIX))


_WIN_RESERVED = {"CON", "PRN", "AUX", "NUL", "CONIN$", "CONOUT$",
                 *(f"COM{i}" for i in range(10)), *(f"LPT{i}" for i in range(10)),
                 "COM¹", "COM²", "COM³", "LPT¹", "LPT²", "LPT³"}
_BAD_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f\x7f]')


# 새 파일로 올리지 않을 것: 프로그램들이 저장할 때 잠깐 만드는 임시·잠금·백업 파일
_NOT_NEW_EXT = {".tmp", ".temp", ".bak", ".wbk", ".asd", ".swp", ".lock", ".lnk", ".ini", ".db",
                ".part", ".partial", ".crdownload", PART_SUFFIX, DEL_SUFFIX}
_HEX8 = re.compile(r"^[0-9A-Fa-f]{8}$")   # Excel 임시 파일 (확장자 없는 16진수 8자리)
_EXTRA_MIME = {".hwp": "application/x-hwp", ".hwpx": "application/hwp+zip"}


def is_new_file_candidate(name: str) -> bool:
    """작업 폴더에 새로 생긴 이 파일을 드라이브에 새 파일로 올릴 만한가."""
    if not name or name[0] in "~.$" or is_office_temp(name):
        return False
    ext = os.path.splitext(name)[1].lower()
    if ext in _NOT_NEW_EXT:
        return False
    if not ext and _HEX8.match(name):
        return False
    return True


def guess_mime(name: str) -> str:
    ext = os.path.splitext(name)[1].lower()
    return _EXTRA_MIME.get(ext) or mimetypes.guess_type(name)[0] or "application/octet-stream"


def _utf16_len(s: str) -> int:
    return len(s.encode("utf-16-le")) // 2


def _cut_utf16(s: str, limit: int) -> str:
    while s and _utf16_len(s) > limit:
        s = s[:-1]
    return s


def safe_filename(name: str, max_len: int = 150) -> str:
    """드라이브 이름을 Windows에서 쓸 수 있는 파일 이름으로 바꾼다.

    드라이브는 / : * ? 같은 글자, 끝의 점·공백, CON 같은 예약어를 허용하지만 Windows는 안 된다.
    너무 긴 이름은 확장자를 살리고 앞부분을 줄인다 (Word/Excel 경로 길이 한계 대비).
    """
    name = _BAD_CHARS.sub("_", name or "").strip()
    name = name.rstrip(". ")
    if not name:
        name = "이름 없음"
    stem, dot, ext = name.rpartition(".")
    if not dot or not stem or len(ext) > 12 or " " in ext:
        stem, ext = name, ""
    else:
        ext = "." + ext
    if stem.split(".")[0].strip().upper() in _WIN_RESERVED:
        stem = "_" + stem
    max_len = max(max_len, 16)
    if _utf16_len(stem + ext) > max_len:
        stem = _cut_utf16(stem, max_len - _utf16_len(ext)).rstrip(". ") or "_"
    return stem + ext


def short_dir_name(file_id: str) -> str:
    # 파일마다 하위 폴더를 따로 둬서 이름이 같아도 섞이지 않게 한다 (경로를 짧게 유지)
    return hashlib.sha1(file_id.encode("utf-8")).hexdigest()[:10]


if os.name == "nt":
    import msvcrt
    from ctypes import wintypes

    _k32 = ctypes.WinDLL("kernel32", use_last_error=True)
    _CreateFileW = _k32.CreateFileW
    _CreateFileW.argtypes = [wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD, wintypes.LPVOID,
                             wintypes.DWORD, wintypes.DWORD, wintypes.HANDLE]
    _CreateFileW.restype = wintypes.HANDLE
    _CloseHandle = _k32.CloseHandle
    _CloseHandle.argtypes = [wintypes.HANDLE]
    _CloseHandle.restype = wintypes.BOOL
    _INVALID = wintypes.HANDLE(-1).value
    _GENERIC_READ = 0x80000000
    _SHARE_ALL = 0x1 | 0x2 | 0x4          # READ | WRITE | DELETE
    _OPEN_EXISTING = 3
    _SEQUENTIAL = 0x08000000
    _SHARING_VIOLATION = 32
    _LOCK_VIOLATION = 33

    def open_shared_read(path):
        """다른 프로그램이 이름을 바꾸거나 지우는 것도 막지 않는 읽기 전용 열기."""
        h = _CreateFileW(str(path), _GENERIC_READ, _SHARE_ALL, None, _OPEN_EXISTING, _SEQUENTIAL, None)
        if h == _INVALID or h is None:
            err = ctypes.get_last_error()
            raise ctypes.WinError(err)
        try:
            fd = msvcrt.open_osfhandle(h, os.O_RDONLY | getattr(os, "O_BINARY", 0))
        except Exception:
            _CloseHandle(h)
            raise
        return os.fdopen(fd, "rb")

    def is_in_use(path) -> bool:
        """다른 프로그램(Word/Excel/한글)이 파일을 열고 있으면 True."""
        h = _CreateFileW(str(path), _GENERIC_READ, 0, None, _OPEN_EXISTING, 0, None)
        if h == _INVALID or h is None:
            return ctypes.get_last_error() in (_SHARING_VIOLATION, _LOCK_VIOLATION)
        _CloseHandle(h)
        return False

    _u32 = ctypes.WinDLL("user32", use_last_error=True)
    _ENUMPROC = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
    _u32.EnumWindows.argtypes = [_ENUMPROC, wintypes.LPARAM]
    _u32.IsWindowVisible.argtypes = [wintypes.HWND]
    _u32.GetWindowTextLengthW.argtypes = [wintypes.HWND]
    _u32.GetWindowTextW.argtypes = [wintypes.HWND, wintypes.LPWSTR, ctypes.c_int]

    def visible_window_titles() -> list[str]:
        titles: list[str] = []

        def cb(hwnd, _lparam):
            try:
                if _u32.IsWindowVisible(hwnd):
                    n = _u32.GetWindowTextLengthW(hwnd)
                    if n > 0:
                        buf = ctypes.create_unicode_buffer(n + 1)
                        _u32.GetWindowTextW(hwnd, buf, n + 1)
                        titles.append(buf.value)
            except Exception:
                pass
            return True
        try:
            _u32.EnumWindows(_ENUMPROC(cb), 0)
        except Exception:
            pass
        return titles
else:
    def open_shared_read(path):
        return open(path, "rb")

    def is_in_use(path) -> bool:
        return False

    def visible_window_titles() -> list[str]:
        return []


def title_mentions(titles: list[str], filename: str) -> bool:
    """창 제목에 파일 이름이 보이면 열려 있는 것으로 본다 (파일을 잠그지 않는 프로그램 대비).

    잘못 짚어도 '안 지우고 남겨 두는' 쪽이라 안전하다.
    """
    name = filename.casefold()
    stem = os.path.splitext(filename)[0].casefold()
    for t in titles:
        low = t.casefold()
        if name in low or (len(stem) >= 4 and stem in low):
            return True
    return False


def file_md5(path) -> str | None:
    h = hashlib.md5()
    try:
        with open_shared_read(path) as f:
            for chunk in iter(lambda: f.read(1 << 20), b""):
                h.update(chunk)
    except OSError:
        return None
    return h.hexdigest()


def _stat_sig(path) -> tuple[int, int] | None:
    try:
        st = os.stat(path)
    except OSError:
        return None
    return (st.st_mtime_ns, st.st_size)


def conflict_name(name: str) -> str:
    stem, dot, ext = name.rpartition(".")
    if not dot or not stem:
        stem, ext = name, ""
    else:
        ext = "." + ext
    return f"{stem} (충돌 사본 {time.strftime('%Y-%m-%d %H%M%S')}){ext}"


class _Changed(Exception):
    """읽는 도중 파일이 바뀜."""


class _Snapshot:
    def __init__(self, md5: str, size: int, data: bytes | None = None, staged: Path | None = None):
        self.md5 = md5
        self.size = size
        self.data = data
        self.staged = staged

    @property
    def source(self):
        return self.data if self.data is not None else self.staged

    def discard(self) -> None:
        if self.staged:
            try:
                self.staged.unlink()
            except OSError:
                pass


# ---------- 추적 관리자 ----------
class SyncManager:
    """작업 폴더의 편집 중인 파일을 추적하고 저장되면 드라이브에 올린다. 스레드 안전."""

    PERSIST = ("rel", "file_id", "name", "mime", "kind", "parent_id", "drive_id",
               "synced_md5", "known_md5", "remote_modified", "create_tried", "fmt")

    def __init__(self, work_dir: Path, state_path: Path, staging_dir: Path,
                 get_drive: Callable[[], object | None],
                 notify: Callable[[str, object], None] = lambda *_: None,
                 debounce: float = config.UPLOAD_DEBOUNCE_SEC,
                 poll: float = config.POLL_SEC,
                 memory_limit: int = config.MEMORY_UPLOAD_LIMIT,
                 use_watchdog: bool = True,
                 new_file_delay: float = config.NEW_FILE_DELAY_SEC) -> None:
        self.work_dir = Path(work_dir)
        self.state_path = Path(state_path)
        self.staging_dir = Path(staging_dir)
        self.get_drive = get_drive
        self.notify = notify
        self.debounce = debounce
        self.poll = poll
        self.memory_limit = memory_limit
        self.use_watchdog = use_watchdog
        self.new_file_delay = new_file_delay
        self._new_seen: dict[str, list] = {}   # 새 파일 후보: norm 경로 -> [경로, 처음 본 시각, 상태]
        self._save_pending = False
        self._lock = threading.RLock()
        self._save_lock = threading.Lock()
        self._recs: dict[str, dict] = {}    # norm 경로 -> 기록
        self._stop = threading.Event()
        self._wake = threading.Event()
        self._observer = None
        self._thread: threading.Thread | None = None
        self._sidecars: dict[str, str] = {}   # 폴더 -> 마지막으로 쓴 기록 사본 내용
        self._holds: list[tuple[float, str]] = []
        self.paused = False                    # 랜섬웨어 의심 → 모든 업로드 멈춤

    @property
    def backup_path(self) -> Path:
        return self.state_path.with_name(self.state_path.stem + ".bak.json")

    # ----- 시작/종료 -----
    def start(self) -> None:
        self.work_dir.mkdir(parents=True, exist_ok=True)
        self.staging_dir.mkdir(parents=True, exist_ok=True)
        self._clean_leftovers()
        self._load_state()
        if self.use_watchdog:
            self._start_watcher()
        self._thread = threading.Thread(target=self._loop, name="sync", daemon=True)
        self._thread.start()
        self.save_state()   # 시작하자마자 기록 두 벌·폴더별 사본을 만들어 둔다 (첫 변경 전 정전 대비)

    def ensure_running(self) -> None:
        """업로드 스레드나 파일 감시가 예기치 않게 멈췄으면 다시 시작한다 (앱이 주기적으로 부름)."""
        if self._stop.is_set():
            return
        if self._thread is None or not self._thread.is_alive():
            log.error("업로드 스레드가 멈춰 있어 다시 시작합니다")
            self._thread = threading.Thread(target=self._loop, name="sync", daemon=True)
            self._thread.start()
        if self.use_watchdog and self._observer is not None and not self._observer.is_alive():
            log.error("파일 감시가 멈춰 있어 다시 시작합니다")
            self._observer = None
            self._start_watcher()

    def stop(self) -> None:
        self._stop.set()
        self._wake.set()
        if self._observer is not None:
            try:
                self._observer.stop()
            except Exception:
                pass
        self.save_state()

    def _clean_leftovers(self) -> None:
        # 지난번에 끊긴 다운로드 조각, 업로드용 임시 사본 정리
        for p in self.work_dir.glob("*/*"):
            if p.name.lower().endswith((PART_SUFFIX, DEL_SUFFIX)):
                try:
                    p.unlink()
                except OSError:
                    pass
        for p in self.staging_dir.glob("*"):
            try:
                p.unlink()
            except OSError:
                pass

    # ----- 상태 저장 -----
    def _read_state_file(self, path: Path) -> list | None:
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
            files = data.get("files")
            return files if isinstance(files, list) else None
        except FileNotFoundError:
            return None
        except Exception as e:
            log.error("%s을(를) 읽지 못했습니다: %s", path.name, e)
            try:
                shutil.copyfile(path, path.with_suffix(".broken.json"))
            except OSError:
                pass
            return None

    def _restore(self, item: dict, path: Path) -> None:
        # (잠금을 잡은 상태에서 호출)
        rec = {k: item.get(k) for k in self.PERSIST}
        rec["rel"] = os.path.relpath(path, self.work_dir)
        rec["known_md5"] = list(rec.get("known_md5") or [])
        self._init_runtime(rec, path)
        rec["create_tried"] = bool(item.get("create_tried"))
        if rec["kind"] == KIND_EDIT:
            rec["status"] = "확인 중..."
            rec["sig"] = None
            rec["due"] = time.monotonic()  # 지난번에 못 올린 변경이 있으면 바로 업로드
        self._recs[rec["key"]] = rec

    def _load_state(self) -> None:
        files = self._read_state_file(self.state_path)
        if files is None and self.backup_path.exists():
            log.warning("state.json 대신 백업(state.bak.json)으로 복원합니다")
            files = self._read_state_file(self.backup_path)
        restored = 0
        with self._lock:
            for item in files or []:
                try:
                    path = self.work_dir / item["rel"]
                    if path.is_file():
                        self._restore(item, path)
                except Exception:
                    log.exception("state 항목 복원 실패: %r", item)
            # 작업 폴더마다 남긴 기록 사본으로 빠진 것을 채운다 (기록이 깨졌거나 지워진 경우)
            for side in self.work_dir.glob("*/" + SIDECAR):
                try:
                    data = json.loads(side.read_text(encoding="utf-8"))
                    for item in data.get("files", []):
                        path = side.parent / str(item.get("file") or "")
                        if item.get("file") and path.is_file() and norm(path) not in self._recs:
                            self._restore(item, path)
                            restored += 1
                except Exception as e:
                    log.warning("기록 사본을 읽지 못함 %s: %s", side, e)
        if restored:
            log.warning("작업 폴더의 기록 사본에서 %d개 파일 기록을 되살렸습니다", restored)
            self.save_state()
        self.notify("changed", None)

    def save_state(self) -> None:
        with self._save_lock:  # 늦게 찍은 스냅샷이 먼저 찍은 것을 덮어쓰도록 순서 보장
            with self._lock:
                recs = [(rec["path"], {k: rec.get(k) for k in self.PERSIST}) for rec in self._recs.values()]
            text = json.dumps({"version": 1, "files": [r for _, r in recs]}, ensure_ascii=False, indent=1)
            for target in (self.state_path, self.backup_path):  # 두 벌: 하나가 깨져도 다른 하나로 복원
                try:
                    config.atomic_write_text(target, text)
                except Exception:
                    log.exception("%s 저장 실패", target.name)
            self._write_sidecars(recs)

    def _write_sidecars(self, recs: list[tuple[str, dict]]) -> None:
        by_folder: dict[str, list[dict]] = {}
        for path, item in recs:
            folder = os.path.dirname(path)
            if os.path.normcase(folder) == os.path.normcase(str(self.work_dir)):
                continue
            by_folder.setdefault(folder, []).append(dict(item, file=os.path.basename(path)))
        for folder, items in by_folder.items():
            text = json.dumps({"version": 1, "files": items}, ensure_ascii=False)
            if self._sidecars.get(folder) == text:
                continue
            try:
                config.atomic_write_text(Path(folder) / SIDECAR, text)
                self._sidecars[folder] = text
            except Exception as e:
                log.warning("기록 사본 저장 실패 %s: %s", folder, e)
        for folder in [f for f in self._sidecars if f not in by_folder]:
            self._sidecars.pop(folder, None)
            try:
                os.unlink(os.path.join(folder, SIDECAR))
            except OSError:
                pass

    # ----- 기록 -----
    def _init_runtime(self, rec: dict, path: Path) -> None:
        rec["path"] = str(path)
        rec["key"] = norm(path)
        rec["sig"] = _stat_sig(path)
        rec["due"] = None
        rec["busy"] = False
        rec["fails"] = 0
        rec["missing"] = 0
        rec["create_tried"] = False
        rec["force"] = False
        rec.setdefault("fmt", None)
        rec["attempted"] = set()
        rec["seq"] = time.monotonic()
        rec["status"] = {KIND_EDIT: STATUS_IDLE, KIND_NATIVE: STATUS_NATIVE,
                         KIND_READONLY: STATUS_READONLY}.get(rec.get("kind"), "")

    def add(self, path: Path, meta: dict, md5: str, kind: str) -> str:
        """내려받은 파일을 추적 목록에 넣는다. 이미 있으면 새 내용 기준으로 바꾼다."""
        path = Path(path)
        rel = os.path.relpath(path, self.work_dir)
        parents = meta.get("parents") or []
        rec = {
            "rel": rel,
            "file_id": meta["id"],
            "name": meta.get("name") or path.name,
            "mime": meta.get("mimeType") or "application/octet-stream",
            "kind": kind,
            "parent_id": parents[0] if parents else None,
            "drive_id": meta.get("driveId"),
            "synced_md5": md5,
            "known_md5": [md5] + ([meta["md5Checksum"]] if meta.get("md5Checksum") else []),
            "remote_modified": meta.get("modifiedTime", ""),
            "fmt": file_format(path) if kind == KIND_EDIT else None,
        }
        self._init_runtime(rec, path)
        with self._lock:
            self._recs[rec["key"]] = rec
        self.save_state()
        self.notify("changed", None)
        return rec["key"]

    def get_by_file_id(self, file_id: str) -> dict | None:
        with self._lock:
            for rec in self._recs.values():
                if rec["file_id"] == file_id:
                    return self._public(rec)
        return None

    def get(self, key: str) -> dict | None:
        with self._lock:
            rec = self._recs.get(key)
            return self._public(rec) if rec else None

    @staticmethod
    def _public(rec: dict) -> dict:
        out = dict(rec)
        out["attempted"] = set(rec["attempted"])
        out["known_md5"] = list(rec["known_md5"])
        return out

    def rows(self) -> list[dict]:
        """화면 표시용 (최근 것이 위)."""
        with self._lock:
            recs = sorted(self._recs.values(), key=lambda r: r["seq"], reverse=True)
            return [{"key": r["key"], "name": r["name"], "status": r["status"], "path": r["path"],
                     "kind": r["kind"], "file_id": r["file_id"]} for r in recs]

    def local_state(self, key: str) -> str:
        """'clean'(드라이브와 같음) / 'dirty'(올릴 변경 있음) / 'busy'(업로드 중·대기) / 'missing'."""
        with self._lock:
            rec = self._recs.get(key)
            if not rec:
                return "missing"
            if rec["busy"] or rec["due"] is not None:
                return "busy"
            path, synced = rec["path"], rec["synced_md5"]
        if not os.path.isfile(path):
            return "missing"
        md5 = file_md5(path)
        if md5 is None:
            return "busy"
        return "clean" if md5 == synced else "dirty"

    def unsynced_names(self) -> list[str]:
        """아직 드라이브에 올라가지 않은 변경이 있는 파일 이름."""
        with self._lock:
            recs = [(r["name"], r["path"], r["synced_md5"], r["busy"] or r["due"] is not None)
                    for r in self._recs.values() if r["kind"] == KIND_EDIT]
        out = []
        for name, path, synced, busy in recs:
            if busy:
                out.append(name)
                continue
            md5 = file_md5(path)
            if md5 is not None and md5 != synced:
                out.append(name)
        return out

    def in_use_names(self) -> list[str]:
        """Word/Excel/한글 등에서 아직 열려 있는 것으로 보이는 편집 파일 이름."""
        with self._lock:
            recs = [(r["name"], r["path"]) for r in self._recs.values() if r["kind"] == KIND_EDIT]
        titles = visible_window_titles() if recs else []
        return [name for name, path in recs
                if os.path.exists(path) and (is_in_use(path) or self._lock_file_in_use(Path(path).parent)
                                             or title_mentions(titles, os.path.basename(path)))]

    def _set_status(self, rec: dict, status: str) -> None:
        with self._lock:
            if rec.get("status") == status:
                return
            rec["status"] = status
        self.notify("changed", None)

    # ----- 감지 -----
    def _start_watcher(self) -> None:
        try:
            from watchdog.events import FileSystemEventHandler
            from watchdog.observers import Observer
        except Exception:
            log.exception("watchdog을 불러오지 못해 주기 확인만 사용합니다")
            return

        mgr = self

        class Handler(FileSystemEventHandler):
            def on_modified(self, event):
                if not event.is_directory:
                    mgr.note_change(event.src_path)

            def on_created(self, event):
                if not event.is_directory:
                    mgr.note_change(event.src_path)

            def on_moved(self, event):
                # Word/Excel은 임시 파일에 쓴 뒤 원래 이름으로 바꾸는 방식으로 저장한다
                if not event.is_directory:
                    mgr.note_change(event.dest_path)

        try:
            obs = Observer()
            obs.schedule(Handler(), str(self.work_dir), recursive=True)
            obs.daemon = True
            obs.start()
            self._observer = obs
        except Exception:
            log.exception("파일 감시 시작 실패 - 주기 확인만 사용합니다")

    def note_change(self, raw_path) -> None:
        try:
            name = os.path.basename(str(raw_path))
            if is_office_temp(name):
                return
            key = norm(raw_path)
        except Exception:
            return
        with self._lock:
            rec = self._recs.get(key)
            if rec is None:
                self._consider_new(str(raw_path), key)
                return
            if rec["kind"] != KIND_EDIT:
                return
            rec["sig"] = _stat_sig(rec["path"])
            rec["missing"] = 0
            rec["due"] = time.monotonic() + self.debounce
        self._wake.set()

    # ----- '다른 이름으로 저장'한 새 파일 -----
    def _tracked_folders(self) -> dict[str, dict]:
        """작업 하위 폴더(norm) -> 그 폴더의 대표 기록 (드라이브 폴더 위치를 알려 줌)."""
        out: dict[str, dict] = {}
        for rec in self._recs.values():
            if rec.get("file_id") is None:
                continue
            folder = norm(os.path.dirname(rec["path"]))
            if folder not in out or (rec.get("parent_id") and not out[folder].get("parent_id")):
                out[folder] = rec
        return out

    def _consider_new(self, path: str, key: str, folders: dict | None = None) -> None:
        # (잠금을 잡은 상태에서 호출)
        if key in self._recs or key in self._new_seen:
            return
        if not is_new_file_candidate(os.path.basename(path)):
            return
        folders = folders if folders is not None else self._tracked_folders()
        if norm(os.path.dirname(path)) not in folders:
            return  # 이 프로그램이 연 파일의 폴더가 아님
        self._new_seen[key] = [path, time.monotonic(), _stat_sig(path)]

    def _promote_new_files(self, now: float) -> bool:
        # (잠금을 잡은 상태에서 호출) 일정 시간 그대로 있는 새 파일 → 업로드 대상 기록으로
        folders = None
        added = False
        for key, item in list(self._new_seen.items()):
            path, first, sig = item
            cur = _stat_sig(path)
            if cur is None:
                del self._new_seen[key]  # 잠깐 생겼다 사라진 임시 파일
                continue
            if cur != sig or cur[1] == 0:
                item[1], item[2] = now, cur  # 아직 쓰는 중 → 다시 기다림
                continue
            if now - first < self.new_file_delay:
                continue
            del self._new_seen[key]
            if key in self._recs:
                continue
            folders = folders if folders is not None else self._tracked_folders()
            base = folders.get(norm(os.path.dirname(path)))
            if base is None:
                continue
            name = os.path.basename(path)
            rec = {"rel": os.path.relpath(path, self.work_dir), "file_id": None, "name": name,
                   "mime": guess_mime(name), "kind": KIND_EDIT, "parent_id": base.get("parent_id"),
                   "drive_id": base.get("drive_id"), "synced_md5": None, "known_md5": [],
                   "remote_modified": ""}
            self._init_runtime(rec, Path(path))
            rec["status"] = "새 파일 - 드라이브에 올리는 중"
            rec["due"] = now
            self._recs[key] = rec
            log.info("새 파일 감지 → 업로드 예정: %s", name)
            added = True
        return added

    def _poll(self) -> None:
        now = time.monotonic()
        with self._lock:
            for rec in self._recs.values():
                if rec["kind"] != KIND_EDIT:
                    continue
                sig = _stat_sig(rec["path"])
                if sig is not None and sig != rec["sig"]:
                    rec["sig"] = sig
                    rec["missing"] = 0
                    rec["due"] = now + self.debounce
            # watchdog가 놓친 '다른 이름으로 저장'도 잡도록 편집 중인 파일의 폴더를 훑는다 (파일 몇 개뿐)
            folders = self._tracked_folders()
            for folder in folders:
                try:
                    names = os.listdir(folder)
                except OSError:
                    continue
                for name in names:
                    path = os.path.join(folder, name)
                    key = norm(path)
                    if key not in self._recs and key not in self._new_seen and os.path.isfile(path):
                        self._consider_new(path, key, folders)
            added = self._promote_new_files(now)
        if added:
            self.save_state()
            self.notify("changed", None)

    # ----- 업로드 루프 -----
    def _loop(self) -> None:
        next_poll = 0.0
        while not self._stop.is_set():
            try:
                now = time.monotonic()
                if now >= next_poll:
                    self._poll()
                    next_poll = now + self.poll
                rec = self._next_due(now)
                if rec is not None:
                    self._process(rec)
                    continue
            except Exception:
                log.exception("업로드 루프 오류")
            self._wake.wait(self._sleep_time())
            self._wake.clear()

    def _sleep_time(self) -> float:
        now = time.monotonic()
        with self._lock:
            dues = [r["due"] for r in self._recs.values() if r["due"] is not None and not r["busy"]]
        t = min([self.poll] + [max(0.0, d - now) for d in dues])
        return max(0.05, t)

    def _next_due(self, now: float) -> dict | None:
        with self._lock:
            if self.paused:
                return None  # 랜섬웨어 의심으로 멈춤 → 사용자가 다시 시작할 때까지 올리지 않음
            ready = [r for r in self._recs.values()
                     if r["due"] is not None and r["due"] <= now and not r["busy"]]
            if not ready:
                return None
            rec = min(ready, key=lambda r: r["due"])
            rec["due"] = None
            rec["busy"] = True
            return rec

    def _retry_later(self, rec: dict, delay: float) -> None:
        with self._lock:
            if rec["due"] is None:  # 그 사이 새 저장이 들어왔으면 그쪽 일정을 따름
                rec["due"] = time.monotonic() + delay

    def _snapshot(self, path: str) -> _Snapshot:
        before = _stat_sig(path)
        if before is None:
            raise FileNotFoundError(path)
        h = hashlib.md5()
        if before[1] <= self.memory_limit:
            with open_shared_read(path) as f:
                data = f.read()
            h.update(data)
            snap = _Snapshot(h.hexdigest(), len(data), data=data)
        else:
            staged = self.staging_dir / uuid.uuid4().hex
            size = 0
            try:
                with open_shared_read(path) as src, open(staged, "wb") as dst:
                    for chunk in iter(lambda: src.read(1 << 20), b""):
                        h.update(chunk)
                        dst.write(chunk)
                        size += len(chunk)
            except BaseException:
                try:
                    staged.unlink()
                except OSError:
                    pass
                raise
            snap = _Snapshot(h.hexdigest(), size, staged=staged)
        after = _stat_sig(path)
        if after != before or snap.size != before[1]:
            snap.discard()
            raise _Changed()
        return snap

    def _process(self, rec: dict) -> None:
        path = rec["path"]
        snap = None
        try:
            try:
                snap = self._snapshot(path)
            except _Changed:
                self._retry_later(rec, self.debounce)
                return
            except FileNotFoundError:
                with self._lock:
                    rec["missing"] += 1
                    gone = rec["missing"] > 5
                    if gone and rec.get("file_id") is None:
                        self._recs.pop(rec["key"], None)  # 올리기 전에 사라진 새 파일 → 잊음
                if gone:
                    # 로컬 파일이 지워짐 → 계속 다시 시도하지 않음 (다시 생기면 주기 확인이 잡음)
                    self._set_status(rec, "로컬 파일이 없습니다")
                    self.save_state()
                    return
                # 저장 도중(이름 바꾸기 사이)일 수 있음 → 잠시 후 다시 확인
                self._retry_later(rec, 2.0)
                return
            except OSError as e:
                log.info("파일 잠김, 재시도 예정: %s (%s)", path, e)
                self._set_status(rec, "파일이 잠겨 있어 기다리는 중...")
                self._retry_later(rec, 1.5)
                return

            with self._lock:
                rec["missing"] = 0
            if snap.md5 == rec["synced_md5"]:
                # 내용이 안 바뀜 (열기만 했거나 같은 내용 저장) → 업로드 안 함
                with self._lock:
                    rec["fails"] = 0
                    status = rec["status"]
                if status == "확인 중...":
                    self._set_status(rec, "드라이브와 같음")
                elif status.startswith(("업로드", "파일이 잠겨", "로그인 필요", "드라이브 연결")):
                    self._set_status(rec, STATUS_IDLE)
                return

            if self._looks_damaged(rec, snap):
                self._hold(rec)
                return

            drive = self.get_drive()
            if drive is None:
                self._set_status(rec, "드라이브 연결 대기 중 (연결되면 자동 업로드)")
                self._retry_later(rec, 3.0)
                return

            self._set_status(rec, "업로드 중...")
            try:
                self._upload(rec, drive, snap)
            except AuthError as e:
                self._set_status(rec, f"로그인 필요 - 변경 내용은 보관 중 ({friendly_error(e)})")
                self._retry_later(rec, 10.0)
            except Exception as e:
                msg = friendly_error(e)
                if is_transient(e):
                    with self._lock:
                        rec["fails"] += 1
                        delay = min(config.RETRY_MAX_SEC, 2 ** (rec["fails"] + 1))
                    log.warning("업로드 실패(재시도 %s초 후): %s: %s", delay, rec["name"], e)
                    self._set_status(rec, f"업로드 실패: {msg} - {delay}초 후 다시 시도")
                    self._retry_later(rec, delay)
                else:
                    log.error("업로드 실패(자동 재시도 안 함): %s: %r", rec["name"], e)
                    self._set_status(rec, f"업로드 실패: {msg} (다시 저장하면 재시도)")
        finally:
            if snap is not None:
                snap.discard()
            with self._lock:
                rec["busy"] = False
            self._wake.set()

    def _upload(self, rec: dict, drive, snap: _Snapshot) -> None:
        if rec.get("file_id") is None:
            self._create_new(rec, drive, snap)
            return
        with self._lock:
            file_id = rec["file_id"]
            synced = rec["synced_md5"]
            ours = set(rec["known_md5"]) | rec["attempted"]
        try:
            remote = drive.get_meta(file_id)
        except Exception as e:
            if not is_not_found(e):
                raise
            remote = None

        remote_md5 = (remote or {}).get("md5Checksum")
        if remote is not None and remote_md5 == snap.md5 and not remote.get("trashed"):
            # 이미 같은 내용이 올라가 있음 (지난 업로드의 응답만 못 받은 경우 등)
            self._mark_synced(rec, snap.md5, remote)
            return

        reason = None
        if remote is None:
            reason = "드라이브에서 원본이 삭제되어"
        elif remote.get("trashed"):
            reason = "원본이 휴지통에 있어"
        elif remote_md5 and synced and remote_md5 != synced and remote_md5 not in ours:
            reason = "다른 곳에서 원본이 수정되어"

        with self._lock:
            rec["attempted"].add(snap.md5)

        if reason is None:
            new = drive.update_content(file_id, snap.source, rec["mime"])
            got = new.get("md5Checksum")
            if got and got != snap.md5:
                with self._lock:
                    rec["attempted"].add(got)  # 우리가 올린 것이므로 재시도 때 충돌로 보지 않음
                raise IntegrityError("업로드한 내용의 검사값이 다릅니다. 다시 시도합니다.")
            self._mark_synced(rec, snap.md5, new)
            return

        # 충돌: 원본을 덮어쓰지 않고 같은 폴더에 사본으로 올린다
        copy_name = conflict_name(rec["name"])
        parent = ((remote or {}).get("parents") or [rec.get("parent_id") or "root"])[0]
        try:
            new = drive.create_file(copy_name, parent, snap.source, rec["mime"])
        except Exception as e:
            if is_transient(e) or isinstance(e, AuthError):
                raise
            log.warning("원래 폴더에 사본을 만들 수 없어 내 드라이브에 만듭니다: %r", e)
            parent = "root"
            new = drive.create_file(copy_name, parent, snap.source, rec["mime"])
        with self._lock:
            rec["file_id"] = new["id"]
            rec["parent_id"] = parent
            rec["drive_id"] = new.get("driveId")
            rec["name"] = copy_name
        self._mark_synced(rec, snap.md5, new, status=f"충돌 사본으로 저장됨 {time.strftime('%H:%M:%S')}")
        self.notify("conflict", {"reason": reason, "name": copy_name, "path": rec["path"]})

    # ----- 손상·암호화 감지 -----
    def _looks_damaged(self, rec: dict, snap: _Snapshot) -> bool:
        if rec.get("force"):
            return False
        if snap.size == 0 and rec.get("synced_md5") not in (None, EMPTY_MD5):
            return True  # 내용이 있던 문서가 갑자기 0바이트
        fmt = rec.get("fmt")
        if not fmt:
            return False
        if snap.data is not None:
            head = snap.data[:32]
        else:
            try:
                with open(snap.staged, "rb") as f:
                    head = f.read(32)
            except OSError:
                return False
        if content_format(head) != fmt:
            return True  # 형식이 바뀜 (예: docx였는데 zip이 아님 → 암호화·손상 의심)
        if fmt == "zip" and snap.data is not None and len(snap.data) <= ZIP_TEST_LIMIT:
            return not _zip_ok(snap.data)
        return False

    def _hold(self, rec: dict) -> None:
        now = time.monotonic()
        with self._lock:
            self._holds = [(t, k) for t, k in self._holds if now - t < 600]
            if rec["key"] not in [k for _, k in self._holds]:
                self._holds.append((now, rec["key"]))
            many = len(self._holds) >= 3 and not self.paused
            if many:
                self.paused = True
            names = [self._recs[k]["name"] for _, k in self._holds if k in self._recs]
        log.warning("업로드 보류(손상·암호화 의심): %s", rec["name"])
        self._set_status(rec, STATUS_HELD)
        self.notify("suspicious", {"name": rec["name"]})
        if many:
            log.error("여러 파일이 한꺼번에 손상/암호화됨 → 모든 업로드 중지: %s", names)
            self.notify("paused", {"names": names})

    def remote_changed(self, file_id: str) -> None:
        """드라이브 쪽 내용이 바뀌었다고 알림 (타임머신으로 이전 버전을 되돌린 경우).
        이 PC의 사본은 낡은 것으로 보고, 그 사본을 저장하면 덮어쓰지 않고 충돌 사본으로 올린다."""
        with self._lock:
            for rec in self._recs.values():
                if rec["file_id"] == file_id:
                    rec["known_md5"] = [rec["synced_md5"]] if rec["synced_md5"] else []
                    rec["attempted"] = set()
        self.save_state()

    def force_upload(self, key: str) -> bool:
        """사용자가 확인하고 '그래도 업로드'를 고른 파일."""
        with self._lock:
            rec = self._recs.get(key)
            if not rec or rec["kind"] != KIND_EDIT:
                return False
            rec["force"] = True
            rec["due"] = time.monotonic()
        self._wake.set()
        return True

    def resume(self) -> None:
        """랜섬웨어 의심으로 멈춘 업로드를 다시 시작한다."""
        with self._lock:
            self.paused = False
            self._holds.clear()
        self._wake.set()
        self.notify("changed", None)

    def _create_new(self, rec: dict, drive, snap: _Snapshot) -> None:
        """'다른 이름으로 저장'한 새 파일을 원래 파일이 있던 드라이브 폴더에 새로 만든다."""
        name = rec["name"]
        parent = rec.get("parent_id") or "root"
        new = None
        if rec.get("create_tried"):
            # 지난번에 만들기 요청을 보냈는데 응답을 못 받았을 수 있음 → 이미 있으면 그걸 씀 (중복 방지)
            for p in dict.fromkeys([parent, "root"]):
                try:
                    found = drive.find_in_folder(name, p)
                except Exception as e:
                    if is_transient(e) or isinstance(e, AuthError):
                        raise
                    found = []
                new = next((m for m in found if m.get("md5Checksum") == snap.md5), None)
                if new is not None:
                    parent = p
                    break
        if new is None:
            with self._lock:
                rec["create_tried"] = True
            self.save_state()  # 응답 전에 꺼져도 다음 실행 때 중복으로 만들지 않게 기록
            try:
                new = drive.create_file(name, parent, snap.source, rec["mime"])
            except Exception as e:
                if is_transient(e) or isinstance(e, AuthError) or parent == "root":
                    raise
                # 원래 폴더에 쓸 권한이 없거나 폴더가 지워짐 → 내 드라이브에 만든다
                log.warning("원래 폴더에 새 파일을 만들 수 없어 내 드라이브에 만듭니다: %r", e)
                parent = "root"
                new = drive.create_file(name, parent, snap.source, rec["mime"])
        with self._lock:
            rec["file_id"] = new["id"]
            rec["parent_id"] = parent
            rec["drive_id"] = new.get("driveId")
            rec["mime"] = new.get("mimeType") or rec["mime"]
            if snap.data is not None:
                fmt = content_format(snap.data[:32])
                rec["fmt"] = fmt if fmt != "zip" or _zip_ok(snap.data) else None
        where = "내 드라이브" if parent == "root" else "원래 파일과 같은 폴더"
        self._mark_synced(rec, snap.md5, new,
                          status=f"새 파일로 드라이브에 저장됨 ({where}) {time.strftime('%H:%M:%S')}")
        self.notify("new_file", {"name": name, "where": where})

    def _mark_synced(self, rec: dict, md5: str, remote: dict, status: str | None = None) -> None:
        with self._lock:
            rec["synced_md5"] = md5
            known = [m for m in rec["known_md5"] if m != md5]
            rec["known_md5"] = ([md5] + known)[:20]
            rec["remote_modified"] = remote.get("modifiedTime", rec.get("remote_modified", ""))
            rec["fails"] = 0
            rec["force"] = False
            rec["status"] = status or f"드라이브에 저장됨 {time.strftime('%H:%M:%S')}"
        self.save_state()
        self.notify("changed", None)

    # ----- 정리 -----
    def cleanup(self) -> tuple[int, list[str]]:
        """올리기가 끝나고 닫힌 파일의 로컬 사본을 지운다. (지운 개수, 남긴 파일 이름)"""
        removed, kept = 0, []
        with self._lock:
            recs = list(self._recs.values())
        titles = visible_window_titles() if recs else []
        for rec in recs:
            name, path = rec["name"], Path(rec["path"])
            with self._lock:
                if rec["busy"] or rec["due"] is not None:
                    kept.append(name)
                    continue
                rec["busy"] = True  # 지우는 동안 업로드 루프가 건드리지 않게
            try:
                if path.exists():
                    if file_md5(path) != rec["synced_md5"]:
                        # 아직 업로드 안 된 변경이 있음 (복사본·보기 전용 파일을 고친 경우도 지우지 않음)
                        kept.append(name)
                        continue
                    if (is_in_use(path) or self._lock_file_in_use(path.parent)
                            or title_mentions(titles, path.name)):
                        kept.append(name)  # Word/Excel/한글 등에서 아직 열려 있음
                        continue
                    if self._has_untracked_new(path.parent):
                        kept.append(name)  # 같은 폴더에 아직 올리지 않은 새 파일이 있음 (다른 이름으로 저장 직후)
                        continue
                    # 이름 바꾸기가 되면 아무도 잡고 있지 않다는 뜻 → 지움
                    tmp = path.with_name(path.name + DEL_SUFFIX)
                    try:
                        os.replace(path, tmp)
                    except OSError:
                        kept.append(name)
                        continue
                    try:
                        os.unlink(tmp)
                    except OSError:
                        try:
                            os.replace(tmp, path)
                        except OSError:
                            pass
                        kept.append(name)
                        continue
                with self._lock:
                    self._recs.pop(rec["key"], None)
                    folder_used = any(os.path.dirname(r["path"]) == str(path.parent)
                                      for r in self._recs.values())
                if not folder_used and path.parent != self.work_dir:
                    try:
                        (path.parent / SIDECAR).unlink()
                    except OSError:
                        pass
                    self._sidecars.pop(str(path.parent), None)
                    try:
                        if not any(path.parent.iterdir()):
                            path.parent.rmdir()
                    except OSError:
                        pass
                removed += 1
            finally:
                with self._lock:
                    rec["busy"] = False
        self.save_state()
        self.notify("changed", None)
        return removed, kept

    def _has_untracked_new(self, folder: Path) -> bool:
        try:
            names = os.listdir(folder)
        except OSError:
            return False
        with self._lock:
            for n in names:
                p = os.path.join(folder, n)
                if is_new_file_candidate(n) and norm(p) not in self._recs and os.path.isfile(p):
                    return True
        return False

    @staticmethod
    def _lock_file_in_use(folder: Path) -> bool:
        # Word/Excel이 만드는 '~$이름' 소유자 파일이 잡혀 있으면 문서가 열려 있는 것
        try:
            for p in folder.iterdir():
                if p.name.startswith("~$") and is_in_use(p):
                    return True
        except OSError:
            pass
        return False
