"""앱 경로와 설정값."""
from __future__ import annotations

import json
import os
import shutil
import sys
import tempfile
import time
from pathlib import Path

APP_NAME = "DriveKeeper"          # 데이터 폴더·exe 이름
APP_TITLE = "드라이브 지킴이"       # 창 제목 (중복 실행 확인에도 씀)
VERSION = "2.0.0"
GITHUB_REPO = "janghwansangai/drive-keeper"
SITE_URL = "https://janghwansangai.github.io/drive-keeper/"
RELEASES_URL = f"https://github.com/{GITHUB_REPO}/releases/latest"

# 드라이브 전체 읽기/쓰기 권한 (파일 열기·수정 저장에 필요)
SCOPES = ["https://www.googleapis.com/auth/drive"]

# 저장 이벤트가 연달아 들어와도 마지막 저장 후 이 시간(초)이 지나면 한 번만 업로드
UPLOAD_DEBOUNCE_SEC = 0.7
# 파일 감시(watchdog)가 놓친 저장도 잡기 위해 편집 중인 파일을 이 주기(초)로 직접 확인
POLL_SEC = 1.0
# '다른 이름으로 저장'으로 생긴 새 파일은 이 시간(초) 동안 그대로 있어야 올린다 (잠깐 생겼다 사라지는 임시 파일 거르기)
NEW_FILE_DELAY_SEC = 2.0
# 업로드 실패 시 재시도 간격(초): 4, 8, 16, 32, 60, 60, ...
RETRY_MAX_SEC = 60
# 이보다 큰 파일은 메모리 대신 임시 사본으로 업로드
MEMORY_UPLOAD_LIMIT = 64 * 1024 * 1024
# 브라우저 로그인을 기다리는 최대 시간(초)
LOGIN_TIMEOUT_SEC = 300
# API 호출 하나의 네트워크 제한 시간(초)
HTTP_TIMEOUT_SEC = 60
# Word/Excel/한글이 다룰 수 있는 전체 경로 길이 (Excel 한계 218자에 맞춤)
MAX_LOCAL_PATH = 215


def _base_dir() -> Path:
    # Windows: C:\Users\<이름>\GDriveExplorer
    # AppData가 아닌 사용자 폴더에 둔다. AppData는 '앱 격리'(MSIX)된 프로그램(예: Claude 앱, Microsoft Store
    # 앱)을 거쳐 실행하면 Windows가 몰래 다른 폴더로 바꿔 저장해서, 더블클릭으로 켰을 때와 기록이 갈라진다.
    if os.name == "nt":
        root = os.environ.get("USERPROFILE") or os.path.expanduser("~")
    else:
        root = os.environ.get("XDG_DATA_HOME") or os.path.join(
            os.path.expanduser("~"), ".local", "share")
    return Path(root) / APP_NAME


def legacy_base_dirs() -> list[Path]:
    """예전 버전이 쓰던 위치들 (가까운 것부터). 옛 이름 'GDriveExplorer' 시절 포함."""
    out = []
    if os.name == "nt":
        home = os.environ.get("USERPROFILE") or os.path.expanduser("~")
        out.append(Path(home) / "GDriveExplorer")
        if os.environ.get("LOCALAPPDATA"):
            out.append(Path(os.environ["LOCALAPPDATA"]) / "GDriveExplorer")
    return out


def _project_dir() -> Path:
    # exe로 빌드했으면 exe가 있는 폴더, 아니면 이 파일이 있는 폴더
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


PROJECT_DIR = _project_dir()
BASE_DIR = _base_dir()
TOKEN_PATH = BASE_DIR / "token.json"
STATE_PATH = BASE_DIR / "state.json"      # 편집 중인 파일 기록 (재시작해도 업로드 이어감)
LOG_PATH = BASE_DIR / "app.log"
LOCK_PATH = BASE_DIR / "app.lock"         # 중복 실행 방지
STAGING_DIR = BASE_DIR / "staging"        # 큰 파일 업로드용 임시 사본
# 편집 중인 파일만 잠깐 내려받는 작업 폴더 (드라이브 전체가 아님)
WORK_DIR = BASE_DIR / "work"


def ensure_dirs() -> None:
    for d in (BASE_DIR, WORK_DIR, STAGING_DIR):
        d.mkdir(parents=True, exist_ok=True)


def migrate_legacy_base() -> Path | None:
    """예전 위치에 있던 로그인·계정·편집 기록·작업 폴더를 새 위치로 복사한다 (새 위치에 계정이 없을 때만).

    예전 위치의 파일은 지우지 않고 그대로 둔다 (만일을 위한 백업). 옮긴 위치를 돌려준다.
    """
    if (BASE_DIR / "accounts.json").exists() or (BASE_DIR / "token.json").exists():
        return None
    for old in legacy_base_dirs():
        try:
            if old.resolve() == BASE_DIR.resolve() or not old.is_dir():
                continue
            if not ((old / "accounts.json").exists() or (old / "token.json").exists()):
                continue
        except OSError:
            continue
        BASE_DIR.mkdir(parents=True, exist_ok=True)
        try:
            shutil.copytree(old, BASE_DIR, dirs_exist_ok=True,
                            ignore=shutil.ignore_patterns("app.lock", "*.gde-part", "*.gde-del", "*.tmp"))
        except shutil.Error:
            pass  # 일부 파일을 못 읽어도 나머지는 옮긴다 (예전 위치에 원본이 남아 있음)
        _rewrite_paths(old)
        try:
            (old / "옮겨짐.txt").write_text(
                f"이 폴더의 내용은 {BASE_DIR} 로 옮겨졌습니다.\n예전 기록의 백업으로 남겨 둔 것이니 지워도 됩니다.\n",
                encoding="utf-8")
        except OSError:
            pass
        return old
    return None


def _rewrite_paths(old: Path) -> None:
    """계정 기록 안의 작업 폴더 경로(예전 위치)를 새 위치로 바꾼다."""
    old_s, new_s = str(old), str(BASE_DIR)
    files = [BASE_DIR / "accounts.json"] + list((BASE_DIR / "accounts").glob("*/account.json"))
    for f in files:
        try:
            data = json.loads(f.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            continue
        items = data.get("accounts", []) if "accounts" in data else [data]
        changed = False
        for a in items:
            wd = str(a.get("work_dir") or "")
            if wd and os.path.normcase(wd).startswith(os.path.normcase(old_s)):
                a["work_dir"] = new_s + wd[len(old_s):]
                changed = True
        if changed:
            atomic_write_text(f, json.dumps(data, ensure_ascii=False, indent=1))


# ---------- 앱 열쇠 (client_secret.json): 사용자마다 자기 구글 클라우드 프로젝트에서 만든 것 ----------
def check_client_secret(path: Path) -> dict:
    """열쇠 파일이 맞는지 확인하고 내용을 돌려준다. 아니면 쉬운 설명과 함께 ValueError."""
    try:
        data = json.loads(Path(path).read_text(encoding="utf-8"))
    except (OSError, ValueError):
        raise ValueError("열쇠 파일을 읽을 수 없습니다. 구글에서 받은 .json 파일인지 확인해 주세요.")
    if not isinstance(data, dict):
        raise ValueError("열쇠 파일이 아닙니다.")
    if "web" in data:
        raise ValueError("이 열쇠는 '웹 애플리케이션' 종류입니다.\n"
                         "열쇠를 만들 때 애플리케이션 유형을 '데스크톱 앱'으로 골라 다시 만들어 주세요.")
    info = data.get("installed")
    if not isinstance(info, dict) or not info.get("client_id") or not info.get("client_secret"):
        raise ValueError("구글 앱 열쇠 파일이 아닙니다.\n"
                         "구글 클라우드에서 '데스크톱 앱' 열쇠를 만들고 받은 client_secret_....json 파일을 골라 주세요.")
    return data


def import_client_secret(path: Path) -> str:
    """사용자가 고른 열쇠 파일을 프로그램 데이터 폴더에 저장한다. 프로젝트 이름을 돌려준다."""
    data = check_client_secret(path)
    atomic_write_text(BASE_DIR / "client_secret.json", json.dumps(data, ensure_ascii=False))
    return str(data["installed"].get("project_id") or "")


def downloads_dir() -> Path:
    """Windows '다운로드' 폴더 (옮겨 두었어도 찾음)."""
    if os.name == "nt":
        try:
            import ctypes
            from ctypes import wintypes

            class GUID(ctypes.Structure):
                _fields_ = [("a", wintypes.DWORD), ("b", wintypes.WORD), ("c", wintypes.WORD),
                            ("d", ctypes.c_ubyte * 8)]
            fid = GUID(0x374DE290, 0x123F, 0x4565, (ctypes.c_ubyte * 8)(0x91, 0x64, 0x39, 0xC4, 0x92, 0x5E, 0x46, 0x7B))
            out = ctypes.c_wchar_p()
            shell32 = ctypes.windll.shell32
            if shell32.SHGetKnownFolderPath(ctypes.byref(fid), 0, None, ctypes.byref(out)) == 0:
                path = out.value
                ctypes.windll.ole32.CoTaskMemFree(out)
                if path:
                    return Path(path)
        except Exception:
            pass
    return Path(os.path.expanduser("~")) / "Downloads"


def find_downloaded_secrets() -> list[Path]:
    """다운로드 폴더·바탕화면에서 구글 열쇠 파일(client_secret*.json)을 찾는다 (최근 것부터)."""
    places = [downloads_dir(), Path(os.path.expanduser("~")) / "Desktop"]
    found: list[Path] = []
    for place in places:
        try:
            for p in place.glob("client_secret*.json"):
                try:
                    check_client_secret(p)
                    found.append(p)
                except ValueError:
                    pass
        except OSError:
            pass
    found.sort(key=lambda p: p.stat().st_mtime, reverse=True)
    return found


def find_client_secret() -> Path | None:
    """프로그램 폴더 → 앱 데이터 폴더 순서로 client_secret.json을 찾는다.

    프로그램 폴더에서 찾으면 앱 데이터 폴더에도 복사해 둔다. 그래야 exe를 다른 곳으로
    옮겨도 7일마다 하는 재로그인이 된다. (token.json에도 같은 값이 들어 있어 노출이 늘지 않음)
    """
    candidates = [PROJECT_DIR / "client_secret.json"]
    if getattr(sys, "frozen", False):
        candidates.append(PROJECT_DIR.parent / "client_secret.json")
    for p in candidates:
        if p.is_file():
            try:
                saved = BASE_DIR / "client_secret.json"
                if not saved.exists() or saved.read_bytes() != p.read_bytes():
                    BASE_DIR.mkdir(parents=True, exist_ok=True)
                    shutil.copyfile(p, saved)
            except OSError:
                pass
            return p
    saved = BASE_DIR / "client_secret.json"
    return saved if saved.is_file() else None


def atomic_write_text(path: Path, text: str) -> None:
    """쓰는 도중 꺼져도 파일이 깨지지 않게 임시 파일에 쓴 뒤 바꿔치기한다."""
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, tmp = tempfile.mkstemp(prefix=path.name + ".", suffix=".tmp", dir=str(path.parent))
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            f.write(text)
            f.flush()
            os.fsync(f.fileno())
        for attempt in range(10):
            try:
                os.replace(tmp, path)
                break
            except PermissionError:
                # 백신/검사 프로그램이 잠깐 잡고 있는 경우
                if attempt == 9:
                    raise
                time.sleep(0.05 * (attempt + 1))
    except BaseException:
        try:
            os.unlink(tmp)
        except OSError:
            pass
        raise
