"""새 버전: GitHub에 올라온 최신 버전과 비교하고, 누르면 한 번에 업데이트한다.

- 확인: 구글 드라이브와 상관없는 GitHub 공개 정보만 읽는다 (사용자 정보는 보내지 않음).
- 업데이트: 이 저장소의 GitHub 릴리스에 올라온 DriveKeeper.zip만 받는다. GitHub가 알려 주는 SHA-256과
  맞는지, 배포 파일 구성이 맞는지 확인한 뒤 풀고, 프로그램이 닫히면 원래 설치 프로그램(install.ps1)으로
  바꿔 넣고 다시 켠다. 설치본(%LOCALAPPDATA%\\Programs\\DriveKeeper)으로 실행할 때만 한다.
"""
from __future__ import annotations

import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
import urllib.error
import urllib.request
import zipfile
from pathlib import Path

import config

API_URL = f"https://api.github.com/repos/{config.GITHUB_REPO}/releases/latest"
ASSET_NAME = "DriveKeeper.zip"
DOWNLOAD_PREFIX = f"https://github.com/{config.GITHUB_REPO}/releases/download/"
TOP = "드라이브 지킴이"          # zip 안 맨 위 폴더 (tools/package_release.py)
MAX_ZIP = 300 * 1024 * 1024


class UpdateError(Exception):
    """사용자에게 그대로 보여 줄 수 있는 업데이트 실패 이유."""


def parse_version(text: str) -> tuple[int, ...]:
    nums = re.findall(r"\d+", text or "")
    return tuple(int(n) for n in nums[:3]) or (0,)


def is_newer(latest: str, current: str = config.VERSION) -> bool:
    return parse_version(latest) > parse_version(current)


def check_latest(timeout: float = 6.0) -> dict | None:
    """{'version', 'url': 다운로드 페이지, 'notes': 바뀐 점, 'zip_url', 'zip_size', 'sha256'} 또는 None(올린 버전이 없음)."""
    req = urllib.request.Request(API_URL, headers={
        "Accept": "application/vnd.github+json",
        "User-Agent": f"{config.APP_NAME}/{config.VERSION}",
    })
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            data = json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        if e.code == 404:
            return None
        raise
    tag = str(data.get("tag_name") or "")
    if not tag:
        return None
    asset = next((a for a in data.get("assets") or [] if a.get("name") == ASSET_NAME), {})
    zip_url = str(asset.get("browser_download_url") or "")
    digest = str(asset.get("digest") or "")
    return {"version": tag.lstrip("vV"), "url": data.get("html_url") or config.RELEASES_URL,
            "notes": str(data.get("body") or "")[:2000],
            "zip_url": zip_url if zip_url.startswith(DOWNLOAD_PREFIX) else "",
            "zip_size": int(asset.get("size") or 0),
            "sha256": digest[7:].lower() if digest.startswith("sha256:") else ""}


def install_dir() -> Path | None:
    """설치본으로 실행 중이면 그 폴더. 아니면 None (개발용 실행·압축만 푼 exe는 한 번에 업데이트하지 않음)."""
    if not getattr(sys, "frozen", False) or os.name != "nt":
        return None
    here = Path(sys.executable).resolve().parent
    want = Path(os.environ.get("LOCALAPPDATA") or "") / "Programs" / config.APP_NAME
    try:
        return here if want.exists() and here == want.resolve() else None
    except OSError:
        return None


def can_self_update(info: dict | None) -> bool:
    return bool(info and info.get("zip_url")) and install_dir() is not None


def work_dir() -> Path:
    return config.BASE_DIR / "update"


def download(info: dict, progress=None) -> Path:
    """새 버전 zip을 받아 검사하고 풀어 둔다. 풀린 설치 프로그램(app\\install.ps1)의 경로를 돌려준다.

    progress(받은 바이트, 전체 바이트)는 받는 동안 여러 번 불린다 (작업 스레드에서).
    """
    url = info.get("zip_url") or ""
    if not url.startswith(DOWNLOAD_PREFIX):
        raise UpdateError("이 버전은 한 번에 업데이트할 수 없어요. 다운로드 페이지에서 받아 주세요.")
    root = work_dir()
    shutil.rmtree(root, ignore_errors=True)
    root.mkdir(parents=True, exist_ok=True)
    zpath = root / ASSET_NAME
    req = urllib.request.Request(url, headers={"User-Agent": f"{config.APP_NAME}/{config.VERSION}"})
    h = hashlib.sha256()
    got = 0
    with urllib.request.urlopen(req, timeout=30) as resp, open(zpath, "wb") as f:
        total = int(resp.headers.get("Content-Length") or info.get("zip_size") or 0)
        if total > MAX_ZIP:
            raise UpdateError("받을 파일이 너무 커요. 다운로드 페이지에서 받아 주세요.")
        while True:
            chunk = resp.read(256 * 1024)
            if not chunk:
                break
            got += len(chunk)
            if got > MAX_ZIP:
                raise UpdateError("받을 파일이 너무 커요. 다운로드 페이지에서 받아 주세요.")
            f.write(chunk)
            h.update(chunk)
            if progress is not None:
                progress(got, total)
    if total and got != total:
        raise UpdateError("새 버전을 끝까지 받지 못했어요. 인터넷 연결을 확인하고 다시 해 주세요.")
    if info.get("sha256") and h.hexdigest() != info["sha256"]:
        raise UpdateError("받은 파일이 GitHub에 올라온 파일과 달라요. 설치하지 않았어요.")
    out = root / "files"
    try:
        with zipfile.ZipFile(zpath) as z:
            names = z.namelist()
            for need in (f"{TOP}/app/{config.APP_NAME}.exe", f"{TOP}/app/install.ps1"):
                if need not in names:
                    raise UpdateError("드라이브 지킴이 배포 파일이 아니에요. 설치하지 않았어요.")
            base = out.resolve()
            for n in names:   # '../' 같은 이름으로 다른 폴더에 파일을 쓰지 못하게
                if not (base / n).resolve().is_relative_to(base):
                    raise UpdateError("배포 파일 안에 이상한 경로가 있어요. 설치하지 않았어요.")
            z.extractall(out)
    except zipfile.BadZipFile:
        raise UpdateError("받은 파일이 깨졌어요. 다시 해 주세요.") from None
    return out / TOP / "app" / "install.ps1"


# 프로그램이 닫히기를 기다렸다가 설치 프로그램을 조용히 돌리고, 설치된 프로그램을 다시 켠다
HELPER_PS1 = r"""param([int]$WaitPid, [string]$Installer, [string]$Exe, [string]$Log)
Wait-Process -Id $WaitPid -Timeout 180 -ErrorAction SilentlyContinue
Start-Sleep -Milliseconds 500
$out = & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $Installer -Quiet 2>&1
$code = $LASTEXITCODE
@("exit=$code") + @($out | ForEach-Object { "$_" }) | Out-File -FilePath $Log -Encoding utf8
if (Test-Path $Exe) { Start-Process -FilePath $Exe }
"""


def launch_installer(installer: Path) -> None:
    """이 프로그램이 끝나면 설치하고 다시 켜는 도우미를 띄운다. 부른 뒤에는 프로그램을 닫아야 한다."""
    target = install_dir()
    if target is None:
        raise UpdateError("설치된 프로그램이 아니라서 한 번에 업데이트할 수 없어요.")
    root = work_dir()
    helper = root / "run_update.ps1"
    helper.write_text(HELPER_PS1, encoding="utf-8-sig")
    flags = 0x00000008 | 0x00000200 | 0x08000000   # DETACHED_PROCESS | NEW_PROCESS_GROUP | NO_WINDOW
    subprocess.Popen(["powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass", "-WindowStyle", "Hidden",
                      "-File", str(helper), "-WaitPid", str(os.getpid()), "-Installer", str(installer),
                      "-Exe", str(target / f"{config.APP_NAME}.exe"), "-Log", str(root / "install.log")],
                     creationflags=flags, close_fds=True)


def finished_message() -> str | None:
    """지난번 업데이트 결과 (다시 켰을 때 한 번 알려 줌). 받아 둔 파일은 지운다."""
    root = work_dir()
    log = root / "install.log"
    if not log.exists():
        return None
    try:
        text = log.read_text(encoding="utf-8-sig", errors="replace")
    except OSError:
        text = ""
    shutil.rmtree(root, ignore_errors=True)
    if text.startswith("exit=0"):
        return f"새 버전 {config.VERSION}으로 업데이트했어요."
    err = next((ln[6:].strip() for ln in text.splitlines() if ln.startswith("ERROR:")), "")
    return "업데이트하지 못했어요" + (f": {err}" if err else "") + ". 다운로드 페이지에서 받아 설치해 주세요."
