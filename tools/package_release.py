"""배포용 압축 파일 만들기: release/DriveKeeper.zip

먼저 build_exe.py로 dist/DriveKeeper 를 만든 뒤 실행.
압축 안 구성 (폴더 '드라이브 지킴이'):
  설치하기.cmd · 지우기.cmd · 처음 읽어 주세요.txt
  드라이브 지킴이 사용설명서.pdf · 드라이브 지킴이 설명서 (AI용).md
  app/  (DriveKeeper.exe, _internal, install.ps1, uninstall.ps1)
"""
from __future__ import annotations

import shutil
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
import config  # noqa: E402

TOP = "드라이브 지킴이"
FORBIDDEN = ("client_secret", "token", "security.json", "accounts.json", "handoff", "claude.md")

INSTALL_CMD = r"""@echo off
rem DriveKeeper installer (no admin rights needed)
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0app\install.ps1"
"""
UNINSTALL_CMD = r"""@echo off
rem DriveKeeper uninstaller (your Google Drive files are not touched)
set "U=%LOCALAPPDATA%\Programs\DriveKeeper\uninstall.ps1"
if not exist "%U%" set "U=%~dp0app\uninstall.ps1"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%U%"
"""


def readme_text() -> str:
    return f"""드라이브 지킴이 {config.VERSION} — 처음 읽어 주세요
==================================================

구글 드라이브를 비밀번호로 지켜 주는 프로그램이에요.

■ 설치는 딱 두 번만 누르면 끝!
  1) 이 폴더의 [설치하기] 를 두 번 눌러요.
     ('Windows의 PC 보호' 파란 창이 뜨면 → [추가 정보] → [실행])
  2) "설치가 끝났어요!" 가 나오면 [예] 를 눌러요.
  → 바탕화면에 '드라이브 지킴이' 아이콘이 생겨요.

■ 처음 켜면
  '시작 도우미' 가 나와요. 1번 칸의 [그림으로 따라 하기] 를 누르면
  한 화면에 한 단계씩 그림으로 보여 줘요. 주황색 동그라미만 따라 누르면
  ① 내 구글 열쇠 만들기  ② 열쇠 넣기  ③ 구글 로그인
  까지 약 5분이면 끝나요.
  (인터넷 주소: {config.GUIDE_URL}
   인쇄해서 보려면 [드라이브 지킴이 사용설명서.pdf] 의 2장)

■ 지우고 싶으면
  [지우기] 를 두 번 누르거나,
  Windows [설정] → [앱] → '드라이브 지킴이' → [제거]
  (구글 드라이브의 파일은 그대로 있어요)

■ 새 버전
  프로그램이 새 버전을 알려 주면 홈페이지에서 받아
  [설치하기] 를 다시 누르면 돼요. 설정과 로그인은 그대로예요.

홈페이지: {config.SITE_URL}
"""


def main() -> int:
    dist = ROOT / "dist" / config.APP_NAME
    exe = dist / f"{config.APP_NAME}.exe"
    if not exe.exists():
        print("먼저 build_exe.py를 실행하세요:", exe)
        return 1
    out = ROOT / "release"
    stage = out / TOP
    if out.exists():
        shutil.rmtree(out)
    app_dir = stage / "app"
    shutil.copytree(dist, app_dir)
    for n in ("install.ps1", "uninstall.ps1"):
        # PowerShell 5.1이 한글을 바르게 읽도록 UTF-8 BOM으로 저장
        text = (ROOT / "installer" / n).read_text(encoding="utf-8-sig")
        (app_dir / n).write_text(text, encoding="utf-8-sig", newline="\r\n")
    (stage / "설치하기.cmd").write_bytes(INSTALL_CMD.replace("\n", "\r\n").encode("ascii"))
    (stage / "지우기.cmd").write_bytes(UNINSTALL_CMD.replace("\n", "\r\n").encode("ascii"))
    (stage / "처음 읽어 주세요.txt").write_text(readme_text(), encoding="utf-8-sig", newline="\r\n")
    shutil.copyfile(ROOT / "docs" / "manual" / "DriveKeeper-Manual.pdf", stage / "드라이브 지킴이 사용설명서.pdf")
    shutil.copyfile(ROOT / "docs" / "manual" / "DriveKeeper-AI-Guide.md", stage / "드라이브 지킴이 설명서 (AI용).md")

    bad = [p for p in stage.rglob("*") if p.is_file() and any(f in p.name.lower() for f in FORBIDDEN)]
    if bad:
        print("배포하면 안 되는 파일이 들어 있어요:", bad)
        return 1

    zpath = out / "DriveKeeper.zip"
    with zipfile.ZipFile(zpath, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as z:
        for p in sorted(stage.rglob("*")):
            if p.is_file():
                z.write(p, p.relative_to(out).as_posix())
    with zipfile.ZipFile(zpath) as z:
        assert z.testzip() is None
        count = len(z.namelist())
    print(f"만듦: {zpath} ({zpath.stat().st_size / 1024 / 1024:.1f}MB, 파일 {count}개)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
