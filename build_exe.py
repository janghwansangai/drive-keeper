"""DriveKeeper.exe 만들기 (PyInstaller, 폴더형 → 시작이 빠름).

실행: build_exe.bat (또는 .venv\\Scripts\\python build_exe.py)
결과: dist\\DriveKeeper\\DriveKeeper.exe
client_secret.json은 절대 넣지 않는다. 사용자가 각자 만든 열쇠를 프로그램에서 넣는다.
"""
from __future__ import annotations

import os
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
NAME = "DriveKeeper"


def write_version_file(path: Path, version: str) -> None:
    """exe 속성에 보이는 버전 정보 (설치 프로그램이 이 값을 읽어 [앱] 목록에 표시)."""
    nums = [int(x) for x in version.split(".")] + [0] * 4
    tup = tuple(nums[:4])
    path.write_text(f"""VSVersionInfo(
  ffi=FixedFileInfo(filevers={tup}, prodvers={tup}, mask=0x3f, flags=0x0, OS=0x40004,
                    fileType=0x1, subtype=0x0, date=(0, 0)),
  kids=[
    StringFileInfo([StringTable('041204B0', [
      StringStruct('CompanyName', 'DriveKeeper'),
      StringStruct('FileDescription', 'DriveKeeper - Google Drive password lock'),
      StringStruct('FileVersion', '{version}'),
      StringStruct('InternalName', '{NAME}'),
      StringStruct('OriginalFilename', '{NAME}.exe'),
      StringStruct('ProductName', 'DriveKeeper'),
      StringStruct('ProductVersion', '{version}')])]),
    VarFileInfo([VarStruct('Translation', [0x0412, 1200])])
  ]
)
""", encoding="utf-8")


def main() -> int:
    import PyInstaller.__main__

    sys.path.insert(0, str(ROOT))
    import branding
    import config

    build = ROOT / "build"
    # --dist 폴더: 실행 중인 exe가 dist를 잡고 있을 때 다른 곳에 만들어 두고 나중에 바꿔 넣기 위함
    dist = Path(sys.argv[sys.argv.index("--dist") + 1]).resolve() if "--dist" in sys.argv else ROOT / "dist"
    build.mkdir(exist_ok=True)
    icon = build / "app.ico"
    icon.write_bytes(branding.ico_bytes())
    verfile = build / "version_info.txt"
    write_version_file(verfile, config.VERSION)
    PyInstaller.__main__.run([
        str(ROOT / "app.py"),
        "--name", NAME,
        "--noconfirm", "--clean", "--windowed", "--onedir",
        "--icon", str(icon),
        "--version-file", str(verfile),
        "--additional-hooks-dir", str(ROOT / "build_hooks"),
        "--hidden-import", "drive_client",
        "--hidden-import", "watchdog.observers.read_directory_changes",
        "--hidden-import", "watchdog.observers.polling",
        "--exclude-module", "win32com",
        "--exclude-module", "pythoncom",
        "--exclude-module", "pywintypes",
        "--distpath", str(dist),
        "--workpath", str(build),
        "--specpath", str(build),
    ])

    out = dist / NAME
    exe = out / f"{NAME}.exe"
    docs = list(out.rglob("drive.v3.json"))
    secrets = list(out.rglob("client_secret*.json")) + list(out.rglob("token*.json"))
    size = sum(p.stat().st_size for p in out.rglob("*") if p.is_file()) / 1024 / 1024
    print(f"\nexe: {exe} ({'있음' if exe.exists() else '없음'})")
    print(f"drive.v3.json 포함: {bool(docs)} / 열쇠·로그인 파일 포함(없어야 정상): {bool(secrets)} / 크기 {size:.1f}MB")
    return 0 if exe.exists() and docs and not secrets else 1


if __name__ == "__main__":
    os.chdir(ROOT)
    sys.exit(main())
