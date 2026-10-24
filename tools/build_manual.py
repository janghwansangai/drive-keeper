"""사람용 설명서 PDF 만들기: manual/manual.html → Edge(또는 Chrome)로 인쇄.

실행: .venv\\Scripts\\python tools\\build_manual.py
결과: docs/manual/DriveKeeper-Manual.pdf
"""
from __future__ import annotations

import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / "manual" / "manual.html"
OUT = ROOT / "docs" / "manual" / "DriveKeeper-Manual.pdf"
BROWSERS = [r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
            r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
            r"C:\Program Files\Google\Chrome\Application\chrome.exe"]


def main() -> int:
    exe = next((b for b in BROWSERS if Path(b).exists()), None)
    if not exe:
        print("Edge나 Chrome을 찾을 수 없습니다.")
        return 1
    OUT.parent.mkdir(parents=True, exist_ok=True)
    # 브라우저 임시 프로필은 AppData가 아닌 build 폴더에 (앱 격리 환경에서 실행해도 Edge가 볼 수 있게)
    prof = ROOT / "build" / "edge_profile"
    prof.mkdir(parents=True, exist_ok=True)
    if True:
        if OUT.exists():
            OUT.unlink()
        cmd = [exe, "--headless", "--disable-gpu", "--no-first-run", f"--user-data-dir={prof}",
               "--no-pdf-header-footer", f"--print-to-pdf={OUT}", SRC.resolve().as_uri()]
        subprocess.run(cmd, timeout=180, capture_output=True)
    if not OUT.exists():
        print("PDF를 만들지 못했습니다.")
        return 1
    print("만듦:", OUT, f"{OUT.stat().st_size / 1024:.0f}KB")
    return 0


if __name__ == "__main__":
    sys.exit(main())
