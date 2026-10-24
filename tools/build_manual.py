"""사람용 설명서 PDF 만들기: 홈페이지의 따라 하기 안내(docs/start.html)를 Edge(또는 Chrome)로 인쇄.

화면용 안내와 PDF가 늘 같은 내용이 되도록 한 파일에서 만든다 (인쇄할 때는 모든 단계를 차례로 보여 줌).

실행: .venv\\Scripts\\python tools\\build_manual.py
결과: docs/manual/DriveKeeper-Manual.pdf
"""
from __future__ import annotations

import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / "docs" / "start.html"
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
        # ?print: 모든 단계를 펼쳐서 보여 준다. 글꼴·그림을 다 받을 때까지 기다린 뒤 인쇄
        cmd = [exe, "--headless", "--disable-gpu", "--no-first-run", f"--user-data-dir={prof}",
               "--no-pdf-header-footer", "--virtual-time-budget=15000", "--run-all-compositor-stages-before-draw",
               f"--print-to-pdf={OUT}", SRC.resolve().as_uri() + "?print"]
        subprocess.run(cmd, timeout=180, capture_output=True)
    if not OUT.exists():
        print("PDF를 만들지 못했습니다.")
        return 1
    print("만듦:", OUT, f"{OUT.stat().st_size / 1024:.0f}KB")
    return 0


if __name__ == "__main__":
    sys.exit(main())
