"""한컴(한글·한워드)이 파일을 연 상태에서 이 프로그램이 읽을 수 있는지 확인.

한글이 파일을 '읽기 금지'로 잠그면 저장해도 문서를 닫을 때까지 업로드할 수 없으므로 중요하다.
앱과 똑같이 os.startfile로 연다 (한글 창이 잠깐 화면에 떴다가 닫힘). 저장은 하지 않는다.
실행: .venv\\Scripts\\python tests\\real_hancom_lock_test.py
"""
from __future__ import annotations

import hashlib
import os
import shutil
import sys
import tempfile
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(Path(__file__).resolve().parent))

import sync  # noqa: E402
from win_helpers import (child_texts, close_process_windows, kill, make_docx_with_word, pids_of,  # noqa: E402
                         windows_of)

HNC = Path(r"C:\Program Files (x86)\Hnc\Office 2020\HOffice110\Bin\Hwp\DAISY\DaisyStyle2.hwp")
RESULTS: list[tuple[str, bool, str]] = []


def check(name: str, ok: bool, detail: str = "") -> None:
    RESULTS.append((name, ok, detail))
    print(("  [통과] " if ok else "  [실패] ") + name + (f" - {detail}" if detail else ""), flush=True)


def run_case(exe: str, path: Path) -> None:
    print(f"\n=== {exe}: {path.name} ===", flush=True)
    before = pids_of(exe)
    if before:
        print(f"  {exe}가 이미 실행 중이라 건너뜀 (사용자 문서를 건드리지 않기 위해)")
        return
    original = hashlib.md5(path.read_bytes()).hexdigest()
    os.startfile(str(path))
    pid = None
    end = time.monotonic() + 45
    while time.monotonic() < end:
        new = pids_of(exe) - before
        if new:
            pid = next(iter(new))
            if any(v and c != "#32770" and t for _, v, c, t in windows_of(pid)):
                break
        time.sleep(0.5)
    if pid is None:
        check(f"{exe} 실행", False, "프로세스를 찾지 못함")
        return
    time.sleep(4)  # 문서를 다 불러올 때까지
    try:
        dialogs = [(t, [x for _, c, x in child_texts(h) if c == "Static" and x])
                   for h, v, c, t in windows_of(pid) if v and c == "#32770"]
        if dialogs:
            print("    떠 있는 대화상자:", dialogs, flush=True)
        titles = [t for _, v, c, t in windows_of(pid) if v and t]
        print("    창 제목:", titles[:4], flush=True)
        in_use = sync.is_in_use(path)
        print(f"    파일 잠금(다른 프로그램이 사용 중): {in_use}", flush=True)
        print("    폴더 내용:", sorted(p.name for p in path.parent.iterdir()), flush=True)
        try:
            with sync.open_shared_read(path) as f:
                data = f.read()
            readable = hashlib.md5(data).hexdigest() == original
            detail = "읽기 성공" if readable else "읽었지만 내용이 다름"
        except OSError as e:
            readable, detail = False, f"읽기 실패: {e}"
        check(f"{exe}가 연 상태에서 내용 읽기 (업로드 가능)", readable, detail)
        by_title = sync.title_mentions(sync.visible_window_titles(), path.name)
        check(f"{exe}가 연 상태를 '사용 중'으로 감지 (정리 버튼이 지우지 않음)",
              in_use or by_title or sync.SyncManager._lock_file_in_use(path.parent),
              f"잠금={in_use}, 창 제목={by_title}")
    finally:
        if not close_process_windows(pid, 15):
            print(f"    {exe}가 닫기에 응답하지 않아 종료함 (이 테스트가 연 빈 문서뿐)", flush=True)
            kill(pid)
        time.sleep(1)
    check(f"{exe} 닫은 뒤 '사용 중' 아님", not sync.is_in_use(path))


def main() -> int:
    tmp = Path(tempfile.mkdtemp(prefix="gde_hnc_"))
    try:
        hwp = tmp / "work" / "a1" / "한글 테스트.hwp"
        hwp.parent.mkdir(parents=True)
        shutil.copyfile(HNC, hwp)
        run_case("Hwp.exe", hwp)
        docx = tmp / "work" / "b2" / "한워드 테스트.docx"
        make_docx_with_word(docx, "한워드 테스트")
        run_case("Hword.exe", docx)
    finally:
        time.sleep(0.5)
        shutil.rmtree(tmp, ignore_errors=True)
    failed = [r for r in RESULTS if not r[1]]
    print(f"\n결과: {len(RESULTS) - len(failed)}/{len(RESULTS)} 통과", flush=True)
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
