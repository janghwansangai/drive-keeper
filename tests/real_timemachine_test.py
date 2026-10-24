"""실제 구글 드라이브로 타임머신 확인 (사용자 허락 후 실행 - 드라이브에 테스트 파일을 만들었다가 휴지통으로 옮김).

프로그램 본체(DriveClient, 통신 차단 장치 포함)로: 휴지통 목록 → 되살리기 → 버전 목록 → 예전 버전 받기 → 되돌리기.
실행: .venv\\Scripts\\python tests\\real_timemachine_test.py
"""
from __future__ import annotations

import hashlib
import shutil
import sys
import tempfile
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(Path(__file__).resolve().parent))

import drive_admin  # noqa: E402
from drive_client import DriveClient  # noqa: E402

NAME = "_GDE_TEST_삭제예정 (타임머신).txt"
RESULTS: list[tuple[str, bool, str]] = []


def check(name: str, ok: bool, detail: str = "") -> bool:
    RESULTS.append((name, ok, detail))
    print(("  [통과] " if ok else "  [실패] ") + name + (f" - {detail}" if detail else ""), flush=True)
    return ok


def md5(b: bytes) -> str:
    return hashlib.md5(b).hexdigest()


def main() -> int:
    tmp = Path(tempfile.mkdtemp(prefix="gde_tm_"))
    drive = None
    fid = None
    try:
        token = tmp / "token.json"
        drive_admin.copy_token_to(token)
        drive = DriveClient(token_path=token, interactive=False)
        v1, v2, v3 = "첫 번째 내용".encode(), "두 번째 내용".encode(), "세 번째 내용".encode()
        meta = drive.create_file(NAME, "root", v1, "text/plain")
        fid = meta["id"]
        time.sleep(1.5)
        drive.update_content(fid, v2, "text/plain")
        time.sleep(1.5)
        drive.update_content(fid, v3, "text/plain")
        check("테스트 파일 만들고 두 번 고침", True, NAME)

        print("\n=== 이전 버전 ===", flush=True)
        t = time.perf_counter()
        revs = drive.list_revisions(fid)
        check("버전 목록 읽기", len(revs) >= 2, f"{len(revs)}개, {time.perf_counter() - t:.2f}초")
        oldest = revs[0]
        got = drive.download_revision(fid, oldest, tmp / "old.txt")
        check("가장 오래된 버전 받기 (검사값 확인)", (tmp / "old.txt").read_bytes() == v1 and got == md5(v1))
        drive.update_content(fid, (tmp / "old.txt").read_bytes(), "text/plain")   # 앱의 '되돌리기'와 같은 동작
        now = drive.get_meta(fid)
        check("예전 버전으로 되돌리기", now["md5Checksum"] == md5(v1))
        revs2 = drive.list_revisions(fid)
        check("되돌린 뒤에도 그전 내용이 기록에 남음", any(r.get("md5Checksum") == md5(v3) for r in revs2),
              f"버전 {len(revs2)}개")

        print("\n=== 지운 파일 되살리기 ===", flush=True)
        drive_admin.trash(drive, fid)            # 사용자가 실수로 지운 상황 (테스트 도우미로)
        time.sleep(1.0)
        t = time.perf_counter()
        first, nxt = drive.list_trash_page()
        check("휴지통 첫 쪽 읽기 (최근 것부터)", fid in {m["id"] for m in first},
              f"{len(first)}개, {time.perf_counter() - t:.2f}초, 다음 쪽 {'있음' if nxt else '없음'}")
        t = time.perf_counter()
        trash, _ = drive.list_trash_page(name="_GDE_TEST_삭제예정 (타임머신)")
        mine = [m for m in trash if m["id"] == fid]
        check("휴지통에서 이름으로 찾기", bool(mine), f"{time.perf_counter() - t:.2f}초")
        if mine:
            check("휴지통 항목 정보", bool(mine[0].get("name")),
                  f"trashedTime={'있음' if mine[0].get('trashedTime') else '없음(내 드라이브는 구글이 안 줌)'}")
        drive.untrash(fid)                        # 통신 차단 장치를 통과해야 함
        back = drive.get_meta(fid)
        check("되살리기 (통신 차단 장치 통과)", not back.get("trashed"))
        check("원래 폴더로 돌아감", fid in {m["id"] for m in drive.list_children("root")})
    except Exception as e:  # noqa: BLE001
        import traceback
        traceback.print_exc()
        check("테스트 실행", False, repr(e))
    finally:
        if drive is not None and fid:
            try:
                drive_admin.trash(drive, fid)
                check("정리: 테스트 파일을 휴지통으로", True)
            except Exception as e:  # noqa: BLE001
                check("정리: 테스트 파일을 휴지통으로", False, repr(e))
        shutil.rmtree(tmp, ignore_errors=True)
    failed = [r for r in RESULTS if not r[1]]
    print(f"\n결과: {len(RESULTS) - len(failed)}/{len(RESULTS)} 통과", flush=True)
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
