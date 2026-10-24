"""실제 Word/Excel로 저장 감지 테스트 (가짜 드라이브 사용, 구글에 접속하지 않음).

Windows + Microsoft Office + pywin32 필요. Office 창은 화면에 띄우지 않는다.
실행: .venv\\Scripts\\python tests\\real_office_test.py
"""
from __future__ import annotations

import shutil
import sys
import tempfile
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(Path(__file__).resolve().parent))

import pythoncom  # noqa: E402
import win32com.client  # noqa: E402

import sync  # noqa: E402
from fake_drive import FakeDrive, wait_until  # noqa: E402
from win_helpers import DialogDismisser, com_retry, kill, pids_of  # noqa: E402

RESULTS: list[tuple[str, bool, str]] = []


def check(name: str, ok: bool, detail: str = "") -> None:
    RESULTS.append((name, ok, detail))
    print(("  [통과] " if ok else "  [실패] ") + name + (f" - {detail}" if detail else ""), flush=True)


class EventLog:
    """작업 폴더에서 일어나는 파일 이벤트를 기록 (각 프로그램의 저장 방식 확인용)."""

    def __init__(self, folder: Path) -> None:
        from watchdog.events import FileSystemEventHandler
        from watchdog.observers import Observer
        self.events: list[str] = []
        log = self.events

        class H(FileSystemEventHandler):
            def on_any_event(self, e):
                if e.is_directory:
                    return
                dest = getattr(e, "dest_path", "")
                log.append(f"{e.event_type}: {Path(e.src_path).name}" + (f" -> {Path(dest).name}" if dest else ""))
        self.obs = Observer()
        self.obs.schedule(H(), str(folder), recursive=True)
        self.obs.start()

    def stop(self) -> None:
        self.obs.stop()


def setup(tmp: Path, debounce=0.4, poll=0.5):
    drive = FakeDrive()
    work = tmp / "work"
    mgr = sync.SyncManager(work, tmp / "state.json", tmp / "staging", get_drive=lambda: drive,
                           debounce=debounce, poll=poll)
    mgr.start()
    return drive, mgr, work


def open_from_drive(drive: FakeDrive, mgr: sync.SyncManager, work: Path, name: str, data: bytes, mime: str):
    meta = drive.add_file(name, data, mime=mime)
    fresh = drive.get_meta(meta["id"])
    path = work / sync.short_dir_name(meta["id"]) / sync.safe_filename(name, 120)
    digest = drive.download(fresh, path)
    mgr.add(path, fresh, digest, sync.KIND_EDIT)
    return meta["id"], path


def same_as_disk(drive: FakeDrive, file_id: str, path: Path) -> bool:
    with sync.open_shared_read(path) as f:
        local = f.read()
    return drive.content.get(file_id) == local


def test_word(tmp: Path) -> None:
    print("\n=== Word ===", flush=True)
    drive, mgr, work = setup(tmp)
    before = pids_of("WINWORD.EXE")
    # 이 PC는 .docx 기본 프로그램이 한컴이라 Word가 시작할 때 '기본 프로그램으로 설정할까요?'를 묻는다.
    # 설정을 바꾸지 않는 '아니요'만 눌러 준다 (이 테스트가 띄운 Word에만).
    dismisser = DialogDismisser(lambda: pids_of("WINWORD.EXE") - before,
                                [("기본 프로그램", "아니요"), ("대화 상자가 열려", "확인")])
    word = win32com.client.DispatchEx("Word.Application")
    ev = None
    try:
        com_retry(lambda: setattr(word, "Visible", False))  # 시작 대화상자가 닫힐 때까지 기다림
        # 형식 정보를 붙여 둬야 Documents.Add() 결과의 속성 접근이 안정적이다
        word = com_retry(lambda: win32com.client.gencache.EnsureDispatch(word))
        com_retry(lambda: setattr(word, "DisplayAlerts", 0))
        src = tmp / "src" / "원본.docx"
        src.parent.mkdir(parents=True, exist_ok=True)
        d = com_retry(lambda: word.Documents.Add())
        com_retry(lambda: d.Range().InsertAfter("처음 내용"))
        com_retry(lambda: d.SaveAs2(str(src), 16))
        com_retry(lambda: d.Close(False))
        fid, path = open_from_drive(drive, mgr, work, "가정통신문 (3월).docx", src.read_bytes(),
                                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        ev = EventLog(work)

        doc = com_retry(lambda: word.Documents.Open(str(path)))
        time.sleep(2.5)
        check("열기만 하면 업로드 안 함", drive.uploads == [], f"업로드 {len(drive.uploads)}회")
        check("Word가 연 파일은 '사용 중'으로 감지", sync.is_in_use(path))

        doc.Range().InsertAfter(" / 수정 1")
        doc.Save()
        ok = wait_until(lambda: len(drive.uploads) >= 1, 15)
        time.sleep(1.5)
        check("Ctrl+S 저장 → 업로드", ok and same_as_disk(drive, fid, path),
              f"업로드 {len(drive.uploads)}회")
        print("    Word 저장 시 파일 이벤트:", ev.events[:12], flush=True)

        n = len(drive.uploads)
        doc.Range().InsertAfter(" / 수정 2")
        doc.Save()
        ok = wait_until(lambda: len(drive.uploads) > n, 15)
        time.sleep(1.5)
        check("두 번째 저장 → 다시 업로드", ok and same_as_disk(drive, fid, path))

        # 빠른 연속 저장: 우리가 읽는 동안에도 Word 저장이 실패하면 안 됨
        errors = 0
        for i in range(12):
            doc.Range().InsertAfter(f" / 연속 {i}")
            try:
                doc.Save()
            except Exception as e:  # noqa: BLE001
                errors += 1
                print("    Word 저장 오류:", e, flush=True)
            time.sleep(0.25)
        ok = wait_until(lambda: same_as_disk(drive, fid, path), 20)
        check("연속 저장 12회 중 Word 저장 오류 없음", errors == 0, f"오류 {errors}회")
        check("연속 저장 후 마지막 내용이 드라이브와 같음", ok, f"업로드 {len(drive.uploads)}회")

        removed, kept = mgr.cleanup()
        check("열려 있는 동안 '작업 폴더 정리'는 지우지 않음", removed == 0 and path.exists(), f"kept={kept}")

        # 다른 이름으로 저장 → 같은 드라이브 폴더에 새 파일로 업로드
        new = path.parent / "가정통신문 수정본.docx"
        com_retry(lambda: doc.SaveAs2(str(new), 16))
        ok = wait_until(lambda: any(c[0] == new.name for c in drive.creates), 15)
        time.sleep(1.0)
        rec = mgr.get(sync.norm(new))
        check("Word '다른 이름으로 저장' → 새 파일로 업로드",
              ok and rec and rec["file_id"] and same_as_disk(drive, rec["file_id"], new),
              f"만든 파일 {[c[0] for c in drive.creates]}")
        check("임시 파일(~$, ~WR*.tmp)은 올리지 않음", len(drive.creates) == 1)
        com_retry(lambda: doc.Range().InsertAfter(" / 수정본에서 저장"))
        com_retry(lambda: doc.Save())
        ok = wait_until(lambda: rec and same_as_disk(drive, rec["file_id"], new), 15)
        check("새 파일에서 다시 저장 → 그 파일에 반영 (또 만들지 않음)", ok and len(drive.creates) == 1)

        doc.Close(False)
        time.sleep(1.5)
        check("닫은 뒤 '사용 중' 아님", not sync.is_in_use(path) and not sync.is_in_use(new))
        removed, kept = mgr.cleanup()
        check("닫은 뒤 정리하면 삭제됨", removed == 2 and not path.exists() and not new.exists(),
              f"removed={removed} kept={kept}")
        if dismisser.clicked:
            print("    자동으로 누른 Word 대화상자:", dismisser.clicked, flush=True)
    finally:
        if ev:
            ev.stop()
        try:
            word.Quit(0)
        except Exception:
            pass
        time.sleep(2)
        for pid in pids_of("WINWORD.EXE") - before:  # 이 테스트가 띄운 Word가 남아 있으면 정리
            kill(pid)
        dismisser.stop()
        mgr.stop()


def test_excel(tmp: Path) -> None:
    print("\n=== Excel ===", flush=True)
    drive, mgr, work = setup(tmp)
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    ev = None
    try:
        src = tmp / "src" / "원본.xlsx"
        src.parent.mkdir(parents=True, exist_ok=True)
        wb = excel.Workbooks.Add()
        wb.Worksheets(1).Cells(1, 1).Value = "처음"
        wb.SaveAs(str(src), FileFormat=51)
        wb.Close(False)
        fid, path = open_from_drive(drive, mgr, work, "2반 명렬표.xlsx", src.read_bytes(),
                                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        ev = EventLog(work)

        wb = excel.Workbooks.Open(str(path))
        time.sleep(2.5)
        check("열기만 하면 업로드 안 함", drive.uploads == [], f"업로드 {len(drive.uploads)}회")
        check("Excel이 연 파일은 '사용 중'으로 감지", sync.is_in_use(path))

        wb.Worksheets(1).Cells(2, 1).Value = "수정 1"
        wb.Save()
        ok = wait_until(lambda: len(drive.uploads) >= 1, 15)
        time.sleep(1.5)
        check("저장 → 업로드", ok and same_as_disk(drive, fid, path), f"업로드 {len(drive.uploads)}회")
        print("    Excel 저장 시 파일 이벤트:", ev.events[:12], flush=True)

        errors = 0
        for i in range(10):
            wb.Worksheets(1).Cells(3 + i, 1).Value = f"연속 {i}"
            try:
                wb.Save()
            except Exception as e:  # noqa: BLE001
                errors += 1
                print("    Excel 저장 오류:", e, flush=True)
            time.sleep(0.25)
        ok = wait_until(lambda: same_as_disk(drive, fid, path), 20)
        check("연속 저장 10회 중 Excel 저장 오류 없음", errors == 0, f"오류 {errors}회")
        check("연속 저장 후 마지막 내용이 드라이브와 같음", ok, f"업로드 {len(drive.uploads)}회")

        new = path.parent / "2반 명렬표 사본.xlsx"
        wb.SaveAs(str(new), FileFormat=51)
        ok = wait_until(lambda: any(c[0] == new.name for c in drive.creates), 15)
        time.sleep(1.0)
        rec = mgr.get(sync.norm(new))
        check("Excel '다른 이름으로 저장' → 새 파일로 업로드",
              ok and rec and rec["file_id"] and same_as_disk(drive, rec["file_id"], new) and len(drive.creates) == 1,
              f"만든 파일 {[c[0] for c in drive.creates]}")

        wb.Close(False)
        time.sleep(1.5)
        removed, kept = mgr.cleanup()
        check("닫은 뒤 정리하면 삭제됨", removed == 2 and not path.exists(), f"removed={removed} kept={kept}")
    finally:
        if ev:
            ev.stop()
        try:
            excel.Quit()
        except Exception:
            pass
        mgr.stop()


def main() -> int:
    pythoncom.CoInitialize()
    tmp = Path(tempfile.mkdtemp(prefix="gde_office_"))
    try:
        for t in (test_word, test_excel):
            sub = tmp / t.__name__
            sub.mkdir()
            try:
                t(sub)
            except Exception as e:  # noqa: BLE001
                import traceback
                traceback.print_exc()
                check(f"{t.__name__} 실행", False, repr(e))
    finally:
        time.sleep(0.5)
        shutil.rmtree(tmp, ignore_errors=True)
    failed = [r for r in RESULTS if not r[1]]
    print(f"\n결과: {len(RESULTS) - len(failed)}/{len(RESULTS)} 통과", flush=True)
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
