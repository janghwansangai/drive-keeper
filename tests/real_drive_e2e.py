"""실제 구글 드라이브 왕복 테스트 (사용자 허락을 받고 실행할 것 - 드라이브에 테스트 파일을 만든다).

- 저장된 로그인(token.json)을 쓴다. 브라우저 로그인은 하지 않는다.
- 드라이브 최상위에 _GDE_TEST_삭제예정.* 파일을 만들고, 탐색기(App)로 열고, Word/Excel로 수정·저장해
  드라이브에 올라간 내용이 로컬 파일과 바이트 단위로 같은지 확인한다.
- 끝나면 만든 파일(충돌 사본 포함)을 모두 휴지통으로 옮긴다.
실행: .venv\\Scripts\\python tests\\real_drive_e2e.py
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

import pythoncom  # noqa: E402
import win32com.client  # noqa: E402

import config  # noqa: E402

PREFIX = "_GDE_TEST_삭제예정"
HWP_A = Path(r"C:\Program Files (x86)\Hnc\Office 2020\HOffice110\Bin\Hwp\DAISY\DaisyStyle2.hwp")
HWP_B = Path(r"C:\Program Files (x86)\Hnc\Office 2020\HOffice110\Bin\HncPrivacy\Summary_Report_Template.hwp")
DOCX = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
RESULTS: list[tuple[str, bool, str]] = []


def check(name: str, ok: bool, detail: str = "") -> bool:
    RESULTS.append((name, ok, detail))
    print(("  [통과] " if ok else "  [실패] ") + name + (f" - {detail}" if detail else ""), flush=True)
    return ok


def md5_file(p: Path) -> str:
    import sync
    with sync.open_shared_read(p) as f:
        return hashlib.md5(f.read()).hexdigest()


def main() -> int:
    pythoncom.CoInitialize()
    tmp = Path(tempfile.mkdtemp(prefix="gde_e2e_"))
    created: list[str] = []
    drive = None
    app = None
    try:
        import drive_admin
        from drive_client import BlockedRequest, DriveClient
        from win_helpers import DialogDismisser, com_retry, kill, make_docx_with_word, pids_of

        print("=== 연결 ===", flush=True)
        token = tmp / "e2e_token.json"          # 실제 토큰을 건드리지 않게 (암호화된) 사본으로 테스트
        drive_admin.copy_token_to(token)
        t = time.perf_counter()
        drive = DriveClient(token_path=token, interactive=False)
        email = drive.about_email()
        check("저장된 로그인으로 연결 (브라우저 없이)", bool(email), f"{email}, {time.perf_counter() - t:.2f}초")

        print("\n=== 테스트 파일 만들기 ===", flush=True)
        src = tmp / "src"
        make_docx_with_word(src / "a.docx", "드라이브 왕복 테스트 문서")
        xl = win32com.client.DispatchEx("Excel.Application")
        try:
            xl.Visible = False
            xl.DisplayAlerts = False
            wb = xl.Workbooks.Add()
            wb.Worksheets(1).Cells(1, 1).Value = "드라이브 왕복 테스트"
            wb.SaveAs(str(src / "a.xlsx"), FileFormat=51)
            wb.Close(False)
        finally:
            xl.Quit()
        docx = drive.create_file(f"{PREFIX}.docx", "root", (src / "a.docx").read_bytes(), DOCX)
        created.append(docx["id"])
        xlsx = drive.create_file(f"{PREFIX}.xlsx", "root", (src / "a.xlsx").read_bytes(), XLSX)
        created.append(xlsx["id"])
        hwp = drive.create_file(f"{PREFIX}.hwp", "root", HWP_A.read_bytes(), "application/x-hwp")
        created.append(hwp["id"])
        from googleapiclient.http import MediaIoBaseUpload
        import io
        gdoc = drive._run(lambda s: s.files().create(
            body={"name": f"{PREFIX} (Google 문서)", "parents": ["root"],
                  "mimeType": "application/vnd.google-apps.document"},
            media_body=MediaIoBaseUpload(io.BytesIO((src / "a.docx").read_bytes()), mimetype=DOCX),
            fields="id,name,mimeType,modifiedTime").execute())
        created.append(gdoc["id"])
        check("테스트 파일 4개 생성", len(created) == 4,
              ", ".join(m["name"] for m in (docx, xlsx, hwp, gdoc)))
        check("업로드 검사값(MD5) 일치", docx.get("md5Checksum") == hashlib.md5((src / "a.docx").read_bytes()).hexdigest())

        print("\n=== 목록·검색 속도 ===", flush=True)
        t = time.perf_counter()
        items = drive.list_children("root")
        dt = time.perf_counter() - t
        ids = {m["id"] for m in items}
        check("내 드라이브 목록에 테스트 파일이 보임", all(i in ids for i in created), f"{len(items)}개 항목, {dt:.2f}초")
        t = time.perf_counter()
        items = drive.list_children("root")
        check("두 번째 목록 (연결 재사용)", True, f"{time.perf_counter() - t:.2f}초")
        t = time.perf_counter()
        found = drive.search(PREFIX)
        check("검색으로 찾기", {m["id"] for m in found} >= set(created),
              f"{len(found)}개, {time.perf_counter() - t:.2f}초")

        print("\n=== 보안: 이 프로그램은 삭제·휴지통·공유 변경 요청을 보낼 수 없음 ===", flush=True)
        attempts = {
            "영구 삭제": lambda s: s.files().delete(fileId=docx["id"]).execute(),
            "휴지통으로 이동": lambda s: s.files().update(fileId=docx["id"], body={"trashed": True}).execute(),
            "휴지통 비우기": lambda s: s.files().emptyTrash().execute(),
            "전체 공개로 공유": lambda s: s.permissions().create(
                fileId=docx["id"], body={"type": "anyone", "role": "reader"}).execute(),
        }
        for label, fn in attempts.items():
            try:
                drive._run(fn)
                blocked = False
            except BlockedRequest:
                blocked = True
            check(f"'{label}' 요청 차단", blocked)
        still = drive.get_meta(docx["id"])
        check("차단 후에도 파일은 드라이브에 그대로", not still.get("trashed"))

        print("\n=== 탐색기(App)로 열기 → Word/Excel로 수정·저장 → 드라이브 반영 ===", flush=True)
        import app as app_mod
        base = tmp / "base"
        config.BASE_DIR = base                    # 실행 중인 프로그램의 작업 폴더와 섞이지 않게
        config.WORK_DIR = base / "work"           # (로그인 토큰은 진짜 것을 그대로 씀)
        config.STATE_PATH = base / "state.json"
        config.STAGING_DIR = base / "staging"
        config.LOG_PATH = base / "app.log"
        opened: list[Path] = []
        dialogs: list[str] = []
        app_mod.open_with_default = lambda p: opened.append(Path(p))
        app_mod.messagebox.showinfo = lambda t, m, **k: dialogs.append(m)
        app_mod.messagebox.showwarning = lambda t, m, **k: dialogs.append(m)
        app_mod.messagebox.showerror = lambda t, m, **k: dialogs.append("오류: " + m)
        import accounts
        import protect
        st = accounts.AccountStore()
        st.load()
        acct, _ = st.add(email)
        protect.write_token(acct.token_path, drive.creds.to_json())
        app = app_mod.App(drive_factory=lambda a, cb: drive)

        def pump(cond, timeout=30.0):
            end = time.monotonic() + timeout
            while time.monotonic() < end:
                app.update()
                if cond():
                    return True
                time.sleep(0.02)
            return bool(cond())

        t = time.perf_counter()
        ok = pump(lambda: docx["id"] in app.item_meta)
        check("탐색기에 내 드라이브 목록 표시", ok, f"{time.perf_counter() - t:.2f}초")

        def status_of(path: Path) -> str:
            import sync
            rec = app.sync.get(sync.norm(path))
            return rec["status"] if rec else ""

        def open_in_app(file_id: str) -> Path:
            n = len(opened)
            t0 = time.perf_counter()
            app.open_item(app.item_meta[file_id])
            pump(lambda: len(opened) > n, 60)
            check(f"열기(내려받기): {app.item_meta[file_id]['name']}", len(opened) > n,
                  f"{time.perf_counter() - t0:.2f}초")
            return opened[-1]

        # --- Word ---
        p_docx = open_in_app(docx["id"])
        check("내려받은 파일 = 드라이브 원본 (MD5)", md5_file(p_docx) == docx["md5Checksum"])
        before = pids_of("WINWORD.EXE")
        dis = DialogDismisser(lambda: pids_of("WINWORD.EXE") - before, [("기본 프로그램", "아니요")])
        word = win32com.client.DispatchEx("Word.Application")
        try:
            com_retry(lambda: setattr(word, "Visible", False))
            word = com_retry(lambda: win32com.client.gencache.EnsureDispatch(word))
            com_retry(lambda: setattr(word, "DisplayAlerts", 0))
            doc = com_retry(lambda: word.Documents.Open(str(p_docx)))
            com_retry(lambda: doc.Range().InsertAfter(" / 학교 PC에서 수정"))
            t0 = time.perf_counter()
            com_retry(lambda: doc.Save())
            ok = pump(lambda: "드라이브에 저장됨" in status_of(p_docx), 60)
            up = time.perf_counter() - t0
            remote = drive.get_meta(docx["id"])
            check("Word 저장 → 드라이브 반영", ok and remote["md5Checksum"] == md5_file(p_docx),
                  f"저장 후 {up:.2f}초 만에 반영")

            # 다른 컴퓨터에서 고친 것처럼 드라이브 원본을 바꾼 뒤 여기서 또 저장 → 충돌 사본
            other = make_other = src / "other.docx"
            shutil.copyfile(src / "a.docx", make_other)
            with open(make_other, "ab") as f:
                f.write(b"\0")  # 내용만 다르게 (zip 뒤에 붙은 바이트는 Word가 무시)
            drive.update_content(docx["id"], other.read_bytes(), DOCX)
            other_md5 = hashlib.md5(other.read_bytes()).hexdigest()
            com_retry(lambda: doc.Range().InsertAfter(" / 충돌 상황 저장"))
            com_retry(lambda: doc.Save())
            ok = pump(lambda: "충돌 사본" in status_of(p_docx), 60)
            import sync
            rec = app.sync.get(sync.norm(p_docx))
            if rec and rec["file_id"] != docx["id"]:
                created.append(rec["file_id"])
            orig_now = drive.get_meta(docx["id"])
            copy_meta = drive.get_meta(rec["file_id"]) if rec else {}
            check("다른 곳에서 바뀐 원본은 덮어쓰지 않음", orig_now["md5Checksum"] == other_md5)
            check("내 변경은 '충돌 사본'으로 저장", ok and copy_meta.get("md5Checksum") == md5_file(p_docx),
                  copy_meta.get("name", ""))
            check("충돌 안내 창 표시", pump(lambda: any("충돌" in d or "사본" in d for d in dialogs), 10))
            com_retry(lambda: doc.Close(False))
        finally:
            try:
                word.Quit(0)
            except Exception:
                pass
            time.sleep(2)
            for pid in pids_of("WINWORD.EXE") - before:
                kill(pid)
            dis.stop()

        # --- Excel ---
        p_xlsx = open_in_app(xlsx["id"])
        xl = win32com.client.DispatchEx("Excel.Application")
        try:
            xl.Visible = False
            xl.DisplayAlerts = False
            wb = xl.Workbooks.Open(str(p_xlsx))
            wb.Worksheets(1).Cells(2, 1).Value = "학교 PC에서 수정"
            t0 = time.perf_counter()
            wb.Save()
            ok = pump(lambda: "드라이브에 저장됨" in status_of(p_xlsx), 60)
            remote = drive.get_meta(xlsx["id"])
            check("Excel 저장 → 드라이브 반영", ok and remote["md5Checksum"] == md5_file(p_xlsx),
                  f"저장 후 {time.perf_counter() - t0:.2f}초 만에 반영")
            wb.Close(False)
        finally:
            xl.Quit()

        # --- 한글 파일 (내용 바꿔 쓰기로 저장 흉내) ---
        p_hwp = open_in_app(hwp["id"])
        p_hwp.write_bytes(HWP_B.read_bytes())
        ok = pump(lambda: "드라이브에 저장됨" in status_of(p_hwp), 60)
        remote = drive.get_meta(hwp["id"])
        check(".hwp 변경 → 드라이브 반영", ok and remote["md5Checksum"] == md5_file(p_hwp))
        dl = tmp / "verify.hwp"
        drive.download(remote, dl)
        check("드라이브에서 다시 받은 .hwp = 로컬 파일 (바이트 단위)", dl.read_bytes() == p_hwp.read_bytes())

        # --- Google 문서 ---
        before_mod = drive.get_meta(gdoc["id"])["modifiedTime"]
        p_g = open_in_app(gdoc["id"])
        check("Google 문서는 .docx로 내보내 열림", p_g.suffix == ".docx" and p_g.stat().st_size > 0,
              f"{p_g.name}, {p_g.stat().st_size} bytes")
        check("'복사본' 안내 창 표시", any("복사본" in d for d in dialogs))
        with open(p_g, "ab") as f:
            f.write(b"\0")
        pump(lambda: False, 4)
        check("복사본을 고쳐도 Google 문서 원본은 그대로",
              drive.get_meta(gdoc["id"])["modifiedTime"] == before_mod)

        # --- 다시 열기는 즉시 (다시 받지 않음) ---
        t0 = time.perf_counter()
        n = len(opened)
        app.open_item(app.item_meta[xlsx["id"]])
        pump(lambda: len(opened) > n, 30)
        check("이미 받은 파일 다시 열기", len(opened) > n, f"{time.perf_counter() - t0:.2f}초")

        errors = [d for d in dialogs if d.startswith("오류")]
        check("오류 창 없음", not errors, "; ".join(errors))
    except Exception as e:  # noqa: BLE001
        import traceback
        traceback.print_exc()
        check("테스트 실행", False, repr(e))
    finally:
        if app is not None:
            try:
                app.shutdown()
            except Exception:
                pass
        if drive is not None and created:
            print("\n=== 정리: 테스트 파일을 휴지통으로 ===", flush=True)
            trashed = 0
            for fid in created:
                try:
                    drive_admin.trash(drive, fid)   # 테스트 전용 정리 (프로그램 본체는 못 함)
                    trashed += 1
                except Exception as e:  # noqa: BLE001
                    print("    휴지통 이동 실패:", fid, e, flush=True)
            left = [m["name"] for m in drive.search(PREFIX)]
            check(f"테스트 파일 {trashed}/{len(created)}개 휴지통으로 이동", trashed == len(created) and not left,
                  f"남은 파일: {left}" if left else "")
        time.sleep(0.5)
        shutil.rmtree(tmp, ignore_errors=True)
    failed = [r for r in RESULTS if not r[1]]
    print(f"\n결과: {len(RESULTS) - len(failed)}/{len(RESULTS)} 통과", flush=True)
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
