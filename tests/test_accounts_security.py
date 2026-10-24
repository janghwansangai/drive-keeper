"""여러 계정, 비밀번호 잠금, 정전·오류 대비, 해킹·악성코드 대비 테스트.

실행: .venv\\Scripts\\python -m unittest tests.test_accounts_security -v
"""
from __future__ import annotations

import io
import json
import os
import shutil
import sys
import tempfile
import threading
import time
import unittest
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(Path(__file__).resolve().parent))

import accounts  # noqa: E402
import app as app_mod  # noqa: E402
import config  # noqa: E402
import protect  # noqa: E402
import security  # noqa: E402
import sync  # noqa: E402
from fake_drive import FakeDrive, wait_until, write_in_place  # noqa: E402

DOCX = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"


def make_zip(text: str) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("word/document.xml", text * 20)
        z.writestr("[Content_Types].xml", "<Types/>")
    return buf.getvalue()


class TempBase(unittest.TestCase):
    """config.BASE_DIR 등을 임시 폴더로 바꿔 진짜 앱 데이터를 건드리지 않는다."""

    KEYS = ("BASE_DIR", "WORK_DIR", "STATE_PATH", "STAGING_DIR", "TOKEN_PATH", "LOG_PATH")

    def setUp(self) -> None:
        self.tmp = Path(tempfile.mkdtemp(prefix="gde_as_"))
        self._saved = {k: getattr(config, k) for k in self.KEYS}
        config.BASE_DIR = self.tmp
        config.WORK_DIR = self.tmp / "work"
        config.STATE_PATH = self.tmp / "state.json"
        config.STAGING_DIR = self.tmp / "staging"
        config.TOKEN_PATH = self.tmp / "token.json"
        config.LOG_PATH = self.tmp / "app.log"
        self._marker = protect.LOCK_MARKER
        protect.LOCK_MARKER = False

    def tearDown(self) -> None:
        for k, v in self._saved.items():
            setattr(config, k, v)
        protect.LOCK_MARKER = self._marker
        time.sleep(0.1)
        shutil.rmtree(self.tmp, ignore_errors=True)


# ======================= 보호(암호화) =======================
class TestProtect(TempBase):
    def test_roundtrip_and_not_plaintext(self):
        stored = protect.protect_text('{"refresh_token": "SECRET-123"}')
        self.assertTrue(stored.startswith(protect.PREFIX))
        if os.name == "nt":
            self.assertNotIn("SECRET-123", stored)
        self.assertEqual(protect.unprotect_text(stored), '{"refresh_token": "SECRET-123"}')

    def test_tampered_text_raises(self):
        stored = protect.protect_text("hello")
        bad = stored[:-6] + ("AAAA==" if not stored.endswith("AAAA==") else "BBBB==")
        if os.name == "nt":
            with self.assertRaises(ValueError):
                protect.unprotect_text(bad)

    def test_token_file_and_legacy(self):
        p = self.tmp / "t.json"
        protect.LOCK_MARKER = True
        protect.write_token(p, '{"a": 1}')
        creds, lock = protect.read_token(p)
        self.assertEqual(json.loads(creds), {"a": 1})
        self.assertTrue(lock)
        if os.name == "nt":
            self.assertNotIn('"a"', p.read_text(encoding="utf-8"))
        p.write_text('{"refresh_token": "x"}', encoding="utf-8")   # 예전 평문 토큰
        creds, lock = protect.read_token(p)
        self.assertEqual(json.loads(creds)["refresh_token"], "x")
        self.assertFalse(lock)


# ======================= 보안 설정 =======================
class TestSecurity(TempBase):
    def test_password_check_and_lockout(self):
        s = security.Security()
        s.load()
        self.assertFalse(s.active)
        with self.assertRaises(ValueError):
            s.set_password("12")
        s.set_password("1234")
        self.assertTrue(s.active)
        self.assertNotIn("1234", (self.tmp / "security.json").read_text(encoding="utf-8"))
        self.assertEqual(s.check("1234"), (True, 0))
        for _ in range(4):
            self.assertEqual(s.check("wrong")[0], False)
        ok, wait = s.check("wrong")          # 5번째
        self.assertFalse(ok)
        self.assertGreaterEqual(wait, 29)
        self.assertEqual(s.check("1234"), (False, s.wait_remaining()))  # 기다리는 동안은 맞아도 안 됨
        s2 = security.Security()
        s2.load()                            # 껐다 켜도 기다림 유지
        self.assertGreater(s2.wait_remaining(), 0)
        s2.data["wait_until"] = 0
        self.assertTrue(s2.check("1234")[0])
        self.assertEqual(s2.fails(), 0)

    def test_settings_persist_and_disable_all(self):
        s = security.Security()
        s.load()
        s.set_password("abcd")
        s.update(True, False, 15, True)
        s2 = security.Security()
        s2.load()
        self.assertEqual((s2.data["lock_on_start"], s2.data["idle_minutes"], s2.data["lock_on_sleep"]),
                         (False, 15, True))
        s2.disable_all()
        s3 = security.Security()
        s3.load()
        self.assertFalse(s3.active)
        self.assertTrue(s3.has_password)

    def test_corrupt_or_deleted_file_requires_reset(self):
        s = security.Security()
        s.load()
        s.set_password("abcd")
        path = self.tmp / "security.json"
        path.write_text('{"enabled": false}', encoding="utf-8")  # 메모장으로 고쳐 잠금 풀기 시도
        s2 = security.Security()
        s2.load()
        if os.name == "nt":
            self.assertTrue(s2.reset_required)   # 암호화 안 된 설정 = 위조 → 본인 확인 필요
        path.unlink()
        s3 = security.Security()
        s3.load(lock_marker_seen=True)       # 토큰에 '잠금 켜짐' 표시가 있는데 설정 파일이 없음
        self.assertTrue(s3.reset_required)
        path.write_text(protect.PREFIX + "깨진내용", encoding="utf-8")
        s4 = security.Security()
        s4.load()
        self.assertTrue(s4.reset_required)

    def test_idle_seconds(self):
        idle = security.system_idle_seconds()
        if os.name == "nt":
            self.assertIsNotNone(idle)
            self.assertGreaterEqual(idle, 0)


# ======================= 계정 목록 =======================
class TestAccountStore(TempBase):
    def test_add_find_remove(self):
        st = accounts.AccountStore()
        st.load()
        a, created = st.add("Teacher@Gmail.com")
        self.assertTrue(created)
        b, created2 = st.add("teacher@gmail.com")   # 대소문자만 다른 같은 계정
        self.assertFalse(created2)
        self.assertEqual(a.id, b.id)
        st.add("second@gmail.com")
        st2 = accounts.AccountStore()
        st2.load()
        self.assertEqual([x.email for x in st2.accounts], ["Teacher@Gmail.com", "second@gmail.com"])
        st2.remove(a.id)
        self.assertEqual(len(accounts.AccountStore().load()), 1)

    def test_broken_index_rebuilt_from_account_folders(self):
        st = accounts.AccountStore()
        st.load()
        a, _ = st.add("a@gmail.com")
        (self.tmp / "accounts.json").write_text("{ broken", encoding="utf-8")
        st2 = accounts.AccountStore()
        got = st2.load()
        self.assertEqual([x.id for x in got], [a.id])

    def test_move_data_out_of_appdata(self):
        """예전 위치(AppData)의 계정·작업 폴더를 새 위치로 복사하고, 작업 폴더 경로를 고친다."""
        old = self.tmp / "old_appdata" / "GDriveExplorer"
        new = self.tmp / "home" / "GDriveExplorer"
        (old / "accounts" / "c0").mkdir(parents=True)
        (old / "work" / "abc").mkdir(parents=True)
        (old / "work" / "abc" / "문서.hwpx").write_bytes(b"work file")
        (old / "app.lock").write_bytes(b"")
        acct = {"id": "c0", "email": "a@gmail.com", "work_dir": str(old / "work")}
        (old / "accounts.json").write_text(json.dumps({"version": 1, "accounts": [acct]}), encoding="utf-8")
        (old / "accounts" / "c0" / "account.json").write_text(json.dumps(acct), encoding="utf-8")
        (old / "accounts" / "c0" / "token.json").write_text("TOKEN", encoding="utf-8")
        orig_legacy = config.legacy_base_dirs
        config.BASE_DIR = new
        config.legacy_base_dirs = lambda: [old]
        try:
            self.assertEqual(config.migrate_legacy_base(), old)
            self.assertEqual((new / "work" / "abc" / "문서.hwpx").read_bytes(), b"work file")
            self.assertEqual((new / "accounts" / "c0" / "token.json").read_text(encoding="utf-8"), "TOKEN")
            self.assertFalse((new / "app.lock").exists())
            got = accounts.AccountStore().load()
            self.assertEqual(got[0].work_dir, new / "work")          # 경로가 새 위치로 바뀜
            self.assertTrue((old / "accounts.json").exists())         # 예전 것은 백업으로 남음
            self.assertIsNone(config.migrate_legacy_base())           # 두 번 하지 않음
        finally:
            config.legacy_base_dirs = orig_legacy

    def test_migrate_legacy_single_account(self):
        (self.tmp / "work" / "abc").mkdir(parents=True)
        (self.tmp / "work" / "abc" / "문서.docx").write_bytes(b"x")
        (self.tmp / "token.json").write_text('{"refresh_token": "r"}', encoding="utf-8")
        (self.tmp / "state.json").write_text(json.dumps({"files": [{"rel": "abc\\문서.docx"}]}), encoding="utf-8")
        st = accounts.AccountStore()
        got = st.load()
        self.assertEqual(len(got), 1)
        acct = got[0]
        self.assertEqual(acct.work_dir, self.tmp / "work")       # 작업 폴더는 그대로
        self.assertTrue(acct.token_path.exists())
        self.assertEqual(json.loads(protect.read_token(acct.token_path)[0])["refresh_token"], "r")
        if os.name == "nt":
            self.assertTrue(acct.token_path.read_text(encoding="utf-8").startswith(protect.PREFIX))
        self.assertTrue(acct.state_path.exists())
        self.assertFalse((self.tmp / "token.json").exists())
        self.assertTrue((self.tmp / "accounts.json").exists())


# ======================= 통신 보안 (삭제 차단) =======================
class TestRequestGuard(unittest.TestCase):
    def setUp(self) -> None:
        import drive_client
        self.dc = drive_client

    def test_allowlist(self):
        ok = [("https://www.googleapis.com/drive/v3/files?q=x", "GET", None, {}),
              ("https://www.googleapis.com/drive/v3/files/abc?alt=media", "GET", None, {}),
              ("https://www.googleapis.com/upload/drive/v3/files/abc?uploadType=media", "PATCH", b"PK..",
               {"content-type": DOCX}),
              ("https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart", "POST",
               b'--BB\r\ncontent-type: application/json\r\n\r\n{"name": "a.docx"}\r\n--BB\r\n\r\nPK\r\n--BB--',
               {"content-type": 'multipart/related; boundary="BB"'})]
        for args in ok:
            self.dc.check_request(*args)
        blocked = [("https://www.googleapis.com/drive/v3/files/abc", "DELETE", None, {}),
                   ("https://www.googleapis.com/drive/v3/files/trash", "DELETE", None, {}),
                   ("https://www.googleapis.com/drive/v3/files/abc", "PATCH", b'{"trashed": true}',
                    {"content-type": "application/json"}),
                   ("https://www.googleapis.com/drive/v3/files/abc/permissions", "POST",
                    b'{"role": "reader", "type": "anyone"}', {"content-type": "application/json"}),
                   ("https://www.googleapis.com/drive/v3/files/abc/copy", "POST", b"{}", {}),
                   ("https://www.googleapis.com/upload/drive/v3/files/abc?uploadType=multipart", "PATCH",
                    b'--BB\r\ncontent-type: application/json\r\n\r\n{"trashed": true}\r\n--BB\r\n\r\nx\r\n--BB--',
                    {"Content-Type": 'multipart/related; boundary="BB"'})]
        for args in blocked:
            with self.assertRaises(self.dc.BlockedRequest, msg=str(args[:2])):
                self.dc.check_request(*args)

    def test_guard_stops_real_client_calls_before_network(self):
        from googleapiclient.discovery import build
        sent = []

        class Inner:
            def request(self, uri, method="GET", body=None, headers=None, **kw):
                sent.append((method, uri))
                raise AssertionError("네트워크로 나가면 안 됨")
        svc = build("drive", "v3", http=self.dc._GuardedHttp(Inner()), cache_discovery=False, static_discovery=True)
        calls = [lambda: svc.files().delete(fileId="x").execute(),
                 lambda: svc.files().emptyTrash().execute(),
                 lambda: svc.files().update(fileId="x", body={"trashed": True}).execute(),
                 lambda: svc.permissions().create(fileId="x", body={"type": "anyone", "role": "reader"}).execute()]
        for call in calls:
            with self.assertRaises(self.dc.BlockedRequest):
                call()
        self.assertEqual(sent, [])

    def test_real_upload_requests_pass_guard(self):
        """실제 구글 라이브러리가 만드는 업로드 요청(단순·여러 부분·이어 올리기)이 차단되지 않는지."""
        import httplib2
        from googleapiclient.discovery import build
        from googleapiclient.http import MediaIoBaseUpload
        sent = []

        class Inner:
            def request(self, uri, method="GET", body=None, headers=None, **kw):
                sent.append((method, uri))
                if "uploadType=resumable" in uri and "upload_id" not in uri:
                    return httplib2.Response({"status": 200, "location": uri + "&upload_id=abc"}), b""
                return httplib2.Response({"status": 200}), b'{"id": "f1", "md5Checksum": "x"}'
        svc = build("drive", "v3", http=self.dc._GuardedHttp(Inner()), cache_discovery=False, static_discovery=True)
        small = MediaIoBaseUpload(io.BytesIO(b"PK" + b"x" * 100), mimetype=DOCX, resumable=False)
        svc.files().update(fileId="f1", media_body=small, supportsAllDrives=True).execute()
        big = MediaIoBaseUpload(io.BytesIO(b"y" * 300_000), mimetype=DOCX, resumable=True, chunksize=256 * 1024)
        svc.files().update(fileId="f1", media_body=big, supportsAllDrives=True).execute()
        new = MediaIoBaseUpload(io.BytesIO(b"z" * 50), mimetype="application/x-hwp", resumable=False)
        svc.files().create(body={"name": "새 파일.hwp", "parents": ["p1"]}, media_body=new).execute()
        svc.files().list(q="'root' in parents").execute()
        self.assertGreaterEqual(len(sent), 5)
        self.assertTrue(all(m in ("GET", "POST", "PATCH", "PUT") for m, _ in sent))

    def test_no_delete_methods_in_client(self):
        self.assertFalse(hasattr(self.dc.DriveClient, "trash"))
        self.assertFalse(hasattr(self.dc.DriveClient, "delete"))

    def test_relogin_with_other_account_rejected(self):
        dc = self.dc
        tmp = Path(tempfile.mkdtemp(prefix="gde_rl_"))
        orig_login, orig_email = dc.browser_login, dc.fetch_email

        class FakeCreds:
            def to_json(self):
                return '{"refresh_token": "r"}'
        try:
            dc.browser_login = lambda *a, **k: FakeCreds()
            dc.fetch_email = lambda creds: "intruder@gmail.com"
            client = dc.DriveClient.__new__(dc.DriveClient)
            client.expected_email = "teacher@gmail.com"
            client._on_browser_login = None
            with self.assertRaises(dc.AuthError):
                client._browser_login()
            dc.fetch_email = lambda creds: "Teacher@gmail.com"
            self.assertIsInstance(client._browser_login(), FakeCreds)
        finally:
            dc.browser_login, dc.fetch_email = orig_login, orig_email
            shutil.rmtree(tmp, ignore_errors=True)


# ======================= 정전·오류·랜섬웨어 대비 (업로드 엔진) =======================
class TestSyncHardening(TempBase):
    def setUp(self) -> None:
        super().setUp()
        self.work = self.tmp / "w"
        self.drive = FakeDrive()
        self.events: list = []
        self.mgr = self.make()

    def make(self) -> sync.SyncManager:
        m = sync.SyncManager(self.work, self.tmp / "state.json", self.tmp / "staging",
                             get_drive=lambda: self.drive, notify=lambda e, p: self.events.append((e, p)),
                             debounce=0.3, poll=0.25)
        m.start()
        return m

    def tearDown(self) -> None:
        self.mgr.stop()
        super().tearDown()

    def open(self, name: str, data: bytes) -> tuple[dict, Path]:
        meta = self.drive.add_file(name, data, mime=DOCX)
        fresh = self.drive.get_meta(meta["id"])
        path = self.work / sync.short_dir_name(meta["id"]) / name
        self.mgr.add(path, fresh, self.drive.download(fresh, path), sync.KIND_EDIT)
        return fresh, path

    def ups(self, fid: str) -> list:
        return [m for f, m in self.drive.uploads if f == fid]

    def test_valid_document_save_uploads(self):
        meta, path = self.open("정상.docx", make_zip("v1"))
        self.assertEqual(self.mgr.get(sync.norm(path))["fmt"], "zip")
        write_in_place(path, make_zip("v2"))
        self.assertTrue(wait_until(lambda: self.ups(meta["id"]), 5))

    def test_encrypted_document_is_held(self):
        meta, path = self.open("암호화됨.docx", make_zip("v1"))
        write_in_place(path, os.urandom(4000))               # 랜섬웨어가 암호화한 것처럼
        self.assertTrue(wait_until(lambda: self.mgr.get(sync.norm(path))["status"] == sync.STATUS_HELD, 5))
        time.sleep(1.0)
        self.assertEqual(self.ups(meta["id"]), [])            # 드라이브 원본은 그대로
        self.assertTrue(any(e == "suspicious" for e, _ in self.events))
        # 사용자가 확인 후 '그래도 업로드'
        self.mgr.force_upload(sync.norm(path))
        self.assertTrue(wait_until(lambda: self.ups(meta["id"]), 5))

    def test_truncated_zip_and_zero_bytes_held(self):
        meta, path = self.open("잘림.docx", make_zip("v1"))
        write_in_place(path, make_zip("v2")[:60])            # 앞부분만 남은 zip
        self.assertTrue(wait_until(lambda: self.mgr.get(sync.norm(path))["status"] == sync.STATUS_HELD, 5))
        meta2, path2 = self.open("빈파일.txt", b"some text")  # 형식 없는 파일도 0바이트는 보류
        write_in_place(path2, b"")
        self.assertTrue(wait_until(lambda: self.mgr.get(sync.norm(path2))["status"] == sync.STATUS_HELD, 5))
        self.assertEqual(self.drive.uploads, [])

    def test_many_damaged_files_pause_everything(self):
        files = [self.open(f"문서{i}.docx", make_zip(f"v{i}")) for i in range(4)]
        for meta, path in files[:3]:
            write_in_place(path, os.urandom(3000))
        self.assertTrue(wait_until(lambda: self.mgr.paused, 8))
        self.assertTrue(any(e == "paused" for e, _ in self.events))
        meta4, path4 = files[3]
        write_in_place(path4, make_zip("good change"))      # 정상 파일도 멈춘 동안은 올리지 않음
        time.sleep(1.5)
        self.assertEqual(self.drive.uploads, [])
        self.mgr.resume()
        self.assertTrue(wait_until(lambda: self.ups(meta4["id"]), 5))

    def test_state_restored_from_backup_then_sidecar(self):
        meta, path = self.open("복원.docx", b"base")
        self.mgr.stop()
        time.sleep(0.2)
        write_in_place(path, b"edited before power loss")
        (self.tmp / "state.json").write_text("{ 정전으로 깨짐", encoding="utf-8")
        self.mgr = self.make()                                 # 백업(state.bak.json)으로 복원
        self.assertTrue(wait_until(lambda: self.ups(meta["id"]), 6))
        self.mgr.stop()
        time.sleep(0.2)
        write_in_place(path, b"edited again")
        (self.tmp / "state.json").unlink()
        (self.tmp / "state.bak.json").unlink()
        self.mgr = self.make()                                 # 두 기록이 다 없어도 작업 폴더의 사본으로 복원
        self.assertTrue(wait_until(lambda: self.drive.content[meta["id"]] == b"edited again", 6))
        self.assertTrue((self.tmp / "state.json").exists())

    def test_dead_upload_thread_is_restarted(self):
        meta, path = self.open("스레드.docx", b"a")
        dead = threading.Thread(target=lambda: None)
        dead.start()
        dead.join()
        self.mgr._stop.set()          # 기존 루프를 멈추고
        time.sleep(0.6)
        self.mgr._stop.clear()
        self.mgr._thread = dead       # 죽은 스레드로 바꿔 놓음
        write_in_place(path, b"b")
        time.sleep(1.0)
        self.assertEqual(self.ups(meta["id"]), [])
        self.mgr.ensure_running()     # 앱이 5초마다 부르는 점검
        self.assertTrue(wait_until(lambda: self.ups(meta["id"]), 5))

    def test_cleanup_removes_sidecar_and_folder(self):
        meta, path = self.open("정리.docx", b"a")
        self.assertTrue((path.parent / sync.SIDECAR).exists())
        removed, kept = self.mgr.cleanup()
        self.assertEqual(removed, 1)
        self.assertFalse(path.parent.exists())


# ======================= 화면: 여러 계정·잠금 =======================
class AppBase(TempBase):
    def setUp(self) -> None:
        super().setUp()
        self.opened: list[Path] = []
        self.dialogs: list[tuple[str, str]] = []
        self.answer = True
        self._orig_open = app_mod.open_with_default
        app_mod.open_with_default = lambda p: self.opened.append(Path(p))
        self._orig_mb = {n: getattr(app_mod.messagebox, n) for n in
                         ("showinfo", "showwarning", "showerror", "askyesno")}
        app_mod.messagebox.showinfo = lambda t, m, **k: self.dialogs.append(("info", m))
        app_mod.messagebox.showwarning = lambda t, m, **k: self.dialogs.append(("warning", m))
        app_mod.messagebox.showerror = lambda t, m, **k: self.dialogs.append(("error", m))
        app_mod.messagebox.askyesno = lambda t, m, **k: (self.dialogs.append(("ask", m)), self.answer)[1]
        self.drives = {"a@gmail.com": FakeDrive("a@gmail.com"), "b@gmail.com": FakeDrive("b@gmail.com")}
        self.drives["a@gmail.com"].add_file("A의 문서.docx", b"from A")
        self.drives["b@gmail.com"].add_file("B의 문서.docx", b"from B")
        self.revoked: list[str] = []
        self.app = None

    def start(self, emails=("a@gmail.com",), add_flow=None) -> None:
        st = accounts.AccountStore()
        st.load()
        for e in emails:
            st.add(e)
        if self.app is not None:
            self.app.shutdown()
        self.app = app_mod.App(drive_factory=lambda acct, cb: self.drives[acct.email],
                               add_account_flow=add_flow or (lambda cb: ('{"refresh_token": "n"}', "b@gmail.com")),
                               revoke=lambda creds: self.revoked.append(creds))
        for sm in self.app.syncs.values():
            sm.debounce, sm.poll = 0.3, 0.25
        self.pump(lambda: all(self.app.drives.get(a.id) for a in self.app.store.accounts))

    def tearDown(self) -> None:
        try:
            if self.app is not None:
                self.app.shutdown()
        except Exception:
            pass
        self.app = None
        import gc
        gc.collect()   # 닫은 창의 객체를 화면 스레드에서 정리 (다른 스레드에서 정리되면 Tk 경고가 남)
        app_mod.open_with_default = self._orig_open
        for n, f in self._orig_mb.items():
            setattr(app_mod.messagebox, n, f)
        super().tearDown()

    def pump(self, cond, timeout: float = 8.0) -> bool:
        end = time.monotonic() + timeout
        while time.monotonic() < end:
            self.app.update()
            if cond():
                return True
            time.sleep(0.01)
        self.app.update()
        return bool(cond())

    def names(self) -> list[str]:
        return [self.app.item_meta[i]["name"] for i in self.app.tree.get_children() if i in self.app.item_meta]

    def acct(self, email: str) -> accounts.Account:
        return self.app.store.find_email(email)

    def open_named(self, name: str) -> None:
        iid = next(i for i, m in self.app.item_meta.items() if m["name"] == name)
        self.app.tree.focus(iid)
        self.app.tree.selection_set(iid)
        self.app.open_selected()


class TestMultiAccountUI(AppBase):
    def test_two_accounts_are_separate(self):
        self.start(("a@gmail.com", "b@gmail.com"))
        a, b = self.acct("a@gmail.com"), self.acct("b@gmail.com")
        self.assertTrue(self.app.nav.exists(f"a:{a.id}") and self.app.nav.exists(f"a:{b.id}"))
        self.assertTrue(self.pump(lambda: "A의 문서.docx" in self.names()))
        self.app.navigate([app_mod.loc_mydrive(b.id)])
        self.assertTrue(self.pump(lambda: "B의 문서.docx" in self.names()))
        self.assertNotIn("A의 문서.docx", self.names())
        self.open_named("B의 문서.docx")
        self.assertTrue(self.pump(lambda: len(self.opened) == 1))
        self.assertTrue(str(self.opened[0]).startswith(str(b.work_dir)))
        write_in_place(self.opened[0], b"edited B")
        db, da = self.drives["b@gmail.com"], self.drives["a@gmail.com"]
        self.assertTrue(self.pump(lambda: db.uploads, 8))
        self.assertEqual(da.uploads, [])                     # 다른 계정으로 올라가지 않음
        self.assertEqual(self.app.track_tree.cget("displaycolumns"), ("account", "status"))

    def test_add_account_and_duplicate(self):
        self.start(("a@gmail.com",))
        self.app.add_account()
        self.assertTrue(self.pump(lambda: self.acct("b@gmail.com") is not None))
        b = self.acct("b@gmail.com")
        self.assertTrue(b.token_path.exists())
        if os.name == "nt":
            self.assertTrue(b.token_path.read_text(encoding="utf-8").startswith(protect.PREFIX))
        self.assertTrue(self.pump(lambda: "B의 문서.docx" in self.names()))
        self.app.add_account()                               # 같은 계정을 또 추가
        self.pump(lambda: not self.app._adding)
        self.assertEqual(len(self.app.store.accounts), 2)

    def test_remove_account_refused_with_unsynced_then_allowed(self):
        self.start(("a@gmail.com", "b@gmail.com"))
        b = self.acct("b@gmail.com")
        protect.write_token(b.token_path, '{"refresh_token": "rb"}')
        self.app.navigate([app_mod.loc_mydrive(b.id)])
        self.pump(lambda: "B의 문서.docx" in self.names())
        self.open_named("B의 문서.docx")
        self.pump(lambda: len(self.opened) == 1)
        self.app.drives.pop(b.id)                             # 업로드 못 하게
        write_in_place(self.opened[0], b"not uploaded")
        self.app.remove_account(b.id)
        self.assertIn("올라가지 않은 변경", self.dialogs[-1][1])
        self.assertIsNotNone(self.acct("b@gmail.com"))
        write_in_place(self.opened[0], b"from B")             # 원래 내용으로 되돌림 → 올릴 것 없음
        time.sleep(0.5)
        self.app.remove_account(b.id)
        self.assertTrue(self.pump(lambda: self.acct("b@gmail.com") is None))
        self.assertFalse(b.token_path.exists())
        self.assertEqual(self.revoked, ['{"refresh_token": "rb"}'])
        self.assertFalse(self.app.nav.exists(f"a:{b.id}"))


class TestLockUI(AppBase):
    def set_password(self, pw="1234", **opts) -> None:
        s = security.Security()
        s.load()
        s.set_password(pw)
        s.update(True, opts.get("start", True), opts.get("idle", 10), opts.get("sleep", True))

    def test_lock_on_start_blocks_until_password(self):
        self.set_password()
        self.start()
        self.assertIsNotNone(self.app.lock)
        self.assertTrue(self.pump(lambda: "A의 문서.docx" in self.names()))   # 목록은 잠금 뒤에서 준비됨
        self.open_named("A의 문서.docx")
        self.pump(lambda: False, 0.5)
        self.assertEqual(self.opened, [])                    # 잠긴 동안은 열리지 않음
        self.app.lock.entry.insert(0, "0000")
        self.app.lock.try_unlock()
        self.assertIsNotNone(self.app.lock)
        self.assertIn("틀렸", self.app.lock.msg.get())
        self.app.lock.entry.insert(0, "1234")
        self.app.lock.try_unlock()
        self.assertIsNone(self.app.lock)

    def test_no_lock_when_disabled(self):
        self.set_password()
        s = security.Security()
        s.load()
        s.disable_all()
        self.start()
        self.assertIsNone(self.app.lock)
        self.app.lock_now("manual")
        self.assertIsNone(self.app.lock)                     # 꺼져 있으면 잠그지 않음

    def test_idle_sleep_and_manual_locks(self):
        self.set_password(start=False, idle=1)
        self.start()
        self.assertIsNone(self.app.lock)
        orig = security.system_idle_seconds
        security.system_idle_seconds = lambda: 61.0         # 1분 넘게 손대지 않음
        try:
            self.app._security_tick()
        finally:
            security.system_idle_seconds = orig
        self.assertEqual(self.app.lock.reason, "idle")
        self.app.unlock()
        self.app._last_wall = time.time() - 120             # 절전에서 돌아온 것처럼 시계가 건너뜀
        self.app._security_tick()
        self.assertEqual(self.app.lock.reason, "sleep")
        self.app.unlock()
        self.app.lock_now("manual")
        self.assertEqual(self.app.lock.reason, "manual")

    def test_dialogs_and_file_open_wait_until_unlock(self):
        self.set_password(start=False)
        self.start()
        self.pump(lambda: "A의 문서.docx" in self.names())
        self.app.lock_now("manual")
        a = self.acct("a@gmail.com")
        self.app._on_sync_event((a.id, "conflict", {"reason": "테스트", "name": "사본.docx"}))
        self.assertEqual(self.dialogs, [])                   # 잠긴 화면 위에 파일 이름이 뜨지 않음
        self.app.lock.entry.insert(0, "1234")
        self.app.lock.try_unlock()
        self.assertTrue(any("사본.docx" in m for _, m in self.dialogs))

    def test_forgot_password_with_registered_google_account(self):
        self.set_password()
        self.start(add_flow=lambda cb: ('{"refresh_token": "new"}', "a@gmail.com"))
        self.assertIsNotNone(self.app.lock)
        self.app.lock.forgot()
        self.assertTrue(self.pump(lambda: self.app.lock is None))
        self.assertFalse(self.app.security.has_password)     # 새 비밀번호를 정하도록
        self.assertIsNotNone(self.app._sec_dialog)

    def test_forgot_password_with_unknown_account_refused(self):
        self.set_password()
        self.start(add_flow=lambda cb: ('{"refresh_token": "x"}', "stranger@gmail.com"))
        self.app.lock.forgot()
        self.assertTrue(self.pump(lambda: "등록된 계정이 아닙니다" in self.app.lock.msg.get()))
        self.assertIsNotNone(self.app.lock)

    def test_deleting_security_file_does_not_unlock(self):
        self.set_password()
        self.start()
        a = self.acct("a@gmail.com")
        protect.write_token(a.token_path, '{"refresh_token": "r"}')
        self.app.unlock()
        self.app.security_changed("테스트")                  # 토큰에 잠금 표시 기록
        self.pump(lambda: protect.read_token(a.token_path)[1], 3)
        self.app.shutdown()
        self.app = None
        (self.tmp / "security.json").unlink()                # 설정 파일을 지워 잠금을 풀려는 시도
        self.start()
        self.assertIsNotNone(self.app.lock)
        self.assertEqual(self.app.lock.reason, "reset")

    def test_security_dialog_creates_password(self):
        self.start()
        self.app.lock_now("manual")                          # 비밀번호가 없으면 설정 창을 연다
        self.assertIsNone(self.app.lock)
        dlg = self.app._sec_dialog
        self.assertIsNotNone(dlg)
        dlg.new1.insert(0, "5678")
        dlg.new2.insert(0, "5678")
        dlg._save()
        self.assertTrue(self.app.security.active)
        self.app.lock_now("manual")
        self.assertIsNotNone(self.app.lock)

    def test_executable_from_drive_is_not_run(self):
        self.drives["a@gmail.com"].add_file("설치.exe", b"MZ fake", mime="application/x-msdownload")
        self.start()
        self.app.refresh()
        self.pump(lambda: "설치.exe" in self.names())
        self.answer = False
        self.open_named("설치.exe")
        self.assertTrue(self.pump(lambda: any("실행 파일" in m for _, m in self.dialogs)))
        self.assertEqual(self.opened, [])


if __name__ == "__main__":
    unittest.main(verbosity=2)
