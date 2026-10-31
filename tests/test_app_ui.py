"""탐색기 화면 동작 테스트 (실제 Tk 창 + 가짜 드라이브)."""
from __future__ import annotations

import shutil
import sys
import tempfile
import time
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(Path(__file__).resolve().parent))

import accounts  # noqa: E402
import app as app_mod  # noqa: E402
import config  # noqa: E402
from fake_drive import FakeDrive, write_in_place  # noqa: E402

DOCX = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"


class AppTest(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp = Path(tempfile.mkdtemp(prefix="gde_ui_"))
        self._saved = {k: getattr(config, k) for k in
                       ("BASE_DIR", "WORK_DIR", "STATE_PATH", "STAGING_DIR", "TOKEN_PATH", "LOG_PATH")}
        config.BASE_DIR = self.tmp
        config.WORK_DIR = self.tmp / "work"
        config.STATE_PATH = self.tmp / "state.json"
        config.STAGING_DIR = self.tmp / "staging"
        config.TOKEN_PATH = self.tmp / "token.json"
        config.LOG_PATH = self.tmp / "app.log"
        self.opened: list[Path] = []
        self._orig_open = app_mod.open_with_default
        app_mod.open_with_default = lambda p: self.opened.append(Path(p))
        self.dialogs: list[tuple[str, str]] = []
        self._orig_mb = {n: getattr(app_mod.messagebox, n) for n in
                         ("showinfo", "showwarning", "showerror", "askyesno")}
        app_mod.messagebox.showinfo = lambda t, m, **k: self.dialogs.append(("info", m))
        app_mod.messagebox.showwarning = lambda t, m, **k: self.dialogs.append(("warning", m))
        app_mod.messagebox.showerror = lambda t, m, **k: self.dialogs.append(("error", m))
        self.answer = True
        app_mod.messagebox.askyesno = lambda t, m, **k: (self.dialogs.append(("ask", m)), self.answer)[1]

        self.drive = FakeDrive()
        self.doc = self.drive.add_file("가정통신문.docx", b"hello")
        self.hwp = self.drive.add_file("학급 2반 명렬표.hwp", b"hwp data", mime="application/x-hwp")
        self.drive.add_file("학급 10반 명렬표.hwp", b"x", mime="application/x-hwp")
        self.gdoc = self.drive.add_file("구글 문서", b"g", mime="application/vnd.google-apps.document")
        self.form = self.drive.add_file("설문", None, mime="application/vnd.google-apps.form")
        self.drive.add_file("안쪽 파일.xlsx", b"in", parent="folder1",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        self.sc = self.drive.add_file("폴더 바로가기", None, mime="application/vnd.google-apps.shortcut",
                                      extra={"shortcutDetails": {"targetId": "folder1",
                                                                 "targetMimeType": "application/vnd.google-apps.folder"}})
        store = accounts.AccountStore()
        store.load()
        self.acct, _ = store.add("tester@example.com")
        self.app = app_mod.App(drive_factory=lambda acct, cb: self.drive)
        self.app.sync.debounce = 0.3
        self.app.sync.poll = 0.25
        self.pump(lambda: self.app.drive is not None and self.names())

    def tearDown(self) -> None:
        try:
            self.app.shutdown()
        except Exception:
            pass
        self.app = None
        import gc
        gc.collect()   # 닫은 창의 객체를 화면 스레드에서 정리
        app_mod.open_with_default = self._orig_open
        for n, f in self._orig_mb.items():
            setattr(app_mod.messagebox, n, f)
        for k, v in self._saved.items():
            setattr(config, k, v)
        time.sleep(0.1)
        shutil.rmtree(self.tmp, ignore_errors=True)

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

    def placeholder(self) -> str:
        t = self.app.tree
        return t.item("__ph__", "text").strip() if t.exists("__ph__") else ""

    def activate(self, file_id: str) -> None:
        self.app.tree.focus(file_id)
        self.app.tree.selection_set(file_id)
        self.app.open_selected()

    # ----- 목록 -----
    def test_root_listing_sorted_folders_first_natural(self):
        names = self.names()
        self.assertEqual(names[0], "내 폴더")
        self.assertEqual(names[1], "폴더 바로가기")
        self.assertLess(names.index("학급 2반 명렬표.hwp"), names.index("학급 10반 명렬표.hwp"))
        self.assertEqual(self.app.account_var.get(), "tester@example.com")

    def test_navigate_into_folder_and_back(self):
        self.activate("folder1")
        self.pump(lambda: self.names() == ["안쪽 파일.xlsx"])
        self.assertEqual([loc.name for loc in self.app.stack], ["내 드라이브", "내 폴더"])
        self.app.go_back()
        self.pump(lambda: "가정통신문.docx" in self.names())
        self.app.go_forward()
        self.pump(lambda: self.names() == ["안쪽 파일.xlsx"])
        self.app.go_up()
        self.pump(lambda: "가정통신문.docx" in self.names())

    def test_shortcut_to_folder_navigates(self):
        self.activate(self.sc["id"])
        self.assertTrue(self.pump(lambda: self.names() == ["안쪽 파일.xlsx"]))

    def test_stale_listing_does_not_overwrite_current_folder(self):
        self.drive.list_delay["folder1"] = 1.0
        self.activate("folder1")        # 느린 요청
        self.app.update()
        self.app.go_back()              # 바로 돌아옴 (루트는 캐시로 즉시 표시)
        self.pump(lambda: "가정통신문.docx" in self.names())
        time.sleep(1.3)                 # 느린 응답 도착
        self.pump(lambda: False, timeout=0.5)
        self.assertIn("가정통신문.docx", self.names())
        self.assertEqual(self.app.stack[-1].id, "root")

    def test_cached_folder_shows_instantly(self):
        self.activate("folder1")
        self.pump(lambda: self.names() == ["안쪽 파일.xlsx"])
        self.app.go_back()
        self.pump(lambda: "가정통신문.docx" in self.names())
        self.drive.list_delay["folder1"] = 2.0
        self.activate("folder1")
        self.app.update()
        self.assertEqual(self.names(), ["안쪽 파일.xlsx"])  # 네트워크를 기다리지 않음

    def test_listing_error_shows_message(self):
        self.drive.fail_list = [ConnectionResetError(10054, "reset")]
        self.activate("folder1")
        self.pump(lambda: "불러오지 못했습니다" in self.placeholder())
        self.assertIn("인터넷 연결", self.placeholder())
        self.app.refresh()
        self.assertTrue(self.pump(lambda: self.names() == ["안쪽 파일.xlsx"]))

    def test_search(self):
        self.app.search_var.set("명렬표")
        self.app._on_search()
        self.assertTrue(self.pump(lambda: sorted(self.names()) == sorted(["학급 2반 명렬표.hwp", "학급 10반 명렬표.hwp"])))
        self.app.search_var.set("it's \\ odd")
        self.app._on_search()
        self.assertTrue(self.pump(lambda: self.placeholder() == "(비어 있음)"))

    def test_sort_columns(self):
        for col in ("size", "kind", "modified", "name", "name"):
            self.app._sort_by(col)
        self.assertTrue(self.app.sort_rev)
        names = self.names()
        self.assertEqual(names[:2], ["폴더 바로가기", "내 폴더"])

    # ----- 열기 -----
    def test_open_file_downloads_once_and_reopens_from_local(self):
        self.activate(self.doc["id"])
        self.activate(self.doc["id"])   # 빠른 두 번 클릭
        self.assertTrue(self.pump(lambda: len(self.opened) == 1))
        self.pump(lambda: False, timeout=0.3)
        self.assertEqual(self.drive.download_count, 1)
        path = self.opened[0]
        self.assertEqual(path.read_bytes(), b"hello")
        self.assertTrue(str(path).startswith(str(self.acct.work_dir)))
        self.activate(self.doc["id"])   # 이미 받은 파일은 다시 받지 않음
        self.assertTrue(self.pump(lambda: len(self.opened) == 2))
        self.assertEqual(self.drive.download_count, 1)
        rows = self.app.track_tree.get_children()
        self.assertEqual(len(rows), 1)

    def test_remote_newer_and_local_clean_redownloads(self):
        self.activate(self.doc["id"])
        self.pump(lambda: len(self.opened) == 1)
        self.drive.remote_edit(self.doc["id"], b"newer from home")
        self.app.refresh()
        self.pump(lambda: self.app.item_meta.get(self.doc["id"], {}).get("md5Checksum")
                  == self.drive.files[self.doc["id"]]["md5Checksum"])
        self.activate(self.doc["id"])
        self.assertTrue(self.pump(lambda: len(self.opened) == 2))
        self.assertEqual(self.opened[1].read_bytes(), b"newer from home")
        self.assertEqual(self.drive.download_count, 2)

    def test_local_edit_uploads_and_status_shows(self):
        self.activate(self.hwp["id"])
        self.pump(lambda: len(self.opened) == 1)
        write_in_place(self.opened[0], b"edited hwp")
        self.assertTrue(self.pump(lambda: self.drive.content[self.hwp["id"]] == b"edited hwp", 8))
        iid = self.app.track_tree.get_children()[0]
        self.assertTrue(self.pump(lambda: "드라이브에 저장됨" in self.app.track_tree.set(iid, "status")))

    def test_google_doc_opens_copy_with_notice(self):
        self.activate(self.gdoc["id"])
        self.assertTrue(self.pump(lambda: len(self.opened) == 1))
        self.assertTrue(self.opened[0].name.endswith(".docx"))
        self.pump(lambda: self.dialogs)
        self.assertEqual(self.dialogs[0][0], "info")
        write_in_place(self.opened[0], b"edit copy")
        time.sleep(1.0)
        self.pump(lambda: False, timeout=0.3)
        self.assertEqual(self.drive.uploads, [])

    def test_form_opens_in_browser(self):
        called = []
        orig = app_mod.webbrowser.open
        app_mod.webbrowser.open = lambda url: called.append(url)
        try:
            self.activate(self.form["id"])
        finally:
            app_mod.webbrowser.open = orig
        self.assertEqual(called, [self.form["webViewLink"]])
        self.assertEqual(self.drive.download_count, 0)

    def test_open_failure_shows_error(self):
        meta = self.drive.add_file("사라질.docx", b"x")
        self.app.refresh()
        self.pump(lambda: meta["id"] in self.app.item_meta)
        self.drive.delete(meta["id"])
        self.activate(meta["id"])
        self.assertTrue(self.pump(lambda: any(k == "error" for k, _ in self.dialogs)))
        self.assertIn("찾을 수 없습니다", self.dialogs[-1][1])
        self.assertNotIn(meta["id"], self.app.downloading)

    def test_illegal_name_downloads(self):
        meta = self.drive.add_file('3/1 회의: "중요"?.docx', b"x")
        self.app.refresh()
        self.pump(lambda: meta["id"] in self.app.item_meta)
        self.activate(meta["id"])
        self.assertTrue(self.pump(lambda: len(self.opened) == 1))
        self.assertEqual(self.opened[0].name, "3_1 회의_ _중요__.docx")

    def test_open_uses_listing_meta_without_extra_request(self):
        before = self.drive.get_meta_count
        self.activate(self.doc["id"])
        self.assertTrue(self.pump(lambda: len(self.opened) == 1))
        self.assertEqual(self.drive.get_meta_count, before)  # 목록 정보로 바로 내려받음

    def test_stale_listing_meta_still_opens_latest(self):
        # 목록을 받은 뒤 다른 곳에서 파일이 바뀐 경우: 최신 내용으로 열고 기록도 최신으로
        self.drive.remote_edit(self.doc["id"], b"changed after listing")
        self.activate(self.doc["id"])
        self.assertTrue(self.pump(lambda: len(self.opened) == 1))
        self.assertEqual(self.opened[0].read_bytes(), b"changed after listing")
        import sync
        rec = self.app.sync.get(sync.norm(self.opened[0]))
        self.assertEqual(rec["synced_md5"], self.drive.files[self.doc["id"]]["md5Checksum"])
        write_in_place(self.opened[0], b"local edit")
        self.assertTrue(self.pump(lambda: self.drive.content[self.doc["id"]] == b"local edit", 8))
        self.assertEqual(self.drive.creates, [])  # 충돌로 잘못 보지 않음

    def test_root_cache_shown_before_connect_and_pending_open(self):
        self.assertTrue(self.pump(lambda: self.acct.root_cache_path.exists()))
        self.app.shutdown()
        import threading
        gate = threading.Event()

        def slow_factory(acct, cb):
            gate.wait(10)
            return self.drive
        self.app = app_mod.App(drive_factory=slow_factory)
        self.app.update()
        self.assertIn("가정통신문.docx", self.names())   # 지난번 목록을 연결 전에 바로 표시
        self.activate(self.doc["id"])                  # 연결 전에 더블클릭
        self.assertIn("연결되면 엽니다", self.app.status_var.get())
        gate.set()
        self.assertTrue(self.pump(lambda: len(self.opened) == 1))

    # ----- 종료 -----
    def test_close_with_unsynced_asks(self):
        self.activate(self.doc["id"])
        self.pump(lambda: len(self.opened) == 1)
        self.app.drive = None  # 업로드 못 하는 상태
        write_in_place(self.opened[0], b"unsaved")
        self.answer = False
        self.app.on_close()
        self.assertFalse(self.app._closing)
        self.assertIn("가정통신문.docx", self.dialogs[-1][1])

    def test_auth_error_shows_login_button(self):
        from errors import AuthError
        self.drive.fail_list = [AuthError("로그인이 필요합니다.")]
        self.activate("folder1")
        self.assertTrue(self.pump(lambda: self.app._login_btn_shown))
        self.app.on_login_button()
        self.assertTrue(self.pump(lambda: self.drive.relogin_count == 1))
        self.assertTrue(self.pump(lambda: self.names() == ["안쪽 파일.xlsx"]))

    def test_connect_failure_then_retry(self):
        self.app.shutdown()
        attempts = {"n": 0}

        def factory(acct, cb):
            attempts["n"] += 1
            if attempts["n"] == 1:
                raise ConnectionResetError(10054, "reset")
            return self.drive
        self.app = app_mod.App(drive_factory=factory)
        self.assertTrue(self.pump(lambda: self.app._login_btn_shown))
        self.assertIn("연결 실패", self.app.status_var.get())
        self.assertIn("가정통신문.docx", self.names())   # 연결 실패해도 지난번 목록은 계속 보임
        self.app.on_login_button()
        self.assertTrue(self.pump(lambda: "가정통신문.docx" in self.names()))


class TestHelpers(unittest.TestCase):
    def test_natural_key(self):
        names = ["10반", "2반", "1반", "x²", "²", "A", "a"]
        self.assertEqual(sorted(names, key=app_mod.natural_key)[:3], ["1반", "2반", "10반"])

    def test_fmt_time(self):
        self.assertRegex(app_mod.fmt_time("2024-03-01T05:12:33.123Z"), r"^2024-03-01 \d\d:12$")
        self.assertEqual(app_mod.fmt_time(""), "")
        self.assertEqual(app_mod.fmt_time("bad"), "bad")

    def test_human_size(self):
        self.assertEqual(app_mod.human_size("0"), "0 B")
        self.assertEqual(app_mod.human_size(1536), "1.5 KB")
        self.assertEqual(app_mod.human_size(None), "")

    def test_track_kind(self):
        import sync
        cases = {"드라이브에 저장됨 14:32:10": "saved", "드라이브와 같음": "saved", "업로드 중...": "uploading",
                 sync.STATUS_IDLE: "open", sync.STATUS_NATIVE: "copy", sync.STATUS_READONLY: "copy",
                 "업로드 실패: 응답 없음 - 4초 후 다시 시도": "wait", "파일이 잠겨 있어 기다리는 중...": "wait",
                 "업로드 실패: 권한 없음 (다시 저장하면 재시도)": "alert", sync.STATUS_HELD: "alert",
                 "충돌 사본으로 저장됨 10:00:00": "alert", "로컬 파일이 없습니다": "missing"}
        for status, kind in cases.items():
            self.assertEqual(app_mod.track_kind(status), kind, status)

    def test_days_left_in_trash(self):
        now = app_mod._dt.datetime.now(app_mod._dt.timezone.utc)

        def ago(days):
            return (now - app_mod._dt.timedelta(days=days, minutes=1)).isoformat().replace("+00:00", "Z")
        self.assertEqual(app_mod.days_left_in_trash(ago(0)), "30일")
        self.assertEqual(app_mod.days_left_in_trash(ago(28)), "2일 (곧 지워져요)")
        self.assertEqual(app_mod.days_left_in_trash(ago(40)), "0일 (곧 지워져요)")
        self.assertEqual(app_mod.days_left_in_trash(""), "")


if __name__ == "__main__":
    unittest.main(verbosity=2)
