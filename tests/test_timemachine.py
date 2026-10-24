"""타임머신(지운 파일 되살리기·이전 버전 되돌리기)과 도움말 테스트.

실행: .venv\\Scripts\\python -m unittest tests.test_timemachine -v
"""
from __future__ import annotations

import sys
import time
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(Path(__file__).resolve().parent))

import app as app_mod  # noqa: E402
from fake_drive import write_in_place  # noqa: E402
from test_accounts_security import AppBase  # noqa: E402


class TestUntrashGuard(unittest.TestCase):
    def test_only_untrash_allowed(self):
        import drive_client as dc
        url = "https://www.googleapis.com/drive/v3/files/abc123?fields=id&supportsAllDrives=true"
        dc.check_request(url, "PATCH", b'{"trashed": false}', {"content-type": "application/json"})
        for body in (b'{"trashed": true}', b'{"trashed": false, "name": "x"}', b'{"name": "x"}', b"", None):
            with self.assertRaises(dc.BlockedRequest, msg=repr(body)):
                dc.check_request(url, "PATCH", body, {"content-type": "application/json"})
        with self.assertRaises(dc.BlockedRequest):   # 다른 경로(공유 설정 등)는 여전히 막힘
            dc.check_request("https://www.googleapis.com/drive/v3/files/abc/permissions/p1", "PATCH",
                             b'{"trashed": false}', {})


class TestTimeMachineUI(AppBase):
    def setUp(self) -> None:
        super().setUp()
        self.d = self.drives["a@gmail.com"]
        self.doc = self.d.add_file("학급 문집.docx", b"v1 first")
        self.d.remote_edit(self.doc["id"], b"v2 second")
        self.d.remote_edit(self.doc["id"], b"v3 third")
        self.gone = self.d.add_file("지운 계획서.hwp", b"plan", mime="application/x-hwp")
        self.d.trash(self.gone["id"])
        self.start()
        self.pump(lambda: "학급 문집.docx" in self.names())

    def tm(self):
        return self.app._tm_dialog

    def test_restore_deleted_file(self):
        self.assertNotIn("지운 계획서.hwp", self.names())
        self.app.open_time_machine()
        self.assertTrue(self.pump(lambda: self.tm().trash_tree.exists(self.gone["id"])))
        self.tm().trash_tree.selection_set(self.gone["id"])
        self.tm().restore_selected()
        self.assertTrue(self.pump(lambda: not self.d.files[self.gone["id"]].get("trashed")))
        self.assertTrue(self.pump(lambda: "지운 계획서.hwp" in self.names()))   # 원래 폴더 목록에 다시 보임
        self.assertFalse(self.tm().trash_tree.exists(self.gone["id"]))

    def test_trash_filter(self):
        other = self.d.add_file("다른 파일.xlsx", b"x")
        self.d.trash(other["id"])
        self.app.open_time_machine()
        self.pump(lambda: len(self.tm().trash_tree.get_children()) == 2)
        self.tm().filter_var.set("계획")
        self.tm().search_trash()
        self.assertTrue(self.pump(lambda: list(self.tm().trash_tree.get_children()) == [self.gone["id"]]))
        self.tm()._reset_search()
        self.assertTrue(self.pump(lambda: len(self.tm().trash_tree.get_children()) == 2))

    def test_trash_paging_more_button(self):
        for i in range(60):
            f = self.d.add_file(f"지운 파일 {i:02d}.txt", b"x")
            self.d.trash(f["id"])
        orig = self.d.list_trash_page
        self.d.list_trash_page = lambda token=None, name=None, page_size=500: orig(token, name, 25)
        self.app.open_time_machine()
        tm = self.tm()
        self.assertTrue(self.pump(lambda: len(tm.trash_tree.get_children()) == 50))   # 40개 넘을 때까지 두 쪽
        self.assertNotIn("disabled", tm.btn_more.state())                            # 더 있음 → 버튼 켜짐
        tm._fetch_trash_page()
        self.assertTrue(self.pump(lambda: len(tm.trash_tree.get_children()) == 61))

    def test_restore_previous_version_and_stale_copy_protection(self):
        self.open_named("학급 문집.docx")                       # PC에 v3 사본이 생김
        self.assertTrue(self.pump(lambda: len(self.opened) == 1))
        local = self.opened[0]
        self.assertEqual(local.read_bytes(), b"v3 third")
        self.app.versions_of_selected()
        tm = self.tm()
        self.assertTrue(self.pump(lambda: len(tm.ver_tree.get_children()) == 3))
        oldest = tm.revs[0]["id"]
        self.assertEqual(tm.ver_tree.set(tm.ver_tree.get_children()[0], "note"), "지금 버전")
        tm.ver_tree.selection_set(oldest)
        tm.restore_version()
        self.assertTrue(self.pump(lambda: self.d.content[self.doc["id"]] == b"v1 first"))
        self.assertTrue(self.pump(lambda: len(tm.ver_tree.get_children()) == 4))  # 지금 내용도 기록에 남음
        # 다시 열면 낡은 PC 사본이 아니라 되돌린 내용으로 열려야 함
        self.app.refresh()
        self.pump(lambda: self.app.item_meta.get(self.doc["id"], {}).get("md5Checksum")
                  == self.d.files[self.doc["id"]]["md5Checksum"])
        self.open_named("학급 문집.docx")
        self.assertTrue(self.pump(lambda: len(self.opened) == 2))
        self.assertEqual(self.opened[1].read_bytes(), b"v1 first")

    def test_saving_stale_copy_after_restore_makes_conflict_copy(self):
        self.open_named("학급 문집.docx")
        self.assertTrue(self.pump(lambda: len(self.opened) == 1))
        local = self.opened[0]
        self.app.versions_of_selected()
        tm = self.tm()
        self.pump(lambda: len(tm.ver_tree.get_children()) == 3)
        tm.ver_tree.selection_set(tm.revs[0]["id"])
        tm.restore_version()
        self.assertTrue(self.pump(lambda: self.d.content[self.doc["id"]] == b"v1 first"))
        write_in_place(local, b"edited old v3 copy")          # 낡은 사본을 고쳐 저장
        self.assertTrue(self.pump(lambda: self.d.creates, 10))
        self.assertEqual(self.d.content[self.doc["id"]], b"v1 first")   # 되돌린 내용은 덮어쓰지 않음
        self.assertIn("충돌 사본", self.d.creates[0][0])

    def test_open_old_version_for_viewing(self):
        meta = next(m for m in self.app.item_meta.values() if m["name"] == "학급 문집.docx")
        self.app.open_time_machine(meta)
        tm = self.tm()
        self.pump(lambda: len(tm.ver_tree.get_children()) == 3)
        tm.ver_tree.selection_set(tm.revs[1]["id"])
        tm.open_version()
        self.assertTrue(self.pump(lambda: len(self.opened) == 1))
        view = self.opened[0]
        self.assertEqual(view.read_bytes(), b"v2 second")
        self.assertIn("이전 버전", view.name)
        self.assertNotIn(str(self.acct("a@gmail.com").work_dir), str(view))   # 작업 폴더 밖 (올라가지 않음)
        write_in_place(view, b"changed view copy")
        time.sleep(1.0)
        self.pump(lambda: False, 0.3)
        self.assertEqual(self.d.uploads, [])

    def test_google_doc_and_folder_have_no_version_list(self):
        g = self.d.add_file("구글 문서", b"g", mime="application/vnd.google-apps.document")
        self.app.open_time_machine(g)
        tm = self.tm()
        self.assertIn("브라우저", tm.ver_status.get())
        self.assertIn("disabled", tm.btn_ver_restore.state())
        tm.show_versions({"id": "folder1", "name": "내 폴더", "mimeType": "application/vnd.google-apps.folder"})
        self.assertIn("폴더", tm.ver_status.get())

    def test_lock_closes_time_machine(self):
        import security
        s = security.Security()
        s.load()
        s.set_password("1234")
        self.app.security.load()
        self.app.security_changed("테스트")
        self.app.open_time_machine()
        self.assertIsNotNone(self.tm())
        self.app.lock_now("manual")
        self.assertIsNone(self.app._tm_dialog)


class TestHelp(AppBase):
    def test_account_help_and_start_button(self):
        self.start()
        self.app.show_help("account")
        dlgs = [w for w in self.app.winfo_children() if isinstance(w, app_mod.HelpDialog)]
        self.assertEqual(len(dlgs), 1)
        text = dlgs[0].text.get("1.0", "end")
        for must in ("1단계", "테스트 사용자", "2단계", "고급", "access_denied", "계정 제거"):
            self.assertIn(must, text)
        dlgs[0]._start_add()                                   # [계정 추가 시작]
        self.assertTrue(self.pump(lambda: self.acct("b@gmail.com") is not None))

    def test_other_help_topics_open(self):
        self.start()
        for topic in ("lock", "timemachine"):
            self.app.show_help(topic)
        self.assertEqual(len([w for w in self.app.winfo_children() if isinstance(w, app_mod.HelpDialog)]), 2)

    def test_access_denied_offers_help(self):
        from errors import AuthError

        def denied(cb):
            raise AuthError("로그인이 취소되었거나 차단되었습니다.\n다른 계정이라면 Google Cloud 콘솔의 '테스트 사용자'에 "
                            "그 계정을 추가해야 합니다 (README '다른 계정 추가').")
        self.start(add_flow=denied)
        self.app.add_account()
        self.assertTrue(self.pump(lambda: any("단계별 안내" in m for _, m in self.dialogs)))
        self.assertTrue(self.pump(lambda: any(isinstance(w, app_mod.HelpDialog) for w in self.app.winfo_children())))

    def test_console_url_has_project(self):
        url = app_mod._console_url()
        self.assertTrue(url.startswith("https://console.cloud.google.com/auth/audience"))


if __name__ == "__main__":
    unittest.main(verbosity=2)
