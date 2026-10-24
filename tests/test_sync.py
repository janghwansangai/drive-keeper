"""저장 감지 → 업로드 시나리오 테스트 (실제 Windows 파일 시스템 + watchdog, 가짜 드라이브).

실행: .venv\\Scripts\\python -m unittest discover -s tests -v
"""
from __future__ import annotations

import os
import shutil
import sys
import tempfile
import threading
import time
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(Path(__file__).resolve().parent))

import sync  # noqa: E402
from fake_drive import FakeDrive, FakeHttpError, md5, wait_until, write_in_place  # noqa: E402

DOCX = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"


class SyncTestBase(unittest.TestCase):
    debounce = 0.3
    poll = 0.25

    def setUp(self) -> None:
        self.tmp = Path(tempfile.mkdtemp(prefix="gde_test_"))
        self.work = self.tmp / "work"
        self.state = self.tmp / "state.json"
        self.staging = self.tmp / "staging"
        self.drive = FakeDrive()
        self.events: list[tuple[str, object]] = []
        self.connected = True
        self.mgr = self.make_manager()

    def make_manager(self, **kw) -> sync.SyncManager:
        m = sync.SyncManager(self.work, self.state, self.staging,
                             get_drive=lambda: self.drive if self.connected else None,
                             notify=lambda ev, p: self.events.append((ev, p)),
                             debounce=self.debounce, poll=self.poll, **kw)
        m.start()
        return m

    def tearDown(self) -> None:
        self.mgr.stop()
        time.sleep(0.1)
        shutil.rmtree(self.tmp, ignore_errors=True)

    def open_remote(self, name: str, content: bytes, mgr=None, **kw) -> tuple[dict, Path]:
        """앱이 파일을 여는 과정을 흉내: 드라이브에서 받아 추적 목록에 넣는다."""
        mgr = mgr or self.mgr
        meta = self.drive.add_file(name, content, mime=kw.pop("mime", DOCX), **kw)
        fresh = self.drive.get_meta(meta["id"])
        folder = self.work / sync.short_dir_name(meta["id"])
        path = folder / sync.safe_filename(name, 150)
        digest = self.drive.download(fresh, path)
        kind = sync.KIND_EDIT if fresh["capabilities"]["canEdit"] else sync.KIND_READONLY
        mgr.add(path, fresh, digest, kind)
        return fresh, path

    def uploads_of(self, file_id: str) -> list[str]:
        return [m for fid, m in self.drive.uploads if fid == file_id]

    def status_of(self, path: Path, mgr=None) -> str:
        rec = (mgr or self.mgr).get(sync.norm(path))
        return rec["status"] if rec else ""


class TestSaveDetection(SyncTestBase):
    def test_open_only_no_upload(self):
        meta, path = self.open_remote("열기만.docx", b"original")
        time.sleep(1.5)
        self.assertEqual(self.drive.uploads, [])

    def test_in_place_write_uploads_once(self):
        meta, path = self.open_remote("제자리.docx", b"original")
        write_in_place(path, b"edited content")
        self.assertTrue(wait_until(lambda: self.uploads_of(meta["id"]), 5))
        time.sleep(1.0)
        self.assertEqual(self.uploads_of(meta["id"]), [md5(b"edited content")])
        self.assertEqual(self.drive.content[meta["id"]], b"edited content")
        self.assertIn("드라이브에 저장됨", self.status_of(path))

    def test_rapid_saves_debounced(self):
        meta, path = self.open_remote("연속저장.docx", b"v0")
        for i in range(1, 8):
            write_in_place(path, f"v{i}".encode())
            time.sleep(0.05)
        self.assertTrue(wait_until(lambda: self.uploads_of(meta["id"]), 5))
        time.sleep(1.0)
        self.assertEqual(self.uploads_of(meta["id"]), [md5(b"v7")])

    def test_word_style_temp_rename_save(self):
        # Word: 임시 파일에 쓰고 → 원본을 다른 이름으로 → 임시 파일을 원본 이름으로 → 옛 파일 삭제
        meta, path = self.open_remote("워드식.docx", b"original")
        folder = path.parent
        tmp = folder / "~WRD0001.tmp"
        tmp.write_bytes(b"word saved content")
        os.replace(path, folder / "~WRL0002.tmp")
        os.replace(tmp, path)
        (folder / "~WRL0002.tmp").unlink()
        self.assertTrue(wait_until(lambda: self.uploads_of(meta["id"]), 5))
        time.sleep(1.0)
        self.assertEqual(self.uploads_of(meta["id"]), [md5(b"word saved content")])

    def test_excel_style_random_temp_save(self):
        # Excel: 확장자 없는 무작위 이름 임시 파일 → 원본 자리로 바꿔치기
        meta, path = self.open_remote("엑셀식.xlsx", b"original")
        tmp = path.parent / "A1B2C3D4"
        tmp.write_bytes(b"excel content")
        os.replace(tmp, path)
        self.assertTrue(wait_until(lambda: self.uploads_of(meta["id"]), 5))
        time.sleep(0.8)
        self.assertEqual(self.uploads_of(meta["id"]), [md5(b"excel content")])

    def test_same_content_not_reuploaded(self):
        meta, path = self.open_remote("같은내용.docx", b"same")
        write_in_place(path, b"same")
        os.utime(path, None)
        time.sleep(1.5)
        self.assertEqual(self.drive.uploads, [])

    def test_office_lock_files_ignored(self):
        meta, path = self.open_remote("잠금파일.docx", b"x")
        (path.parent / "~$잠금파일.docx").write_bytes(b"owner")
        (path.parent / "something.tmp").write_bytes(b"tmp")
        time.sleep(1.2)
        self.assertEqual(self.drive.uploads, [])

    def test_native_copy_never_uploaded(self):
        meta = self.drive.add_file("구글문서", b"gdoc", mime="application/vnd.google-apps.document")
        path = self.work / sync.short_dir_name(meta["id"]) / "구글문서.docx"
        digest = self.drive.download(meta, path)
        self.mgr.add(path, self.drive.get_meta(meta["id"]), digest, sync.KIND_NATIVE)
        write_in_place(path, b"changed copy")
        time.sleep(1.5)
        self.assertEqual(self.drive.uploads, [])
        self.assertEqual(self.drive.creates, [])

    def test_readonly_never_uploaded(self):
        meta, path = self.open_remote("보기전용.docx", b"ro", can_edit=False)
        write_in_place(path, b"changed")
        time.sleep(1.5)
        self.assertEqual(self.drive.uploads, [])
        self.assertEqual(self.status_of(path), sync.STATUS_READONLY)

    def test_change_during_upload_uploads_again(self):
        meta, path = self.open_remote("업로드중변경.docx", b"v0")
        self.drive.upload_delay = 1.0
        write_in_place(path, b"v1")
        self.assertTrue(wait_until(lambda: "업로드 중" in self.status_of(path), 5))
        write_in_place(path, b"v2")
        self.assertTrue(wait_until(lambda: len(self.uploads_of(meta["id"])) >= 2, 8))
        self.assertEqual(self.uploads_of(meta["id"])[-1], md5(b"v2"))
        self.assertEqual(self.drive.content[meta["id"]], b"v2")

    def test_poll_catches_change_without_watchdog(self):
        mgr = self.make_manager(use_watchdog=False)
        try:
            meta, path = self.open_remote("감시없음.docx", b"a", mgr=mgr)
            time.sleep(0.4)
            write_in_place(path, b"polled change")
            self.assertTrue(wait_until(lambda: self.uploads_of(meta["id"]), 5))
        finally:
            mgr.stop()

    def test_large_file_uses_staging(self):
        mgr = self.make_manager(memory_limit=1024)
        try:
            meta, path = self.open_remote("큰파일.docx", b"small", mgr=mgr)
            big = os.urandom(300_000)
            write_in_place(path, big)
            self.assertTrue(wait_until(lambda: self.uploads_of(meta["id"]), 5))
            self.assertEqual(self.drive.content[meta["id"]], big)
            time.sleep(0.3)
            self.assertEqual(list(self.staging.iterdir()), [])  # 임시 사본은 지워짐
        finally:
            mgr.stop()

    def test_path_case_insensitive(self):
        meta, path = self.open_remote("CaseTest.docx", b"a")
        self.mgr.note_change(str(path).upper())
        rec = self.mgr.get(sync.norm(path))
        self.assertIsNotNone(rec)
        write_in_place(path, b"b")
        self.assertTrue(wait_until(lambda: self.uploads_of(meta["id"]), 5))


@unittest.skipUnless(os.name == "nt", "Windows 전용")
class TestLockedFiles(SyncTestBase):
    def _open_exclusive(self, path: Path):
        import ctypes
        from ctypes import wintypes
        k32 = ctypes.WinDLL("kernel32", use_last_error=True)
        k32.CreateFileW.restype = wintypes.HANDLE
        k32.CreateFileW.argtypes = [wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD, wintypes.LPVOID,
                                    wintypes.DWORD, wintypes.DWORD, wintypes.HANDLE]
        h = k32.CreateFileW(str(path), 0x80000000 | 0x40000000, 0, None, 3, 0, None)  # 공유 안 함
        self.assertNotEqual(h, wintypes.HANDLE(-1).value)
        return k32, h

    def test_locked_file_waits_then_uploads(self):
        meta, path = self.open_remote("잠김.docx", b"a")
        write_in_place(path, b"locked content")
        k32, h = self._open_exclusive(path)
        try:
            self.assertTrue(sync.is_in_use(path))
            time.sleep(1.5)
            self.assertEqual(self.drive.uploads, [])
            self.assertIn("잠겨", self.status_of(path))
        finally:
            k32.CloseHandle(h)
        self.assertTrue(wait_until(lambda: self.uploads_of(meta["id"]), 6))
        self.assertEqual(self.drive.content[meta["id"]], b"locked content")

    def test_shared_read_allows_rename_during_read(self):
        # 우리가 읽고 있는 동안에도 Word처럼 이름을 바꿀 수 있어야 한다
        p = self.tmp / "share.bin"
        p.write_bytes(b"x" * 1000)
        with sync.open_shared_read(p) as f:
            f.read(10)
            os.replace(p, self.tmp / "renamed.bin")
            os.replace(self.tmp / "renamed.bin", p)

    def test_cleanup_keeps_open_and_unsynced(self):
        m1, p1 = self.open_remote("닫힘.docx", b"1")
        m2, p2 = self.open_remote("열림.docx", b"2")
        m3, p3 = self.open_remote("수정됨.docx", b"3")
        self.connected = False  # 업로드 못 하게
        write_in_place(p3, b"3 changed")
        time.sleep(0.8)
        k32, h = self._open_exclusive(p2)
        try:
            removed, kept = self.mgr.cleanup()
        finally:
            k32.CloseHandle(h)
        self.assertEqual(removed, 1)
        self.assertFalse(p1.exists())
        self.assertFalse(p1.parent.exists())
        self.assertTrue(p2.exists())
        self.assertTrue(p3.exists())
        self.assertEqual(sorted(kept), sorted(["열림.docx", "수정됨.docx"]))
        self.assertIsNone(self.mgr.get_by_file_id(m1["id"]))
        self.assertIsNotNone(self.mgr.get_by_file_id(m2["id"]))


class TestFailuresAndConflicts(SyncTestBase):
    def test_transient_failure_retried(self):
        meta, path = self.open_remote("재시도.docx", b"a")
        self.drive.fail_upload = [ConnectionResetError(10054, "reset"), FakeHttpError(503, "backendError")]
        write_in_place(path, b"retry me")
        self.assertTrue(wait_until(lambda: "업로드 실패" in self.status_of(path), 5))
        self.assertTrue(wait_until(lambda: self.uploads_of(meta["id"]), 20))
        self.assertEqual(self.drive.content[meta["id"]], b"retry me")
        self.assertIn("드라이브에 저장됨", self.status_of(path))

    def test_permanent_failure_not_retried_until_next_save(self):
        meta, path = self.open_remote("권한없음.docx", b"a")
        self.drive.fail_upload = [FakeHttpError(403, "insufficientFilePermissions")]
        write_in_place(path, b"b")
        self.assertTrue(wait_until(lambda: "권한" in self.status_of(path), 5))
        time.sleep(2.5)
        self.assertEqual(self.drive.uploads, [])
        write_in_place(path, b"c")  # 다시 저장하면 재시도
        self.assertTrue(wait_until(lambda: self.uploads_of(meta["id"]), 5))

    def test_offline_waits_then_uploads(self):
        meta, path = self.open_remote("오프라인.docx", b"a")
        self.connected = False
        write_in_place(path, b"offline edit")
        self.assertTrue(wait_until(lambda: "연결 대기" in self.status_of(path), 5))
        self.connected = True
        self.assertTrue(wait_until(lambda: self.uploads_of(meta["id"]), 8))

    def test_remote_changed_creates_conflict_copy(self):
        meta, path = self.open_remote("충돌.docx", b"base")
        self.drive.remote_edit(meta["id"], b"edited at home")
        write_in_place(path, b"edited at school")
        self.assertTrue(wait_until(lambda: self.drive.creates, 5))
        self.assertEqual(self.drive.content[meta["id"]], b"edited at home")  # 원본은 그대로
        name, parent, digest = self.drive.creates[0]
        self.assertIn("충돌 사본", name)
        self.assertEqual(parent, "root")
        self.assertEqual(digest, md5(b"edited at school"))
        self.assertTrue(any(ev == "conflict" for ev, _ in self.events))
        # 이후 저장은 사본에 반영
        rec = self.mgr.get(sync.norm(path))
        copy_id = rec["file_id"]
        self.assertNotEqual(copy_id, meta["id"])
        write_in_place(path, b"second save")
        self.assertTrue(wait_until(lambda: self.uploads_of(copy_id), 5))
        self.assertEqual(self.drive.content[meta["id"]], b"edited at home")
        self.assertEqual(self.drive.content[copy_id], b"second save")

    def test_remote_deleted_creates_copy(self):
        meta, path = self.open_remote("삭제됨.docx", b"base")
        self.drive.delete(meta["id"])
        write_in_place(path, b"keep my work")
        self.assertTrue(wait_until(lambda: self.drive.creates, 5))
        self.assertEqual(self.drive.creates[0][2], md5(b"keep my work"))

    def test_lost_response_not_treated_as_conflict(self):
        # 업로드는 됐는데 응답만 못 받은 경우: 재시도 때 충돌 사본을 만들면 안 됨
        meta, path = self.open_remote("응답유실.docx", b"base")
        orig = self.drive.update_content
        calls = {"n": 0}

        def flaky(file_id, source, mime):
            calls["n"] += 1
            result = orig(file_id, source, mime)
            if calls["n"] == 1:
                raise TimeoutError("응답 없음")
            return result
        self.drive.update_content = flaky
        write_in_place(path, b"saved once")
        self.assertTrue(wait_until(lambda: "드라이브에 저장됨" in self.status_of(path), 15))
        self.assertEqual(self.drive.creates, [])
        write_in_place(path, b"saved twice")
        self.assertTrue(wait_until(lambda: self.drive.content[meta["id"]] == b"saved twice", 8))
        self.assertEqual(self.drive.creates, [])


class TestSaveAsNewFile(SyncTestBase):
    """Word/한글에서 '다른 이름으로 저장'으로 같은 폴더에 생긴 새 파일."""

    def setUp(self) -> None:
        super().setUp()
        self.mgr.new_file_delay = 0.5

    def creates_named(self, name: str) -> list[tuple[str, str, str]]:
        return [c for c in self.drive.creates if c[0] == name]

    def test_save_as_uploads_to_same_drive_folder(self):
        meta, path = self.open_remote("원본.hwpx", b"orig", parent="folder1")
        new = path.parent / "다른 이름.hwp"
        new.write_bytes(b"saved as hwp")
        self.assertTrue(wait_until(lambda: self.creates_named("다른 이름.hwp"), 6))
        name, parent, digest = self.creates_named("다른 이름.hwp")[0]
        self.assertEqual(parent, "folder1")
        self.assertEqual(digest, md5(b"saved as hwp"))
        self.assertTrue(wait_until(lambda: any(ev == "new_file" for ev, _ in self.events), 3))
        rec = self.mgr.get(sync.norm(new))
        self.assertTrue(rec["file_id"])
        self.assertIn("새 파일로 드라이브에 저장됨", rec["status"])
        # 이후 저장은 그 새 파일에 반영 (또 만들지 않음)
        write_in_place(new, b"second save")
        self.assertTrue(wait_until(lambda: self.uploads_of(rec["file_id"]), 5))
        self.assertEqual(self.drive.content[rec["file_id"]], b"second save")
        self.assertEqual(len(self.creates_named("다른 이름.hwp")), 1)
        self.assertEqual(self.drive.content[meta["id"]], b"orig")  # 원본은 그대로

    def test_temp_and_lock_files_not_uploaded(self):
        meta, path = self.open_remote("문서.docx", b"x")
        for n in ["~$문서.docx", "~WRD0001.tmp", "$tm8B82.tmp", "A1B2C3D4", ".~lock.문서.docx#",
                  "문서의 백업.wbk", "자동복구.asd", "desktop.ini", "x.docx.gde-part", "빈 파일.docx"]:
            (path.parent / n).write_bytes(b"" if n.startswith("빈") else b"junk")
        time.sleep(2.0)
        self.assertEqual(self.drive.creates, [])

    def test_short_lived_file_not_uploaded(self):
        meta, path = self.open_remote("문서.docx", b"x")
        tmp = path.parent / "잠깐.docx"
        tmp.write_bytes(b"temp")
        time.sleep(0.1)
        tmp.unlink()
        time.sleep(1.5)
        self.assertEqual(self.drive.creates, [])

    def test_growing_file_waits_until_done(self):
        meta, path = self.open_remote("문서.docx", b"x")
        new = path.parent / "천천히 저장.docx"
        with open(new, "wb") as f:
            for i in range(8):
                f.write(b"chunk%d;" % i)
                f.flush()
                os.fsync(f.fileno())
                time.sleep(0.2)
        self.assertTrue(wait_until(lambda: self.creates_named("천천히 저장.docx"), 6))
        self.assertEqual(self.creates_named("천천히 저장.docx")[0][2], md5(new.read_bytes()))
        self.assertEqual(len(self.creates_named("천천히 저장.docx")), 1)

    def test_existing_untracked_file_uploaded_after_restart(self):
        meta, path = self.open_remote("원본.docx", b"x", parent="folder1")
        self.mgr.stop()
        time.sleep(0.3)
        (path.parent / "꺼진 사이 저장.docx").write_bytes(b"made while closed")
        self.mgr = self.make_manager(new_file_delay=0.5)
        self.assertTrue(wait_until(lambda: self.creates_named("꺼진 사이 저장.docx"), 6))
        self.assertEqual(self.creates_named("꺼진 사이 저장.docx")[0][1], "folder1")

    def test_no_permission_on_folder_falls_back_to_my_drive(self):
        meta, path = self.open_remote("보기전용.docx", b"x", parent="folder1", can_edit=False)
        self.drive.readonly_parents = {"folder1"}
        (path.parent / "내 사본.docx").write_bytes(b"my copy")
        self.assertTrue(wait_until(lambda: self.creates_named("내 사본.docx"), 6))
        self.assertEqual(self.creates_named("내 사본.docx")[0][1], "root")

    def test_lost_create_response_does_not_duplicate(self):
        meta, path = self.open_remote("원본.docx", b"x")
        orig = self.drive.create_file
        calls = {"n": 0}

        def flaky(name, parent, source, mime):
            calls["n"] += 1
            result = orig(name, parent, source, mime)
            if calls["n"] == 1:
                raise TimeoutError("응답 없음")
            return result
        self.drive.create_file = flaky
        new = path.parent / "한 번만.docx"
        new.write_bytes(b"once")
        self.assertTrue(wait_until(lambda: "새 파일로 드라이브에 저장됨" in self.status_of(new), 15))
        same = [m for m in self.drive.files.values() if m["name"] == "한 번만.docx"]
        self.assertEqual(len(same), 1)
        self.assertEqual(calls["n"], 1)

    def test_cleanup_keeps_folder_with_pending_new_file(self):
        meta, path = self.open_remote("원본.docx", b"x")
        self.mgr.new_file_delay = 30  # 아직 올리기 전 상태 유지
        (path.parent / "방금 저장.docx").write_bytes(b"new work")
        removed, kept = self.mgr.cleanup()
        self.assertEqual(removed, 0)
        self.assertTrue(path.exists())
        self.assertTrue((path.parent / "방금 저장.docx").exists())

    def test_missing_tracked_file_stops_retrying(self):
        meta, path = self.open_remote("지워질.docx", b"x")
        path.unlink()
        self.mgr.note_change(str(path))
        self.assertTrue(wait_until(lambda: self.status_of(path) == "로컬 파일이 없습니다", 20))
        rec = self.mgr.get(sync.norm(path))
        self.assertIsNone(rec["due"])

    def test_candidate_rules(self):
        ok = ["보고서.hwp", "가정통신문(1).hwpx", "README", "사진.jpg", "2반 명렬표.xlsx"]
        bad = ["~$보고서.hwp", "$tm1.tmp", "0D840210", ".hidden", "a.bak", "a.TMP", "a.gde-part"]
        for n in ok:
            self.assertTrue(sync.is_new_file_candidate(n), n)
        for n in bad:
            self.assertFalse(sync.is_new_file_candidate(n), n)
        self.assertEqual(sync.guess_mime("a.hwp"), "application/x-hwp")
        self.assertEqual(sync.guess_mime("a.unknownext"), "application/octet-stream")


class TestPersistence(SyncTestBase):
    def test_restart_uploads_pending_changes(self):
        meta, path = self.open_remote("재시작.docx", b"base")
        self.mgr.stop()
        time.sleep(0.3)
        write_in_place(path, b"edited while app closed")
        self.mgr = self.make_manager()
        self.assertTrue(wait_until(lambda: self.uploads_of(meta["id"]), 6))
        self.assertEqual(self.drive.content[meta["id"]], b"edited while app closed")

    def test_restart_without_changes_no_upload(self):
        meta, path = self.open_remote("그대로.docx", b"base")
        self.mgr.stop()
        self.mgr = self.make_manager()
        self.assertTrue(wait_until(lambda: self.status_of(path) == "드라이브와 같음", 5))
        self.assertEqual(self.drive.uploads, [])

    def test_broken_state_file_is_tolerated(self):
        self.mgr.stop()
        self.state.write_text("{ broken", encoding="utf-8")
        self.mgr = self.make_manager()
        self.assertEqual(self.mgr.rows(), [])

    def test_leftover_parts_cleaned(self):
        self.mgr.stop()
        d = self.work / "abc"
        d.mkdir(parents=True)
        (d / "x.docx.gde-part").write_bytes(b"partial")
        self.mgr = self.make_manager()
        self.assertFalse((d / "x.docx.gde-part").exists())


class TestHelpers(unittest.TestCase):
    def test_safe_filename(self):
        f = sync.safe_filename
        self.assertEqual(f("회의록 2024/03/01.docx"), "회의록 2024_03_01.docx")
        self.assertEqual(f('a<b>c:d"e|f?g*h.hwp'), "a_b_c_d_e_f_g_h.hwp")
        self.assertEqual(f("끝에 점..."), "끝에 점")
        self.assertEqual(f("CON.docx"), "_CON.docx")
        self.assertEqual(f("nul"), "_nul")
        self.assertEqual(f("com1.tar.gz"), "_com1.tar.gz")
        self.assertEqual(f(""), "이름 없음")
        self.assertEqual(f("   "), "이름 없음")
        long = "가" * 300 + ".xlsx"
        out = f(long, 100)
        self.assertTrue(out.endswith(".xlsx"))
        self.assertLessEqual(len(out), 100)
        emoji = "📁" * 100 + ".docx"
        out = f(emoji, 50)
        self.assertLessEqual(len(out.encode("utf-16-le")) // 2, 50)
        self.assertTrue(out.endswith(".docx"))
        self.assertEqual(f("확장자 없음"), "확장자 없음")
        self.assertEqual(f("탭\t문자.txt"), "탭_문자.txt")

    def test_safe_filename_creatable(self):
        tmp = Path(tempfile.mkdtemp(prefix="gde_name_"))
        try:
            budget = 215 - len(str(tmp)) - 1  # 앱과 같은 방식으로 폴더 길이를 빼고 이름 길이를 정함
            for name in ["a/b\\c.docx", "CON", "aux.txt", "끝 .", "x" * 400 + ".hwp", "?*?", "LPT1.pdf"]:
                p = tmp / sync.safe_filename(name, budget)
                self.assertLessEqual(len(str(p)), 215)
                p.write_bytes(b"ok")
                self.assertTrue(p.exists(), name)
        finally:
            shutil.rmtree(tmp, ignore_errors=True)

    def test_is_office_temp(self):
        for n in ["~$문서.docx", "~WRL0001.tmp", "x.docx.gde-part", ".~lock.a.docx#"]:
            self.assertTrue(sync.is_office_temp(n), n)
        for n in ["문서.docx", "보고서.hwp", "~문서.hwp"]:
            self.assertFalse(sync.is_office_temp(n), n)

    def test_title_mentions(self):
        titles = ["가정통신문 (3월).docx - Word", "2반 명렬표 - Excel", "한글 테스트.hwp [C:\\x\\] - 한글"]
        self.assertTrue(sync.title_mentions(titles, "가정통신문 (3월).docx"))
        self.assertTrue(sync.title_mentions(titles, "2반 명렬표.xlsx"))      # 확장자 숨김 표시
        self.assertTrue(sync.title_mentions(titles, "한글 테스트.hwp"))
        self.assertFalse(sync.title_mentions(titles, "다른 파일.docx"))
        self.assertFalse(sync.title_mentions(["a - 메모장"], "a.txt"))        # 너무 짧은 이름은 확장자까지 일치해야
        self.assertTrue(sync.title_mentions(["a.txt - 메모장"], "a.txt"))

    def test_conflict_name(self):
        self.assertRegex(sync.conflict_name("보고서.hwp"), r"^보고서 \(충돌 사본 \d{4}-\d\d-\d\d \d{6}\)\.hwp$")
        self.assertRegex(sync.conflict_name("README"), r"^README \(충돌 사본 .+\)$")

    def test_concurrent_note_change_is_safe(self):
        tmp = Path(tempfile.mkdtemp(prefix="gde_conc_"))
        drive = FakeDrive()
        mgr = sync.SyncManager(tmp / "w", tmp / "s.json", tmp / "st", get_drive=lambda: drive,
                               debounce=0.2, poll=0.1)
        mgr.start()
        try:
            meta = drive.add_file("동시.docx", b"a")
            path = tmp / "w" / "d" / "동시.docx"
            mgr.add(path, drive.get_meta(meta["id"]), drive.download(meta, path), sync.KIND_EDIT)
            write_in_place(path, b"final")
            threads = [threading.Thread(target=lambda: [mgr.note_change(str(path)) for _ in range(200)])
                       for _ in range(8)]
            for t in threads:
                t.start()
            for t in threads:
                t.join()
            self.assertTrue(wait_until(lambda: drive.content[meta["id"]] == b"final", 5))
            time.sleep(0.5)
            self.assertEqual(len(drive.uploads), 1)
        finally:
            mgr.stop()
            time.sleep(0.1)
            shutil.rmtree(tmp, ignore_errors=True)


if __name__ == "__main__":
    unittest.main(verbosity=2)
