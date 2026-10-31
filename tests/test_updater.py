"""새 버전 확인·한 번에 업데이트 테스트 (인터넷 없이 가짜 응답으로)."""
from __future__ import annotations

import hashlib
import io
import json
import shutil
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import config  # noqa: E402
import updater  # noqa: E402

ZIP_URL = updater.DOWNLOAD_PREFIX + "v9.9.9/DriveKeeper.zip"


class FakeResponse(io.BytesIO):
    def __init__(self, data: bytes) -> None:
        super().__init__(data)
        self.headers = {"Content-Length": str(len(data))}

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        self.close()


def make_zip(names: dict[str, bytes]) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as z:
        for name, data in names.items():
            z.writestr(name, data)
    return buf.getvalue()


GOOD = {f"{updater.TOP}/app/DriveKeeper.exe": b"exe", f"{updater.TOP}/app/install.ps1": b"ps1",
        f"{updater.TOP}/설치하기.cmd": b"cmd"}


class UpdaterTest(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp = Path(tempfile.mkdtemp(prefix="gde_upd_"))
        self._base = config.BASE_DIR
        config.BASE_DIR = self.tmp
        self._urlopen = updater.urllib.request.urlopen
        self.served = b""
        updater.urllib.request.urlopen = lambda req, timeout=0: FakeResponse(self.served)

    def tearDown(self) -> None:
        updater.urllib.request.urlopen = self._urlopen
        config.BASE_DIR = self._base
        shutil.rmtree(self.tmp, ignore_errors=True)

    def test_is_newer(self):
        self.assertTrue(updater.is_newer("2.1.0", "2.0.1"))
        self.assertFalse(updater.is_newer("v2.0.1", "2.0.1"))
        self.assertTrue(updater.is_newer("10.0.0", "9.9.9"))

    def test_check_latest_reads_zip_asset_and_digest(self):
        self.served = json.dumps({"tag_name": "v9.9.9", "html_url": "https://example.com/rel", "body": "바뀐 점",
                                  "assets": [{"name": "other.zip", "browser_download_url": ZIP_URL + "x"},
                                             {"name": "DriveKeeper.zip", "browser_download_url": ZIP_URL,
                                              "size": 123, "digest": "sha256:ABCDEF"}]}).encode()
        info = updater.check_latest()
        self.assertEqual(info["version"], "9.9.9")
        self.assertEqual(info["zip_url"], ZIP_URL)
        self.assertEqual(info["zip_size"], 123)
        self.assertEqual(info["sha256"], "abcdef")

    def test_check_latest_ignores_zip_outside_this_repo(self):
        self.served = json.dumps({"tag_name": "v9.9.9", "assets": [
            {"name": "DriveKeeper.zip", "browser_download_url": "https://evil.example.com/DriveKeeper.zip"}]}).encode()
        self.assertEqual(updater.check_latest()["zip_url"], "")

    def test_download_checks_hash_and_extracts(self):
        self.served = make_zip(GOOD)
        info = {"zip_url": ZIP_URL, "sha256": hashlib.sha256(self.served).hexdigest()}
        seen = []
        ps1 = updater.download(info, lambda got, total: seen.append((got, total)))
        self.assertEqual(ps1.read_bytes(), b"ps1")
        self.assertTrue(ps1.is_relative_to(self.tmp / "update"))
        self.assertEqual(seen[-1], (len(self.served), len(self.served)))

    def test_download_refuses_wrong_hash(self):
        self.served = make_zip(GOOD)
        with self.assertRaises(updater.UpdateError):
            updater.download({"zip_url": ZIP_URL, "sha256": "0" * 64})

    def test_download_refuses_other_zip_or_bad_paths(self):
        for names in ({f"{updater.TOP}/app/install.ps1": b"x"},                 # exe 없음
                      dict(GOOD, **{"../evil.txt": b"x"})):                      # 폴더 밖으로 쓰기
            self.served = make_zip(names)
            with self.assertRaises(updater.UpdateError):
                updater.download({"zip_url": ZIP_URL})
        self.assertFalse((self.tmp / "evil.txt").exists())

    def test_download_refuses_url_outside_this_repo(self):
        with self.assertRaises(updater.UpdateError):
            updater.download({"zip_url": "https://evil.example.com/DriveKeeper.zip"})

    def test_not_installed_copy_cannot_self_update(self):
        self.assertIsNone(updater.install_dir())     # 개발용 실행 (exe 아님)
        self.assertFalse(updater.can_self_update({"zip_url": ZIP_URL}))

    def test_finished_message(self):
        self.assertIsNone(updater.finished_message())
        root = updater.work_dir()
        root.mkdir(parents=True)
        (root / "install.log").write_text("exit=0\n", encoding="utf-8-sig")
        self.assertTrue(updater.finished_message().startswith("새 버전"))
        self.assertFalse(root.exists())             # 받아 둔 파일은 지움
        root.mkdir(parents=True)
        (root / "install.log").write_text("exit=1\nERROR: 파일 복사 실패\n", encoding="utf-8-sig")
        self.assertIn("파일 복사 실패", updater.finished_message())


if __name__ == "__main__":
    unittest.main(verbosity=2)
