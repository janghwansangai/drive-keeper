"""테스트용 가짜 구글 드라이브 (DriveClient와 같은 메서드를 메모리로 흉내)."""
from __future__ import annotations

import hashlib
import json
import os
import threading
import time
import uuid
from pathlib import Path

import gtypes


class FakeResp:
    def __init__(self, status: int) -> None:
        self.status = status


class FakeHttpError(Exception):
    """googleapiclient.errors.HttpError 처럼 resp.status 와 content 를 가진다."""

    def __init__(self, status: int, reason: str = "") -> None:
        super().__init__(f"HTTP {status} {reason}")
        self.resp = FakeResp(status)
        self.content = json.dumps({"error": {"errors": [{"reason": reason}]}}).encode()


class FakeDrive:
    def __init__(self, email: str = "tester@example.com") -> None:
        self.email = email
        self.lock = threading.Lock()
        self.files: dict[str, dict] = {}
        self.content: dict[str, bytes] = {}
        self.uploads: list[tuple[str, str]] = []     # (file_id, md5)
        self.creates: list[tuple[str, str, str]] = []  # (name, parent, md5)
        self.fail_upload: list[BaseException] = []    # 다음 업로드 때 차례로 던질 예외
        self.fail_list: list[BaseException] = []
        self.list_delay: dict[str, float] = {}         # 폴더 id -> 목록 지연(초)
        self.upload_delay = 0.0
        self.download_count = 0
        self.relogin_count = 0
        self.get_meta_count = 0
        self.readonly_parents: set[str] = set()
        self.revisions: dict[str, list[dict]] = {}   # 파일 id -> 버전 목록 (오래된 것부터)
        self.add_file("내 폴더", None, file_id="folder1", mime=gtypes.FOLDER_MIME)

    # ----- 준비 -----
    def add_file(self, name: str, content: bytes | None, parent: str = "root", file_id: str | None = None,
                 mime: str = "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                 can_edit: bool = True, extra: dict | None = None) -> dict:
        file_id = file_id or uuid.uuid4().hex[:28]
        meta = {"id": file_id, "name": name, "mimeType": mime, "parents": [parent],
                "modifiedTime": self._now(), "capabilities": {"canEdit": can_edit, "canDownload": True},
                "webViewLink": f"https://example.invalid/{file_id}"}
        if extra:
            meta.update(extra)
        with self.lock:
            self.files[file_id] = meta
            if content is not None:
                self._set_content(file_id, content)
        return dict(meta)

    def remote_edit(self, file_id: str, content: bytes) -> None:
        """다른 컴퓨터에서 고친 것처럼 내용을 바꾼다."""
        with self.lock:
            self._set_content(file_id, content)

    def delete(self, file_id: str) -> None:
        with self.lock:
            self.files.pop(file_id, None)
            self.content.pop(file_id, None)

    @staticmethod
    def _now() -> str:
        t = time.time()
        return time.strftime("%Y-%m-%dT%H:%M:%S", time.gmtime(t)) + f".{int(t * 1000) % 1000:03d}Z"

    def _set_content(self, file_id: str, data: bytes) -> None:
        meta = self.files[file_id]
        self.content[file_id] = data
        if not gtypes.is_google_native(meta):
            meta["md5Checksum"] = hashlib.md5(data).hexdigest()
            meta["size"] = str(len(data))
        time.sleep(0.002)
        meta["modifiedTime"] = self._now()
        if not gtypes.is_google_native(meta):
            revs = self.revisions.setdefault(file_id, [])
            revs.append({"id": f"r{len(revs) + 1}", "modifiedTime": meta["modifiedTime"], "size": str(len(data)),
                         "md5Checksum": hashlib.md5(data).hexdigest(), "keepForever": False,
                         "lastModifyingUser": {"displayName": "테스트", "emailAddress": self.email},
                         "_data": data})

    # ----- DriveClient 인터페이스 -----
    def about_email(self) -> str:
        return self.email

    def relogin(self) -> None:
        self.relogin_count += 1

    def _children(self, pred) -> list[dict]:
        with self.lock:
            return [dict(m) for m in self.files.values() if pred(m) and not m.get("trashed")]

    def list_children(self, folder_id: str = "root", drive_id=None) -> list[dict]:
        delay = self.list_delay.get(folder_id, 0)
        if delay:
            time.sleep(delay)
        if self.fail_list:
            raise self.fail_list.pop(0)
        return self._children(lambda m: folder_id in (m.get("parents") or []))

    def list_shared_with_me(self) -> list[dict]:
        return self._children(lambda m: m.get("shared"))

    def search(self, text: str, limit: int = 2000) -> list[dict]:
        return self._children(lambda m: text.lower() in m["name"].lower())

    def list_drives(self) -> list[dict]:
        return []

    def get_meta(self, file_id: str) -> dict:
        self.get_meta_count += 1
        with self.lock:
            if file_id not in self.files:
                raise FakeHttpError(404, "notFound")
            return dict(self.files[file_id])

    def download(self, meta: dict, dest: Path, progress=None) -> str:
        with self.lock:
            if meta["id"] not in self.files:
                raise FakeHttpError(404, "notFound")
            data = self.content.get(meta["id"], b"")
            if gtypes.is_google_native(self.files[meta["id"]]):
                data = b"EXPORTED:" + data + str(time.time()).encode()  # 내보내기는 매번 달라짐
            self.download_count += 1
        dest = Path(dest)
        dest.parent.mkdir(parents=True, exist_ok=True)
        part = dest.with_name(dest.name + ".gde-part")
        with open(part, "wb") as f:
            f.write(data)
        if progress:
            progress(1.0)
        from drive_client import _replace
        try:
            _replace(part, dest)
        finally:
            if part.exists():
                part.unlink()
        return hashlib.md5(data).hexdigest()

    @staticmethod
    def _read(source) -> bytes:
        if isinstance(source, (bytes, bytearray)):
            return bytes(source)
        return Path(source).read_bytes()

    def update_content(self, file_id: str, source, mime: str) -> dict:
        if self.upload_delay:
            time.sleep(self.upload_delay)
        if self.fail_upload:
            raise self.fail_upload.pop(0)
        data = self._read(source)
        with self.lock:
            if file_id not in self.files:
                raise FakeHttpError(404, "notFound")
            self._set_content(file_id, data)
            self.uploads.append((file_id, hashlib.md5(data).hexdigest()))
            return dict(self.files[file_id])

    def create_file(self, name: str, parent_id: str, source, mime: str) -> dict:
        if self.fail_upload:
            raise self.fail_upload.pop(0)
        if parent_id in self.readonly_parents:
            raise FakeHttpError(403, "insufficientFilePermissions")
        data = self._read(source)
        meta = self.add_file(name, data, parent=parent_id, mime=mime)
        with self.lock:
            self.creates.append((name, parent_id, hashlib.md5(data).hexdigest()))
        return meta

    def find_in_folder(self, name: str, parent_id: str) -> list[dict]:
        return self._children(lambda m: m["name"] == name and parent_id in (m.get("parents") or []))

    def trash(self, file_id: str) -> None:
        with self.lock:
            self.files[file_id]["trashed"] = True
            self.files[file_id]["explicitlyTrashed"] = True

    # ----- 타임머신 -----
    def list_trash_page(self, page_token=None, name=None, page_size=500):
        with self.lock:
            items = [dict(m) for m in self.files.values() if m.get("trashed") and m.get("explicitlyTrashed", True)
                     and (not name or name.lower() in m["name"].lower())]
        items.sort(key=lambda m: m.get("modifiedTime", ""), reverse=True)
        start = int(page_token or 0)
        page = items[start:start + page_size]
        nxt = str(start + page_size) if start + page_size < len(items) else None
        return page, nxt

    def untrash(self, file_id: str) -> dict:
        with self.lock:
            if file_id not in self.files:
                raise FakeHttpError(404, "notFound")
            self.files[file_id]["trashed"] = False
            self.files[file_id].pop("explicitlyTrashed", None)
            return dict(self.files[file_id])

    def list_revisions(self, file_id: str) -> list[dict]:
        with self.lock:
            if file_id not in self.files:
                raise FakeHttpError(404, "notFound")
            return [{k: v for k, v in r.items() if k != "_data"} for r in self.revisions.get(file_id, [])]

    def download_revision(self, file_id: str, revision: dict, dest: Path) -> str:
        with self.lock:
            rev = next(r for r in self.revisions[file_id] if r["id"] == revision["id"])
            data = rev["_data"]
        dest = Path(dest)
        dest.parent.mkdir(parents=True, exist_ok=True)
        dest.write_bytes(data)
        return hashlib.md5(data).hexdigest()


def md5(data: bytes) -> str:
    return hashlib.md5(data).hexdigest()


def wait_until(cond, timeout: float = 10.0, interval: float = 0.05) -> bool:
    end = time.monotonic() + timeout
    while time.monotonic() < end:
        if cond():
            return True
        time.sleep(interval)
    return bool(cond())


def write_in_place(path: Path, data: bytes) -> None:
    with open(path, "r+b") as f:
        f.seek(0)
        f.write(data)
        f.truncate()
        f.flush()
        os.fsync(f.fileno())
