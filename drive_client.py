"""Google Drive API 래퍼: 로그인(토큰 유지), 목록, 다운로드, 업로드.

- 연결(http) 묶음을 돌려 쓴다. 호출 하나가 연결 하나를 빌려 쓰고 돌려준다.
  httplib2는 스레드 안전하지 않아 동시에 같은 연결을 쓰면 안 되고, 하나의 잠금으로 묶으면
  큰 파일 업로드 중에 폴더 목록이 멈추기 때문. 이미 열린 연결을 재사용해 TLS 연결 시간(약 0.4초)을 아낀다.
- 모든 호출은 네트워크 오류·서버 과부하 시 자동 재시도(num_retries)한다.
- 로그인이 만료되면(테스트 상태 앱은 7일마다) 브라우저 로그인을 한 번 띄우고 이어서 진행한다.
- 계정마다 DriveClient 하나. 다시 로그인했을 때 다른 계정으로 로그인했으면 거부한다 (파일이 섞이는 사고 방지).
- 보안: 이 프로그램이 보낼 수 있는 요청을 '읽기'와 '파일 내용 올리기/새 파일 만들기'로만 제한한다.
  삭제·휴지통 이동·휴지통 비우기·공유 설정 변경 요청은 통신 단계에서 막는다 (_GuardedHttp).
"""
from __future__ import annotations

import hashlib
import io
import json
import logging
import os
import re
import threading
import time
from pathlib import Path
from typing import Callable
from urllib.parse import urlsplit

import google_auth_httplib2
import httplib2
from google.auth.exceptions import RefreshError
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow

try:
    from google_auth_oauthlib.flow import WSGITimeoutError
except ImportError:  # 옛 버전
    class WSGITimeoutError(AttributeError):
        pass
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload

import config
import gtypes
import protect
from errors import AuthError, FileLockedError, IntegrityError, is_network_error
from gtypes import FOLDER_MIME, GOOGLE_EXPORT
from sync import PART_SUFFIX   # 내려받는 중인 파일 (끝나면 원래 이름으로 바꿈)

log = logging.getLogger(__name__)

# 목록에 편집 권한까지 받아 두면 파일을 열 때 정보 요청 없이 바로 내려받을 수 있다
LIST_FIELDS = ("id,name,mimeType,modifiedTime,size,md5Checksum,driveId,parents,webViewLink,"
               "shortcutDetails(targetId,targetMimeType),capabilities(canEdit,canDownload)")
META_FIELDS = LIST_FIELDS + ",trashed"
POOL_MAX = 6

RETRIES = 4                       # 네트워크/서버 오류 시 자동 재시도 횟수 (지수 백오프)
CHUNK = 8 * 1024 * 1024           # 다운로드/업로드 조각 크기
SIMPLE_UPLOAD_LIMIT = 5 * 1024 * 1024   # 이하면 요청 1번으로 끝나는 단순 업로드

LOGIN_DONE_HTML = (
    "<html><head><meta charset='utf-8'><title>로그인 완료</title></head>"
    "<body style='font-family:Malgun Gothic,sans-serif;padding:40px'>"
    "<h2>로그인이 완료되었습니다.</h2><p>이 창을 닫고 드라이브 지킴이로 돌아가세요.</p>"
    "</body></html>")


# 브라우저 로그인은 한 번에 하나씩 (계정 여러 개가 동시에 만료돼도 창이 겹치지 않게)
_BROWSER_LOCK = threading.Lock()


class BlockedRequest(Exception):
    """보안상 이 프로그램이 보내지 않는 종류의 요청."""


def _header(headers: dict | None, name: str) -> str:
    for k, v in (headers or {}).items():
        if k.lower() == name:
            return str(v)
    return ""


def _metadata_part(body, headers) -> bytes | None:
    """업로드 요청에서 파일 정보(JSON) 부분만 꺼낸다. 파일 내용만 보내는 요청이면 None."""
    if body is None:
        return None
    data = body if isinstance(body, (bytes, bytearray)) else str(body).encode("utf-8", "replace")
    ctype = _header(headers, "content-type").lower()
    if ctype.startswith("application/json"):
        return bytes(data)
    if ctype.startswith("multipart/related") and "boundary=" in ctype:
        boundary = _header(headers, "content-type").split("boundary=", 1)[1].strip().strip('"')
        parts = bytes(data).split(b"--" + boundary.encode("ascii", "replace"))
        return parts[1] if len(parts) > 1 else bytes(data[:4096])
    return None


_FILE_PATH = re.compile(r"^/drive/v3/files/[A-Za-z0-9_-]+$")


def _is_untrash(body) -> bool:
    """본문이 정확히 {"trashed": false} 인가 (휴지통에서 꺼내기만 허용)."""
    if body is None:
        return False
    try:
        data = json.loads(body if isinstance(body, str) else bytes(body).decode("utf-8"))
    except Exception:
        return False
    return data == {"trashed": False}


def check_request(uri: str, method: str, body=None, headers=None) -> None:
    """허용 목록: 읽기(GET), 파일 올리기·새 파일 만들기(/upload/...), 휴지통에서 꺼내기.
    삭제·휴지통 이동·휴지통 비우기·공유 변경 등 나머지는 모두 막는다."""
    m = (method or "GET").upper()
    path = urlsplit(uri).path
    if m == "GET":
        return
    if m in ("POST", "PATCH", "PUT") and path.startswith("/upload/drive/v3/files"):
        meta = _metadata_part(body, headers)
        if meta is not None and b'"trashed"' in meta:
            raise BlockedRequest("보안상 허용되지 않은 요청: 휴지통 이동")
        return
    if m == "PATCH" and _FILE_PATH.match(path) and _is_untrash(body):
        return  # 타임머신: 지운 파일을 휴지통에서 되살리기 (지우는 방향은 불가)
    raise BlockedRequest(f"보안상 허용되지 않은 요청: {m} {path}")


class _GuardedHttp:
    """이 프로그램이 보내는 모든 드라이브 요청이 지나가는 문. check_request 를 통과해야만 나간다."""

    def __init__(self, inner) -> None:
        self._inner = inner

    def __getattr__(self, name):
        return getattr(self._inner, name)

    def request(self, uri, method="GET", body=None, headers=None, *args, **kwargs):
        check_request(uri, method, body, headers)
        return self._inner.request(uri, method, body, headers, *args, **kwargs)


def _client_config() -> dict:
    secret = config.find_client_secret()
    if not secret:
        raise FileNotFoundError(
            "앱 열쇠(client_secret.json)가 아직 없어요.\n"
            "위쪽 [도움말] → [시작 도우미]에서 내 구글 열쇠를 만들고 불러와 주세요.")
    try:
        info = json.loads(secret.read_text(encoding="utf-8"))
    except (OSError, ValueError) as e:
        raise AuthError(f"client_secret.json을 읽을 수 없습니다: {e}") from e
    if "installed" not in info:
        raise AuthError(
            "client_secret.json이 '데스크톱 앱' 유형이 아닙니다.\n"
            "Google Cloud 콘솔 → 클라이언트 만들기 → 애플리케이션 유형 '데스크톱 앱'으로 다시 만들어 주세요.")
    return info


def browser_login(on_browser_login: Callable[[], None] | None = None, login_hint: str | None = None,
                  select_account: bool = False) -> Credentials:
    """브라우저로 구글 로그인. 한 번에 하나씩만 진행한다."""
    info = _client_config()
    with _BROWSER_LOCK:
        if on_browser_login:
            try:
                on_browser_login()
            except Exception:
                log.exception("on_browser_login 콜백 오류")
        flow = InstalledAppFlow.from_client_config(info, config.SCOPES)
        kwargs = {"access_type": "offline",
                  # offline + consent 로 받아야 refresh_token이 나와서 다음부터 자동 로그인된다
                  "prompt": "consent select_account" if select_account else "consent"}
        if login_hint:
            kwargs["login_hint"] = login_hint
        try:
            return flow.run_local_server(
                port=0, timeout_seconds=config.LOGIN_TIMEOUT_SEC, success_message=LOGIN_DONE_HTML,
                authorization_prompt_message="", **kwargs)
        except WSGITimeoutError as e:
            raise AuthError("로그인 시간이 초과되었습니다. '다시 로그인'을 눌러 주세요.") from e
        except Exception as e:
            if "access_denied" in str(e):
                raise AuthError(
                    "로그인이 취소되었거나 차단되었습니다.\n"
                    "다른 계정이라면 Google Cloud 콘솔의 '테스트 사용자'에 그 계정을 추가해야 합니다 "
                    "(README '다른 계정 추가').") from e
            raise


def fetch_email(creds: Credentials) -> str:
    http = _GuardedHttp(google_auth_httplib2.AuthorizedHttp(
        creds, http=httplib2.Http(timeout=config.HTTP_TIMEOUT_SEC)))
    svc = build("drive", "v3", http=http, cache_discovery=False, static_discovery=True)
    resp = svc.about().get(fields="user(emailAddress)").execute(num_retries=RETRIES)
    return (resp.get("user") or {}).get("emailAddress", "")


def login_new_account(on_browser_login: Callable[[], None] | None = None) -> tuple[str, str]:
    """'계정 추가'·본인 확인: 브라우저에서 계정을 골라 로그인 → (로그인 정보 JSON, 이메일)."""
    creds = browser_login(on_browser_login, select_account=True)
    email = fetch_email(creds)
    if not email:
        raise AuthError("계정 이메일을 확인할 수 없습니다.")
    return creds.to_json(), email


def revoke_token(creds_json: str) -> None:
    """계정을 제거할 때 구글 쪽 접근 권한도 취소한다 (실패해도 무시)."""
    try:
        import requests
        info = json.loads(creds_json)
        token = info.get("refresh_token") or info.get("token")
        if token:
            requests.post("https://oauth2.googleapis.com/revoke", data={"token": token},
                          headers={"content-type": "application/x-www-form-urlencoded"}, timeout=15)
    except Exception as e:
        log.warning("토큰 취소 실패(무시): %r", e)


class _HashingWriter:
    """내려받으면서 MD5를 같이 계산한다 (다 받은 뒤 파일을 다시 읽지 않아도 됨)."""

    def __init__(self, fh, h):
        self._fh = fh
        self._h = h

    def write(self, data):
        self._h.update(data)
        return self._fh.write(data)


class DriveClient:
    def __init__(self, token_path: Path | str | None = None, expected_email: str | None = None,
                 on_browser_login: Callable[[], None] | None = None, interactive: bool = True) -> None:
        config.ensure_dirs()
        self.token_path = Path(token_path) if token_path else config.TOKEN_PATH
        self.expected_email = expected_email or None
        self._on_browser_login = on_browser_login
        self._pool: list[tuple[int, object]] = []   # (로그인 세대, 서비스) - 열린 연결 재사용
        self._pool_lock = threading.Lock()
        self._auth_lock = threading.Lock()
        self._gen = 0                  # 로그인이 바뀔 때마다 증가 → 스레드별 연결 다시 만들기
        self._auth_broken = False      # 재로그인 실패 → 사용자가 '다시 로그인'을 누를 때까지 브라우저를 또 띄우지 않음
        self.creds = self._load_credentials(interactive)

    # ---------- 로그인 ----------
    def _save(self, creds: Credentials) -> None:
        protect.write_token(self.token_path, creds.to_json())

    def _load_credentials(self, interactive: bool) -> Credentials:
        creds = None
        if self.token_path.exists():
            try:
                creds_json, _lock = protect.read_token(self.token_path)
                creds = Credentials.from_authorized_user_info(json.loads(creds_json), config.SCOPES)
                if not creds.has_scopes(config.SCOPES):
                    creds = None
            except Exception as e:  # 깨졌거나 다른 PC에서 만든 토큰 → 새로 로그인
                log.warning("토큰을 읽을 수 없어 다시 로그인합니다: %s", e)
                creds = None

        if creds and creds.valid:
            return creds

        if creds and creds.refresh_token:
            # 저장된 토큰으로 자동 갱신 → 다시 로그인할 필요 없음
            # (네트워크 오류는 그대로 올려보낸다: 인터넷 문제로 브라우저를 띄우면 안 됨)
            try:
                creds.refresh(Request())
                self._save(creds)
                return creds
            except RefreshError as e:
                # 앱이 '테스트' 상태면 7일마다 토큰이 만료된다 → 브라우저 로그인으로 다시 받는다
                log.info("토큰 갱신 실패, 브라우저 로그인으로 전환: %s", e)

        if not interactive:
            raise AuthError("로그인이 필요합니다.")
        creds = self._browser_login()
        self._save(creds)
        return creds

    def _browser_login(self) -> Credentials:
        creds = browser_login(self._on_browser_login, login_hint=self.expected_email)
        if self.expected_email:
            # 다른 계정으로 로그인했으면 받아들이지 않는다 (이 계정의 파일을 다른 계정으로 올리는 사고 방지)
            actual = fetch_email(creds)
            if actual.casefold() != self.expected_email.casefold():
                raise AuthError(f"{actual} 계정으로 로그인했습니다.\n"
                                f"{self.expected_email} 계정으로 다시 로그인해 주세요.")
        return creds

    def relogin(self) -> None:
        """사용자가 '다시 로그인'을 눌렀을 때: 브라우저 로그인을 강제로 다시 한다."""
        with self._auth_lock:
            creds = self._browser_login()
            self._save(creds)
            self.creds = creds
            self._gen += 1
            self._auth_broken = False

    def _relogin_after_expiry(self, seen_gen: int) -> None:
        with self._auth_lock:
            if self._gen != seen_gen:
                return  # 다른 스레드가 이미 다시 로그인함
            if self._auth_broken:
                raise AuthError("로그인이 필요합니다. 아래 '다시 로그인'을 눌러 주세요.")
            try:
                creds = self._browser_login()
                self._save(creds)
            except Exception as e:
                self._auth_broken = True
                if isinstance(e, AuthError):
                    raise
                raise AuthError(f"다시 로그인하지 못했습니다: {e}") from e
            self.creds = creds
            self._gen += 1

    @property
    def auth_broken(self) -> bool:
        return self._auth_broken

    # ---------- 호출 공통 ----------
    def _borrow(self) -> tuple[int, object]:
        with self._pool_lock:
            while self._pool:
                gen, svc = self._pool.pop()
                if gen == self._gen:
                    return gen, svc
            gen = self._gen
        http = _GuardedHttp(google_auth_httplib2.AuthorizedHttp(
            self.creds, http=httplib2.Http(timeout=config.HTTP_TIMEOUT_SEC)))
        # static_discovery: 설명 문서를 패키지에 든 파일에서 읽음 (네트워크 요청 없음)
        return gen, build("drive", "v3", http=http, cache_discovery=False, static_discovery=True)

    def _give_back(self, gen: int, svc) -> None:
        with self._pool_lock:
            if gen == self._gen and len(self._pool) < POOL_MAX:
                self._pool.append((gen, svc))

    def _call(self, fn):
        gen, svc = self._borrow()
        keep = True
        try:
            return fn(svc)
        except BaseException as e:
            keep = not (is_network_error(e) or isinstance(e, RefreshError))  # 끊긴 연결은 버림
            raise
        finally:
            if keep:
                self._give_back(gen, svc)

    def _run(self, fn):
        """fn(service)를 실행한다. 로그인이 만료됐으면 한 번 다시 로그인하고 재시도."""
        if self._auth_broken:
            raise AuthError("로그인이 필요합니다. 아래 '다시 로그인'을 눌러 주세요.")
        gen = self._gen
        try:
            return self._call(fn)
        except RefreshError:
            log.info("API 호출 중 로그인 만료 감지 → 재로그인")
            self._relogin_after_expiry(gen)
            return self._call(fn)

    # 예전 코드와의 호환용
    is_folder = staticmethod(gtypes.is_folder)
    is_google_native = staticmethod(gtypes.is_google_native)

    # ---------- 계정 ----------
    def about_email(self) -> str:
        resp = self._run(lambda s: s.about().get(fields="user(emailAddress)").execute(num_retries=RETRIES))
        return (resp.get("user") or {}).get("emailAddress", "")

    # ---------- 목록 (파일을 내려받지 않고 정보만 가져옴) ----------
    def _list_files(self, limit: int | None = None, **kw) -> list[dict]:
        items: list[dict] = []
        page_token = None
        while True:
            token = page_token
            resp = self._run(lambda s: s.files().list(
                pageToken=token, pageSize=1000, fields=f"nextPageToken, files({LIST_FIELDS})",
                supportsAllDrives=True, includeItemsFromAllDrives=True, **kw,
            ).execute(num_retries=RETRIES))
            items.extend(resp.get("files", []))
            page_token = resp.get("nextPageToken")
            if not page_token or (limit and len(items) >= limit):
                break
        return items

    def list_children(self, folder_id: str = "root", drive_id: str | None = None) -> list[dict]:
        kw = {"q": f"'{_q(folder_id)}' in parents and trashed = false"}
        if drive_id:
            kw.update(corpora="drive", driveId=drive_id)
        return self._list_files(**kw)

    def list_shared_with_me(self) -> list[dict]:
        return self._list_files(q="sharedWithMe = true and trashed = false")

    def search(self, text: str, limit: int = 2000) -> list[dict]:
        q = f"name contains '{_q(text)}' and trashed = false"
        return self._list_files(limit=limit, q=q)

    def list_drives(self) -> list[dict]:
        """공유 드라이브 목록 (폴더처럼 보여줌)."""
        items: list[dict] = []
        page_token = None
        while True:
            token = page_token
            resp = self._run(lambda s: s.drives().list(
                pageToken=token, pageSize=100, fields="nextPageToken, drives(id,name)",
            ).execute(num_retries=RETRIES))
            for d in resp.get("drives", []):
                items.append({"id": d["id"], "name": d.get("name", ""), "mimeType": FOLDER_MIME,
                              "driveId": d["id"]})
            page_token = resp.get("nextPageToken")
            if not page_token:
                break
        return items

    def get_meta(self, file_id: str) -> dict:
        return self._run(lambda s: s.files().get(
            fileId=file_id, fields=META_FIELDS, supportsAllDrives=True,
        ).execute(num_retries=RETRIES))

    # ---------- 다운로드 (선택한 파일 하나만) ----------
    def download(self, meta: dict, dest: Path,
                 progress: Callable[[float], None] | None = None) -> str:
        """meta 파일을 dest로 내려받고 내용의 MD5를 돌려준다.

        임시 이름(.gde-part)으로 받은 뒤 검사값을 확인하고 원래 이름으로 바꾼다.
        그래서 받다가 끊겨도 반쯤 받은 파일이 열리는 일이 없다.
        """
        mime = meta["mimeType"]
        file_id = meta["id"]
        if mime in GOOGLE_EXPORT:
            export_mime = GOOGLE_EXPORT[mime][0]

            def make(s):
                return s.files().export_media(fileId=file_id, mimeType=export_mime)
            expected = None  # 내보내기 결과는 매번 달라서 검사값이 없음
        else:
            def make(s):
                return s.files().get_media(fileId=file_id, supportsAllDrives=True)
            expected = meta.get("md5Checksum")

        dest = Path(dest)
        dest.parent.mkdir(parents=True, exist_ok=True)
        part = dest.with_name(dest.name + PART_SUFFIX)

        def work(s):
            h = hashlib.md5()
            with open(part, "wb") as fh:
                downloader = MediaIoBaseDownload(_HashingWriter(fh, h), make(s), chunksize=CHUNK)
                done = False
                while not done:
                    status, done = downloader.next_chunk(num_retries=RETRIES)
                    if progress and status is not None:
                        try:
                            progress(status.progress())
                        except Exception:
                            pass
                fh.flush()
                os.fsync(fh.fileno())
            return h.hexdigest()

        try:
            md5 = None
            for attempt in range(3):
                md5 = self._run(work)
                if not expected or md5 == expected:
                    break
                # 목록을 받은 뒤에 파일이 바뀌었을 수 있음 → 최신 검사값으로 다시 비교
                latest = self.get_meta(file_id).get("md5Checksum")
                if latest == md5:
                    break
                log.warning("다운로드 MD5 불일치 (%s): %s != %s", meta.get("name"), md5, latest)
                expected = latest
                if attempt == 2:
                    raise IntegrityError("내려받은 파일이 손상되었습니다. 다시 시도해 주세요.")
            _replace(part, dest)
            return md5
        finally:
            try:
                part.unlink()
            except OSError:
                pass

    # ---------- 업로드 ----------
    @staticmethod
    def _media(source: bytes | Path, mime: str, fh=None):
        mime = mime or "application/octet-stream"
        if isinstance(source, (bytes, bytearray)):
            return MediaIoBaseUpload(io.BytesIO(source), mimetype=mime, chunksize=CHUNK,
                                     resumable=len(source) > SIMPLE_UPLOAD_LIMIT)
        return MediaIoBaseUpload(fh, mimetype=mime, chunksize=CHUNK, resumable=True)

    def _with_media(self, source: bytes | Path, mime: str, call):
        def work(s):
            if isinstance(source, (bytes, bytearray)):
                return call(s, self._media(source, mime))
            with open(source, "rb") as fh:  # 큰 파일: 임시 사본에서 조금씩 읽어 올림
                return call(s, self._media(source, mime, fh))
        return self._run(work)

    def update_content(self, file_id: str, source: bytes | Path, mime: str) -> dict:
        """수정한 내용을 원래 파일에 덮어쓴다. 새 메타데이터(md5Checksum 등)를 돌려준다."""
        return self._with_media(source, mime, lambda s, media: s.files().update(
            fileId=file_id, media_body=media, fields=META_FIELDS, supportsAllDrives=True,
        ).execute(num_retries=RETRIES))

    def create_file(self, name: str, parent_id: str, source: bytes | Path, mime: str) -> dict:
        body = {"name": name, "parents": [parent_id]}
        return self._with_media(source, mime, lambda s, media: s.files().create(
            body=body, media_body=media, fields=META_FIELDS, supportsAllDrives=True,
        ).execute(num_retries=RETRIES))

    def find_in_folder(self, name: str, parent_id: str) -> list[dict]:
        """폴더 안에서 이름이 정확히 같은 파일 (중복 생성 방지 확인용)."""
        q = f"name = '{_q(name)}' and '{_q(parent_id)}' in parents and trashed = false"
        return self._list_files(q=q)

    # ---------- 타임머신: 휴지통·이전 버전 ----------
    def list_trash_page(self, page_token: str | None = None, name: str | None = None,
                        page_size: int = 500) -> tuple[list[dict], str | None]:
        """휴지통 한 쪽: (직접 지운 항목들, 다음 쪽 표시). 최근에 고친 것부터.

        휴지통에는 지운 폴더 안의 파일까지 하나하나 들어 있어 수만 개일 수 있다.
        그래서 한꺼번에 다 읽지 않고 쪽마다 읽고, 이름으로 찾을 때는 드라이브에서 바로 찾는다.
        """
        q = "trashed = true"
        if name:
            q += f" and name contains '{_q(name)}'"
        fields = f"nextPageToken, files({LIST_FIELDS},trashed,trashedTime,explicitlyTrashed)"
        resp = self._run(lambda s: s.files().list(
            q=q, pageToken=page_token, pageSize=page_size, fields=fields, orderBy="modifiedTime desc",
            supportsAllDrives=True, includeItemsFromAllDrives=True,
        ).execute(num_retries=RETRIES))
        items = [m for m in resp.get("files", []) if m.get("explicitlyTrashed", True)]
        return items, resp.get("nextPageToken")

    def untrash(self, file_id: str) -> dict:
        """휴지통에서 꺼내 원래 폴더로 되돌린다."""
        return self._run(lambda s: s.files().update(
            fileId=file_id, body={"trashed": False}, fields=META_FIELDS, supportsAllDrives=True,
        ).execute(num_retries=RETRIES))

    def list_revisions(self, file_id: str) -> list[dict]:
        """이전 버전 목록 (오래된 것 → 최신). 구글은 일반 파일의 이전 버전을 30일·최대 100개 보관."""
        fields = ("nextPageToken, revisions(id,modifiedTime,size,md5Checksum,keepForever,"
                  "originalFilename,lastModifyingUser(displayName,emailAddress))")
        revs: list[dict] = []
        page_token = None
        while True:
            token = page_token
            resp = self._run(lambda s: s.revisions().list(
                fileId=file_id, pageToken=token, pageSize=200, fields=fields,
            ).execute(num_retries=RETRIES))
            revs.extend(resp.get("revisions", []))
            page_token = resp.get("nextPageToken")
            if not page_token:
                break
        return revs

    def download_revision(self, file_id: str, revision: dict, dest: Path) -> str:
        """이전 버전 하나를 dest로 내려받고 MD5를 돌려준다 (검사값 확인)."""
        dest = Path(dest)
        dest.parent.mkdir(parents=True, exist_ok=True)
        part = dest.with_name(dest.name + PART_SUFFIX)
        rev_id = revision["id"]

        def work(s):
            h = hashlib.md5()
            with open(part, "wb") as fh:
                req = s.revisions().get_media(fileId=file_id, revisionId=rev_id)
                downloader = MediaIoBaseDownload(_HashingWriter(fh, h), req, chunksize=CHUNK)
                done = False
                while not done:
                    _status, done = downloader.next_chunk(num_retries=RETRIES)
                fh.flush()
                os.fsync(fh.fileno())
            return h.hexdigest()

        try:
            md5 = self._run(work)
            expected = revision.get("md5Checksum")
            if expected and md5 != expected:
                raise IntegrityError("이전 버전을 받는 중 파일이 손상되었습니다. 다시 시도해 주세요.")
            _replace(part, dest)
            return md5
        finally:
            try:
                part.unlink()
            except OSError:
                pass

    # 삭제·휴지통으로 보내는 기능은 일부러 두지 않는다. (테스트 정리는 tests/drive_admin.py 가 따로 함)


def _q(value: str) -> str:
    """검색어 안의 따옴표·역슬래시가 쿼리를 깨지 않게 처리."""
    return value.replace("\\", "\\\\").replace("'", "\\'")


def _replace(src: Path, dest: Path) -> None:
    """src를 dest로 바꿔치기. 검사 프로그램이 잠깐 잡고 있으면 조금 기다렸다가 다시 시도."""
    for attempt in range(6):
        try:
            os.replace(src, dest)
            return
        except PermissionError as e:
            if attempt == 5:
                raise FileLockedError(
                    f"'{dest.name}' 파일이 다른 프로그램에서 열려 있어 새 내용으로 바꾸지 못했습니다.") from e
            time.sleep(0.05 * (2 ** attempt))

