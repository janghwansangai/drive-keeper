"""오류 분류와 한국어 안내 문구.

구글 라이브러리를 import하지 않고 속성으로만 판별한다. (창을 먼저 빨리 띄우기 위해)
"""
from __future__ import annotations

import json
import socket
import ssl


class AuthError(Exception):
    """로그인이 필요하다 (토큰 만료·취소, 브라우저 로그인 실패/시간 초과)."""


class FileLockedError(Exception):
    """작업 폴더의 파일이 다른 프로그램(Word/Excel/한글)에서 열려 있어 바꿀 수 없다."""


class IntegrityError(Exception):
    """내려받거나 올린 내용의 검사값(MD5)이 드라이브와 다르다."""


def http_status(e: BaseException) -> int | None:
    resp = getattr(e, "resp", None)
    status = getattr(resp, "status", None)
    try:
        return int(status) if status is not None else None
    except (TypeError, ValueError):
        return None


def http_reason(e: BaseException) -> str:
    """HttpError 본문의 reason (예: rateLimitExceeded, exportSizeLimitExceeded)."""
    content = getattr(e, "content", b"")
    try:
        data = json.loads(content.decode("utf-8") if isinstance(content, bytes) else content)
        err = data.get("error", {})
        errors = err.get("errors") or []
        if errors and errors[0].get("reason"):
            return str(errors[0]["reason"])
        details = err.get("details") or []
        for d in details:
            if d.get("reason"):
                return str(d["reason"])
        return str(err.get("status", ""))
    except Exception:
        return ""


def _is_named(e: BaseException, *names: str) -> bool:
    return any(cls.__name__ in names for cls in type(e).__mro__)


def is_auth_error(e: BaseException) -> bool:
    return isinstance(e, AuthError) or _is_named(e, "RefreshError")


def is_network_error(e: BaseException) -> bool:
    if isinstance(e, (socket.timeout, TimeoutError, ConnectionError, ssl.SSLError, socket.gaierror)):
        return True
    if _is_named(e, "ServerNotFoundError", "TransportError", "RedirectMissingLocation",
                 "ConnectionError", "Timeout", "ReadTimeout", "ConnectTimeout", "ProxyError"):
        return True
    # httplib2/소켓 계층의 OSError (WinError 10054/10060/10065 등)
    return isinstance(e, OSError) and getattr(e, "winerror", None) in (
        10050, 10051, 10052, 10053, 10054, 10060, 10061, 10064, 10065, 11001, 11002, 11004)


def is_not_found(e: BaseException) -> bool:
    return http_status(e) == 404


def is_transient(e: BaseException) -> bool:
    """잠시 후 다시 하면 될 가능성이 높은 오류인가 (네트워크, 서버 과부하)."""
    if is_network_error(e) or isinstance(e, IntegrityError):
        return True
    status = http_status(e)
    if status is None:
        return False
    if status in (408, 429) or status >= 500:
        return True
    if status == 403 and http_reason(e) in ("rateLimitExceeded", "userRateLimitExceeded",
                                            "sharingRateLimitExceeded", "backendError"):
        return True
    return False


def friendly_error(e: BaseException) -> str:
    if isinstance(e, AuthError):
        return str(e) or "로그인이 필요합니다."
    if _is_named(e, "RefreshError"):
        return "로그인이 만료되었습니다. 다시 로그인해 주세요."
    if isinstance(e, FileLockedError):
        return str(e) or "파일이 다른 프로그램에서 열려 있습니다."
    if isinstance(e, IntegrityError):
        return str(e) or "전송 중 파일이 손상되었습니다."
    if isinstance(e, FileNotFoundError) and "client_secret" in str(e):
        return str(e)
    if is_network_error(e):
        return "인터넷 연결을 확인해 주세요. (구글 서버에 연결할 수 없음)"
    status = http_status(e)
    if status is not None:
        reason = http_reason(e)
        if status == 404:
            return "드라이브에서 찾을 수 없습니다. (삭제되었거나 접근 권한이 없음)"
        if status == 403:
            if reason == "exportSizeLimitExceeded":
                return "Google 문서가 너무 커서(10MB 초과) 오피스 파일로 내보낼 수 없습니다. 브라우저에서 여세요."
            if reason in ("cannotDownloadFile", "cannotCopyFile"):
                return "소유자가 다운로드를 막아 둔 파일입니다. 브라우저에서 여세요."
            if reason in ("insufficientFilePermissions", "insufficientPermissions", "forbidden"):
                return "이 파일을 수정할 권한이 없습니다. (보기 전용)"
            if reason == "storageQuotaExceeded":
                return "구글 드라이브 저장 공간이 가득 찼습니다."
            if reason == "fileNotDownloadable":
                return "이 형식은 내려받을 수 없습니다. 브라우저에서 여세요."
            if reason in ("rateLimitExceeded", "userRateLimitExceeded"):
                return "요청이 너무 많습니다. 잠시 후 다시 시도해 주세요."
            if reason == "accessNotConfigured":
                return "Google Cloud 프로젝트에서 Google Drive API가 꺼져 있습니다. (README 2단계)"
            return f"접근이 거부되었습니다. ({reason or 403})"
        if status == 401:
            return "로그인이 만료되었습니다. 다시 로그인해 주세요."
        if status == 429 or status >= 500:
            return "구글 서버가 바쁩니다. 잠시 후 다시 시도해 주세요."
        return f"구글 드라이브 오류 {status} ({reason})" if reason else f"구글 드라이브 오류 {status}"
    if isinstance(e, PermissionError):
        return "파일에 접근할 수 없습니다. 다른 프로그램에서 열려 있는지 확인해 주세요."
    if isinstance(e, OSError) and getattr(e, "winerror", None) == 1155:
        return "이 형식의 파일을 열 프로그램이 설치되어 있지 않습니다."
    text = str(e).strip()
    return text or type(e).__name__
