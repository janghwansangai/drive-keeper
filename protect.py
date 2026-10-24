"""비밀 파일 보호: Windows DPAPI로 암호화 (이 PC의 이 Windows 사용자만 풀 수 있음).

- 로그인 토큰(token.json)과 보안 설정(security.json)에 쓴다. 파일을 다른 PC·다른 사용자로
  복사해 가도 쓸 수 없다. (같은 사용자로 실행되는 악성코드까지 막지는 못함 → Windows 로그인 암호가 중요)
- 비밀번호 잠금이 켜져 있으면 토큰 파일 안에도 그 표시(lock)를 남긴다. 누가 security.json을
  지워서 잠금을 풀려고 해도, 토큰에 남은 표시를 보고 구글 계정 본인 확인을 요구한다.
"""
from __future__ import annotations

import base64
import ctypes
import json
import os
from pathlib import Path

import config

PREFIX = "GDE-DPAPI1:"
_ENTROPY = b"GDriveExplorer/v1"

# 비밀번호 잠금이 켜져 있는지 (App이 보안 설정을 읽은 뒤 맞춰 둔다)
LOCK_MARKER = False


if os.name == "nt":
    from ctypes import wintypes

    class _Blob(ctypes.Structure):
        _fields_ = [("cbData", wintypes.DWORD), ("pbData", ctypes.POINTER(ctypes.c_char))]

    _crypt32 = ctypes.WinDLL("crypt32", use_last_error=True)
    _kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    _crypt32.CryptProtectData.argtypes = [ctypes.POINTER(_Blob), wintypes.LPCWSTR, ctypes.POINTER(_Blob),
                                          ctypes.c_void_p, ctypes.c_void_p, wintypes.DWORD,
                                          ctypes.POINTER(_Blob)]
    _crypt32.CryptProtectData.restype = wintypes.BOOL
    _crypt32.CryptUnprotectData.argtypes = [ctypes.POINTER(_Blob), ctypes.c_void_p, ctypes.POINTER(_Blob),
                                            ctypes.c_void_p, ctypes.c_void_p, wintypes.DWORD,
                                            ctypes.POINTER(_Blob)]
    _crypt32.CryptUnprotectData.restype = wintypes.BOOL
    _kernel32.LocalFree.argtypes = [ctypes.c_void_p]
    _kernel32.LocalFree.restype = ctypes.c_void_p
    _UI_FORBIDDEN = 0x1

    def _make_blob(data: bytes):
        buf = ctypes.create_string_buffer(data, len(data))
        return _Blob(len(data), ctypes.cast(buf, ctypes.POINTER(ctypes.c_char))), buf

    def _call(fn, data: bytes) -> bytes:
        blob_in, keep1 = _make_blob(data)
        blob_ent, keep2 = _make_blob(_ENTROPY)
        blob_out = _Blob()
        if fn is _crypt32.CryptProtectData:
            ok = fn(ctypes.byref(blob_in), "GDriveExplorer", ctypes.byref(blob_ent), None, None,
                    _UI_FORBIDDEN, ctypes.byref(blob_out))
        else:
            ok = fn(ctypes.byref(blob_in), None, ctypes.byref(blob_ent), None, None,
                    _UI_FORBIDDEN, ctypes.byref(blob_out))
        if not ok:
            raise ctypes.WinError(ctypes.get_last_error())
        try:
            return ctypes.string_at(blob_out.pbData, blob_out.cbData)
        finally:
            _kernel32.LocalFree(ctypes.cast(blob_out.pbData, ctypes.c_void_p))

    def dpapi_protect(data: bytes) -> bytes:
        return _call(_crypt32.CryptProtectData, data)

    def dpapi_unprotect(data: bytes) -> bytes:
        return _call(_crypt32.CryptUnprotectData, data)
else:  # 개발용 (Mac/Linux): 암호화 없이 그대로
    def dpapi_protect(data: bytes) -> bytes:
        return data

    def dpapi_unprotect(data: bytes) -> bytes:
        return data


def protect_text(text: str) -> str:
    return PREFIX + base64.b64encode(dpapi_protect(text.encode("utf-8"))).decode("ascii")


def unprotect_text(stored: str) -> str:
    """암호화된 글이면 풀고, 예전 방식(평문)이면 그대로 돌려준다. 풀 수 없으면 ValueError."""
    stored = stored.strip()
    if not stored.startswith(PREFIX):
        return stored
    try:
        return dpapi_unprotect(base64.b64decode(stored[len(PREFIX):])).decode("utf-8")
    except Exception as e:
        raise ValueError(f"암호화된 파일을 풀 수 없습니다 (다른 PC·다른 사용자에서 만든 파일?): {e}") from e


def write_token(path: Path, creds_json: str) -> None:
    payload = json.dumps({"gde_token": 1, "creds": creds_json, "lock": bool(LOCK_MARKER)})
    config.atomic_write_text(Path(path), protect_text(payload))
    try:
        os.chmod(path, 0o600)
    except OSError:
        pass


def read_token(path: Path) -> tuple[str, bool]:
    """(구글 로그인 정보 JSON, 잠금 표시). 예전 평문 token.json도 읽는다."""
    text = unprotect_text(Path(path).read_text(encoding="utf-8"))
    try:
        obj = json.loads(text)
    except ValueError as e:
        raise ValueError(f"토큰 파일이 손상되었습니다: {e}") from e
    if isinstance(obj, dict) and obj.get("gde_token"):
        return str(obj.get("creds") or ""), bool(obj.get("lock"))
    return text, False
