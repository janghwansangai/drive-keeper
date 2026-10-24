"""비밀번호 잠금: 설정 저장, 비밀번호 확인, 자리 비움(유휴) 시간.

- 비밀번호는 저장하지 않고 scrypt(소금+느린 해시) 결과만 저장한다.
- 설정 파일은 DPAPI로 암호화해 메모장으로 고칠 수 없다. 손상·삭제되면(토큰에 잠금 표시가 있을 때)
  구글 계정 본인 확인을 거쳐야 새 비밀번호를 정할 수 있다.
- 비밀번호를 5번 틀리면 30초부터 최대 15분까지 기다리게 한다 (껐다 켜도 유지).
"""
from __future__ import annotations

import base64
import ctypes
import hashlib
import hmac
import json
import logging
import os
import time
from pathlib import Path

import config
import protect

log = logging.getLogger(__name__)

DEFAULTS = {
    "enabled": False,        # 잠금 사용 ('모든 잠금 끄기' → False)
    "lock_on_start": True,   # 프로그램을 켤 때(부팅 후 처음 켤 때) 잠금
    "idle_minutes": 10,      # 이 시간 동안 컴퓨터를 쓰지 않으면 잠금 (0 = 끔)
    "lock_on_sleep": True,   # 절전 모드에서 돌아오면 잠금
}
_SCRYPT = {"n": 2 ** 14, "r": 8, "p": 1}
MIN_PASSWORD = 4
MAX_FAILS_FREE = 5


def _hash(password: str, salt: bytes, n: int, r: int, p: int) -> bytes:
    return hashlib.scrypt(password.encode("utf-8"), salt=salt, n=n, r=r, p=p, dklen=32,
                          maxmem=64 * 1024 * 1024)


class Security:
    def __init__(self) -> None:
        self.data: dict = dict(DEFAULTS)
        self.reset_required = False   # 설정 파일이 사라졌거나 손상됨 → 구글 본인 확인 필요

    @property
    def path(self) -> Path:
        return config.BASE_DIR / "security.json"

    # ----- 읽기/쓰기 -----
    def load(self, lock_marker_seen: bool = False) -> None:
        """lock_marker_seen: 토큰 파일 중 하나라도 '잠금이 켜져 있었다'는 표시가 있는가."""
        self.data = dict(DEFAULTS)
        self.reset_required = False
        try:
            raw = self.path.read_text(encoding="utf-8")
        except FileNotFoundError:
            if lock_marker_seen:
                log.warning("보안 설정 파일이 없는데 토큰에 잠금 표시가 있음 → 본인 확인 필요")
                self.reset_required = True
            return
        except OSError as e:
            log.error("보안 설정을 읽을 수 없음: %s", e)
            self.reset_required = True
            return
        try:
            if os.name == "nt" and not raw.strip().startswith(protect.PREFIX):
                raise ValueError("암호화되지 않은 설정 파일 (직접 고친 흔적)")
            data = json.loads(protect.unprotect_text(raw))
            if not isinstance(data, dict):
                raise ValueError("형식 오류")
            for k in DEFAULTS:
                if k in data:
                    self.data[k] = data[k]
            for k in ("salt", "hash", "kdf", "fails", "wait_until"):
                if k in data:
                    self.data[k] = data[k]
            self.data["idle_minutes"] = max(0, min(int(self.data.get("idle_minutes") or 0), 240))
        except Exception as e:
            log.error("보안 설정 파일이 손상됨 → 본인 확인 필요: %s", e)
            self.data = dict(DEFAULTS)
            self.reset_required = True

    def save(self) -> None:
        config.atomic_write_text(self.path, protect.protect_text(json.dumps(self.data)))

    # ----- 상태 -----
    @property
    def has_password(self) -> bool:
        return bool(self.data.get("hash") and self.data.get("salt"))

    @property
    def active(self) -> bool:
        """잠금이 실제로 동작하는가."""
        return bool(self.data.get("enabled")) and self.has_password and not self.reset_required

    # ----- 비밀번호 -----
    def set_password(self, password: str) -> None:
        if len(password) < MIN_PASSWORD:
            raise ValueError(f"비밀번호는 {MIN_PASSWORD}자 이상이어야 합니다.")
        salt = os.urandom(16)
        self.data["salt"] = base64.b64encode(salt).decode("ascii")
        self.data["kdf"] = dict(_SCRYPT)
        self.data["hash"] = base64.b64encode(_hash(password, salt, **_SCRYPT)).decode("ascii")
        self.data["enabled"] = True
        self.data["fails"] = 0
        self.data["wait_until"] = 0
        self.reset_required = False
        self.save()

    def wait_remaining(self) -> int:
        return max(0, int(float(self.data.get("wait_until") or 0) - time.time() + 0.999))

    def check(self, password: str) -> tuple[bool, int]:
        """(맞는가, 기다려야 할 초). 틀린 횟수가 쌓이면 기다리게 한다."""
        wait = self.wait_remaining()
        if wait:
            return False, wait
        if not self.has_password:
            return False, 0
        kdf = self.data.get("kdf") or _SCRYPT
        try:
            salt = base64.b64decode(self.data["salt"])
            expected = base64.b64decode(self.data["hash"])
            got = _hash(password, salt, int(kdf["n"]), int(kdf["r"]), int(kdf["p"]))
        except Exception:
            log.exception("비밀번호 확인 오류")
            return False, 0
        if hmac.compare_digest(got, expected):
            if self.data.get("fails"):
                self.data["fails"] = 0
                self.data["wait_until"] = 0
                self._save_quiet()
            return True, 0
        fails = int(self.data.get("fails") or 0) + 1
        self.data["fails"] = fails
        if fails >= MAX_FAILS_FREE:
            delay = min(900, 30 * 2 ** (fails - MAX_FAILS_FREE))
            self.data["wait_until"] = time.time() + delay
        self._save_quiet()
        return False, self.wait_remaining()

    def fails(self) -> int:
        return int(self.data.get("fails") or 0)

    def _save_quiet(self) -> None:
        try:
            self.save()
        except Exception:
            log.exception("보안 설정 저장 실패")

    # ----- 설정 -----
    def update(self, enabled: bool, lock_on_start: bool, idle_minutes: int, lock_on_sleep: bool) -> None:
        self.data["enabled"] = bool(enabled) and self.has_password
        self.data["lock_on_start"] = bool(lock_on_start)
        self.data["idle_minutes"] = max(0, min(int(idle_minutes), 240))
        self.data["lock_on_sleep"] = bool(lock_on_sleep)
        self.save()

    def disable_all(self) -> None:
        """'모든 잠금 끄기': 비밀번호는 남겨 두고 잠금만 모두 끈다 (다시 켤 때 비밀번호 필요)."""
        self.data["enabled"] = False
        self.save()


# ----- 자리 비움 시간 (컴퓨터 전체 기준: 마지막으로 키보드·마우스를 쓴 뒤 지난 초) -----
if os.name == "nt":
    from ctypes import wintypes

    class _LastInput(ctypes.Structure):
        _fields_ = [("cbSize", wintypes.UINT), ("dwTime", wintypes.DWORD)]

    _u32 = ctypes.WinDLL("user32", use_last_error=True)
    _k32 = ctypes.WinDLL("kernel32", use_last_error=True)
    _u32.GetLastInputInfo.argtypes = [ctypes.POINTER(_LastInput)]
    _u32.GetLastInputInfo.restype = wintypes.BOOL
    _k32.GetTickCount.restype = wintypes.DWORD

    def system_idle_seconds() -> float | None:
        info = _LastInput(ctypes.sizeof(_LastInput), 0)
        if not _u32.GetLastInputInfo(ctypes.byref(info)):
            return None
        ms = (_k32.GetTickCount() - info.dwTime) & 0xFFFFFFFF
        return ms / 1000.0
else:
    def system_idle_seconds() -> float | None:
        return None
