"""여러 구글 계정: 계정 목록과 계정별 파일 위치.

계정마다 따로 둔다 (섞이면 A 계정 파일을 B 계정으로 올리는 사고가 나므로):
  accounts/<id>/token.json       로그인 토큰 (DPAPI 암호화)
  accounts/<id>/state.json       편집 중인 파일 기록
  accounts/<id>/root_cache.json  시작할 때 바로 보여 줄 '내 드라이브' 목록
  accounts/<id>/staging/         큰 파일 업로드용 임시 사본
  작업 폴더: accounts/<id>/work  (예전 버전에서 옮겨 온 첫 계정은 기존 work 폴더를 그대로 씀)
accounts.json은 목록, 계정 폴더의 account.json은 그 계정 정보 (목록이 깨지면 이것으로 다시 만든다).
"""
from __future__ import annotations

import json
import logging
import os
import re
import secrets
import shutil
import threading
from pathlib import Path

import config

log = logging.getLogger(__name__)

_ID = re.compile(r"^[0-9a-zA-Z_-]{1,40}$")


def new_id() -> str:
    return secrets.token_hex(3)


class Account:
    __slots__ = ("id", "email", "work_dir")

    def __init__(self, id: str, email: str = "", work_dir: str | Path | None = None) -> None:
        self.id = id
        self.email = email or ""
        self.work_dir = Path(work_dir) if work_dir else self.dir / "work"

    @property
    def dir(self) -> Path:
        return config.BASE_DIR / "accounts" / self.id

    @property
    def token_path(self) -> Path:
        return self.dir / "token.json"

    @property
    def state_path(self) -> Path:
        return self.dir / "state.json"

    @property
    def root_cache_path(self) -> Path:
        return self.dir / "root_cache.json"

    @property
    def staging_dir(self) -> Path:
        return self.dir / "staging"

    @property
    def label(self) -> str:
        return self.email or "계정 확인 중..."

    @property
    def short(self) -> str:
        return self.email.split("@")[0] if self.email else "계정"

    def ensure_dirs(self) -> None:
        for d in (self.dir, self.work_dir, self.staging_dir):
            d.mkdir(parents=True, exist_ok=True)

    def to_json(self) -> dict:
        return {"id": self.id, "email": self.email, "work_dir": str(self.work_dir)}

    @classmethod
    def from_json(cls, d: dict) -> "Account":
        if not _ID.match(str(d.get("id", ""))):
            raise ValueError(f"잘못된 계정 id: {d.get('id')!r}")
        return cls(str(d["id"]), str(d.get("email") or ""), d.get("work_dir") or None)


class AccountStore:
    def __init__(self) -> None:
        self.accounts: list[Account] = []
        self._lock = threading.RLock()

    @property
    def path(self) -> Path:
        return config.BASE_DIR / "accounts.json"

    # ----- 읽기 -----
    def load(self) -> list[Account]:
        with self._lock:
            try:
                data = json.loads(self.path.read_text(encoding="utf-8"))
                self.accounts = [Account.from_json(d) for d in data.get("accounts", [])]
            except FileNotFoundError:
                self.accounts = self._from_dirs()
                if not self.accounts:
                    self._migrate_legacy()
                else:
                    self.save()
            except Exception as e:
                log.error("accounts.json이 손상되어 계정 폴더에서 다시 만듭니다: %s", e)
                try:
                    shutil.copyfile(self.path, self.path.with_suffix(".broken.json"))
                except OSError:
                    pass
                self.accounts = self._from_dirs()
                self.save()
            return list(self.accounts)

    def _from_dirs(self) -> list[Account]:
        found = []
        for f in sorted((config.BASE_DIR / "accounts").glob("*/account.json")):
            try:
                found.append(Account.from_json(json.loads(f.read_text(encoding="utf-8"))))
            except Exception as e:
                log.warning("계정 정보 파일을 읽지 못함 %s: %s", f, e)
        return found

    def _migrate_legacy(self) -> None:
        """예전 버전(계정 하나)의 token.json·state.json을 첫 계정으로 옮긴다."""
        base = config.BASE_DIR
        legacy = [base / "token.json", base / "state.json", base / "root_cache.json"]
        if not any(p.exists() for p in legacy[:2]):
            return
        acct = Account(new_id(), "", base / "work")   # 작업 폴더는 그대로 (열려 있는 파일이 있을 수 있음)
        acct.ensure_dirs()
        token = legacy[0]
        if token.exists():
            # 예전 평문 토큰 → 암호화해서 옮긴다
            try:
                import protect
                creds, _ = protect.read_token(token)
                protect.write_token(acct.token_path, creds)
                token.unlink()
            except Exception as e:
                log.warning("토큰을 암호화해 옮기지 못해 그대로 옮깁니다: %s", e)
                try:
                    os.replace(token, acct.token_path)
                except OSError:
                    shutil.copyfile(token, acct.token_path)
        for src, dst in zip(legacy[1:], (acct.state_path, acct.root_cache_path)):
            if not src.exists():
                continue
            try:
                os.replace(src, dst)
            except OSError:
                shutil.copyfile(src, dst)
        log.info("예전 로그인 정보를 계정 %s로 옮김", acct.id)
        self.accounts = [acct]
        self.save()

    # ----- 쓰기 -----
    def save(self) -> None:
        with self._lock:
            data = {"version": 1, "accounts": [a.to_json() for a in self.accounts]}
            config.atomic_write_text(self.path, json.dumps(data, ensure_ascii=False, indent=1))
            for a in self.accounts:
                try:
                    a.dir.mkdir(parents=True, exist_ok=True)
                    config.atomic_write_text(a.dir / "account.json", json.dumps(a.to_json(), ensure_ascii=False))
                except OSError as e:
                    log.warning("계정 정보 파일 저장 실패: %s", e)

    def get(self, account_id: str | None) -> Account | None:
        with self._lock:
            return next((a for a in self.accounts if a.id == account_id), None)

    def find_email(self, email: str) -> Account | None:
        e = (email or "").casefold()
        with self._lock:
            return next((a for a in self.accounts if a.email and a.email.casefold() == e), None)

    def add(self, email: str) -> tuple[Account, bool]:
        """(계정, 새로 만들었는가). 같은 이메일이 이미 있으면 그 계정을 돌려준다."""
        with self._lock:
            existing = self.find_email(email)
            if existing:
                return existing, False
            acct = Account(new_id(), email)
            while self.get(acct.id) is not None:
                acct = Account(new_id(), email)
            acct.ensure_dirs()
            self.accounts.append(acct)
            self.save()
            return acct, True

    def set_email(self, account_id: str, email: str) -> None:
        with self._lock:
            acct = self.get(account_id)
            if acct and email and acct.email != email:
                acct.email = email
                self.save()

    def remove(self, account_id: str) -> Account | None:
        with self._lock:
            acct = self.get(account_id)
            if acct:
                self.accounts.remove(acct)
                self.save()
            return acct
