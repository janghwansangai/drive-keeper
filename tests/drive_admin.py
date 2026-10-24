"""테스트 전용: 실제 드라이브 테스트 뒤 정리 (휴지통으로 옮기기).

프로그램 본체(drive_client)는 보안상 삭제·휴지통 요청을 보낼 수 없게 막혀 있다.
그래서 테스트가 만든 파일을 치울 때만 이 도우미가 따로 연결을 만들어 휴지통으로 옮긴다.
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import config  # noqa: E402
import protect  # noqa: E402


def real_token_source() -> Path:
    """이 PC의 실제 로그인 토큰 위치 (예전 방식 token.json 또는 첫 계정)."""
    legacy = config.BASE_DIR / "token.json"
    if legacy.exists():
        return legacy
    data = json.loads((config.BASE_DIR / "accounts.json").read_text(encoding="utf-8"))
    first = data["accounts"][0]["id"]
    return config.BASE_DIR / "accounts" / first / "token.json"


def copy_token_to(dest: Path) -> None:
    """실제 토큰을 (암호화된 채로) 테스트용 위치에 복사. 원래 파일은 건드리지 않는다."""
    creds, _ = protect.read_token(real_token_source())
    protect.write_token(dest, creds)


def trash(drive, file_id: str) -> None:
    from googleapiclient.discovery import build
    svc = build("drive", "v3", credentials=drive.creds, cache_discovery=False, static_discovery=True)
    svc.files().update(fileId=file_id, body={"trashed": True}, supportsAllDrives=True,
                       fields="id").execute(num_retries=4)
