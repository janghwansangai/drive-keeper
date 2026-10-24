"""새 버전 알림: GitHub에 올라온 최신 버전과 지금 버전을 비교한다.

- 내려받거나 실행하지는 않는다. 새 버전이 있으면 알려 주고, 사용자가 누르면 다운로드 페이지를 연다.
- 구글 드라이브와 상관없는 GitHub 공개 정보만 읽는다 (사용자 정보는 보내지 않음).
"""
from __future__ import annotations

import json
import re
import urllib.error
import urllib.request

import config

API_URL = f"https://api.github.com/repos/{config.GITHUB_REPO}/releases/latest"


def parse_version(text: str) -> tuple[int, ...]:
    nums = re.findall(r"\d+", text or "")
    return tuple(int(n) for n in nums[:3]) or (0,)


def is_newer(latest: str, current: str = config.VERSION) -> bool:
    return parse_version(latest) > parse_version(current)


def check_latest(timeout: float = 6.0) -> dict | None:
    """{'version': '2.1.0', 'url': 다운로드 페이지, 'notes': 바뀐 점} 또는 None(올린 버전이 없음)."""
    req = urllib.request.Request(API_URL, headers={
        "Accept": "application/vnd.github+json",
        "User-Agent": f"{config.APP_NAME}/{config.VERSION}",
    })
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            data = json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        if e.code == 404:
            return None
        raise
    tag = str(data.get("tag_name") or "")
    if not tag:
        return None
    return {"version": tag.lstrip("vV"), "url": data.get("html_url") or config.RELEASES_URL,
            "notes": str(data.get("body") or "")[:2000]}
