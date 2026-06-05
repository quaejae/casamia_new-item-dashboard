# -*- coding: utf-8 -*-
"""공개(링크 공유) Google Sheets 로더 — Google Cloud/인증 불필요.

시트를 '링크가 있는 모든 사용자 - 뷰어' 또는 '웹에 게시' 로 공유하면,
export URL 로 워크북 전체를 xlsx 로 받아 기존 xlsx 로더로 그대로 읽는다.

장점: 결제·서비스계정·API 설정 전혀 불필요.
단점: 시트가 링크로 열람 가능해야 함(비공개 유지하려면 sheets_loader.py 의
      서비스계정 방식 사용 — 그래도 무료).

이미지: Google Sheets 의 in-cell 이미지/=IMAGE() 는 xlsx export 시 임베드되지
        않을 수 있어, 본 방식에서는 텍스트 데이터 위주로 반영된다(연결 후 확인).
"""
from __future__ import annotations

import urllib.request
from typing import Dict, List, Optional

import config
import data_loader
from schema import RawProduct

EXPORT_URL = "https://docs.google.com/spreadsheets/d/{key}/export?format=xlsx"


def load(spreadsheet_key: Optional[str] = None) -> Dict[str, List[RawProduct]]:
    key = spreadsheet_key or getattr(config, "SHEETS_KEY", None)
    if not key:
        raise RuntimeError("config.SHEETS_KEY 가 필요합니다.")
    url = EXPORT_URL.format(key=key)
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    try:
        with urllib.request.urlopen(req, timeout=20) as resp:
            data = resp.read()
    except urllib.error.HTTPError as e:
        if e.code in (401, 403):
            raise RuntimeError(
                "시트에 접근할 수 없습니다(비공개). 시트를 '링크가 있는 모든 "
                "사용자 - 뷰어'로 공유하거나 '웹에 게시'한 뒤 다시 시도하세요.") from e
        raise
    if data[:2] != b"PK":   # xlsx(zip) 시그니처가 아니면 로그인 페이지(HTML)
        raise RuntimeError(
            "xlsx 가 아닌 응답을 받았습니다. 시트 공개 설정을 확인하세요.")
    return data_loader.load_bytes(data)


if __name__ == "__main__":
    import io, sys
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
    d = load(sys.argv[1] if len(sys.argv) > 1 else None)
    for owner, prods in d.items():
        print(f"[{owner}] 제품 {len(prods)}개")
