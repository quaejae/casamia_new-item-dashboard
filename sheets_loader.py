# -*- coding: utf-8 -*-
"""2단계: Google Sheets 로더 (드롭인 교체용 스캐폴드).

data_loader.load() 와 동일하게 {담당자: [RawProduct, ...]} 를 반환하므로,
transform / pptx_export / app 코드는 한 줄도 바꾸지 않고 그대로 동작한다.

⚠️ 이 모듈은 실제 Google 인증정보(서비스 계정 키)와 시트가 있어야 검증된다.
   README 의 'Google Sheets 연결' 절차대로 설정 후 사용할 것.

이미지 처리:
  - 권장: 이미지 열 셀에 =IMAGE("https://...") 수식 → 본 로더가 URL 을 받아 다운로드.
  - UI '셀에 이미지 삽입'(임베드) 방식은 Sheets API 로 바이트를 직접 받기 어려워,
    Sheets 단계에서는 =IMAGE() URL 방식을 표준으로 권장한다(README 참고).
"""
from __future__ import annotations

import re
from typing import Dict, List, Optional

import config
from schema import RawProduct
from data_loader import HEADER_ALIASES, _norm, _txt  # 컬럼 매핑 규칙 재사용


def _open_spreadsheet(spreadsheet_key: str, creds_path: str):
    import gspread
    from google.oauth2.service_account import Credentials

    scopes = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
    creds = Credentials.from_service_account_file(creds_path, scopes=scopes)
    gc = gspread.authorize(creds)
    return gc.open_by_key(spreadsheet_key)


def _download_image(url: str) -> Optional[tuple]:
    """=IMAGE() URL 에서 이미지 바이트와 확장자를 받아온다."""
    try:
        import urllib.request
        with urllib.request.urlopen(url, timeout=10) as resp:
            data = resp.read()
        ext = url.rsplit(".", 1)[-1].lower()
        if ext not in ("png", "jpg", "jpeg", "gif", "webp"):
            ext = "png"
        return data, ext
    except Exception:
        return None


def _parse_image_formula(formula: str) -> Optional[str]:
    m = re.search(r'IMAGE\(\s*"([^"]+)"', formula or "", re.IGNORECASE)
    return m.group(1) if m else None


def _load_worksheet(ws) -> List[RawProduct]:
    # 값(표시 텍스트)과 수식을 함께 읽어 이미지 URL 추출
    values = ws.get_all_values()                    # 2차원 텍스트
    formulas = ws.get_all_values(value_render_option="FORMULA")

    def get(r, c):  # 0-base 안전 접근
        return values[r][c] if r < len(values) and c < len(values[r]) else ""

    # 헤더(HEADER_ROW) 로 컬럼 매핑
    hr = config.HEADER_ROW - 1
    field_to_col: Dict[str, int] = {}
    header = values[hr] if hr < len(values) else []
    for c, raw in enumerate(header):
        key = _norm(raw)
        if key in HEADER_ALIASES:
            field_to_col[HEADER_ALIASES[key]] = c

    # 타임라인: MONTH_ROW 의 숫자 셀 연속 구간
    mr = config.MONTH_ROW - 1
    month_row = values[mr] if mr < len(values) else []
    timeline_cols = [c for c, v in enumerate(month_row) if re.fullmatch(r"\d+", str(v).strip())]

    def cell(r, field):
        c = field_to_col.get(field)
        return get(r, c) if c is not None else ""

    def num(r, field):
        v = re.sub(r"[^\d.]", "", str(cell(r, field)))
        return float(v) if v else None

    products: List[RawProduct] = []
    for r in range(config.DATA_START_ROW - 1, len(values)):
        project = _txt(cell(r, "project"))
        category = _txt(cell(r, "category"))
        if not project and not category:
            continue

        timeline = [_txt(get(r, c)) for c in timeline_cols]

        # 이미지: 이미지 열의 =IMAGE() 수식 → URL → 다운로드
        image = image_ext = None
        img_col = field_to_col.get("image")
        if img_col is not None and r < len(formulas) and img_col < len(formulas[r]):
            url = _parse_image_formula(formulas[r][img_col])
            if url:
                got = _download_image(url)
                if got:
                    image, image_ext = got

        products.append(RawProduct(
            no=int(num(r, "no")) if num(r, "no") is not None else None,
            category=category, brand=_txt(cell(r, "brand")), project=project,
            series=_txt(cell(r, "series")), owner=_txt(cell(r, "owner")),
            channel=_txt(cell(r, "channel")),
            size_count=num(r, "size_count"), color_count=num(r, "color_count"),
            sku_count=num(r, "sku_count"), target_revenue=_txt(cell(r, "target_revenue")),
            vendor=_txt(cell(r, "vendor")), vendor_country=_txt(cell(r, "vendor_country")),
            material=_txt(cell(r, "material")), target_price_raw=_txt(cell(r, "target_price_raw")),
            timeline=timeline, note=_txt(cell(r, "note")),
            image=image, image_ext=image_ext,
        ))
    return products


def load(spreadsheet_key: Optional[str] = None,
         creds_path: Optional[str] = None) -> Dict[str, List[RawProduct]]:
    """Google Sheets 를 읽어 {담당자: [RawProduct]} 반환 (data_loader.load 와 호환)."""
    spreadsheet_key = spreadsheet_key or getattr(config, "SHEETS_KEY", None)
    creds_path = creds_path or getattr(config, "SHEETS_CREDS", None)
    if not spreadsheet_key or not creds_path:
        raise RuntimeError(
            "Google Sheets 설정이 필요합니다. config.SHEETS_KEY / config.SHEETS_CREDS 를 "
            "지정하거나 load(key, creds_path) 인자를 전달하세요. (README 참고)")

    sh = _open_spreadsheet(spreadsheet_key, creds_path)
    result: Dict[str, List[RawProduct]] = {}
    for ws in sh.worksheets():
        products = _load_worksheet(ws)
        if products:
            owner = products[0].owner or ws.title
            result[owner] = products
    return result
