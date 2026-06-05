# -*- coding: utf-8 -*-
"""표준 데이터 모델.

data_loader 가 raw data 를 읽어 RawProduct 리스트로 표준화하고,
transform 이 이를 가공해 DashboardTable(화면·PPTX 공용)로 변환한다.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Optional


@dataclass
class RawProduct:
    """raw data 한 행(제품 1개)을 표준화한 모델."""
    no: Optional[int]               # No. (시리즈 번호)
    category: str                   # 카테고리
    brand: str                      # 브랜드
    project: str                    # 프로젝트명
    series: str                     # 시리즈명 ('-' 가능)
    owner: str                      # 담당자
    channel: str                    # 오프/온라인
    size_count: Optional[float]     # 사이즈 수
    color_count: Optional[float]    # 컬러 수
    sku_count: Optional[float]      # SKU 수
    target_revenue: str             # 목표매출 (예 '0.7억')
    vendor: str                     # 업체
    vendor_country: str             # 업체 국가
    material: str                   # 주요소재
    target_price_raw: str           # 목표 판매가 원본 (예 '590,000원')
    timeline: list                  # 13개월 일정 문자열 리스트(없으면 '')
    note: str                       # 비고
    image: Optional[bytes] = None   # 셀에 삽입된 이미지(바이트). 없으면 None
    image_ext: Optional[str] = None # 이미지 확장자 (png/jpeg 등)


@dataclass
class DashboardRow:
    """대시보드 표의 데이터 행 1개(=제품 1개)."""
    no_label: str           # 병합 표시용 NO ('' 이면 위 행과 병합)
    series_cell: str        # 시리즈명 셀 (시리즈명 + 프로젝트명, 줄바꿈)
    image: Optional[bytes]
    image_ext: Optional[str]
    material: str
    target_price: str       # 만원 단위 변환값 (예 '59')
    target_revenue: str
    color_count: str
    sku_count: str
    maker: str              # 제조처 (업체 + 국가)
    owner: str
    timeline: list          # 13개월 셀 문자열
    note: str


@dataclass
class DashboardTable:
    """담당자 1명 = 슬라이드 1장에 해당하는 표 전체."""
    owner: str                      # 담당자 (소제목)
    main_title: str                 # 주제목
    show_q_label: bool              # 목표판매가 헤더에 (Q) 표기 여부
    rows: list = field(default_factory=list)        # DashboardRow 리스트
    summary: Optional["SummaryRow"] = None          # 하단 요약행


@dataclass
class SummaryRow:
    """하단 요약행: 'N시리즈 (총 SKU)' + 분기별 출시 집계."""
    label: str                      # 예 '6시리즈 (32 SKU)'
    per_quarter: list               # 분기별 셀 문자열 (QUARTERS 길이만큼)
    series_count: int = 0           # 총 시리즈 수
    total_sku: int = 0              # 총 SKU 수
