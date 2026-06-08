# -*- coding: utf-8 -*-
"""대시보드 전역 설정.

이 파일의 값만 바꾸면 주제목·단위·색상·카테고리 규칙 등을 조정할 수 있습니다.
화면(app.py)과 PPTX export(pptx_export.py)가 모두 이 설정을 공유합니다.
"""
from __future__ import annotations

# ─────────────────────────────────────────────────────────────────────────────
# 1. 텍스트 / 제목
# ─────────────────────────────────────────────────────────────────────────────
# 주제목: 설정에서 변경 가능한 고정 텍스트 (요구사항 확인 1번)
MAIN_TITLE = "26SS,FW 신상품 출시 스케줄"

# 소제목 접두 기호 (예: "• 담당자1")
SUBTITLE_PREFIX = "• "
# 소제목 글자색 (양식 예시: 주황 accent6)
SUBTITLE_COLOR = "F79646"

# 주제목 아래 장식선
TITLE_RULE_ENABLED = True
TITLE_RULE_COLOR = "181818"     # 거의 검정
TITLE_RULE_WIDTH_PT = 1.0

# 하단 푸터: 장식선 + 브랜드 텍스트(슬라이드마다 고정)
FOOTER_ENABLED = True
FOOTER_TEXT = "SHINSEGAE CASA"
FOOTER_FONT = 9                 # pt
FOOTER_COLOR = "404040"
FOOTER_LINE_MARGIN_MM = 9       # 하단 장식선이 슬라이드 밑변에서 떨어진 거리
PAGE_GAP_MM = 2                 # 표와 하단 장식선 사이 최소 여백

# ─────────────────────────────────────────────────────────────────────────────
# 2. 데이터 소스 (1단계: 로컬 xlsx)
# ─────────────────────────────────────────────────────────────────────────────
# 데이터 소스: "xlsx"(로컬 파일) 또는 "sheets"(Google Sheets)
DATA_SOURCE = "sheets"

# 기본 로컬 raw data 파일명 (working dir 기준 상대경로 또는 절대경로)
LOCAL_XLSX = "신제품 출시 스케줄 관리_law data.xlsx"

# ── Google Sheets 연결 (DATA_SOURCE="sheets" 일 때 사용) ─────────────────────
# 시트 URL 의 /d/ 와 /edit 사이 문자열
SHEETS_KEY = "18yYM6Spbpmn2zEa-IyLuhWZV8xTzNvaqrd_a7cNTQi0"

# 연결 방식:
#   "public"          = 결제·인증 불필요. 시트를 '링크가 있는 사용자-뷰어'로 공유 필요.
#   "service_account" = 비공개 유지. 서비스 계정 JSON 키 필요(무료, Cloud 설정 필요).
SHEETS_MODE = "public"

# service_account 방식일 때만 사용: JSON 키 파일 경로
SHEETS_CREDS = "service_account.json"

# raw data 시트에서 데이터가 시작하는 행/헤더 행 (1-base)
HEADER_ROW = 3          # 컬럼 헤더 (No., 카테고리, ...)
QUARTER_ROW = 4         # 분기 헤더 (25 / 26 1Q ...)
MONTH_ROW = 5           # 월 헤더 (12, 1, 2, ... 12)
DATA_START_ROW = 6      # 제품 데이터 시작 행

# ─────────────────────────────────────────────────────────────────────────────
# 3. 단위 / 변환
# ─────────────────────────────────────────────────────────────────────────────
# 목표판매가: 원 단위 → 만원 단위 (590,000원 → 59)
PRICE_DIVISOR = 10_000

# (Q) = Q(퀸)사이즈 기준가 표기. 매트리스·베드프레임 제품에만 적용 (확인 2번).
#  - 카테고리가 아래 집합에 속하거나
#  - 프로젝트명에 아래 키워드가 포함되면 Q사이즈 적용으로 간주
Q_SIZE_CATEGORIES = {"매트리스"}
Q_SIZE_KEYWORDS = ("베드프레임",)
Q_SIZE_LABEL = "(Q)"

# 시리즈명 셀의 프로젝트명 축약 규칙 (예시: '베드프레임' → '베드')
PROJECT_ABBREV = {
    "베드프레임": "베드",
}

# ─────────────────────────────────────────────────────────────────────────────
# 4. 타임라인 / 마일스톤
# ─────────────────────────────────────────────────────────────────────────────
# 13개월 타임라인: '25.12 ~ '26.12. (연도, 월) 순서대로
TIMELINE_MONTHS = [
    (2025, 12),
    (2026, 1), (2026, 2), (2026, 3),
    (2026, 4), (2026, 5), (2026, 6),
    (2026, 7), (2026, 8), (2026, 9),
    (2026, 10), (2026, 11), (2026, 12),
]

# 분기 구분: (분기 라벨, 시작 월 인덱스(0-base), 길이)
QUARTERS = [
    ("25", 0, 1),       # '25.12
    ("26 1Q", 1, 3),    # 1~3월
    ("26 2Q", 4, 3),    # 4~6월
    ("26 3Q", 7, 3),    # 7~9월
    ("26 4Q", 10, 3),   # 10~12월
]

# 출시 마일스톤으로 인식할 키워드 (요약행 집계용)
LAUNCH_KEYWORD = "출시"
# 진행 화살표 시작 마일스톤 키워드 (기획에서 시작)
PLAN_KEYWORD = "기획"

# 오늘 날짜의 월에 해당하는 열 강조 (매일 자동 갱신)
CURRENT_MONTH_HIGHLIGHT = True
CURRENT_MONTH_COLOR = "002060"      # 강조 테두리색(네이비)
CURRENT_MONTH_WIDTH_PT = 2.25

# 셀 배경색 (RGB hex, '#' 없이) — 예시 PPTX와 동일.
HEADER_FILL = "E7E7E6"          # 헤더 2행
SUMMARY_FILL = "ECE8E2"         # 하단 요약행(베이지)

# 표 테두리 (양식 예시: 회색)
BORDER_COLOR = "BFBFBF"
BORDER_WIDTH_PT = 1.0

# 일정 마일스톤 중 주황 텍스트로 강조할 항목 (발주/출시)
ORANGE_TEXT_COLOR = "984807"
ORANGE_TEXT_EXACT = ("발주",)     # 정확히 일치할 때만(‘1차 샘플 발주’ 제외)
ORANGE_TEXT_CONTAINS = ("출시",)  # 포함하면 강조

# 진행기간 화살표 (기획→출시 직전)
ARROW_ENABLED = True
ARROW_COLOR_FROM = "FFFFFF"     # PPTX 화살표 시작색(기획 쪽)
ARROW_COLOR_TO = "EEE3D6"       # PPTX 화살표 끝색(출시 쪽)
ARROW_HEIGHT_RATIO = 0.45       # 데이터 행 높이 대비 화살표 높이 비율

# 웹 화면용 진행밴드 색 — 텍스트가 잘 보이도록 연하게. 방향은 화살촉으로 표시.
WEB_ARROW_BAND_FROM = "FEF8EF"  # 기획 쪽(아주 연한 베이지)
WEB_ARROW_BAND_TO = "F6E6CC"    # 출시 쪽(연한 베이지)
WEB_ARROW_HEAD = "D38A2E"       # 화살촉 색(또렷한 주황)
# 마일스톤 종류별 강조색. 키워드 부분일치, 매칭 없으면 채움 없음(흰색).
# 예시는 '출시'만 강조(FFF2CC). 필요 시 항목 추가 가능.
MILESTONE_COLORS = {
    "출시": "FFF2CC",
}

# ─────────────────────────────────────────────────────────────────────────────
# 5. 페이지 규격 (PPTX export) — A4 가로 (확인 3번)
# ─────────────────────────────────────────────────────────────────────────────
PAGE_SIZE = "A4_LANDSCAPE"
A4_LANDSCAPE_WIDTH_MM = 297
A4_LANDSCAPE_HEIGHT_MM = 210
PAGE_MARGIN_MM = 8

# 표 24열 상대 너비 가중치 (사용 가능 폭에 맞춰 자동 정규화).
# [NO, 시리즈명, 이미지, 소재, 목표판매가, 목표매출, 컬러, SKU수, 제조처, 담당자] + 13개월 + 비고
COL_WEIGHTS = (
    0.55, 2.0, 1.4, 1.4, 0.9, 0.85, 0.6, 0.7, 1.3, 0.9,   # 정보 10열
    *([0.95] * 13),                                         # 13개월
    1.2,                                                    # 비고
)

# 폰트 크기 (pt)
FONT_TITLE = 16        # 주제목
FONT_SUBTITLE = 12     # 소제목
FONT_HEADER = 7        # 표 헤더
FONT_BODY = 7          # 표 본문
FONT_SUMMARY = 9       # 요약행(양식 예시: 9pt)

# 행 높이 (pt)
ROW_H_HEADER = 14
ROW_H_MONTH = 12
ROW_H_DATA = 34        # 이미지가 들어갈 수 있어 넉넉히
ROW_H_SUMMARY = 21     # 요약행(양식 예시: 21pt)

# 표 기본 글꼴
FONT_NAME = "맑은 고딕"
