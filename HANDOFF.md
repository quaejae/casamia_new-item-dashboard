# 신제품 출시 스케줄 대시보드 — 핸드오프 문서

> 새 세션에서 이 문서만 읽으면 작업을 이어갈 수 있도록 정리한 인수인계 문서입니다.
> 코드 전문이 아니라 **기능별 핵심 로직과 위치**를 요약했습니다. 실제 수정 전에는 해당 파일을 Read 하세요.

---

## 1. 한 줄 요약

까사미아(신세계까사) 가구 신상품 출시 스케줄을, **공개 Google Sheets(law data)** 에서 실시간으로 읽어
**Streamlit 웹 대시보드**로 보여주고 **A4 가로 PPTX**로 내보내는 프로그램.
화면과 PPTX는 같은 가공 로직(`transform.py`)을 공유해 항상 일치한다.

- **실행:** `streamlit run app.py` (working dir = `dashboard/`)
- **배포 링크:** https://casamia-new-item-dashboard.streamlit.app
- **GitHub:** github.com/quaejae/casamia_new-item-dashboard (repo root = `dashboard/` 내용, public, main). `git push` 하면 자동 재배포.
- **데이터:** 공개 시트(`SHEETS_KEY`) → `export?format=xlsx` 다운로드 → openpyxl 파싱. 인증·결제 불필요.

---

## 2. 데이터 흐름 & 파일 역할

```
공개 Google Sheets ──export xlsx──▶ data_loader (RawProduct 목록)
                                         │
                  ┌──────────────────────┴───────────────────────┐
          transform.build_summary                    transform.build_groupings
          (전체요약 매트릭스 dict)                     ({탭명: [DashboardTable]})
                  │                                              │
        ┌─────────┴─────────┐                        ┌───────────┴───────────┐
  render_html.build_       pptx_export.build_   render_html.build_     pptx_export.
  summary_html (웹)        summary_pptx (PPTX)  html_table (웹)        build_pptx (PPTX)
```

| 파일 | 역할 |
|------|------|
| `app.py` | Streamlit 진입점. 사이드바(소스/새로고침/내보내기), 탭 렌더 |
| `config.py` | **모든 설정**(제목·색상·단위·카테고리 규칙·페이지 규격). 값만 바꿔 조정 |
| `schema.py` | `RawProduct` / `DashboardRow` / `DashboardTable` / `SummaryRow` dataclass |
| `data_loader.py` | xlsx → `RawProduct`. 헤더명으로 열 탐지, 셀 이미지 앵커 추출 |
| `sheets_public_loader.py` | 공개 시트 export xlsx 다운로드 → `data_loader.load_bytes` |
| `sheets_loader.py` | (대안) 서비스계정 방식. 현재 미사용 |
| `transform.py` | **핵심 가공**. 정렬·그룹화·요약 매트릭스 |
| `render_html.py` | `DashboardTable`/요약 dict → HTML (웹) |
| `pptx_export.py` | 동일 데이터 → A4 가로 PPTX |

---

## 3. 데이터 모델 (`schema.py`)

- **`RawProduct`**: law data 한 행. `no, category, brand, project, series, owner, channel, sku_count, target_revenue('0.7억' 문자열), vendor, vendor_country, material, target_price_raw('590,000원'), timeline(13개월 문자열 리스트), note, image(bytes), image_ext`.
- **`DashboardTable`**: 소제목 1개(담당자/그룹) = 표 1개. `owner(라벨), show_q_label, rows[DashboardRow], summary[SummaryRow]`.
- **요약(`build_summary`)은 dataclass가 아니라 dict** 반환 (구조는 §6 참조).

---

## 4. 데이터 로딩 (`data_loader.py`)

- 열은 **헤더 텍스트로 탐지**(`HEADER_ALIASES`) → 열 순서/추가에 강함.
- 타임라인 13개월 = `MONTH_ROW`(5행)에서 숫자 셀이 연속된 구간 자동 인식.
- **셀 이미지**: `_extract_images`가 openpyxl `ws._images`의 앵커 행(`anchor._from.row`)으로 `{행: (bytes, ext)}` 매핑. 구글 시트 '셀 내 이미지 삽입'이 xlsx export 시 임베드되어 그대로 읽힘 — **코드 수정 불필요**.
- 행 시작 = `DATA_START_ROW`(6). project·category 둘 다 비면 스킵.
- `load(path)`(로컬) / `load_bytes(data)`(시트 export) 둘 다 `{담당자: [RawProduct]}` 반환.

---

## 5. 상세 탭 (정렬·그룹화) — `transform.py`

**탭 구성** (`TAB_BUILDERS`, app에선 전체요약 다음에 배치):
①카테고리별 ②브랜드별 ③온라인 전용 ④담당자별.

- **출시일 정렬(`_sort_by_launch`)**: 모든 표는 출시 빠른 순(위)→늦은 순(아래). 시리즈((담당자,No.)) 블록을 깨지 않고 블록 단위 정렬. 정렬 후 NO 1,2,3 재번호(`build_table(renumber=True)`).
- **동적 그룹화(`_ordered_unique`)**: `config.CATEGORY_ORDER`/`BRAND_ORDER` 우선 + 데이터의 **신규 카테고리/브랜드를 뒤에 자동 추가**(누락 없음).
- **`SPLIT_BY_BRAND_CATEGORIES`**(매트리스·베드룸)만 카테고리별 탭에서 브랜드로 세분.
- 빈 값 라벨 = `UNCLASSIFIED_LABEL`("(미분류)").
- 요약행(`_build_summary`): 'N시리즈 (총 SKU)' + 분기별 출시 `시리즈수(SKU)`. 분기는 달력 기준(`_month_to_quarter_idx`).

**핵심 함수 위치:** `build_table`(L193), `build_by_category`(L283), `build_by_brand`(L301), `build_online_only`(L313), `build_by_owner`(L249), `build_groupings`(L510).

---

## 6. 전체 요약 탭 (브랜드×카테고리×월 매트릭스)

가장 복잡한 기능. 첫 화면(맨 왼쪽 탭). 한 장 슬라이드로 export.

### 6-1. 데이터: `transform.build_summary(data)` → dict (L411)
반환 dict 구조:
```python
{
  "main_title", "months", "quarters", "seasons",
  "brands": [ {"brand", "rows":[{"label","cells","total"}], "total_q":[...], "total"} ],
  "grand_q":[...], "grand_season":[...], "grand_total"
}
```
- `cells` = 월컬럼(12개) 리스트, 각 칸은 **entry dict 리스트**:
  `{"key": "owner#no", "abbr": 첫단어, "rev": "X.X억", "sku": int}`
- **월 컬럼 매핑(`_launch_month_col`)**: 출시 타임라인 인덱스 → `SUMMARY_MONTHS`(12열, '25.12+1월 합침, 12월=27' 의미).
- **시리즈 병합(`_merge_cell`, L351)**: 같은 셀 내 동일 (담당자,No.)는 1개로 — 시리즈명 첫단어 라벨 + 매출(`_sum_rev` 억 합산) + SKU 합산. 단독은 프로젝트 첫단어.
- **브랜드별 카테고리 행(`_brand_row_specs`, L376)**: `config.SUMMARY_BRAND_CATEGORIES`로 정의. `"__rest__"`=명시 외 전부. **미정의 브랜드(예 라메종)는 실제 카테고리로 동적 구성** + '온라인' 행.
- 소계: 표시는 병합 기준, **집계(프로젝트수/SKU)는 실제 제품 기준** 유지.
- 시즌 SS=1Q+2Q, FW=3Q+4Q, 27'SS=12월.

### 6-2. 프로젝트명 색상 + 클릭 토글 (★ 최근 기능)
- **기본 네이비(`SUMMARY_NAME_COLOR`=1F3864), 클릭 시 빨강(`SUMMARY_NAME_IMPORTANT`=FF0000, 중요표시), 재클릭 시 네이비. 매출/SKU는 항상 흑색. PPTX에도 동일 반영.**
- **상태 저장 = URL 쿼리파라미터** (`?imp=key&imp=key2...`).
  - ⚠️ **중요한 교훈**: 처음엔 `session_state`에 저장했으나, `<a href target="_self">` 클릭이 **전체 페이지 이동 → Streamlit 세션 리셋** → 한 방향만 토글됨. URL에 담아 해결.
  - `app.py`: `important = set(st.query_params.get_all("imp"))` (L25).
- **웹(`render_html.py`)**: `_summary_cell_html`(L197)이 각 이름을 링크로 — 클릭 시 그 key를 **토글한 전체 집합**을 `_imp_href`(L188)로 인코딩한 href 생성(양방향 토글).
- **PPTX(`pptx_export.py`)**: `_set_summary_cell`(L67)이 `important` 받아 이름 run은 빨강/네이비, 매출/SKU run은 흑색.
- export도 같은 `important`를 읽음: `build_summary_pptx_bytes(summary, important)` (app L99).

### 6-3. PPTX 1슬라이드 맞춤: `_add_summary_slide` (L470)
- 무조건 한 장. **행별 최대 줄 수 합산 → 폰트 자동 산정**: `font = avail_pt / (total_lines*1.32)`, `[SUMMARY_FONT_MIN=4, MAX=7]`로 클램프. 한 셀 줄수 `_celllines = 2*len(entries)`.
- 행 높이는 행별 줄 수 비례(비균일).
- 표 아래 축약 규칙 안내문 `SUMMARY_ABBR_NOTE`(웹·PPTX 둘 다).

### 6-4. 웹 가독성 (PPTX와 별개)
`SUMMARY_WEB_FONT/CELL_FONT/PAD/LINE/MIN_WIDTH` — 웹 표만 키움. PPTX 1슬라이드 4~7pt는 유지.

---

## 7. PPTX 상세 슬라이드 (`pptx_export.py`)

A4 가로(297×210mm). 소제목 1개 = 슬라이드 1장(길면 분할).

- **양식(`_add_title_block` L169, `_add_footer` L193)**: 주제목 좌상단 고정 + 아래 장식선(181818,1pt) → 소제목('• '+주황 F79646, 표 바로 위). 하단 장식선 + 'SHINSEGAE CASA'(`FOOTER_*`).
- **표(`_build_table` L236)**: 24열(정보10 + 13월 + 비고). 헤더 2행 병합, NO 세로 병합, 회색 테두리(`_set_cell_border` L150, tcPr 맨 앞 lnL/R/T/B 삽입).
- **발주/출시 주황 텍스트(`_text_color` L140)**: `ORANGE_TEXT_EXACT`('발주' 정확히)·`ORANGE_TEXT_CONTAINS`('출시' 포함). '1차 샘플 발주'는 검정 유지.
- **진행 화살표(`_add_launch_arrows` L342)**: 기획→출시 직전, 표 '뒤'(z-order를 graphicFrame 앞으로 이동). 그라데이션은 `_apply_gradient`(L94) — XML 순서 주의(gradFill은 geometry 뒤·ln 앞).
- **오늘 월 강조(`_add_month_highlight` L323)**: `current_month_index()`로 오늘 (연,월)→해당 열에 네이비(002060) 사각 테두리. 매 실행 자동 갱신, 범위(2025.12~2026.12) 밖이면 없음.
- **페이지 분할(`_paginate_rows` L399)**: 상/하 장식선 사이 용량(capacity) 초과 시 시리즈 블록째 다음 슬라이드. 소제목에 (i/N), 요약행은 마지막 페이지만, NO 연속.

---

## 8. 주요 설정값 (`config.py`)

| 키 | 값 / 의미 |
|----|-----------|
| `MAIN_TITLE` | "26SS,FW 신상품 출시 스케줄" (고정 텍스트, 변경 가능) |
| `DATA_SOURCE` / `SHEETS_MODE` | "sheets" / "public" |
| `SHEETS_KEY` | 18yYM6Spbpmn2zEa-IyLuhWZV8xTzNvaqrd_a7cNTQi0 |
| `PRICE_DIVISOR` | 10000 (원→만원) |
| `CATEGORY_ORDER` | 매트리스·베드룸·수납장·홈오피스·키즈 |
| `BRAND_ORDER` | 마테라소·까사미아 |
| `SPLIT_BY_BRAND_CATEGORIES` | {매트리스, 베드룸} |
| `SUMMARY_BRAND_ORDER` | 마테라소·라메종·까사미아 |
| `SUMMARY_BRAND_CATEGORIES` | 마테라소: 매트리스/가구(__rest__) · 까사미아: 베드룸/매트리스/드레스룸←수납장/홈오피스주니어룸←홈오피스+키즈 |
| `SUMMARY_MONTHS` | 12열, '25.12+1월 합침, 12월=27' |
| `SUMMARY_NAME_COLOR` / `_IMPORTANT` / `SUMMARY_DETAIL_COLOR` | 1F3864(네이비) / FF0000(빨강) / 000000(흑) |
| `SUMMARY_FONT_MAX/MIN` | 7 / 4 |
| `HEADER_FILL` / `SUMMARY_FILL` / `BORDER_COLOR` | E7E7E6 / ECE8E2 / BFBFBF |
| `ORANGE_TEXT_COLOR` / `CURRENT_MONTH_COLOR` | 984807 / 002060 |
| `FOOTER_TEXT` | "SHINSEGAE CASA" |

---

## 9. 배포 & 운영

- 코드 수정 → `git push` → Streamlit Cloud 자동 재배포.
- 팀원이 시트 수정 → 앱에서 **'🔄 데이터 새로고침'**(캐시 클리어 + nonce 증가).
- 공개 시트라 비밀키 불필요. `service_account.json`·`*.bak.xlsx`·`*.pptx`는 `.gitignore`.
- 새 이미지 반영도 '데이터 새로고침'으로.

### 보안 제약 (대신 수행 불가 — 사용자가 직접)
GitHub/Streamlit 로그인·비밀번호 입력·계정 생성·OAuth 승인은 사용자가 직접. 서비스계정 키는 절대 커밋 금지.

---

## 10. 과거에 밟은 함정 (재발 방지)

1. **session_state 토글 실패** → 링크 클릭=전체이동=세션리셋. URL 쿼리파라미터로 해결(§6-2).
2. **python-pptx 음수 인덱싱** 실패 → `n_rows-1` 사용.
3. **그라데이션이 파랑(테마)으로** → XML 순서(gradFill: geometry 뒤·ln 앞), fillRef 제거 시 파일 깨짐(`_apply_gradient`).
4. **Streamlit 모듈 핫리로드가 stale config** 표시 → 서버 전체 재시작 필요.
5. **cwd가 dashboard/** 라 파일 못 찾음 → `_resolve_path`(cwd→스크립트→상위 탐색).
6. 회색 테두리/장식선 색은 예시 PPTX 픽셀 샘플값(181818, BFBFBF).

---

## 11. 현재 상태

- M1~M4 + 배포 + 전체요약(시리즈 병합·프로젝트명 색상 토글)까지 완료·검증됨.
- 최근 커밋: `cd3ad80` (전체요약 시리즈 병합 + 축약 안내문).
- **열린 작업 없음.** 다음 사용자 요청 대기.
