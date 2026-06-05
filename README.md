# 신제품 출시 스케줄 대시보드

law data(스프레드시트)를 읽어 **담당자별 출시 스케줄 대시보드**를 실시간 시각화하고,
팀장이 수작업으로 수정할 수 있는 **A4 가로 PPTX**로 내보내는 프로그램입니다.

```
[law data (xlsx / Google Sheet)]
        │  data_loader / sheets_loader  (행별 표준화 + 셀 이미지 추출)
        ▼
   transform  (시리즈 그룹화·만원 변환·제조처 결합·분기 요약 …)  ← 화면·PPTX 공용
        ├──────────────┐
        ▼              ▼
   app.py (웹 대시보드)   pptx_export.py (A4 가로 PPTX)
```

화면과 PPTX는 **동일한 가공 결과(transform)** 를 사용하므로 항상 일치합니다.

---

## 1. 설치

```powershell
cd dashboard
pip install -r requirements.txt
```

## 2. 실행 (웹 대시보드)

```powershell
streamlit run app.py
```

브라우저가 자동으로 열립니다.
- **담당자 선택**: 사이드바에서 표시할 담당자 선택(기본 전체)
- **🔄 데이터 새로고침**: law data 파일이 갱신되면 클릭해 최신 반영(파일 수정 시각 자동 감지)
- **📥 PPTX 내보내기**: 선택한 담당자들을 A4 가로 PPTX로 다운로드

## 3. PPTX만 생성 (CLI)

```powershell
python pptx_export.py "..\신제품 출시 스케줄 관리_law data.xlsx" "대시보드_export.pptx"
```

---

## 4. law data 양식 규칙

각 **담당자 = 시트 1개**. 시트 구조(행 번호는 `config.py`에서 변경 가능):

| 행 | 내용 |
|----|------|
| 3행 | 컬럼 헤더 (`No.`, `카테고리`, `브랜드`, `프로젝트명`, `시리즈명`, `담당자`, `오프/온라인`, `사이즈 수`, `컬러 수`, `SKU 수`, `목표매출`, `업체`, `업체 국가`, `주요소재`, `목표 판매가`, `일정`, `비고`, **`이미지`**) |
| 4행 | 분기 헤더 (`25`, `26 1Q`, `26 2Q`, `26 3Q`, `26 4Q`) |
| 5행 | 월 헤더 (`12`, `1`, … `12` = 13개월) |
| 6행~ | 제품 데이터. 일정 열(13개월)에 마일스톤(`기획`, `1차 샘플 발주`, `출시` 등) 입력 |

- **컬럼은 헤더 텍스트로 인식**하므로 열 순서/추가에 영향받지 않습니다.
- **`이미지` 열**: 해당 제품 행의 이미지 열 셀에 이미지를 삽입하면 대시보드·PPTX에 반영됩니다.
  - 로컬 xlsx: Excel에서 셀에 이미지 삽입(임베드) → 행 앵커로 자동 매핑
  - Google Sheets: `=IMAGE("이미지URL")` 수식 권장(아래 6절)
- **같은 `No.` 가 연속**되면 한 시리즈로 묶여 NO 셀이 세로 병합됩니다.

### 자동 가공 규칙 (예시 대시보드 기준)
| 대시보드 | 가공 |
|---|---|
| 시리즈명 | `시리즈명` + `프로젝트명` (시리즈명이 `-`면 프로젝트명만) |
| 목표판매가 | 원 → **만원** (590,000원 → 59). 매트리스·베드프레임은 `(Q)` 표기 |
| 제조처 | `업체` + `(업체 국가)` |
| 요약행 | `N시리즈 (총 SKU)` + 분기별 `출시 시리즈수(SKU)` |

설정값은 모두 [`config.py`](config.py)에서 변경합니다(주제목, 단위, 색상, A4 여백, 폰트 등).

---

## 5. 파일 구성

| 파일 | 역할 |
|------|------|
| `config.py` | 전역 설정(주제목·단위·색상·페이지·폰트) |
| `schema.py` | 표준 데이터 모델 |
| `data_loader.py` | **로컬 xlsx** 읽기 + 셀 이미지 추출 |
| `sheets_public_loader.py` | **공개 링크** Google Sheets 읽기(인증 불필요) |
| `sheets_loader.py` | **서비스 계정** Google Sheets 읽기(비공개) |
| `transform.py` | 가공 로직 (화면·PPTX 공용) |
| `render_html.py` | 화면용 HTML 표 |
| `pptx_export.py` | A4 가로 PPTX 생성 |
| `app.py` | Streamlit 웹 앱 |

---

## 6. Google Sheets 연결

`config.py` 의 `DATA_SOURCE`(또는 앱 사이드바의 **데이터 소스** 토글)로 전환합니다.
두 가지 방식 모두 **무료**입니다.

### 방식 A) 공개 링크 — 가장 간단 (결제·Cloud 불필요) ★기본
1. 구글 시트 **공유 → 일반 액세스 → "링크가 있는 모든 사용자(뷰어)"**
2. `config.py`:
   ```python
   DATA_SOURCE = "sheets"
   SHEETS_MODE = "public"
   SHEETS_KEY  = "스프레드시트_ID"   # URL 의 /d/ 와 /edit 사이 값
   ```
3. 끝. 앱이 export URL 로 시트를 xlsx 로 받아 그대로 읽습니다(`sheets_public_loader.py`).
   - ⚠️ 링크를 아는 사람은 시트를 열람할 수 있습니다(내부 데이터 주의).

### 방식 B) 서비스 계정 — 비공개 유지 (무료, Cloud 설정 필요)
1. **Google Cloud 콘솔** → 프로젝트 생성 → *Google Sheets API* 사용 설정
2. **서비스 계정** 생성 → JSON 키 다운로드 → `dashboard/service_account.json`
3. 대상 시트를 **서비스 계정 이메일과 공유**(뷰어)
4. `config.py`:
   ```python
   DATA_SOURCE  = "sheets"
   SHEETS_MODE  = "service_account"
   SHEETS_KEY   = "스프레드시트_ID"
   SHEETS_CREDS = "service_account.json"
   ```

### 이미지 (Google Sheets)
- 공개 링크(방식 A): 시트의 in-cell 이미지/`=IMAGE()` 는 xlsx export 시 임베드되지
  않아 **텍스트 위주**로 반영됩니다. 제품 이미지가 필요하면 로컬 xlsx 운용 또는
  별도 처리(추후 지원)를 권장합니다.
- 텍스트 데이터(일정·가격·제조처 등)는 시트에서 실시간 그대로 반영됩니다.

---

## 참고 / 확정 대기 항목
- **분기 요약 SKU 분배**: 9월 출시 제품의 분기 경계 처리(달력 분기 기준). 시리즈 수·총합은 예시와 일치.
- **발주→출시 화살표**(예시의 그라데이션 바)는 장식 요소로, 현재는 평면 셀로 표현(선택적 추가 가능).
