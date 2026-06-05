# -*- coding: utf-8 -*-
"""신제품 출시 스케줄 대시보드 — Streamlit 웹 앱.

실행:  streamlit run app.py
  - law data(시트/xlsx)를 읽어 여러 기준(탭)으로 대시보드를 실시간 시각화
  - '데이터 새로고침'으로 최신 시트 반영
  - 'PPTX 내보내기'로 선택한 기준의 A4 가로 PPTX 다운로드
"""
from __future__ import annotations

import os

import streamlit as st

import config
import data_loader
import transform
import pptx_export
from render_html import build_html_table

st.set_page_config(page_title="신제품 출시 스케줄 대시보드", layout="wide")


@st.cache_data(show_spinner=False)
def _load_data_xlsx(path: str, mtime: float):
    """경로+수정시각을 키로 캐싱. mtime 이 바뀌면 자동 재로딩."""
    return data_loader.load(path)


@st.cache_data(show_spinner=False)
def _load_data_sheets(key: str, mode: str, creds: str, _nonce: int):
    if mode == "service_account":
        import sheets_loader
        return sheets_loader.load(key, creds)
    import sheets_public_loader
    return sheets_public_loader.load(key)


def _resolve_path(path: str) -> str:
    if os.path.isabs(path):
        return path
    here = os.path.dirname(os.path.abspath(__file__))
    for base in (os.getcwd(), here, os.path.dirname(here)):
        cand = os.path.join(base, path)
        if os.path.exists(cand):
            return cand
    return os.path.join(os.getcwd(), path)


# ── 사이드바: 데이터 소스 ─────────────────────────────────────────────────────
st.sidebar.header("⚙️ 설정")
source = st.sidebar.radio("데이터 소스", ["xlsx", "sheets"],
                          index=0 if config.DATA_SOURCE == "xlsx" else 1,
                          format_func=lambda s: "로컬 xlsx" if s == "xlsx" else "Google Sheets")

if "refresh_nonce" not in st.session_state:
    st.session_state.refresh_nonce = 0
if st.sidebar.button("🔄 데이터 새로고침", use_container_width=True):
    _load_data_xlsx.clear(); _load_data_sheets.clear()
    st.session_state.refresh_nonce += 1
    st.rerun()

if source == "xlsx":
    src_path = st.sidebar.text_input("law data 파일 경로", value=config.LOCAL_XLSX)
    abs_path = _resolve_path(src_path)
    if not os.path.exists(abs_path):
        st.error(f"파일을 찾을 수 없습니다: {abs_path}")
        st.stop()
    mtime = os.path.getmtime(abs_path)
    data = _load_data_xlsx(abs_path, mtime)
    st.sidebar.caption(f"마지막 수정: {__import__('datetime').datetime.fromtimestamp(mtime):%Y-%m-%d %H:%M:%S}")
else:
    if not config.SHEETS_KEY:
        st.error("config.py 의 SHEETS_KEY 를 먼저 설정하세요. (README 6절)")
        st.stop()
    try:
        data = _load_data_sheets(config.SHEETS_KEY, config.SHEETS_MODE,
                                 config.SHEETS_CREDS, st.session_state.refresh_nonce)
    except Exception as e:
        st.error(f"Google Sheets 연결 오류: {e}")
        st.stop()
    mode_label = "공개 링크" if config.SHEETS_MODE == "public" else "서비스 계정"
    st.sidebar.caption(f"출처: Google Sheets ({mode_label})")

# ── 탭별 그룹화 ──────────────────────────────────────────────────────────────
groupings = transform.build_groupings(data)   # {탭이름: [DashboardTable,...]}
tab_names = list(groupings.keys())

# ── 사이드바: PPTX 내보내기 ──────────────────────────────────────────────────
st.sidebar.divider()
export_key = st.sidebar.selectbox("📥 PPTX 내보내기 기준", tab_names,
                                  index=len(tab_names) - 1)  # 기본 담당자별
export_tables = groupings.get(export_key, [])
if export_tables:
    pptx_bytes = pptx_export.build_pptx_bytes(export_tables)
    st.sidebar.download_button(
        f"PPTX 내보내기 ({export_key}, A4 가로)",
        data=pptx_bytes,
        file_name=f"신제품_출시스케줄_대시보드_{export_key}.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        use_container_width=True,
    )

# ── 본문 ─────────────────────────────────────────────────────────────────────
st.title(config.MAIN_TITLE)
st.markdown(
    f"<hr style='border:none;border-top:1.5px solid #{config.TITLE_RULE_COLOR};"
    f"margin:0 0 6px 0;'>", unsafe_allow_html=True)
st.caption("정렬 기준 탭을 선택해 보세요 · 실시간 미리보기")

tabs = st.tabs(tab_names)
for tab, name in zip(tabs, tab_names):
    with tab:
        tables = groupings[name]
        if not tables:
            st.info("표시할 프로젝트가 없습니다.")
            continue
        for t in tables:
            st.markdown(
                f"<h3 style='color:#{config.SUBTITLE_COLOR};margin:10px 0 4px 0;'>"
                f"{config.SUBTITLE_PREFIX}{t.owner}</h3>", unsafe_allow_html=True)
            st.markdown(build_html_table(t), unsafe_allow_html=True)
            st.write("")
