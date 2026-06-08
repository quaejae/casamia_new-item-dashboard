# -*- coding: utf-8 -*-
"""DashboardTable → HTML 표 (Streamlit 화면용).

PPTX export 와 동일한 가공 결과(transform)를 사용하므로 화면과 PPTX가 일치한다.
예시 색상(헤더 E7E7E6, 출시 FFF2CC)·헤더 병합·NO 세로병합·이미지·요약행을 재현.
"""
from __future__ import annotations

import base64
import html
from typing import Optional

import config
from schema import DashboardTable
from transform import launch_arrow_span, current_month_index


def _month_hl(active: bool, top: bool = False, bottom: bool = False) -> str:
    """현재 월 열 강조용 box-shadow 스타일(좌우 + 선택적 상/하 네이비 선)."""
    if not active:
        return ""
    c = config.CURRENT_MONTH_COLOR
    parts = [f"inset 2px 0 0 #{c}", f"inset -2px 0 0 #{c}"]
    if top:
        parts.append(f"inset 0 2px 0 #{c}")
    if bottom:
        parts.append(f"inset 0 -2px 0 #{c}")
    return "box-shadow:" + ",".join(parts) + ";"

INFO_HEADERS = ["NO", "시리즈명", "이미지", "소재", "목표<br>판매가",
                "목표<br>매출", "컬러", "SKU수", "제조처", "담당자"]


def _milestone_fill(text: str) -> Optional[str]:
    for kw, color in config.MILESTONE_COLORS.items():
        if kw in text:
            return color
    return None


def _is_orange(text: str) -> bool:
    """발주(정확히)/출시(포함) 마일스톤이면 True."""
    t = (text or "").strip()
    if t in config.ORANGE_TEXT_EXACT:
        return True
    return any(kw in t for kw in config.ORANGE_TEXT_CONTAINS)


def _esc(s: str) -> str:
    return html.escape(str(s)).replace("\n", "<br>")


def _lerp_hex(t: float, a: str, b: str) -> str:
    """색 a→b 사이를 t(0~1) 비율로 보간한 hex."""
    ar, ag, ab = int(a[0:2], 16), int(a[2:4], 16), int(a[4:6], 16)
    br, bg, bb = int(b[0:2], 16), int(b[2:4], 16), int(b[4:6], 16)
    r = round(ar + (br - ar) * t)
    g = round(ag + (bg - ag) * t)
    bl = round(ab + (bb - ab) * t)
    return f"{r:02X}{g:02X}{bl:02X}"


def _img_tag(data: bytes, ext: Optional[str]) -> str:
    b64 = base64.b64encode(data).decode("ascii")
    mime = f"image/{(ext or 'png').replace('jpg', 'jpeg')}"
    return f'<img src="data:{mime};base64,{b64}" style="max-width:60px;max-height:48px;object-fit:contain;" />'


def build_html_table(table: DashboardTable) -> str:
    css = """
    <style>
    .dash-wrap { overflow-x:auto; }
    table.dash { border-collapse:collapse; font-family:'맑은 고딕',sans-serif;
                 font-size:11px; width:100%; table-layout:fixed; }
    table.dash th, table.dash td { border:1px solid #BFBFBF; padding:2px 3px;
                 text-align:center; vertical-align:middle; word-break:break-all; }
    table.dash th { background:#__HEADER__; font-weight:700; }
    table.dash td.summary { background:#__SUMMARY__; font-weight:700;
        font-size:13px; height:26px; }
    table.dash td.no { font-weight:700; }
    table.dash td.orange { color:#__ORANGE__; font-weight:700; }
    table.dash td.launch { position:relative; }
    table.dash td.launch .ah { position:absolute; left:2px; top:50%;
        transform:translateY(-50%); width:0; height:0;
        border-top:6px solid transparent; border-bottom:6px solid transparent;
        border-left:9px solid #__HEAD__; z-index:3; }
    table.dash td.launch .lab { position:relative; z-index:2; }
    </style>
    """.replace("__HEAD__", config.WEB_ARROW_HEAD) \
       .replace("__HEADER__", config.HEADER_FILL) \
       .replace("__SUMMARY__", config.SUMMARY_FILL) \
       .replace("__ORANGE__", config.ORANGE_TEXT_COLOR)
    months = config.TIMELINE_MONTHS
    n_month = len(months)
    cur = current_month_index() if config.CURRENT_MONTH_HIGHLIGHT else None

    # ── 헤더 2행
    head = ['<tr>']
    for h in INFO_HEADERS:
        label = h
        if h == "목표<br>판매가" and table.show_q_label:
            label = h + "<br>" + config.Q_SIZE_LABEL
        head.append(f'<th rowspan="2">{label}</th>')
    for qlabel, _start, length in config.QUARTERS:
        head.append(f'<th colspan="{length}">{_esc(qlabel)}</th>')
    head.append('<th rowspan="2">비고</th>')
    head.append('</tr>')
    head.append('<tr>')
    for mi, (_y, m) in enumerate(months):
        hl = _month_hl(mi == cur, top=True)
        style = f' style="{hl}"' if hl else ''
        head.append(f'<th{style}>{m}</th>')
    head.append('</tr>')

    # ── NO 세로병합 rowspan 계산
    rows = table.rows
    rowspans = [1] * len(rows)
    i = 0
    while i < len(rows):
        j = i + 1
        while j < len(rows) and rows[j].no_label == "":
            j += 1
        rowspans[i] = j - i
        i = j

    body = []
    for idx, r in enumerate(rows):
        body.append('<tr>')
        if r.no_label != "" or idx == 0:
            body.append(f'<td class="no" rowspan="{rowspans[idx]}">{_esc(r.no_label)}</td>')
        body.append(f'<td>{_esc(r.series_cell)}</td>')
        cell_img = _img_tag(r.image, r.image_ext) if r.image else ""
        body.append(f'<td>{cell_img}</td>')
        body.append(f'<td>{_esc(r.material)}</td>')
        body.append(f'<td>{_esc(r.target_price)}</td>')
        body.append(f'<td>{_esc(r.target_revenue)}</td>')
        body.append(f'<td>{_esc(r.color_count)}</td>')
        body.append(f'<td>{_esc(r.sku_count)}</td>')
        body.append(f'<td>{_esc(r.maker)}</td>')
        body.append(f'<td>{_esc(r.owner)}</td>')
        # 진행기간 화살표 구간(기획→출시 직전) — 셀 뒤 그라데이션 밴드로 표현
        span = launch_arrow_span(r.timeline) if config.ARROW_ENABLED else None
        if span:
            s0, s1 = span
            total = s1 - s0 + 1
        is_last = (idx == len(rows) - 1)
        for mi in range(n_month):
            txt = r.timeline[mi] if mi < len(r.timeline) else ""
            fill = _milestone_fill(txt)
            ocls = " orange" if _is_orange(txt) else ""
            hl = _month_hl(mi == cur, bottom=is_last)
            if fill:
                # 출시 칸: 화살촉을 칸 안 왼쪽(출시 텍스트 왼쪽)에 배치
                arrow = '<span class="ah"></span>' if (span and mi == s1 + 1) else ''
                body.append(
                    f'<td class="launch{ocls}" style="background:#{fill};{hl}">'
                    f'{arrow}<span class="lab">{_esc(txt)}</span></td>')
            elif span and s0 <= mi <= s1:
                bf, bt = config.WEB_ARROW_BAND_FROM, config.WEB_ARROW_BAND_TO
                c1 = _lerp_hex((mi - s0) / total, bf, bt)
                c2 = _lerp_hex((mi - s0 + 1) / total, bf, bt)
                body.append(
                    f'<td class="{ocls.strip()}" style="background:linear-gradient(to right,#{c1},#{c2});{hl}">'
                    f'{_esc(txt)}</td>')
            else:
                body.append(f'<td class="{ocls.strip()}" style="{hl}">{_esc(txt)}</td>')
        body.append(f'<td>{_esc(r.note)}</td>')
        body.append('</tr>')

    # ── 요약행
    s = table.summary
    summ = ['<tr>']
    summ.append(f'<td class="summary" colspan="10">{_esc(s.label)}</td>')
    # 분기별 칸 병합(colspan)
    for qi, (_q, _start, length) in enumerate(config.QUARTERS):
        summ.append(f'<td class="summary" colspan="{length}">{_esc(s.per_quarter[qi])}</td>')
    summ.append(f'<td class="summary">{s.series_count}({s.total_sku})</td>')
    summ.append('</tr>')

    table_html = (
        '<div class="dash-wrap"><table class="dash">'
        + "".join(head) + "".join(body) + "".join(summ)
        + '</table></div>'
    )
    return css + table_html


def build_summary_html(S: dict) -> str:
    """전체 요약 매트릭스(브랜드×카테고리×월) → HTML 표."""
    css = """
    <style>
    .sum-wrap { overflow-x:auto; }
    table.sum { border-collapse:collapse; font-family:'맑은 고딕',sans-serif;
                font-size:10px; width:100%; table-layout:fixed; }
    table.sum th, table.sum td { border:1px solid #BFBFBF; padding:1px 2px;
                text-align:center; vertical-align:middle; word-break:break-all;
                line-height:1.15; }
    table.sum th { background:#E7E7E6; font-weight:700; }
    table.sum td.brand { background:#EFEFEF; font-weight:700; }
    table.sum td.cat { background:#F6F6F6; font-weight:700; }
    table.sum td.tot { background:#ECE8E2; font-weight:700; }
    table.sum td.mcell { font-size:9px; }
    </style>
    """
    months = S["months"]; quarters = S["quarters"]; seasons = S["seasons"]
    h = ['<div class="sum-wrap"><table class="sum">']
    # 헤더 2행
    h.append('<tr><th rowspan="2">브랜드</th><th rowspan="2">카테고리</th>')
    for qlbl, _st, sp in quarters:
        h.append(f'<th colspan="{sp}">{_esc(qlbl)}</th>')
    h.append('<th rowspan="2">계</th></tr>')
    h.append('<tr>')
    for mlbl, _idx in months:
        h.append(f'<th>{_esc(mlbl)}</th>')
    h.append('</tr>')
    # 브랜드 블록
    for b in S["brands"]:
        nblock = len(b["rows"]) + 1
        first = True
        for r in b["rows"]:
            h.append('<tr>')
            if first:
                h.append(f'<td rowspan="{nblock}" class="brand">{_esc(b["brand"])}</td>')
                first = False
            h.append(f'<td class="cat">{_esc(r["label"])}</td>')
            for c in r["cells"]:
                h.append(f'<td class="mcell">{_esc(c)}</td>')
            h.append(f'<td class="tot">{_esc(r["total"])}</td>')
            h.append('</tr>')
        # 브랜드 계 행(분기별)
        h.append('<tr><td class="cat tot">계</td>')
        for qi, (_q, _st, sp) in enumerate(quarters):
            h.append(f'<td colspan="{sp}" class="tot">{_esc(b["total_q"][qi])}</td>')
        h.append(f'<td class="tot">{_esc(b["total"])}</td></tr>')
    # 하단 그랜드: 분기 / 시즌
    h.append('<tr><td rowspan="2" colspan="2" class="tot">총계</td>')
    for qi, (qlbl, _st, sp) in enumerate(quarters):
        h.append(f'<td colspan="{sp}" class="tot">{_esc(qlbl)} {_esc(S["grand_q"][qi])}</td>')
    h.append(f'<td rowspan="2" class="tot">{_esc(S["grand_total"])}</td></tr>')
    h.append('<tr>')
    for si, (slbl, _st, sp) in enumerate(seasons):
        h.append(f'<td colspan="{sp}" class="tot">{_esc(slbl)} {_esc(S["grand_season"][si])}</td>')
    h.append('</tr>')
    h.append('</table></div>')
    return css + "".join(h)
