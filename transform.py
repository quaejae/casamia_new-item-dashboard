# -*- coding: utf-8 -*-
"""가공 로직 — raw data(RawProduct) → 대시보드 표(DashboardTable).

화면(app.py)과 PPTX export(pptx_export.py)가 이 모듈을 공용으로 사용하므로,
화면에 보이는 표와 내보낸 PPTX가 항상 동일하게 유지된다.

가공 규칙(예시 PPTX 역설계 + 협의 결과):
  - 시리즈 그룹화: 연속된 동일 No. → 한 시리즈, NO 셀 병합(첫 행만 표기)
  - 시리즈명 셀: 시리즈명 + 프로젝트명(줄바꿈). 시리즈명이 '-'면 프로젝트명만
  - 목표판매가: 원 → 만원 (590,000원 → 59). Q사이즈 적용 시 헤더에 (Q)
  - 제조처: 업체 + (국가)
  - 타임라인: 13개월 그대로
  - 요약행: 'N시리즈 (총 SKU)' + 분기별 출시 (시리즈수(SKU))
"""
from __future__ import annotations

import datetime
import re
from typing import Dict, List, Optional

import config
from schema import DashboardRow, DashboardTable, RawProduct, SummaryRow


def current_month_index(today: Optional[datetime.date] = None) -> Optional[int]:
    """오늘 날짜의 (연,월)이 타임라인의 몇 번째 월 칸인지 반환(없으면 None)."""
    today = today or datetime.date.today()
    key = (today.year, today.month)
    for i, ym in enumerate(config.TIMELINE_MONTHS):
        if tuple(ym) == key:
            return i
    return None


def _price_to_manwon(raw: str) -> str:
    """'590,000원' → '59' (만원 단위). 숫자 없으면 원본 반환."""
    digits = re.sub(r"[^\d]", "", raw or "")
    if not digits:
        return raw or ""
    won = int(digits)
    val = won / config.PRICE_DIVISOR
    return str(int(val)) if val == int(val) else f"{val:g}"


def _abbrev(name: str) -> str:
    for full, short in config.PROJECT_ABBREV.items():
        name = name.replace(full, short)
    return name


def _series_cell(p: RawProduct) -> str:
    project = _abbrev(p.project)
    series = (p.series or "").strip()
    if series and series != "-":
        return f"{series}\n{project}"
    return project


def launch_arrow_span(timeline):
    """진행 기간 화살표 구간 계산: 기획 → 출시 직전 칸.

    반환 (start_idx, end_idx): start_idx=기획(없으면 첫 마일스톤),
    end_idx=출시 바로 앞 칸. 화살표 본체는 start_idx ~ end_idx 전체를 덮는다.
    출시 없음/구간 없음이면 None.
    """
    launch_idx = None
    for i, c in enumerate(timeline):
        if c and config.LAUNCH_KEYWORD in c:
            launch_idx = i
    if launch_idx is None or launch_idx == 0:
        return None
    # 시작: '기획' 칸 우선, 없으면 출시 이전 첫 마일스톤
    start_idx = None
    for i in range(launch_idx):
        if timeline[i] and config.PLAN_KEYWORD in timeline[i]:
            start_idx = i
            break
    if start_idx is None:
        for i in range(launch_idx):
            if timeline[i]:
                start_idx = i
                break
    if start_idx is None:
        return None
    end_idx = launch_idx - 1
    if end_idx < start_idx:
        return None
    return (start_idx, end_idx)


def _maker_cell(p: RawProduct) -> str:
    if p.vendor and p.vendor_country:
        return f"{p.vendor}\n({p.vendor_country})"
    return p.vendor or p.vendor_country or ""


def _num_str(v) -> str:
    if v is None:
        return ""
    return str(int(v)) if float(v).is_integer() else str(v)


def _is_q_size(p: RawProduct) -> bool:
    """매트리스·베드프레임 제품인지(목표판매가 Q사이즈 기준 여부)."""
    if p.category in config.Q_SIZE_CATEGORIES:
        return True
    return any(kw in p.project for kw in config.Q_SIZE_KEYWORDS)


def _month_to_quarter_idx(month_idx: int) -> int:
    """타임라인 월 인덱스(0~12) → QUARTERS 인덱스."""
    for qi, (_, start, length) in enumerate(config.QUARTERS):
        if start <= month_idx < start + length:
            return qi
    return len(config.QUARTERS) - 1


def _series_key(p: RawProduct, idx: int):
    """시리즈 식별 키 = (담당자, No.). No. 없으면 행마다 고유."""
    return (p.owner, p.no) if p.no is not None else ("__u__", idx)


def _launch_sort_key(p: RawProduct):
    """출시일 정렬 키 = (출시 월 인덱스, 일). 출시 없으면 맨 뒤로."""
    mi = None
    day = 0
    for i, c in enumerate(p.timeline):
        if c and config.LAUNCH_KEYWORD in c:
            mi = i
            m = re.search(r"\((\d+)/(\d+)\)", c)
            if m:
                day = int(m.group(2))
    return (99, 99) if mi is None else (mi, day)


def _sort_by_launch(products: List[RawProduct]) -> List[RawProduct]:
    """출시일 빠른 순(위)→늦은 순(아래)으로 정렬.

    시리즈(연속 동일 (담당자,No.)) 블록을 깨지 않고, 블록 단위로 정렬한다.
    (이 데이터의 모든 시리즈는 구성원 출시월이 동일해 행 단위 정렬과 결과가 같음)
    """
    blocks: List[List[RawProduct]] = []
    cur: List[RawProduct] = []
    cur_key = object()
    for i, p in enumerate(products):
        k = _series_key(p, i)
        if cur and k != cur_key:
            blocks.append(cur); cur = []
        cur.append(p); cur_key = k
    if cur:
        blocks.append(cur)
    blocks.sort(key=lambda b: min(_launch_sort_key(p) for p in b))  # stable
    out: List[RawProduct] = []
    for b in blocks:
        out.extend(b)
    return out


def _build_summary(products: List[RawProduct]) -> SummaryRow:
    # 시리즈 수 = 연속된 서로 다른 (담당자,No.) 그룹 수
    series_count = 0
    prev = object()
    for i, p in enumerate(products):
        k = _series_key(p, i)
        if k != prev:
            series_count += 1
            prev = k

    total_sku = sum(int(p.sku_count) for p in products if p.sku_count)
    label = f"{series_count}시리즈 ({total_sku} SKU)"

    # 분기별 출시 집계: 출시 마일스톤이 있는 월의 분기에 (시리즈수, SKU) 누적
    q_series = [set() for _ in config.QUARTERS]
    q_sku = [0 for _ in config.QUARTERS]
    for i, p in enumerate(products):
        for mi, cell in enumerate(p.timeline):
            if cell and config.LAUNCH_KEYWORD in cell:
                qi = _month_to_quarter_idx(mi)
                q_series[qi].add(_series_key(p, i))
                q_sku[qi] += int(p.sku_count) if p.sku_count else 0
                break  # 제품당 출시 1회만 집계

    per_quarter = []
    for qi in range(len(config.QUARTERS)):
        if q_series[qi]:
            per_quarter.append(f"{len(q_series[qi])}({q_sku[qi]})")
        else:
            per_quarter.append("-")
    return SummaryRow(label=label, per_quarter=per_quarter,
                      series_count=series_count, total_sku=total_sku)


def build_table(label: str, products: List[RawProduct],
                renumber: bool = True) -> DashboardTable:
    """제품 목록 → DashboardTable.

    label: 소제목(담당자명 또는 그룹명).
    항상 출시일 빠른 순(위→아래)으로 정렬한다.
    renumber=True(기본) 면 정렬된 순서대로 NO를 1,2,3… 재부여.
    """
    products = _sort_by_launch(products)
    rows: List[DashboardRow] = []
    prev_key = object()
    seq = 0
    for i, p in enumerate(products):
        key = _series_key(p, i)
        no_label = ""
        if key != prev_key:
            seq += 1
            no_label = str(seq) if renumber else (
                str(p.no) if p.no is not None else "")
            prev_key = key

        rows.append(DashboardRow(
            no_label=no_label,
            series_cell=_series_cell(p),
            image=p.image,
            image_ext=p.image_ext,
            material=p.material,
            target_price=_price_to_manwon(p.target_price_raw),
            target_revenue=p.target_revenue,
            color_count=_num_str(p.color_count),
            sku_count=_num_str(p.sku_count),
            maker=_maker_cell(p),
            owner=p.owner,
            timeline=list(p.timeline),
            note=p.note,
        ))

    show_q = any(_is_q_size(p) for p in products)
    return DashboardTable(
        owner=label,
        main_title=config.MAIN_TITLE,
        show_q_label=show_q,
        rows=rows,
        summary=_build_summary(products),
    )


def _all_products(data: Dict[str, List[RawProduct]]) -> List[RawProduct]:
    """담당자 순서·행 순서를 보존하며 모든 제품을 평탄화(시리즈 인접 유지)."""
    out: List[RawProduct] = []
    for prods in data.values():
        out.extend(prods)
    return out


# ── 탭별 그룹화 ──────────────────────────────────────────────────────────────
def build_by_owner(data: Dict[str, List[RawProduct]]) -> List[DashboardTable]:
    """담당자별: {담당자: [...]} → 담당자당 1표 (기존 동작)."""
    return [build_table(owner, prods) for owner, prods in data.items()]


# 하위 호환 별칭
build_all = build_by_owner


def _group(allp, ordered_specs) -> List[DashboardTable]:
    """ordered_specs = [(label, predicate), ...] 순서대로 비어있지 않은 그룹 표 생성."""
    tables = []
    for label, pred in ordered_specs:
        sel = [p for p in allp if pred(p)]
        if sel:
            tables.append(build_table(label, sel, renumber=True))
    return tables


def _ordered_unique(values, preferred):
    """preferred 순서대로(존재하는 것만) + 그 외 데이터 등장 순으로 추가(중복 제거).

    데이터에 새 값이 생겨도 누락 없이 자동 포함된다.
    """
    seen = list(dict.fromkeys(v for v in values))   # 첫 등장 순 유니크
    out = [v for v in preferred if v in seen]
    out += [v for v in seen if v not in preferred]
    return out


def _lbl(v) -> str:
    return v if v else config.UNCLASSIFIED_LABEL


def build_by_category(data) -> List[DashboardTable]:
    """카테고리별(설정의 SPLIT 카테고리는 브랜드로 세분). 신규 카테고리/브랜드 자동 포함."""
    allp = _all_products(data)
    cats = _ordered_unique((p.category for p in allp), config.CATEGORY_ORDER)
    brands = _ordered_unique((p.brand for p in allp), config.BRAND_ORDER)
    specs = []
    for cat in cats:
        if cat in config.SPLIT_BY_BRAND_CATEGORIES:
            present = [b for b in brands
                       if any(p.category == cat and p.brand == b for p in allp)]
            for b in present:
                specs.append((f"{_lbl(cat)}_{_lbl(b)}",
                              (lambda c, br: lambda p: p.category == c and p.brand == br)(cat, b)))
        else:
            specs.append((_lbl(cat), (lambda c: lambda p: p.category == c)(cat)))
    return _group(allp, specs)


def build_by_brand(data) -> List[DashboardTable]:
    """브랜드별. 신규 브랜드 자동 포함."""
    allp = _all_products(data)
    brands = _ordered_unique((p.brand for p in allp), config.BRAND_ORDER)
    specs = [(_lbl(b), (lambda br: lambda p: p.brand == br)(b)) for b in brands]
    return _group(allp, specs)


def _is_online_only(p: RawProduct) -> bool:
    return (p.channel or "").strip() == "온라인"


def build_online_only(data) -> List[DashboardTable]:
    """온라인 전용(채널=온라인) 제품을 카테고리별로. 신규 카테고리 자동 포함."""
    allp = [p for p in _all_products(data) if _is_online_only(p)]
    cats = _ordered_unique((p.category for p in allp), config.CATEGORY_ORDER)
    specs = [(_lbl(c), (lambda cc: lambda p: p.category == cc)(c)) for c in cats]
    return _group(allp, specs)


# ── 전체 요약(브랜드×카테고리×월 매트릭스) ──────────────────────────────────
def _launch_month_col(p: RawProduct):
    """제품의 출시(law 타임라인 인덱스)를 요약 월컬럼(0~11)으로 매핑. 없으면 None."""
    idx = None
    for i, c in enumerate(p.timeline):
        if c and config.LAUNCH_KEYWORD in c:
            idx = i
    if idx is None:
        return None
    for ci, (_lbl, idxs) in enumerate(config.SUMMARY_MONTHS):
        if idx in idxs:
            return ci
    return None


def _abbr(project: str) -> str:
    """프로젝트명 첫 단어."""
    return project.split()[0] if project else ""


def _brand_row_specs(brand, brand_products):
    """브랜드의 카테고리 행 사양 [(label, kind, value)] 반환(온라인 제외).

    kind: 'set'(value=카테고리 집합) / 'rest'(value=명시된 집합, 그 외 전부).
    누락 방지: 정의에 없는 카테고리는 동적 행으로 추가.
    """
    cats_present = _ordered_unique((p.category for p in brand_products
                                    if not _is_online_only(p)), config.CATEGORY_ORDER)
    spec = config.SUMMARY_BRAND_CATEGORIES.get(brand)
    rows = []
    if spec:
        explicit = set()
        has_rest = False
        for lbl, cats in spec:
            if cats != "__rest__":
                explicit.update(cats)
        for lbl, cats in spec:
            if cats == "__rest__":
                rows.append((lbl, "rest", set(explicit))); has_rest = True
            else:
                rows.append((lbl, "set", set(cats)))
        if not has_rest:   # rest 없으면 미정의 카테고리를 동적 행으로
            for c in cats_present:
                if c not in explicit:
                    rows.append((_lbl(c), "set", {c}))
    else:
        for c in cats_present:
            rows.append((_lbl(c), "set", {c}))
    return rows


def _ct(n, sku) -> str:
    return f"{n}({sku})"


def build_summary(data) -> dict:
    """모든 프로젝트를 브랜드×카테고리×월 매트릭스로 요약."""
    allp = _all_products(data)
    nmon = len(config.SUMMARY_MONTHS)
    brands = _ordered_unique((p.brand for p in allp), config.SUMMARY_BRAND_ORDER)

    brand_blocks = []
    # 분기/시즌 그랜드 집계용
    def empty_qs():
        return [[0, 0] for _ in config.SUMMARY_QUARTERS]
    grand_q = empty_qs()
    grand_total = [0, 0]

    def q_of(moncol):   # 월컬럼 → 분기 인덱스
        for qi, (_l, st, sp) in enumerate(config.SUMMARY_QUARTERS):
            if st <= moncol < st + sp:
                return qi
        return len(config.SUMMARY_QUARTERS) - 1

    for brand in brands:
        bps = [p for p in allp if p.brand == brand]
        specs = _brand_row_specs(brand, bps)
        # 행 라벨 목록 + 온라인 행
        row_labels = [s[0] for s in specs] + [config.SUMMARY_ONLINE_ROW]
        # 행별 월셀 엔트리: label -> [ [entries]*nmon ]
        cells = {lbl: [[] for _ in range(nmon)] for lbl in row_labels}

        for p in bps:
            mc = _launch_month_col(p)
            if mc is None:
                continue
            entry = {"key": f"{p.owner}|{p.project}", "abbr": _abbr(p.project),
                     "rev": p.target_revenue, "sku": int(p.sku_count or 0)}
            if _is_online_only(p):
                target = config.SUMMARY_ONLINE_ROW
            else:
                target = None
                for lbl, kind, val in specs:
                    if (kind == "set" and p.category in val) or \
                       (kind == "rest" and p.category not in val):
                        target = lbl; break
                if target is None:
                    target = specs[0][0] if specs else config.SUMMARY_ONLINE_ROW
            cells[target][mc].append(entry)

        # 행 구성 + 행별 계 (cells = 월별 엔트리 리스트 그대로 보관)
        rows_out = []
        brand_q = empty_qs()
        brand_tot = [0, 0]
        for lbl in row_labels:
            row_cells = [cells[lbl][m] for m in range(nmon)]
            n = sum(len(cells[lbl][m]) for m in range(nmon))
            sku = sum(e["sku"] for m in range(nmon) for e in cells[lbl][m])
            rows_out.append({"label": lbl, "cells": row_cells, "total": _ct(n, sku)})
            # 브랜드 분기/총계
            for m in range(nmon):
                for e in cells[lbl][m]:
                    qi = q_of(m)
                    brand_q[qi][0] += 1; brand_q[qi][1] += e["sku"]
                    brand_tot[0] += 1; brand_tot[1] += e["sku"]
                    grand_q[qi][0] += 1; grand_q[qi][1] += e["sku"]
                    grand_total[0] += 1; grand_total[1] += e["sku"]
        brand_blocks.append({
            "brand": brand,
            "rows": rows_out,
            "total_q": [_ct(*brand_q[qi]) for qi in range(len(brand_q))],
            "total": _ct(*brand_tot),
        })

    # 시즌 그랜드 = 분기 그랜드를 시즌 범위(월컬럼)로 합산
    season_acc = [[0, 0] for _ in config.SUMMARY_SEASONS]
    for si, (_l, st, sp) in enumerate(config.SUMMARY_SEASONS):
        for qi, (_ql, qst, qsp) in enumerate(config.SUMMARY_QUARTERS):
            if st <= qst < st + sp:
                season_acc[si][0] += grand_q[qi][0]
                season_acc[si][1] += grand_q[qi][1]

    return {
        "main_title": config.SUMMARY_MAIN_TITLE,
        "months": config.SUMMARY_MONTHS,
        "quarters": config.SUMMARY_QUARTERS,
        "seasons": config.SUMMARY_SEASONS,
        "brands": brand_blocks,
        "grand_q": [_ct(*grand_q[qi]) for qi in range(len(grand_q))],
        "grand_season": [_ct(*season_acc[si]) for si in range(len(season_acc))],
        "grand_total": _ct(*grand_total),
    }


# 탭 정의: (탭 이름, 빌더 함수)
TAB_BUILDERS = [
    ("카테고리별", build_by_category),
    ("브랜드별", build_by_brand),
    ("온라인 전용", build_online_only),
    ("담당자별", build_by_owner),
]


def build_groupings(data) -> Dict[str, List[DashboardTable]]:
    return {name: fn(data) for name, fn in TAB_BUILDERS}


if __name__ == "__main__":
    import sys
    import data_loader
    tables = build_all(data_loader.load(sys.argv[1] if len(sys.argv) > 1 else None))
    for t in tables:
        print(f"\n=== {config.SUBTITLE_PREFIX}{t.owner}  (Q표기={t.show_q_label}) ===")
        print(f"    요약: {t.summary.label} | 분기별: {t.summary.per_quarter}")
        for r in t.rows[:3]:
            sc = r.series_cell.replace(chr(10), '/')
            mk = r.maker.replace(chr(10), '')
            print(f"    NO[{r.no_label or '〃'}] {sc} | {r.target_price}만 | {mk} | "
                  f"색{r.color_count} SKU{r.sku_count}")
