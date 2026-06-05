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


def build_by_category(data) -> List[DashboardTable]:
    """카테고리별(매트리스·베드룸은 브랜드로 세분)."""
    allp = _all_products(data)
    specs = [
        ("매트리스_마테라소", lambda p: p.category == "매트리스" and p.brand == "마테라소"),
        ("매트리스_까사미아", lambda p: p.category == "매트리스" and p.brand == "까사미아"),
        ("베드룸_마테라소", lambda p: p.category == "베드룸" and p.brand == "마테라소"),
        ("베드룸_까사미아", lambda p: p.category == "베드룸" and p.brand == "까사미아"),
        ("수납장", lambda p: p.category == "수납장"),
        ("홈오피스", lambda p: p.category == "홈오피스"),
        ("키즈", lambda p: p.category == "키즈"),
    ]
    return _group(allp, specs)


def build_by_brand(data) -> List[DashboardTable]:
    """브랜드별."""
    allp = _all_products(data)
    specs = [(b, (lambda b: lambda p: p.brand == b)(b)) for b in ("마테라소", "까사미아")]
    return _group(allp, specs)


def _is_online_only(p: RawProduct) -> bool:
    return (p.channel or "").strip() == "온라인"


def build_online_only(data) -> List[DashboardTable]:
    """온라인 전용(채널=온라인) 제품을 카테고리별로."""
    allp = [p for p in _all_products(data) if _is_online_only(p)]
    specs = [(c, (lambda c: lambda p: p.category == c)(c))
             for c in ("매트리스", "베드룸", "수납장", "홈오피스", "키즈")]
    return _group(allp, specs)


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
