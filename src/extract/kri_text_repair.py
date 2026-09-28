# -*- coding: utf-8 -*-
"""
KRI(코리아리서치인터내셔널) 전용: PDF 텍스트(레이아웃 유지) 기반 행 보정.

목표
- LLM이 표 행 라벨을 누락/병합(예: '서울/인천/경기')하는 케이스를 PDF 텍스트로 복구
- 응답항목(컬럼) 수가 일치하는 경우에만 값까지 덮어써서 안전하게 적용
- 다른 기관에는 영향 0 (extractor에서 KRI일 때만 호출)

주의
- 정당지지도(PSR)는 KRI 원표가 '기본소득당/사회민주당/없다/모름/없음+모름' 등
  추가 컬럼을 가지는 경우가 많아, 응답항목 수 불일치로 보정이 자동으로 스킵될 수 있음.
  (이 경우 기존 LLM 결과를 그대로 사용)
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Dict, List, Optional, Tuple

import re

import pdfplumber

from ..normalize.labels import normalize_label
# GE/PSR은 템플릿 응답 헤더를 고정 유지.
# ISSUE는 질문/응답지가 매번 달라서 동적 유지.
from ..normalize.responses import CANON_GE_HEADERS, CANON_PSR_HEADERS


@dataclass
class _Row:
    group: str
    label: str
    base_n: int
    values: List[Optional[float]]


_SECTION_TO_GROUP = {
    "성": "성별",
    "연령": "연령별",
    "권역": "지역별",
    "지역": "지역별",
    "직업": "직업별",
    "최종학력": "학력별",
    "이념성향": "이념성향별",
    "정당지지도": "정당지지도별",
    "국정운영평가": "국정운영평가별",
}

# '모름/무응답'이 여러 분류에 중복 등장할 수 있어서, 직업 파트에서는 템플릿에 맞게 강제 매핑
_JOB_UNKNOWN_MARKERS = {
    "모름/무응답",
    "모름.무응답",
    "무응답",
    "모름/기타",
    "기타/모름",
}


def _compact(s: str) -> str:
    return re.sub(r"\s+", "", str(s or ""))


def _extract_section_marker(line: str) -> Optional[str]:
    """'◈   권 역   ◈' 같은 라인에서 섹션명(권역/성/연령/직업...)만 뽑는다."""
    if "◈" not in line:
        return None
    t = _compact(line)
    t = t.replace("◈", "")
    # 빈 문자열이면 무시
    return t or None


def _marker_to_group(marker: str) -> Optional[str]:
    # marker가 '권역'처럼 딱 맞는 경우도, '권역◈' 같은 잔여가 남는 경우도 있어서 방어적으로 처리
    m = marker.strip()
    if not m:
        return None
    # 가능한 키 중 가장 긴 것부터 포함 매칭
    for k in sorted(_SECTION_TO_GROUP.keys(), key=len, reverse=True):
        if k in m:
            return _SECTION_TO_GROUP[k]
    return None


def _group_allowed(group: str, include_groups: List[str], exclude_groups: List[str]) -> bool:
    if group in (exclude_groups or []):
        return False
    if include_groups:
        return group in include_groups
    return True


def _parse_numeric_parens(s: str) -> List[re.Match]:
    return list(re.finditer(r"\(\s*(\d[\d,]*)\s*\)", s))


def _parse_row_line(line: str) -> Optional[Tuple[str, int, List[Optional[float]]]]:
    """표의 한 행을 (라벨, base_n, 값들)로 파싱."""
    if not line:
        return None

    # 숫자 괄호가 없는 라인은 스킵
    parens = _parse_numeric_parens(line)
    if not parens:
        return None

    first = parens[0]
    raw_label = line[: first.start()].strip()
    if not raw_label:
        return None

    try:
        base_n = int(first.group(1).replace(",", ""))
    except Exception:
        return None

    # 첫 괄호 이후 부분
    tail = line[first.end() :].strip()

    # 마지막 괄호가 라인 끝에 붙어 있으면(가중값 적용 사례수) 제거
    if len(parens) >= 2:
        last = parens[-1]
        if line[last.end() :].strip() == "":
            tail = line[first.end() : last.start()].strip()

    # 잔여 괄호(혹시 있을 수 있는) 제거
    tail = re.sub(r"\(\s*\d[\d,]*\s*\)", " ", tail).strip()
    if not tail:
        return None

    tokens = [t for t in re.split(r"\s+", tail) if t]
    values: List[Optional[float]] = []
    for t in tokens:
        t = t.strip()
        if t in {"-", "–", "—"}:
            values.append(None)
            continue
        t = t.replace("%", "")
        try:
            values.append(float(t))
        except Exception:
            # 숫자가 아닌 토큰이 섞이면 안전하게 포기
            return None

    return raw_label, base_n, values


def _normalize_label_by_group(raw_label: str, group: str) -> str:
    """KRI 원표의 라벨을 템플릿/정규화 규칙에 맞춰 정리."""
    # 기본 정규화(공백/분리된 한글 등)
    n = normalize_label(raw_label)

    if group == "직업별":
        # 직업 파트의 '모름/무응답'은 다른 분류(이념/정당 등)에도 반복 등장해
        # 단순 정규화 시 카테고리 충돌이 발생하기 쉬움.
        #
        # KRI 템플릿(엑셀)에서는 직업 파트에 '모름.무응답' 행이 없고,
        # 대신 매우 작은 표본(N=1~2 등)을 담는 '밝힐 수 없음' 행을 사용한다.
        #
        # 따라서 KRI에 한해 직업 파트의 모름/무응답 계열을 '밝힐 수 없음'으로 안전하게 분리한다.
        if _compact(n).replace(".", "/") in {_compact(x).replace(".", "/") for x in _JOB_UNKNOWN_MARKERS}:
            return "밝힐 수 없음"

    return n


def _extract_rows_from_page_text(
    text: str,
    include_groups: List[str],
    exclude_groups: List[str],
) -> List[_Row]:
    lines = [ln.rstrip("\n") for ln in (text or "").splitlines()]
    # KRI 표는 대시 라인이 여러 번 등장 (상단, 헤더 하단, 표 하단)
    dash_idx = [i for i, ln in enumerate(lines) if ln.count("-") >= 20]
    if len(dash_idx) < 2:
        return []

    # 두 번째 대시 라인 이후가 데이터 영역
    start = dash_idx[1] + 1
    end = dash_idx[-1] if dash_idx[-1] > start else len(lines)

    cur_group = "전체"
    out: List[_Row] = []

    for ln in lines[start:end]:
        if not ln or not ln.strip():
            continue

        mk = _extract_section_marker(ln)
        if mk:
            g = _marker_to_group(mk)
            if g:
                cur_group = g
            continue

        parsed = _parse_row_line(ln)
        if not parsed:
            continue

        raw_label, base_n, values = parsed
        group = cur_group or "전체"

        if not _group_allowed(group, include_groups, exclude_groups):
            continue

        label = _normalize_label_by_group(raw_label, group)
        out.append(_Row(group=group, label=label, base_n=base_n, values=values))

    return out


def _has_summary_columns(page_text: str) -> bool:
    """표에 '합산/종합' 열(예: 긍정/부정/모름)이 명시되어 있는지 감지."""
    t = _compact(page_text)
    if not t:
        return False

    # KRI 스타일 표에서 흔히 보이는 패턴:
    # - '종합 결과' + '긍정/부정'
    # - '(ⓐ+ⓑ)' '(ⓒ+ⓓ)' 마커
    if ("종합결과" in t) and (("긍정" in t) or ("부정" in t)):
        return True
    if ("ⓐ+ⓑ" in page_text) or ("ⓒ+ⓓ" in page_text):
        return True

    return False


def _looks_like_psr_table(resp_items: List[Any]) -> bool:
    """응답 헤더(정당명)로 PSR(정당지지도) 표 여부를 감지."""
    rs = {_compact(x) for x in (resp_items or [])}
    return ("더불어민주당" in rs) and ("국민의힘" in rs)


def _select_three_cols(values: List[Optional[float]], has_summary_cols: bool) -> Optional[List[Optional[float]]]:
    """LLM이 3개 열을 예상하지만 PDF에 더 많을 때 3개 값을 선택/계산."""
    if not values or len(values) < 3:
        return None

    # 표에 종합 열이 있으면 보통 마지막 3개 숫자가 이에 해당
    if has_summary_cols and len(values) >= 3:
        return values[-3:]

    # 아니면 앞쪽 5개 열에서 (a+b), (c+d), (e)를 계산
    if len(values) >= 5:
        a, b, c, d, e = values[0], values[1], values[2], values[3], values[4]
        if any(x is None for x in [a, b, c, d, e]):
            return None
        return [round(float(a) + float(b), 1), round(float(c) + float(d), 1), e]

    return None


def _map_kri_psr_11(values: List[Optional[float]]) -> Optional[List[Optional[float]]]:
    """KRI PSR 원본 11열 레이아웃을 템플릿 8열 레이아웃으로 매핑.

    가장 흔한 원본 순서(MBC 시리즈 기준):
      민주, 국힘, 조국, 개혁, 진보, 기본, 사회, 그외정당, 없다, 모름/무응답, 없음/모름/무응답
    """
    if not values or len(values) != 11:
        return None

    dem, ppl, cho, reform, jinbo, basic, social, other, none_, unk, _combined = values
    other_sum = 0.0
    for x in [basic, social, other]:
        if x is not None:
            other_sum += float(x)

    return [dem, ppl, cho, jinbo, reform, round(other_sum, 1), none_, unk]



def try_repair_kri_table_data(
    pdf_path: str,
    pages: List[int],
    cand: Dict[str, Any],
    side_cfg: Any,  # SideConfig (타입 의존 줄이기 위해 Any)
) -> Dict[str, Any]:
    """LLM 파싱 결과(cand)를 PDF 텍스트로 보정해서 반환.

    - 응답항목 수가 row 값 수와 일치하는 행만 안전하게 덮어쓴다.
    - 새로운 행은 추가하고, 기존 행은 덮어쓸 수 있다.
    - '서울/인천/경기' 같은 병합 라벨은, 개별 라벨이 확보되면 제거한다.
    """
    if not pdf_path or not pages:
        return cand

    try:
        # 참고:
        # - pages[]는 프로젝트 전반과 동일하게 1부터 시작하는 PDF 페이지 번호 리스트
        # - pdfplumber의 pdf.pages[]는 0-based
        # 여기서 변환을 잊으면 다음 페이지를 읽어 모델 출력이 다른 표로 덮어쓰이는
        # 전형적인 off-by-one 버그가 발생할 수 있음
        page_num = int(pages[0])
        page_idx = page_num - 1
    except Exception:
        return cand

    # 중요: GE/PSR 헤더는 고정(비동적) 유지해 컬럼이 흔들리지 않도록 함.
    # ISSUE는 질문별로 응답지가 달라 동적 유지.
    item_type = str(cand.get("항목유형") or "").upper()

    resp_items = cand.get("응답항목")
    if item_type == "GE":
        resp_items = CANON_GE_HEADERS.copy()
        cand["응답항목"] = resp_items
    elif item_type == "PSR":
        resp_items = CANON_PSR_HEADERS.copy()
        cand["응답항목"] = resp_items

    if not isinstance(resp_items, list) or not resp_items:
        return cand
    expected_cols = len(resp_items)

    include_groups = getattr(side_cfg, "include_groups", []) or []
    exclude_groups = getattr(side_cfg, "exclude_groups", []) or []

    try:
        with pdfplumber.open(pdf_path) as pdf:
            if page_idx < 0 or page_idx >= len(pdf.pages):
                return cand
            page = pdf.pages[page_idx]
            text = page.extract_text(layout=True) or ""
    except Exception:
        return cand

    rows = _extract_rows_from_page_text(text, include_groups, exclude_groups)
    if not rows:
        return cand

    data = cand.get("데이터")
    if not isinstance(data, dict):
        data = {}
    else:
        data = dict(data)  # copy

    has_summary_cols = _has_summary_columns(text)
    is_psr = _looks_like_psr_table(resp_items)

    applied = False
    for r in rows:
        use_values: Optional[List[Optional[float]]] = None

        if len(r.values) == expected_cols:
            use_values = r.values
        elif expected_cols == 3:
            if item_type == "GE":
                # GE(국정운영평가)는 템플릿이 3열(잘함/못함/모름) 고정.
                # KRI 표에는 보통 '종합 결과(긍/부/모름)'가 명시되어 있으므로,
                # 그 *명시된* 마지막 3열만 사용한다.
                # (a+b, c+d 같은 계산 기반 폴백은 동적 처리로 간주되어 오탐 위험이 있음)
                use_values = r.values[-3:] if (has_summary_cols and len(r.values) >= 3) else None
            else:
                use_values = _select_three_cols(r.values, has_summary_cols)
        elif expected_cols == 8 and is_psr:
            use_values = _map_kri_psr_11(r.values)

        if not use_values or len(use_values) != expected_cols:
            continue

        row_dict: Dict[str, Any] = {"조사완료": r.base_n}
        for i, col_name in enumerate(resp_items):
            row_dict[str(col_name)] = use_values[i]
        data[r.label] = row_dict
        applied = True

    if applied:
        # 병합 라벨 제거(개별 라벨이 동시에 존재할 때만)
        def has_label(compact_target: str) -> bool:
            return any(_compact(k) == compact_target for k in data.keys())

        has_seoul = has_label("서울")
        has_incheon_gyeonggi = any(_compact(k) in {"인천/경기", "인천경기"} for k in data.keys())

        if has_seoul and has_incheon_gyeonggi:
            for k in list(data.keys()):
                ck = _compact(k)
                if ck in {"서울/인천/경기", "서울/경기/인천", "서울/인천/경기권"}:
                    data.pop(k, None)

        cand["데이터"] = data

    return cand
