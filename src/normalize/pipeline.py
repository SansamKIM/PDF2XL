"""정규화 파이프라인.

단일 표 JSON(dict)을 받아 Excel 작성에 안정적인 정규화 dict로 변환한다.

이 단계에서 수행하는 작업:
- 기관별 용어 매핑(YAML)
- 공통 라벨 정규화
- 항목 유형별 응답 항목 정규화
- 카테고리 충돌 병합
- 선택적 권역 병합(YAML)
- 지방 권역 키 흔들림 보정
"""

from __future__ import annotations

from copy import deepcopy
import re
from typing import Any, Dict, List, Optional, Tuple

from ..config.org import OrgConfig
from .labels import apply_term_mapping, normalize_label
from .responses import normalize_response_items_by_type, normalize_value_keys_by_type
from .collisions import merge_table_by_category
from .regions import extract_region_list, normalize_region_table_keys, merge_regions_weighted


def _is_age_by_gender_crosstab(label: str) -> bool:
    """휴리스틱: '성 by 연령' 교차 행 여부(예: '18~29세 남성', '40대 여성').

    이런 행은 JSON 크기만 늘리고 토큰/시간을 낭비하며, 템플릿이 이미 별도로
    성별/연령별을 받으므로 사용되지 않는다.
    """

    if label is None:
        return False
    s = str(label)
    if not s:
        return False

    has_gender = ("남성" in s) or ("여성" in s)
    if not has_gender:
        return False

    # 연령 토큰 패턴: 구간(18~29세), 30대, 70세 이상 등
    has_age = False
    if re.search(r"\b\d{1,2}\s*[~-]\s*\d{1,2}\s*세?\b", s):
        has_age = True
    elif re.search(r"\b\d{2}\s*대\b", s):
        has_age = True
    elif re.search(r"\b\d{2}\s*세\s*이상\b", s):
        has_age = True

    return bool(has_age and has_gender)


def _iter_region_merge_rules(org: OrgConfig) -> List[Dict[str, Any]]:
    rules = org.region_merges or []
    out = []
    for r in rules:
        if not isinstance(r, dict):
            continue
        # 기대 형태: {from: ["강원", "제주"], to: "강원/제주"}
        out.append(r)
    return out


def apply_region_merges(data: dict, org: OrgConfig) -> Tuple[dict, List[str]]:
    msgs: List[str] = []
    for rule in _iter_region_merge_rules(org):
        src = rule.get("from") or rule.get("sources")
        dst = rule.get("to") or rule.get("target")
        if not (isinstance(src, list) and len(src) == 2 and isinstance(dst, str)):
            continue
        a, b = src[0], src[1]
        before_keys = set((data.get("데이터") or {}).keys()) if isinstance(data.get("데이터"), dict) else set()
        data = merge_regions_weighted(data, str(a), str(b), str(dst))
        after_keys = set((data.get("데이터") or {}).keys()) if isinstance(data.get("데이터"), dict) else set()
        if before_keys != after_keys:
            msgs.append(f"merge_regions: {a}+{b}->{dst}")
    return data, msgs


def normalize_table_json(
    raw: dict,
    org: OrgConfig,
    survey_type: str,
    region_info: Any = None,
) -> Tuple[dict, Dict[str, Any]]:
    """단일 표 JSON을 정규화.

    반환: (normalized_json, report)
    """

    report: Dict[str, Any] = {
        "collisions": [],
        "region_merges": [],
    }

    data = deepcopy(raw) if isinstance(raw, dict) else {}

    # 1) 기관 용어 매핑
    # org.term_mapping 형태:
    #  - 평면 dict: {from: to}
    #  - 섹션 dict: {"분류": {...}, "응답": {...}, "all": {...}}
    # 적용 전에 평면 매핑으로 정규화
    term_map: Dict[str, str] = {}
    tm = getattr(org, "term_mapping", None) or {}
    if isinstance(tm, dict):
        if isinstance(tm.get("all"), dict):
            term_map.update({str(k): str(v) for k, v in (tm.get("all") or {}).items() if k})
        for _section, mp in tm.items():
            if _section == "all":
                continue
            if isinstance(mp, dict):
                term_map.update({str(k): str(v) for k, v in mp.items() if k})

    if term_map:
        data = apply_term_mapping(data, term_map)

    item_type = (data.get("항목유형") or "ISSUE").upper()

    # 2) 응답항목 리스트와 행별 값 키 정규화
    if isinstance(data.get("응답항목"), list):
        data["응답항목"] = normalize_response_items_by_type(item_type, data.get("응답항목") or [])

    if isinstance(data.get("데이터"), dict):
        new_table: Dict[str, dict] = {}

        # 일부 카테고리 라벨(특히 '모름/무응답')이 한 표 안에 여러 번 등장할 수 있어
        # (예: 이념 블록, 직업 블록), 납작한 JSON 구조에서 충돌을 막기 위해 구분자를 붙인다.
        ideology_keys = {"진보", "중도", "보수"}
        occupation_keys = {
            "농/임/어업",
            "농림어업",
            "자영업",
            "화이트칼라",
            "블루칼라",
            "전업주부",
            "학생",
            "기타",
            "은퇴.무직",
            "밝힐 수 없음",
            "모름/무응답(직업)",
        }

        last_group: str | None = None

        def _update_group(ncat: str) -> str | None:
            if ncat in ideology_keys:
                return "ideology"
            if ncat in occupation_keys:
                return "occupation"
            return None

        for cat, values in data["데이터"].items():
            ncat = normalize_label(cat)

            # '성 by 연령' 교차 행은 건너뜀
            # - 템플릿에서 사용하지 않음
            # - JSON 크기/처리시간만 증가
            if _is_age_by_gender_crosstab(ncat) or _is_age_by_gender_crosstab(cat):
                continue

            # 직업 블록의 모름/무응답이 이념 블록 것과 충돌하지 않도록 블록별로 구분
            if ncat in {"모름.무응답", "잘 모름"} and last_group == "occupation":
                ncat = "모름/무응답(직업)"

            g = _update_group(ncat)
            if g:
                last_group = g

            nvals = normalize_value_keys_by_type(item_type, values if isinstance(values, dict) else {})

            # 중복 행은 병합
            if ncat in new_table:
                from .collisions import merge_category_rows

                new_table[ncat] = merge_category_rows(new_table[ncat], nvals)
                report["collisions"].append(ncat)
            else:
                new_table[ncat] = nvals

        # 마지막 병합 패스
        merged_table, merge_report = merge_table_by_category(new_table)
        if merge_report.get("merged"):
            report["collisions"].extend(merge_report["merged"])
        data["데이터"] = merged_table

    
    # 2.5) PSR: 정당 키를 템플릿에 맞춰 축소
    if item_type == "PSR" and isinstance(data.get("데이터"), dict):
        from .psr import normalize_psr_table

        data = normalize_psr_table(data, survey_type)


    # 2.6) GE: 원본이 (긍정 합산)+(부정 합산)만 있고 '잘 모름' 열이 없을 때 보정
    #
    # 리서치뷰 계열 표에서 종종 마지막 열이 '모름/기타/무응답'인데,
    # 비전 추출이 이를 놓치면 GE가 2열처럼 보이면서 템플릿과 충돌합니다.
    # 이 단계는 **누락된 경우에만** 보수적으로 보정합니다.
    if item_type == "GE" and isinstance(data.get("데이터"), dict):

        def _to_float(v: Any) -> float | None:
            if v is None:
                return None
            if isinstance(v, (int, float)):
                return float(v)
            try:
                return float(str(v).strip())
            except Exception:
                return None

        table2: Dict[str, Any] = {}
        for cat, row in (data.get("데이터") or {}).items():
            if not isinstance(row, dict):
                table2[cat] = row
                continue

            a = _to_float(row.get("잘하고 있다"))
            b = _to_float(row.get("잘 못하고 있다"))
            c = _to_float(row.get("잘 모름"))

            if c is None and a is not None and b is not None:
                rem = 100.0 - (a + b)
                # 반올림 오차로 살짝 초과(예: 100.1)한 경우 허용
                if rem < 0 and rem >= -0.6:
                    rem = 0.0
                # 가드레일: 모름 버킷은 대체로 크지 않음
                if 0.0 <= rem <= 30.0:
                    row2 = dict(row)
                    row2["잘 모름"] = round(rem, 1)
                    table2[cat] = row2
                    continue

            table2[cat] = row

        data["데이터"] = table2

    # 3) YAML 권역 통합 규칙 적용(주로 전국 권역 통합)
    data, merge_msgs = apply_region_merges(data, org)
    report["region_merges"] = merge_msgs

    # 4) 지방: 권역 키 흔들림 보정
    if survey_type == "지방":
        region_info_use = region_info or data.get("권역설명")
        # 표에 실제 행 라벨(구/군/시 등)이 있으면 그 목록을 우선 사용
        # (권역설명 1개 + areas 문자열을 무조건 확장해 매핑이 깨지는 것을 방지)
        from .regions import extract_region_list_for_table

        table_cats = list((data.get("데이터") or {}).keys()) if isinstance(data.get("데이터"), dict) else []
        region_list = extract_region_list_for_table(region_info_use, table_cats)
        if region_list and isinstance(data.get("데이터"), dict):
            data["데이터"] = normalize_region_table_keys(data["데이터"], region_list)

    return data, report
