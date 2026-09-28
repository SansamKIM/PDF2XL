"""권역 유틸리티(병합 규칙, 지방 권역 키 보정 등)."""

from __future__ import annotations

import re
from typing import Dict, List, Optional, Tuple


# 권역설명에서 전체/합계로 쓰이는 이름들.
#
# LLM 출력(또는 일부 자료)에서 "전체" 행이 끼어 있는 경우가 있는데,
# 이는 실제 권역 행이 아니므로 Excel 권역설명 블록에 표시되면 기관 간 일관성이 깨진다.
#
# 항목이 이것 하나뿐인 경우에만 유지한다.
_OVERALL_REGION_NAMES = {"전체", "총계", "합계"}


def is_overall_region_name(name: str) -> bool:
    """전체/합계 권역명인지 여부 반환(예: 전체)."""

    if not isinstance(name, str):
        return False
    s = re.sub(r"\s+", "", name).strip()
    if not s:
        return False
    if s in _OVERALL_REGION_NAMES:
        return True
    # 영어 표현도 허용
    if s.upper() in {"ALL", "TOTAL"}:
        return True
    return False


def cleanup_region_info(region_info):
    """권역설명에서 '전체' 등 전체 항목을 정리.

    구조:
      - {'권역': [...]}  (권장)
      - [...]            (레거시)

    다른 권역이 함께 있으면 전체 항목을 제거하고, 전체만 있으면 유지한다.
    """

    items = normalize_region_items(region_info)
    if not items:
        return region_info

    non_overall = [x for x in items if not is_overall_region_name(x.get("name", ""))]
    # 전체를 빼면 비게 되면 원본을 유지
    filtered = non_overall if non_overall else items

    if isinstance(region_info, dict):
        out = dict(region_info)
        if "권역" in out and isinstance(out.get("권역"), list):
            out["권역"] = filtered
            return out
        # 예상 형태가 아니면 원본 유지
        return region_info

    if isinstance(region_info, list):
        return filtered

    return region_info


_ADMIN_SUFFIXES = ("구", "군", "시", "도", "읍", "면")


def _is_admin_area_name(name: str) -> bool:
    if not isinstance(name, str):
        return False
    s = name.strip()
    if not s:
        return False
    # 연령/기타 숫자 포함 항목은 제외
    if any(ch.isdigit() for ch in s):
        return False
    if "/" in s:
        return False
    return s.endswith(_ADMIN_SUFFIXES)


def _levenshtein_distance(a: str, b: str) -> int:
    if a == b:
        return 0
    if not a:
        return len(b)
    if not b:
        return len(a)

    la, lb = len(a), len(b)
    prev = list(range(lb + 1))
    for i in range(1, la + 1):
        cur = [i] + [0] * lb
        ca = a[i - 1]
        for j in range(1, lb + 1):
            cb = b[j - 1]
            cost = 0 if ca == cb else 1
            cur[j] = min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + cost)
        prev = cur
    return prev[lb]


def _closest_region_name(name: str, region_list: List[str]) -> Optional[str]:
    if not region_list or not isinstance(name, str):
        return None

    s = name.strip()
    if not s:
        return None

    best = None
    best_dist = None
    tie = False

    for cand in region_list:
        if not isinstance(cand, str):
            continue
        c = cand.strip()
        if not c:
            continue

        # 접미사와 길이가 같을 때만 후보로 사용해 과한 매핑을 방지
        if len(c) != len(s):
            continue
        if c[-1] != s[-1]:
            continue

        dist = _levenshtein_distance(s, c)
        threshold = 1 if len(s) <= 3 else 2
        if dist > threshold:
            continue

        if best_dist is None or dist < best_dist:
            best = c
            best_dist = dist
            tie = False
        elif dist == best_dist:
            tie = True

    if tie:
        return None
    return best


def normalize_region_table_keys(table_data: Dict, region_list: List[str]) -> Dict:
    """알려진 region_list를 사용해 권역 키의 미세한 OCR/LLM 흔들림을 보정."""

    if not isinstance(table_data, dict) or not region_list:
        return table_data

    # "권역1" 류를 복원하기 위한 보조 맵.
    # 일부 추출(또는 예전 정규화)에서 '권역' 접두어가 빠져 숫자만 남을 수 있다.
    # region_list에 "권역1", "권역2"...가 있으면 원래 라벨로 복원 가능.
    num_map: Dict[str, str] = {}
    for r in region_list:
        if not isinstance(r, str):
            continue
        rr = r.strip()
        if not rr:
            continue
        compact = re.sub(r"\s+", "", rr)
        m = re.match(r"^권역(?P<n>\d+)$", compact)
        if m:
            num_map[m.group("n")] = rr
            continue
        m2 = re.match(r"^(?P<n>\d+)권역$", compact)
        if m2:
            num_map[m2.group("n")] = rr

    new_data: Dict = {}
    for k, v in table_data.items():
        key = k
        if isinstance(k, str):
            k_strip = k.strip()
            k_compact = re.sub(r"\s+", "", k_strip)

            # 1) 권역 숫자 라벨 복원: "1" -> "권역1" (region_list 기반)
            if k_compact.isdigit() and k_compact in num_map:
                key = num_map[k_compact]
            # 2) "권역 1" -> "권역1" (정확 일치가 없을 때)
            elif k_compact and (k_strip not in region_list):
                m = re.match(r"^권역(?P<n>\d+)$", k_compact)
                if m and m.group("n") in num_map:
                    key = num_map[m.group("n")]
                else:
                    m2 = re.match(r"^(?P<n>\d+)권역$", k_compact)
                    if m2 and m2.group("n") in num_map:
                        key = num_map[m2.group("n")]

            if _is_admin_area_name(k_strip) and k_strip not in region_list:
                mapped = _closest_region_name(k_strip, region_list)
                if mapped and mapped in region_list:
                    key = mapped

        if key in new_data:
            # 충돌 시 처음 값을 유지(추후 병합으로 처리 가능)
            continue
        new_data[key] = v

    return new_data


def normalize_region_items(region_info) -> List[dict]:
    """정규화된 권역 dict 리스트 반환.

    region_info 허용 형태:
    - {'권역': [...]}
    - [...]
    - None
    """
    if not region_info:
        return []
    if isinstance(region_info, list):
        return [x for x in region_info if isinstance(x, dict)]
    if isinstance(region_info, dict):
        items = region_info.get("권역", [])
        if isinstance(items, list):
            return [x for x in items if isinstance(x, dict)]
        return []
    return []


def extract_region_list(region_info) -> List[str]:
    """Excel 지방 템플릿의 '지역별' 행 매핑에 쓸 지역 리스트를 추출.

    지방선거 PDF에는 권역설명 전용 페이지가 없고 응답자 특성 표를 재사용하는 경우가 많다:
        {"권역": [{"name": "광주광역시", "areas": "광산구, 남구, ..."}]}

    쓰기 시에는 도시명이 아니라 표에 실제로 쓰이는 행 라벨(광산구, 남구 등)이 필요하므로,
    {name, areas} 구조가 1개뿐이면 행 라벨처럼 보이는 areas를 콤마 기준으로 확장한다.
    """

    items = normalize_region_items(region_info)
    if not items:
        return []

    # 특수 처리: 지역 1개에 콤마/중점으로 여러 구역이 붙은 경우
    if len(items) == 1 and isinstance(items[0], dict):
        areas = items[0].get("areas")
        if isinstance(areas, str) and areas.strip():
            # 구분자 정규화
            s = areas.replace("·", ",").replace("/", ",")
            parts = [p.strip() for p in s.split(",") if p.strip()]
            if len(parts) >= 2 and all(_is_admin_area_name(p) for p in parts):
                return parts

    out: List[str] = []
    for r in items:
        name = r.get("name")
        if isinstance(name, str) and name.strip():
            out.append(name.strip())
    return out


def extract_region_list_for_table(region_info, table_categories: List[str]) -> List[str]:
    """표 정보를 활용해 지방 템플릿용 지역 리스트를 추출.

    배경
    - 지방선거 PDF에 {"권역": [{"name": "광주광역시", "areas": "광산구, 남구, ..."}]}처럼
      권역설명이 1개만 있는 경우, 표 행은 보통 구/군/시(세분 지역)이다.

    휴리스틱(순서 중요)
    1) 표에 행정구역 라벨(구/군/시/도/읍/면)이 2개 이상 있으면 이를 region_list로 사용.
    2) 아니면 extract_region_list(region_info)로 폴백.
       - 단, 1개 권역의 areas를 리스트로 확장하는 것은 표에 그 구역이 최소 2개 이상 실제로 있을 때만 수행.

    이렇게 하면 권역설명은 있지만 표 라벨이 상위 지역(예: 광주광역시)인 자료에서 불필요한 확장을 막는다.
    """

    # 1) 표 기반: 세분화된 행정구역 라벨이 있으면 우선 사용
    cats = [c for c in (table_categories or []) if isinstance(c, str)]
    admin_cats = [c.strip() for c in cats if _is_admin_area_name(c)]
    # 순서 유지하며 중복 제거
    seen = set()
    admin_ordered: List[str] = []
    for c in admin_cats:
        if c and c not in seen:
            seen.add(c)
            admin_ordered.append(c)

    if len(admin_ordered) >= 2:
        return admin_ordered

    # 2) 권역설명 기반(레거시)
    items = normalize_region_items(region_info)
    if not items:
        return []

    # {name, areas} 한 항목 구조일 때 조건부 확장
    if len(items) == 1 and isinstance(items[0], dict):
        areas = items[0].get("areas")
        if isinstance(areas, str) and areas.strip():
            s = areas.replace("·", ",").replace("/", ",")
            parts = [p.strip() for p in s.split(",") if p.strip()]
            if len(parts) >= 2 and all(_is_admin_area_name(p) for p in parts):
                # 표에 해당 라벨이 실제로 있을 때만 확장
                hit = sum(1 for p in parts if p in set(cats))
                if hit >= 2:
                    return parts

    # 폴백: 권역 이름 목록만 사용
    return extract_region_list(region_info)


def merge_regions_weighted(data: dict, region1: str, region2: str, merged_name: str) -> dict:
    """두 권역 행을 조사완료(표본수) 가중치로 병합."""

    if not isinstance(data, dict) or "데이터" not in data:
        return data

    table = data.get("데이터")
    if not isinstance(table, dict):
        return data

    if region1 not in table or region2 not in table:
        return data

    d1 = table.get(region1) or {}
    d2 = table.get(region2) or {}

    n1 = (d1.get("조사완료") or 0) if isinstance(d1, dict) else 0
    n2 = (d2.get("조사완료") or 0) if isinstance(d2, dict) else 0
    total = n1 + n2
    if total <= 0:
        return data

    merged = {"조사완료": total}
    keys = set()
    if isinstance(d1, dict):
        keys |= set(d1.keys())
    if isinstance(d2, dict):
        keys |= set(d2.keys())

    for k in keys:
        if k == "조사완료":
            continue
        v1 = d1.get(k) if isinstance(d1, dict) else None
        v2 = d2.get(k) if isinstance(d2, dict) else None
        if v1 is None and v2 is None:
            merged[k] = None
        elif v1 is None:
            merged[k] = v2
        elif v2 is None:
            merged[k] = v1
        else:
            try:
                merged[k] = round((n1 * float(v1) + n2 * float(v2)) / total, 1)
            except Exception:
                merged[k] = v1

    # 기존 항목을 병합 결과로 교체
    table.pop(region1, None)
    table.pop(region2, None)
    table[merged_name] = merged

    return data
