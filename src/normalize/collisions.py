"""정규화 후 충돌(collision) 처리.

서로 다른 원본 라벨이 정규화 과정에서 동일한 표준 라벨로 합쳐질 때,
대부분은 '같은 행이 서로 다른 표기/별칭으로 중복 추출된 경우'다.

KRI 스타일 표에서 관측된 예:
- '대전/충청(세종)' vs '대전/세종/충청'
- '모름/무응답' vs '모름.무응답'
- '주부' vs '전업주부' (별칭 매핑 후)

이전 구현은 '조사완료'를 합산하고 퍼센트는 가중평균을 사용했다.
퍼센트는 비슷해 보일 수 있지만, 조사완료(표본수)를 합산하면 표본이 2배로 잡혀
엑셀 출력의 '조사완료'가 잘못된다.

현재 모듈은 충돌을 보수적으로 중복 제거한다:
- 충돌 항목의 '조사완료'는 절대 합산하지 않는다.
- 값이 더 많이 채워진 행(완성도가 높은 행)을 우선한다.
- 가능한 경우 다른 행에서 빈 셀만 채운다.
- '조사완료'는 max(n1, n2)로 유지한다.
"""

from __future__ import annotations

from typing import Dict, Tuple


def _non_null_count(values: dict) -> int:
    if not isinstance(values, dict):
        return 0
    c = 0
    for k, v in values.items():
        if k == "조사완료":
            continue
        if v is not None:
            c += 1
    return c


def _to_int(x) -> int:
    try:
        return int(x)
    except Exception:
        try:
            return int(str(x).replace(",", ""))
        except Exception:
            return 0


def _is_number(x) -> bool:
    if x is None:
        return False
    if isinstance(x, (int, float)):
        return True
    try:
        float(x)
        return True
    except Exception:
        return False


def _approx_equal(a, b, tol: float = 0.05) -> bool:
    """수치 셀(퍼센트)에 대한 느슨한 동등 비교.

    작은 반올림 오차(예: 59 vs 59.0 또는 59.1)를 허용한다.
    """

    if a is None and b is None:
        return True
    if a is None or b is None:
        return False

    if _is_number(a) and _is_number(b):
        try:
            return abs(float(a) - float(b)) <= tol
        except Exception:
            return False

    return str(a).strip() == str(b).strip()


def _rows_effectively_equal(existing: dict, incoming: dict) -> bool:
    if not isinstance(existing, dict) or not isinstance(incoming, dict):
        return False
    keys = set(existing.keys()) | set(incoming.keys())
    keys.discard("조사완료")
    for k in keys:
        if not _approx_equal(existing.get(k), incoming.get(k)):
            return False
    return True


def merge_category_rows(existing: dict, incoming: dict) -> dict:
    """카테고리 행 2개를 하나로 병합(충돌 중복 제거).

    규칙:
    - 둘 다 조사완료가 있으면 조사완료는 max(n1, n2)로 유지(절대 합산하지 않음)
    - 값이 더 많이 채워진 행을 우선
    - 다른 행에서 비어 있는 셀만 채움
    """

    if not isinstance(existing, dict) or not isinstance(incoming, dict):
        return existing or incoming

    n1 = _to_int(existing.get("조사완료"))
    n2 = _to_int(incoming.get("조사완료"))

    # 더 채워진 행을 기준 행으로 선택
    if _non_null_count(incoming) > _non_null_count(existing):
        base = dict(incoming)
        other = existing
    else:
        base = dict(existing)
        other = incoming

    # 조사완료는 합산하지 않고 최댓값 유지
    if n1 > 0 or n2 > 0:
        base["조사완료"] = max(n1, n2)

    # 두 행이 실질적으로 동일하면 누락된 키만 채워서 반환
    # (별칭 중복에서 자주 발생)
    if _rows_effectively_equal(existing, incoming):
        for k, v in other.items():
            if k == "조사완료":
                continue
            if base.get(k) is None and v is not None:
                base[k] = v
        return base

    # 그 외에는 누락된 셀만 보수적으로 채움
    for k, v in other.items():
        if k == "조사완료":
            continue
        if base.get(k) is None and v is not None:
            base[k] = v

    return base


def merge_table_by_category(table: Dict[str, dict]) -> Tuple[Dict[str, dict], Dict[str, list]]:
    """카테고리 키 충돌을 병합해 테이블을 중복 제거.

    반환: (merged_table, report)
      report: {"merged": [카테고리, ...], "dropped": [...]} 
    """

    if not isinstance(table, dict):
        return table, {"merged": [], "dropped": []}

    out: Dict[str, dict] = {}
    merged_report = []

    for cat, values in table.items():
        if cat not in out:
            out[cat] = values
            continue
        before = out[cat]
        out[cat] = merge_category_rows(before, values)
        merged_report.append(cat)

    return out, {"merged": merged_report, "dropped": []}
