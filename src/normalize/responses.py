"""응답 항목 정규화 및 정리.

참고
- PSR/GE 출력은 고정 템플릿에 쓰이므로 열 흔들림을 막기 위해 표준 헤더로 맞추고
  여분을 적극적으로 제거/병합한다.
"""

from __future__ import annotations

import re
from typing import List, Optional, Sequence

from .labels import normalize_label


# 템플릿 고정 헤더(순서 중요)
CANON_PSR_HEADERS: List[str] = [
    "더불어민주당",
    "국민의힘",
    "조국혁신당",
    "개혁신당",
    "진보당",
    "그 외 다른 정당",
    "지지정당없음",
    "잘 모름",
]

CANON_GE_HEADERS: List[str] = [
    "잘하고 있다",
    "잘 못하고 있다",
    "잘 모름",
]


def _is_non_response_header(label: str) -> bool:
    """기본/표본 열처럼 보이는지 여부(응답 옵션 아님)."""
    s = str(label or "").strip()
    if not s:
        return True

    s_norm = normalize_label(s)
    if s_norm in {"조사완료", "BASE"}:
        return True
    # 표본/가중치 계열 헤더
    if any(k in s_norm for k in ("사례수", "표본", "가중", "가중값", "기준")):
        return True
    return False


def _to_float(x) -> Optional[float]:
    if x is None:
        return None
    if isinstance(x, (int, float)):
        return float(x)
    if isinstance(x, str):
        s = x.strip()
        if not s:
            return None
        # 콤마 허용
        s = s.replace(",", "")
        try:
            return float(s)
        except Exception:
            return None
    return None


def _sum_numeric(a, b):
    fa, fb = _to_float(a), _to_float(b)
    if fa is None:
        return b
    if fb is None:
        return a
    # 소수 첫째 자리 유지(표 스타일 맞춤)
    return round(fa + fb, 1)


def cleanup_response_items(response_items: Sequence[str]) -> List[str]:
    """응답 항목 정리.

    - 중복 제거(순서 유지)
    - '없음/모름'이 있으면 개별 없음/모름 라벨을 줄여 열 흔들림 완화
    - GE 계열은 모름을 끝으로 보내고 긍정→부정 순서를 보장
    """

    if not response_items:
        return []

    # 1) 정규화 + 중복 제거
    seen = set()
    items: List[str] = []
    for x in response_items:
        if x is None:
            continue
        s = str(x).strip()
        if not s:
            continue
        s = normalize_label(s)
        if s in seen:
            continue
        seen.add(s)
        items.append(s)

    # 2) '없음/모름' 합산열 처리
    # 일부 PDF에 '없다/모름/무응답' 같은 합산열이 있음.
    # 합산열과 개별 열이 모두 있으면 합산열은 파생값이므로 드롭해 100% 초과와 열 흔들림을 방지.
    combined = "없음/모름"
    if combined in items:
        no_support_labels = {
            "없음", "없다", "지지정당없음", "지지 정당 없음", "지지정당 없음", "지지정당이 없다", "지지정당이없다",
        }
        unknown_labels = {
            "모름",
            "잘 모름",
            "모름/기타",
            "모름.무응답",
            "모름/무응답",
            "무응답",
            "잘모름",
            "밝힐 수 없음",
        }
        has_separate = any(x in items for x in (no_support_labels | unknown_labels))
        if has_separate:
            items = [x for x in items if x != combined]
        else:
            # 결합형만 있으면 그대로 둠
            pass

    # 3) GE 정렬 규칙
    unknown_set = {
        "모름.무응답",
        "모름/기타",
        "모름",
        "무응답",
        "잘 모름",
        "모름/무응답",
        "밝힐 수 없음",
    }
    if ("잘함" in items and "잘못함" in items) or ("잘하고 있다" in items and "잘 못하고 있다" in items):
        unknowns = [x for x in items if x in unknown_set]
        rest = [x for x in items if x not in unknown_set]

        ordered: List[str] = []
        if "잘함" in rest and "잘못함" in rest:
            ordered.extend([x for x in ("잘함", "잘못함") if x in rest])
            ordered.extend([x for x in rest if x not in ("잘함", "잘못함")])
        elif "잘하고 있다" in rest and "잘 못하고 있다" in rest:
            ordered.extend([x for x in ("잘하고 있다", "잘 못하고 있다") if x in rest])
            ordered.extend([x for x in rest if x not in ("잘하고 있다", "잘 못하고 있다")])
        else:
            ordered = rest

        items = ordered + unknowns

    return items


def normalize_psr_response_label(label: str) -> str:
    """정당지지도(PSR) 응답 라벨을 템플릿 표준에 맞게 정규화."""
    s = normalize_label(label)

    # ------------------------------------------------------------------
    # 편의용 결합 버킷(파생 열)
    #
    # 일부 통계표에는 "없음+모름"처럼 합산한 파생 열이 들어있음:
    #   - "없음/모름"
    #   - "없음/모름/무응답"
    #   - "지지정당없음/모름"
    #   - "지지 정당 없음/잘 모름"
    #
    # 이는 실제 응답항목이 아니므로 '그 외 다른 정당'으로 분류하면 안 됨.
    # PSR 정규화기(src/normalize/psr.py)가 결정적으로 제거/라우팅할 수 있도록
    # 별도 키로 유지.
    # ------------------------------------------------------------------
    compact = re.sub(r"\s+", "", s)
    compact_sep = (
        compact
        .replace("·", "/")
        .replace("ㆍ", "/")
        .replace("‧", "/")
        .replace("∙", "/")
        .replace("•", "/")
        .replace("／", "/")
        .replace("|", "/")
    )
    has_no = (
        ("없음" in compact_sep)
        or ("없다" in compact_sep)
        or ("지지정당없" in compact_sep)
        or ("지지하는정당이없" in compact_sep)
    )
    has_unk = (
        ("모름" in compact_sep)
        or ("무응답" in compact_sep)
        or ("기타" in compact_sep)
        or ("응답거절" in compact_sep)
        or ("밝힐수없" in compact_sep)
    )
    if has_no and has_unk:
        return "없음/모름"

    # 무당층/없음 처리
    # 기관 표기 차이 때문에 '지지하는 정당이 없다' 같은 변형이 자주 발생합니다.
    # PSR에서는 이런 라벨을 모두 템플릿의 '지지정당없음'으로 통일해야 합니다.
    if s in {
        "없음",
        "없다",
        "지지 정당 없음",
        "지지정당 없음",
        "지지정당없음",
        "지지정당이없다",
        "지지정당이 없다",
        "지지하는 정당이 없다",
        "지지하는 정당이없다",
        "지지하는정당이없다",
    }:
        return "지지정당없음"

    # 휴리스틱(보수적): 라벨이 명확히 "지지정당 없음"을 뜻하면 통일
    # 포착하고 싶은 예:
    # - 지지하는 정당이 없다
    # - 지지하는 정당 없음
    # - 지지정당 없다
    # 참고: compact는 위에서 계산함.
    if ("지지" in compact and "정당" in compact and ("없음" in compact or "없다" in compact or "없" in compact)):
        # '없음/모름' 같은 결합 버킷을 무당층으로 오인 흡수하지 않도록 방지
        if "모름" not in compact and "무응답" not in compact:
            return "지지정당없음"

    # 모름/무응답 처리
    if s in {
        "모름",
        "잘 모름",
        "잘모름",
        "무응답",
        "모름.무응답",
        "모름/기타",
        "모름/무응답",
        "밝힐 수 없음",
    }:
        return "잘 모름"

    # '기타' 명시 버킷
    if s in {"기타정당", "기타 정당", "그 외", "그외", "그 외 정당", "기타", "그 외 다른 정당"}:
        return "그 외 다른 정당"

    # 흔한 약칭
    if s == "민주당":
        return "더불어민주당"
    if s in {"국힘", "국민의 힘"}:
        return "국민의힘"

    # 템플릿에 있는 정당명은 그대로 유지
    if s in set(CANON_PSR_HEADERS):
        return s

    # 그 외 정당명은 '그 외 다른 정당'으로 묶어 템플릿을 안정적으로 유지
    return "그 외 다른 정당"


def normalize_ge_response_label(label: str) -> str:
    """국정운영평가(GE) 응답 라벨을 템플릿 표준에 맞게 정규화."""
    s = normalize_label(label)

    if s in {"잘함", "잘 함", "잘하고 있음", "잘하고 있다", "긍정", "긍정평가", "긍정적평가", "긍정적 평가", "긍정 평가"}:
        return "잘하고 있다"
    if s in {"잘못함", "잘 못함", "잘못하고 있음", "잘 못하고 있다", "부정", "부정평가", "부정적평가", "부정적 평가", "부정 평가"}:
        return "잘 못하고 있다"
    if s in {"모름.무응답", "모름/기타", "모름", "무응답", "잘 모름", "밝힐 수 없음", "잘모름"}:
        return "잘 모름"

    if s in set(CANON_GE_HEADERS):
        return s

    return s


def normalize_issue_response_label(label: str) -> str:
    """ISSUE 응답 라벨을 가볍게 정규화.

    과한 재작성은 피하되 모름/무응답 변형은 공통으로 정리합니다.
    """
    # 참고:
    # ISSUE는 질문/응답지가 매번 달라서 가능한 한 "표에 나온 헤더 그대로"를 유지하는 편이 안전합니다.
    # 특히 '없다/모름/무응답' 같은 **편의 합산열** 또는 **합성 라벨**은
    # 이를 임의로 '잘 모름' 등으로 흡수하면 (1) PDF에 없는 헤더가 생기거나
    # (2) 별도 컬럼이 덮어써지는 부작용이 발생할 수 있습니다.
    s = normalize_label(label)

    # 합산 버킷: 있는 그대로 유지(모름/무응답으로 매핑하지 않음)
    compact = s.replace(" ", "")
    has_no = ("없다" in compact) or ("없음" in compact)
    has_unk = ("모름" in compact) or ("무응답" in compact) or ("응답거절" in compact)
    if has_no and has_unk:
        return s

    # ISSUE에서는 모름/무응답 라벨을 과도하게 합치지 않음
    # (중복 정리는 cleanup_response_items에서 처리)
    return s


def normalize_response_items_by_type(item_type: str, response_items: Sequence[str]) -> List[str]:
    it = (item_type or "").upper()
    if it == "PSR":
        # 고정 템플릿: 원본 변동과 무관하게 안정된 헤더 집합 유지
        return CANON_PSR_HEADERS.copy()

    if it == "GE":
        return CANON_GE_HEADERS.copy()

    if not response_items:
        return []

    mapped: List[str] = []
    for x in response_items:
        if x is None:
            continue
        s = str(x).strip()
        if not s:
            continue
        if _is_non_response_header(s):
            continue
        mapped.append(normalize_issue_response_label(s))

    return cleanup_response_items(mapped)


def normalize_value_keys_by_type(item_type: str, values: dict) -> dict:
    """단일 카테고리 행의 값 dict 내부 키를 유형별로 정규화."""
    if not isinstance(values, dict):
        return values

    it = (item_type or "").upper()
    out: dict = {}

    for k, v in values.items():
        if k is None:
            continue

        k_str = str(k).strip()
        if not k_str:
            continue

        if k_str == "조사완료":
            out["조사완료"] = v
            continue

        if _is_non_response_header(k_str):
            # 모델이 sample/base 열을 생성했더라도 제거
            continue

        if it == "PSR":
            nk = normalize_psr_response_label(k_str)
            if nk in out:
                out[nk] = _sum_numeric(out[nk], v)
            else:
                out[nk] = v
        elif it == "GE":
            nk = normalize_ge_response_label(k_str)
            # 충돌 시 None이 아닌 값을 우선(합산 금지: 파생 열이 이중 집계될 수 있음)
            if nk in out and out[nk] is not None:
                continue
            out[nk] = v
        else:
            nk = normalize_issue_response_label(k_str)
            if nk in out and out[nk] is not None:
                continue
            out[nk] = v

    # PSR/GE에 대해 고정 템플릿 헤더 강제
    if it == "PSR":
        # 일부 구/파생 키(예: '없음/모름')를 잠시 유지해
        # PSR 전용 정규화기가 결정적으로 제거/라우팅하도록 함
        allowed = set(CANON_PSR_HEADERS) | {"조사완료", "없음/모름"}
        out = {k: v for k, v in out.items() if k in allowed}
    elif it == "GE":
        allowed = set(CANON_GE_HEADERS) | {"조사완료"}
        out = {k: v for k, v in out.items() if k in allowed}

    return out
