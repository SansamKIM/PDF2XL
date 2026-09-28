"""PSR(정당 지지도) 전용 정규화.

PSR 열은 템플릿에서 **고정**이며 최신 통합 기준으로 전국/지방이 동일 헤더를 쓴다:
  더불어민주당, 국민의힘, 조국혁신당, 개혁신당, 진보당,
  그 외 다른 정당, 지지정당없음, 잘 모름

필수 처리:
- 주요 정당은 그대로 유지
- 그 외 정당 이름은 '그 외 다른 정당'에 합산
- '없음/모름(/응답거절)' 같은 합산 버킷은 결정적으로 분배
  (분리 값이 없으면 '지지정당없음'으로 보내고 '잘 모름'은 비움)
- 템플릿과 맞는 안정적 응답 항목 목록을 출력
"""

from __future__ import annotations

from typing import Any, Dict, List


MAJOR_PARTIES: List[str] = [
    "더불어민주당",
    "국민의힘",
    "조국혁신당",
    "개혁신당",
    "진보당",
]


def psr_template_response_items(survey_type: str | None = None) -> List[str]:
    """템플릿이 기대하는 PSR 응답 헤더 반환."""
    _ = survey_type  # 하위 호환 유지용
    return MAJOR_PARTIES + ["그 외 다른 정당", "지지정당없음", "잘 모름"]


def _is_combined_none_unknown(label: str) -> bool:
    """휴리스틱: '없음/모름(/응답거절)' 같은 합산 버킷 감지."""
    s = str(label or "")
    return ("없음" in s and "모름" in s) or ("지지정당" in s and "모름" in s)


def _to_float(v: Any) -> float | None:
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).strip())
    except Exception:
        return None


def _round1(x: float) -> float:
    # 대부분 표가 소수 첫째 자리이므로 일관 유지(부동소수 오차 방지)
    return round(float(x), 1)


def normalize_psr_table(data: dict, survey_type: str) -> dict:
    """data['데이터'] 내 PSR 응답 키를 정규화.

    전제: 파이프라인에서 이미 다음을 적용함
    - 기관 용어 매핑
    - normalize_label(cat)
    - normalize_value_keys_by_type('PSR', ...)

    동작 원칙(결정적·보수적):
    - 주요 정당은 유지
    - 나머지는 '그 외 다른 정당'으로 합산
    - '지지정당없음'과 '잘 모름'은 분리 유지
      (합산 버킷만 있으면 '지지정당없음'으로 보냄)
    """

    if not isinstance(data, dict):
        return data

    table = data.get("데이터")
    if not isinstance(table, dict):
        return data

    major_set = set(MAJOR_PARTIES)

    # 템플릿 버킷(통일)
    other_key = "그 외 다른 정당"
    no_support_key = "지지정당없음"
    unknown_key = "잘 모름"
    combined_key = "없음/모름"  # 일부 PDF의 레거시 라벨

    new_table: Dict[str, Dict[str, Any]] = {}

    for cat, values in table.items():
        if not isinstance(values, dict):
            # 그대로 유지(PSR에서는 드물지만 안전하게 처리)
            new_table[cat] = values
            continue

        out: Dict[str, Any] = {}

        # 조사완료(사례수) 유지
        if "조사완료" in values:
            out["조사완료"] = values.get("조사완료")

        has_separate = (no_support_key in values) or (unknown_key in values)

        other_sum = 0.0
        other_seen = False

        for k, v in values.items():
            if k == "조사완료":
                continue
            if v is None:
                continue

            # 주요 정당
            if k in major_set:
                out[k] = v
                continue

            # 무당층/모름 열이 명시돼 있으면 유지
            if k == no_support_key:
                out[no_support_key] = v
                continue
            if k == unknown_key:
                out[unknown_key] = v
                continue

            # 결합 버킷 변형 -> 무당층으로 매핑(이미 분리 열이 있으면 중복 집계를 피하기 위해 건너뜀)
            if k == combined_key or _is_combined_none_unknown(k):
                if has_separate:
                    continue
                out[no_support_key] = v
                continue

            # 기타 버킷 동의어
            if k in {"그 외 다른 정당", "기타 정당"}:
                other_seen = True
                num = _to_float(v)
                if num is not None:
                    other_sum += num
                continue

            # 그 외 정당명 -> 기타 버킷
            other_seen = True
            num = _to_float(v)
            if num is not None:
                other_sum += num
            continue

        if other_seen:
            out[other_key] = _round1(other_sum)

        new_table[cat] = out

    out_data = dict(data)
    out_data["데이터"] = new_table
    out_data["응답항목"] = psr_template_response_items(survey_type)
    return out_data
