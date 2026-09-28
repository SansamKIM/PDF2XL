"""추출된 표 JSON에 대한 경량 검증.

값을 '교정'하지는 않으며, 결과가 충분히 그럴듯한지 혹은 다른 이미지/프롬프트로 재시도할지 판단만 한다.

필요 이유
- 저렴한 비전 모델은 가끔 열 누락, 행 밀림, 비수치 문자열을 낸다.
- 무작정 재시도하면 호출 비용이 커서, 강한 신호가 있을 때만 재시도한다.

설계
- 보수적: 오탐을 최소화(과도한 재시도 방지).
- 일반적: 기관별 하드코딩 없음.
"""

from __future__ import annotations

import math
import re
from typing import Any, Dict, List, Optional


def _is_number(v: Any) -> bool:
    return isinstance(v, (int, float)) and not (isinstance(v, float) and (math.isnan(v) or math.isinf(v)))



def validate_table_data(item_type: str, table_data: Dict[str, Any]) -> List[str]:
    """문제 목록을 반환(빈 리스트면 수용 가능)."""

    issues: List[str] = []

    if not isinstance(table_data, dict):
        return ["table_data is not dict"]

    resp = table_data.get("응답항목")
    data = table_data.get("데이터")

    # 응답항목이 없어도 행 dict가 있으면 첫 행으로 헤더를 유도해 불필요한 재시도 방지
    if (not isinstance(resp, list) or not resp) and isinstance(data, dict) and data:
        try:
            first_row = next((v for v in data.values() if isinstance(v, dict) and v), None)
            if isinstance(first_row, dict):
                derived = []
                for k in first_row.keys():
                    ks = str(k).strip() if k is not None else ""
                    if not ks:
                        continue
                    ks_norm = ks.replace(" ", "")
                    if ks_norm in {"조사완료", "BASE"}:
                        continue
                    if any(x in ks_norm for x in ("사례수", "표본", "가중", "기준")):
                        continue
                    derived.append(ks)
                if derived:
                    table_data["응답항목"] = derived
                    resp = derived
        except Exception:
            pass

    if not isinstance(resp, list) or not resp:
        issues.append("missing 응답항목")

    if not isinstance(data, dict) or not data:
        issues.append("missing 데이터")
        return issues

    def _norm_overall_key(x: Any) -> str:
        s = str(x or "")
        # 전체 라벨 주변에 붙는 공백/박스/불릿 등을 제거(예: "전 체", "▣전체▣")
        s = re.sub(r"[\s▣□■●○]+", "", s)
        return s

    # PSR/GE는 '전체' 행이 있는 것이 바람직
    itype = (item_type or "").upper()
    if itype in {"PSR", "GE"} and "전체" not in data:
        # 가능한 한 OCR 변형을 '전체'로 통일
        for k in list(data.keys()):
            if _norm_overall_key(k) == "전체":
                try:
                    data["전체"] = data.pop(k)
                except Exception:
                    pass
                break

    if itype in {"PSR", "GE"} and "전체" not in data:
        issues.append("missing 전체 row")

    # ISSUE: 후보명 임의 치환(후보1/후보2/...) 방지
    # - 후보명은 표에 있는 그대로가 원칙이며, '후보N'은 거의 항상 잘못된 추출 신호
    # - 크롭으로 헤더가 잘렸거나 모델이 임의 축약한 경우이므로, 검증 단계에서 실패시켜 재시도(full/strict) 유도
    if itype == "ISSUE" and isinstance(resp, list):
        ph = []
        for x in resp:
            if not isinstance(x, str):
                continue
            s = x.strip()
            if re.fullmatch(r"후보\s*\d+", s) or re.fullmatch(r"candidate\s*\d+", s, flags=re.I):
                ph.append(s)
        # 보수적으로, 플레이스홀더가 여러 개일 때만 트리거
        if len(ph) >= 2:
            issues.append("placeholder 후보명 detected: " + ", ".join(ph[:8]))

    # 수치 검증: 음수/100 초과 금지(소수 반올림은 허용)
    # 참고: "조사완료"는 사례수(N)라 100을 넘는 것이 정상이며 퍼센트/합계 검증 대상 아님
    numeric_cells = 0
    bad_cells = 0

    for _, row_val in data.items():
        if not isinstance(row_val, dict):
            continue
        for k, v in row_val.items():
            # "조사완료"/사례수/BASE 류 컬럼은 퍼센트 검증 대상 아님
            k_str = str(k) if k is not None else ""
            if k_str in {"조사완료", "BASE"} or ("사례수" in k_str) or ("표본" in k_str):
                continue

            if _is_number(v):
                numeric_cells += 1
                if v < -0.5 or v > 100.5:
                    bad_cells += 1

    if numeric_cells < 6:
        issues.append("too few numeric cells")

    if bad_cells > 0:
        issues.append(f"out-of-range numbers: {bad_cells}")

    # '전체' 합계 점검(약함): 충분한 수치값이 있을 때만 수행
    if "전체" in data and isinstance(data.get("전체"), dict):
        row = data["전체"]
        # 합계 검증도 사례수 컬럼은 제외
        nums = []
        for k, v in row.items():
            k_str = str(k) if k is not None else ""
            if k_str in {"조사완료", "BASE"} or ("사례수" in k_str) or ("표본" in k_str):
                continue
            if _is_number(v):
                nums.append(v)
        if len(nums) >= 6:
            total = sum(nums)
            if itype == "GE":
                # GE는 100 근처(여유를 넓게 설정)
                if total < 80 or total > 120:
                    issues.append(f"전체 합계 비정상(GE): {total:.1f}")
            elif itype == "PSR":
                # PSR은 파생 열(예: 무당층)이 포함될 수 있어 허용 범위를 넓게 잡음
                if total < 60 or total > 160:
                    issues.append(f"전체 합계 비정상(PSR): {total:.1f}")

    return issues


def validate_expected_groups(
    table_data: Dict[str, Any],
    include: Any = None,
    region_info: Any = None,
) -> List[str]:
    """주요 행 그룹(지역/연령/성별)이 기대될 때 존재 여부를 검증.

    매우 보수적으로 동작해, 다른 주요 그룹이 하나라도 있을 때만 누락을 표시하여
    '전체'만 있는 합법적 표를 과도하게 재시도하지 않게 함.
    """

    issues: List[str] = []

    if not isinstance(table_data, dict):
        return ["table_data is not dict"]

    data = table_data.get("데이터")
    if not isinstance(data, dict) or not data:
        return []

    # 키 정규화(공백/도형 기호 제거)
    def _kn(x: Any) -> str:
        s = str(x or "")
        s = re.sub(r"[\s▣□■●○]+", "", s)
        return s.strip()

    keys_norm = {_kn(k) for k in data.keys()}

    # --- include 설정으로 기대 그룹 추출 ---
    if include is None:
        inc_items: List[str] = []
    elif isinstance(include, str):
        inc_items = [include]
    elif isinstance(include, list):
        inc_items = [str(x) for x in include if str(x).strip()]
    else:
        inc_items = [str(include)]

    inc_blob = " ".join(inc_items)

    expect_region = any(tok in inc_blob for tok in ["지역", "권역", "거주", "권역별", "지역별"])
    expect_age = any(tok in inc_blob for tok in ["연령", "나이", "연령별", "연령대"])
    expect_gender = any(tok in inc_blob for tok in ["성별", "남녀"])

    # --- 실제 그룹 존재 여부 감지 ---
    def _has_gender() -> bool:
        for g in ("남성", "여성", "남자", "여자"):
            if g in keys_norm:
                return True
        # '남'/'여'가 라벨로 쓰이는 경우
        if "남" in keys_norm and "여" in keys_norm:
            return True
        return False

    age_re = re.compile(
        r"("  # 그룹 시작
        r"\d{1,2}~\d{1,2}세"  # 18~29세
        r"|\d{1,2}세이상"     # 70세이상
        r"|\d{1,2}세이하"     # 18세이하
        r"|\d{1,2}대"         # 30대
        r"|\d{1,2}0대"        # 20대(드문 변형)
        r")"
    )

    def _has_age() -> bool:
        for k in keys_norm:
            if age_re.search(k):
                return True
            # 공백 변형 패턴
            if re.search(r"\d{1,2}\s*~\s*\d{1,2}\s*세", str(k)):
                return True
        return False

    def _region_names_from_region_info() -> List[str]:
        names: List[str] = []
        ri = region_info
        if isinstance(ri, dict):
            rlist = ri.get("권역") or ri.get("regions") or []
        elif isinstance(ri, list):
            rlist = ri
        else:
            rlist = []

        if isinstance(rlist, list):
            for r in rlist:
                if isinstance(r, dict):
                    nm = r.get("name") or r.get("권역명") or r.get("권역")
                else:
                    nm = r
                nm_s = _kn(nm)
                if nm_s:
                    names.append(nm_s)
        return names

    def _has_region() -> bool:
        region_names = _region_names_from_region_info()
        # region_info에 명시된 목록이 있으면 우선 사용
        if region_names:
            for nm in region_names:
                if nm in keys_norm:
                    return True
            # 느슨한 매칭: 키에 지역명이 포함되면 인정(접두/접미 처리)
            for k in keys_norm:
                for nm in region_names:
                    if nm and (nm in k or k in nm):
                        return True
        # 폴백 휴리스틱: '권/권역' 접미사(예: 서북권)
        for k in keys_norm:
            if k.endswith("권") or k.endswith("권역"):
                return True
        # 또 다른 폴백: 흔한 광역시/도 라벨
        common = ["서울", "부산", "대구", "인천", "광주", "대전", "울산", "세종",
                  "경기", "강원", "충북", "충남", "전북", "전남", "경북", "경남", "제주"]
        if any(c in keys_norm for c in common):
            return True
        return False

    has_gender = _has_gender()
    has_age = _has_age()
    has_region = _has_region()

    # 주요 그룹이 최소 하나 있을 때만 누락을 표시
    major_present = (has_gender or has_age or has_region)

    if expect_region and major_present and not has_region:
        issues.append("missing 지역/권역 rows")
    if expect_age and major_present and not has_age:
        issues.append("missing 연령 rows")
    if expect_gender and major_present and not has_gender:
        issues.append("missing 성별 rows")

    # 행 단위 데이터 존재 여부 점검:
    # - 사례수가 충분한데(>50) 퍼센트 값이 하나도 없으면 누락으로 간주
    def _row_sample(row: Dict[str, Any]) -> Optional[float]:
        if not isinstance(row, dict):
            return None
        for rk, rv in row.items():
            nk = _kn(rk)
            if nk in {"조사완료", "사례수", "표본크기"}:
                try:
                    return float(rv)
                except Exception:
                    continue
        return None

    for row_name, row in data.items():
        if not isinstance(row, dict):
            continue
        sample_raw = _row_sample(row)
        sample = sample_raw or 0.0
        if sample <= 50 and sample_raw is not None:
            # 사례수가 매우 적은 그룹은 공란 허용
            continue
        numeric_vals = []
        for ck, cv in row.items():
            nk = _kn(ck)
            if nk in {"조사완료", "사례수", "표본크기"}:
                continue
            if _is_number(cv):
                numeric_vals.append(cv)
        if not numeric_vals:
            # 사례수는 있는데 값이 없는 경우
            if sample_raw is not None and sample > 50:
                issues.append(f"row '{row_name}' has no data (n={int(sample)})")
            # 사례수 자체도 없고 값도 없는 경우
            elif sample_raw is None:
                issues.append(f"row '{row_name}' has no data and no sample")

    return issues


def validate_psr_strict(
    table_data: Dict[str, Any],
    required_cols: List[str] | None = None,
    tol: float = 5.0,
    max_row_checks: int = 25,
) -> List[str]:
    """촘촘한 PSR(정당지지도) 표에 대한 강화 검증.

    동기:
    - 촘촘한 PSR 표는 간혹 '열이 한 칸 밀리는(one-column shift)' 오류가 발생한다.
      이 경우 전체 합계는 여전히 100 근처로 보일 수 있지만,
      특정 필수 열(예: 지지정당없음)이 통째로 사라질 수 있다.

    이 검증은 전역으로 항상 켜는 용도가 아니라,
    추출기에서 기관/항목 범위로 선택적으로 활성화하도록 설계됐다.
    """

    issues: List[str] = []

    if not isinstance(table_data, dict):
        return ["table_data is not dict"]

    data = table_data.get("데이터")
    if not isinstance(data, dict) or not data:
        return ["missing 데이터"]

    # '전체' 행이 있는지 보장
    overall = data.get("전체")
    if not isinstance(overall, dict) or not overall:
        # 흔한 OCR 변형(예: "전 체")을 먼저 시도
        for k in list(data.keys()):
            if re.sub(r"[\s▣□■●○]+", "", str(k or "")) == "전체":
                try:
                    data["전체"] = data.pop(k)
                except Exception:
                    pass
                break
        overall = data.get("전체")

    if not isinstance(overall, dict) or not overall:
        issues.append("missing 전체 row")
        return issues

    def _k_norm(x: Any) -> str:
        s = str(x or "")
        s = re.sub(r"\s+", "", s)
        s = re.sub(r"[^0-9A-Za-z가-힣]", "", s)
        return s.lower()

    # 필수 컬럼 존재 여부 점검(휴리스틱)
    req = required_cols or []
    if isinstance(req, list):
        row_keys_norm = [_k_norm(k) for k in overall.keys()]
        for r in req:
            rn = _k_norm(r)
            if not rn:
                continue

            found = False

            # 1) 직접 일치 / 부분 문자열 일치
            for kn in row_keys_norm:
                if rn == kn or rn in kn or kn in rn:
                    found = True
                    break

            # 2) '지지정당없음' 동의어 휴리스틱(특수 규칙)
            if not found and rn in {"지지정당없음", "지지정당없다", "지지정당 없음"}:
                for kn in row_keys_norm:
                    # 예: '지지하는정당이없다', '지지정당없음', '지지정당이없다'
                    if ("지지" in kn) and ("없" in kn):
                        found = True
                        break

            if not found:
                issues.append(f"missing required column in 전체 row: {r}")

    # 행 합계 점검(~100 ± tol)
    checked = 0
    for row_name, row in data.items():
        if checked >= max_row_checks:
            break
        if not isinstance(row, dict):
            continue

        nums: List[float] = []
        for k, v in row.items():
            k_str = str(k) if k is not None else ""
            if k_str in {"조사완료", "BASE"} or ("사례수" in k_str) or ("표본" in k_str):
                continue

            if v is None:
                nums.append(0.0)
                continue
            if _is_number(v):
                nums.append(float(v))

        # 의미 있으려면 최소한 열이 몇 개 있어야 함
        if len(nums) < 6:
            continue

        total = sum(nums)
        if total < (100.0 - float(tol)) or total > (100.0 + float(tol)):
            issues.append(f"row sum out of range: {row_name}={total:.1f}")
            # 경고가 너무 많아지지 않도록 제한
            if len(issues) >= 6:
                break

        checked += 1

    return issues
