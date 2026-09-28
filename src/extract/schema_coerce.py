"""LLM이 추출한 표 JSON을 표준 스키마로 강제/정규화.

모델이 기대하는 표준 형태:

  {
    "질문": "...",
    "응답항목": ["...", ...],
    "데이터": {"전체": {"조사완료": 800, "...": 12.3, ...}, ...}
  }

실제 실행에서 자주 보이는 변형:

1) 상위 키 동의어
   - 응답항목: 헤더/열헤더/응답지/columns/headers/choices/...
   - 데이터: data/rows/values/result/표/...
   - 질문: 문항/question/title/...

2) row 리스트 스타일
   - 데이터가 list[dict]이고 각 row에 "분류"/"구분"/"category" 등으로 카테고리명이 들어있는 경우
   - 데이터가 dict[str, list] 형태(열 헤더 리스트와 값 배열)인 경우

이 모듈은 *최대한* 결정적으로 강제 변환해:
- 검증을 더 탄탄하게 하고
- 후속 정규화/Excel 작성이 안정된 스키마에서 동작하도록 한다.
"""

from __future__ import annotations

import re
from typing import Any, Dict, Iterable, List, Optional


_RE_KEEP = re.compile(r"[^0-9A-Za-z가-힣]+")


def _norm_key(s: Any) -> str:
    """퍼지 매칭을 위한 dict 키 정규화."""
    if s is None:
        return ""
    t = str(s).strip()
    if not t:
        return ""
    t = _RE_KEEP.sub("", t)
    return t.lower()


def _pick_key_by_norm(d: Dict[str, Any], candidates: Iterable[str]) -> Optional[str]:
    cset = {str(x).lower() for x in candidates if x}
    for k in d.keys():
        if _norm_key(k) in cset:
            return k
    return None


def _unwrap_nested(raw: Dict[str, Any]) -> Dict[str, Any]:
    """자주 보이는 중첩 패턴을 풀어냄(최대한)."""
    if not isinstance(raw, dict):
        return raw

    # 모델이 {"result": {...}}나 {"표": {...}} 등을 반환한 경우
    for wrapper_norm in (
        "result",
        "results",
        "output",
        "table",
        "tabledata",
        "표",
        "표데이터",
        "결과",
        "응답",
    ):
        k = _pick_key_by_norm(raw, [wrapper_norm])
        if not k:
            continue
        v = raw.get(k)
        if not isinstance(v, dict):
            continue

        # 내부가 이미 표준일 수도 있고 동의어 키를 쓸 수도 있음
        has_schema = any(x in v for x in ("응답항목", "데이터"))
        if not has_schema:
            has_schema = _pick_key_by_norm(v, ["응답항목", "응답지", "헤더", "columns", "headers"]) is not None
            has_schema = has_schema or (_pick_key_by_norm(v, ["데이터", "data", "rows", "values", "table", "result"]) is not None)

        if has_schema:
            # 병합: 바깥 키를 폴백으로 두고 내부 값을 우선
            merged = dict(raw)
            merged.update(v)
            return merged
    return raw


def _coerce_columns_values_schema(payload: Dict[str, Any]) -> Optional[Dict[str, Any]]:
    """'columns + rows(values array)' 스키마를 표준 표 스키마로 강제.

    기대 입력(키는 변형 가능):
      {
        "columns": [...],
        "rows": [
          {"section": "...", "label": "...", "조사완료": 519, "values": [...]},
          ...
        ]
      }

    출력:
      {
        "응답항목": [...],
        "데이터": {"전체": {"조사완료": 1022, "...": 42.0, ...}, ...}
      }

    메모:
    - 'label'을 카테고리 키로 유지하되, 섹션마다 중복되는 '모름/무응답'은
      직업 섹션에 '(직업)' 접미어를 붙여 구분한다.
    """

    if not isinstance(payload, dict):
        return None

    # columns 위치 찾기
    k_cols = _pick_key_by_norm(payload, ["columns", "cols", "headers", "header", "응답항목", "헤더", "열헤더"])
    columns = payload.get(k_cols) if k_cols else None
    if not isinstance(columns, list) or not columns:
        return None

    # rows 위치 찾기
    k_rows = _pick_key_by_norm(payload, ["rows", "row", "데이터", "data", "values", "table"])
    rows = payload.get(k_rows) if k_rows else None
    if not isinstance(rows, list) or not rows:
        return None

    # values 배열을 가진 행 탐지
    has_values = False
    for r in rows:
        if isinstance(r, dict):
            kv = _pick_key_by_norm(r, ["values", "vals", "value", "값", "배열"])
            if kv and isinstance(r.get(kv), list):
                has_values = True
                break
    if not has_values:
        return None

    def _is_unknown_label(label: Any) -> bool:
        s = str(label or "").strip()
        if not s:
            return False
        sn = s.replace(" ", "")
        return ("모름" in sn) or ("무응답" in sn) or ("응답거절" in sn)

    def _is_job_section(section: Any) -> bool:
        s = str(section or "").strip()
        if not s:
            return False
        return "직업" in s

    def _is_ideology_section(section: Any) -> bool:
        s = str(section or "").strip()
        if not s:
            return False
        return ("이념" in s) or ("정치" in s and "성향" in s)

    def _make_category_key(section: Any, label: Any) -> Optional[str]:
        lab = str(label or "").strip()
        if not lab:
            return None
        sec = str(section or "").strip()

        # 전체 라벨의 공백/기호 아티팩트 정규화
        # 촘촘한 표에서 "전체"가 "전 체"나 도형과 함께 나올 수 있어 통일
        lab_compact = re.sub(r"[\s▣□■●○]+", "", lab)
        if lab_compact == "전체":
            lab = "전체"
        # 섹션마다 중복되는 '모름/무응답' 구분
        if _is_unknown_label(lab):
            # 직업 섹션은 전용 템플릿 행으로 매핑
            if _is_job_section(sec):
                return "모름/무응답(직업)"
            # 이념 섹션의 모름은 템플릿에 이미 있어 접미어 없음
            if _is_ideology_section(sec):
                return lab
            # 기타 섹션은 덮어쓰기 방지용 접미어 추가
            if sec:
                return f"{lab}({sec})"
        return lab

    # 열은 문자열로 표준화(원문 유지, 후속 정규화가 매핑 처리)
    cols_out: List[str] = []
    for c in columns:
        if c is None:
            continue
        s = str(c).strip()
        if not s:
            continue
        cols_out.append(s)

    table: Dict[str, Dict[str, Any]] = {}

    for r in rows:
        if not isinstance(r, dict):
            continue

        k_label = _pick_key_by_norm(r, ["label", "분류", "구분", "category", "name", "항목", "행", "row"])
        k_section = _pick_key_by_norm(r, ["section", "group", "대분류", "블록", "구간"])
        label = r.get(k_label) if k_label else None
        section = r.get(k_section) if k_section else None

        cat = _make_category_key(section, label)
        if not cat:
            continue

        # 사례수
        k_n = _pick_key_by_norm(r, ["조사완료", "사례수", "표본수", "n", "base"])
        n_val = r.get(k_n) if k_n else None

        # values 배열
        k_vals = _pick_key_by_norm(r, ["values", "vals", "value", "값", "배열"])
        arr = r.get(k_vals) if k_vals else None
        if not isinstance(arr, list):
            continue

        # 길이 맞추기: columns 길이에 패딩/자르기
        arr2 = list(arr[: len(cols_out)])
        if len(arr2) < len(cols_out):
            arr2.extend([None] * (len(cols_out) - len(arr2)))

        row_vals: Dict[str, Any] = {}
        if n_val is not None:
            row_vals["조사완료"] = n_val

        for i, h in enumerate(cols_out):
            row_vals[h] = arr2[i]

        # 충돌 시 기존 널 아님 값을 우선 병합
        if cat in table:
            existing = table[cat]
            if "조사완료" not in existing and "조사완료" in row_vals:
                existing["조사완료"] = row_vals.get("조사완료")
            for kk, vv in row_vals.items():
                if kk == "조사완료":
                    continue
                if kk not in existing or existing.get(kk) is None:
                    existing[kk] = vv
        else:
            table[cat] = row_vals

    if not table:
        return None

    out: Dict[str, Any] = {}
    # 질문 키가 있으면 유지
    kq = _pick_key_by_norm(payload, ["질문", "문항", "question", "title", "q"])
    if kq and isinstance(payload.get(kq), str):
        out["질문"] = payload.get(kq)

    out["응답항목"] = cols_out
    out["데이터"] = table
    return out


def _coerce_rows_list(rows: List[Any]) -> Dict[str, Dict[str, Any]]:
    """row dict 리스트를 표준 카테고리 -> 행 dict 테이블로 변환."""
    out: Dict[str, Dict[str, Any]] = {}
    if not rows:
        return out

    cat_keys_norm = {
        "분류",
        "구분",
        "category",
        "cat",
        "label",
        "name",
        "row",
        "행",
        "항목",
        "항목명",
    }
    # 성능을 위해 한 번만 정규화
    cat_keys_norm = {str(x).lower() for x in cat_keys_norm}

    for r in rows:
        if not isinstance(r, dict):
            continue

        cat: Optional[str] = None
        for k in r.keys():
            if _norm_key(k) in cat_keys_norm:
                v = r.get(k)
                if v is not None:
                    s = str(v).strip()
                    if s:
                        cat = s
                        break

        if not cat:
            # 카테고리 키를 찾지 못하면 행 건너뜀
            continue

        row_vals: Dict[str, Any] = {}
        for k, v in r.items():
            if _norm_key(k) in cat_keys_norm:
                continue

            # 사례수 동의어를 조사완료로 강제
            nk = _norm_key(k)
            if nk in {"조사완료", "사례수", "표본수", "n", "base"}:
                row_vals["조사완료"] = v
            else:
                row_vals[str(k)] = v

        # 충돌 시 기존 널 아님 값을 우선 병합
        if cat in out:
            for kk, vv in row_vals.items():
                if kk not in out[cat] or out[cat][kk] is None:
                    out[cat][kk] = vv
        else:
            out[cat] = row_vals

    return out


def _derive_headers_from_data(data: Dict[str, Any]) -> List[str]:
    """첫 행 dict로부터 그럴듯한 응답항목 리스트 유도."""
    if not isinstance(data, dict) or not data:
        return []
    first_row = None
    for v in data.values():
        if isinstance(v, dict) and v:
            first_row = v
            break
    if not isinstance(first_row, dict) or not first_row:
        return []

    headers: List[str] = []
    for k in first_row.keys():
        if k is None:
            continue
        ks = str(k).strip()
        if not ks:
            continue
        # 응답항목에서는 조사완료 제외(후속 정규화도 제거)
        nk = _norm_key(ks)
        if nk in {"조사완료", "base"}:
            continue
        if any(x in nk for x in ("사례수", "표본", "가중", "기준")):
            continue
        headers.append(ks)
    return headers


def coerce_extracted_table_schema(raw: Any) -> Dict[str, Any]:
    """최대한의 강제 변환. 변환 실패 시 빈 dict 반환."""

    # 0) 리스트 페이로드 감싸기
    if isinstance(raw, list):
        raw = {"데이터": raw}

    if not isinstance(raw, dict):
        return {}

    raw = _unwrap_nested(raw)
    out: Dict[str, Any] = dict(raw)

    # 0.5) 특수: columns+rows(values[]) 스키마(촘촘한 표)
    try:
        coerced = _coerce_columns_values_schema(out)
        if isinstance(coerced, dict) and coerced.get("데이터") and coerced.get("응답항목"):
            merged = dict(out)
            merged.update(coerced)
            out = merged
    except Exception:
        # 최선 시도만 하고 실패 무시
        pass

    # 1) 상위 키 표준화
    if "질문" not in out or not isinstance(out.get("질문"), str):
        kq = _pick_key_by_norm(out, ["질문", "문항", "question", "title", "q"])
        if kq and isinstance(out.get(kq), str):
            out["질문"] = out.get(kq)

    if "응답항목" not in out or not isinstance(out.get("응답항목"), list):
        kr = _pick_key_by_norm(
            out,
            [
                "응답항목",
                "응답지",
                "응답",
                "헤더",
                "열헤더",
                "column",
                "columns",
                "headers",
                "choices",
                "options",
                "answeroptions",
                "responses",
            ],
        )
        if kr and isinstance(out.get(kr), list):
            out["응답항목"] = out.get(kr)

    if "데이터" not in out or not isinstance(out.get("데이터"), (dict, list)):
        kd = _pick_key_by_norm(
            out,
            [
                "데이터",
                "data",
                "rows",
                "values",
                "result",
                "results",
                "table",
                "tabledata",
                "표",
                "표데이터",
                "결과",
            ],
        )
        if kd and isinstance(out.get(kd), (dict, list)):
            out["데이터"] = out.get(kd)

    # 2) row 리스트를 dict 테이블로 변환
    data = out.get("데이터")
    if isinstance(data, list):
        out["데이터"] = _coerce_rows_list(data)

    # 3) dict[str, list] 행을 헤더로 강제 변환
    data = out.get("데이터")
    headers = out.get("응답항목")
    if isinstance(data, dict) and data and isinstance(headers, list) and headers:
        if all(isinstance(v, list) for v in data.values()):
            new_table: Dict[str, Dict[str, Any]] = {}
            for cat, arr in data.items():
                if not isinstance(arr, list):
                    continue
                row: Dict[str, Any] = {}
                for i, h in enumerate(headers):
                    if i >= len(arr):
                        break
                    row[str(h)] = arr[i]
                new_table[str(cat)] = row
            out["데이터"] = new_table

    # 4) 데이터가 있는데 헤더가 없으면 유도
    if (not isinstance(out.get("응답항목"), list)) or (not out.get("응답항목")):
        if isinstance(out.get("데이터"), dict) and out.get("데이터"):
            out["응답항목"] = _derive_headers_from_data(out["데이터"])

    return out
