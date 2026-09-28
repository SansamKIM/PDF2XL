"""PDF -> JSON 추출 상위 파이프라인.
주요 변경점:
- 기관 설정을 src/config.org에서 로드/검증
- PDF 렌더링 결과를 페이지/해상도별로 캐시
- 중간 산출물을 재사용해 이미 JSON이 있으면 모델 호출을 건너뜀

JSON 출력 형식은 기존 엑셀 라이터와 호환됩니다.
- 메타데이터.json
- 권역설명.json (지방 전용)
- 항목명, 항목유형, 페이지, 메타데이터, (권역설명), 질문, 응답항목, 데이터가 담긴 표 항목별 JSON 1개
"""

from __future__ import annotations

import json
import os
import re
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Tuple

from datetime import datetime

from PIL import Image

from .json_parse import parse_json_response
from .schema_coerce import coerce_extracted_table_schema
from .openai_client import OpenAIChatClient, OpenAIOptions
from .pdf_renderer import PDFPageRenderer, RenderOptions, default_cache_dir
from ..util.cache import enforce_cache_quota, pdf_cache_base_dir, run_cache_dir, json_output_dir
from ..util.cancel import UserCancelled
from .table_cropper import PDFTableCropper
from .validators import validate_table_data, validate_psr_strict, validate_expected_groups
from ..config.org import OrgConfig, load_org_config, validate_org_config
from ..normalize import normalize_table_json
from ..normalize.regions import cleanup_region_info
from ..util.unicode_name import decode_hashu
from ..util.paths import project_root as _project_root


def detect_survey_type(pdf_filename: str) -> str:
    """파일명으로 전국/지방 판별 (레거시 규칙).

    참고: 휴리스틱 폴백입니다. 메타데이터 지역 등 더 확실한 규칙이 있으면 그걸 쓰세요.
    """

    filename = decode_hashu(Path(pdf_filename).stem)

    local_keywords = [
        "지방선거", "지방자치", "지선",
        "서울", "부산", "대구", "인천", "광주", "대전", "울산", "세종",
        "경기", "강원", "충북", "충남", "전북", "전남", "경북", "경남", "제주",
        "경상북도", "경상남도", "충청북도", "충청남도", "전라북도", "전라남도", "강원도", "경기도", "제주도",
        "시장", "도지사", "구청장", "군수",
    ]

    for kw in local_keywords:
        if kw in filename:
            return "지방"

    # 폴백: 앞부분 텍스트에서 지역/지방 신호를 스캔
    try:
        import pdfplumber  # type: ignore

        text_chunks: List[str] = []
        with pdfplumber.open(pdf_filename) as pdf:
            for p in pdf.pages[:3]:
                t = p.extract_text() or ""
                if t:
                    text_chunks.append(t)

                text_guess = _survey_type_from_text("\n".join(text_chunks), local_keywords)
                if text_guess:
                    return text_guess
    except Exception:
        # 텍스트 판별 실패는 무시하고 파일명 규칙으로 폴백
        pass

    if "결과표" in filename:
        return "전국"

    return "전국"


def _survey_type_from_text(text: str, local_keywords: List[str]) -> Optional[str]:
    """초반 텍스트로 조사유형을 추정하는 휴리스틱."""

    compact = re.sub(r"\s+", "", text or "")
    if not compact:
        return None

    # 전국 명시가 있으면 우선(콜론/대시/공백 허용)
    if re.search(r"(조사지역|조사대상)\s*[:：\-]?\s*전국", text):
        return "전국"

    # 지방 신호(선거 유형/직위 등) 우선 탐지
    strong_local_tokens = [
        "지역여론",
        "지역조사",
        "조사지역:",
        "조사지역：",
        "조사대상:",
        "조사대상：",
        "지방선거",
        "광역단체장",
        "기초단체장",
        "광역의원",
        "기초의원",
        "시장",
        "도지사",
        "구청장",
        "군수",
    ]
    # 지방 토큰 중 '조사지역/조사대상: 전국'은 전국으로 처리(콜론 없이도 허용).
    if re.search(r"(조사지역|조사대상)\s*[:：]?\s*전국", text):
        return "전국"
    if any(tok in compact for tok in strong_local_tokens):
        return "지방"

    # 지역명이 여러 번 언급되면 지방으로 추정(전국 표기는 제외).
    local_hits = sum(1 for kw in local_keywords if kw and kw in compact)
    if local_hits >= 2 and "전국" not in compact:
        return "지방"

    return None


def _disable_crop_for_org(org: OrgConfig) -> bool:
    """해당 기관에서는 표 크롭을 건너뛰어야 할 때 True."""
    try:
        blob = f"{org.id} {org.name} {' '.join(org.aliases or [])}".lower()
    except Exception:
        blob = f"{getattr(org, 'id', '')} {getattr(org, 'name', '')}".lower()

    # 리서치뷰는 헤더가 꼭대기에 붙는 경우가 많아 크롭 시 제목이 잘릴 수 있음
    if "리서치뷰" in blob or "researchview" in blob:
        return True
    return False


def _safe_filename(text: str, limit: int = 80) -> str:
    s = re.sub(r"[\\/*?:\"<>|\n\r]", "", text or "")
    s = s.strip()
    if not s:
        s = "item"
    return s[:limit]


def _split_vertical_image(im: Any, max_parts: int = 3, overlap_ratio: float = 0.15) -> List[Any]:
    """세로로 긴 표 이미지를 겹쳐 자르기.

    이유:
    - PSR처럼 촘촘한 표는 너무 길어 단일 이미지로 보내면 모델 리사이즈 후 글자가 작아짐
    - 겹쳐 나누면 헤더를 유지하면서 숫자 인식률이 보통 개선됨
    """

    try:
        w = int(getattr(im, "width", 0))
        h = int(getattr(im, "height", 0))
    except Exception:
        return [im]

    if w <= 0 or h <= 0:
        return [im]

    # 충분히 길 때만 분할
    if h <= w * 1.35:
        return [im]

    n = int(max_parts or 2)
    # 아주 길면 3조각 이상 사용
    if h > w * 2.2:
        n = max(n, 3)
    n = max(2, min(n, 4))

    overlap = int(h * float(overlap_ratio or 0.0))
    step = h / float(n)

    out: List[Any] = []
    for i in range(n):
        y0 = int(i * step) - (overlap if i > 0 else 0)
        y1 = int((i + 1) * step) + (overlap if i < n - 1 else 0)
        y0 = max(0, y0)
        y1 = min(h, y1)
        if y1 - y0 < 80:
            continue
        try:
            out.append(im.crop((0, y0, w, y1)))
        except Exception:
            return [im]

    return out or [im]


def _crop_top_portion(im: Any, ratio: float = 0.45) -> Any:
    """이미지 상단만 잘라 열 헤더에 집중."""
    try:
        if not isinstance(im, Image.Image):
            return im
        w, h = im.width, im.height
        if w <= 0 or h <= 0:
            return im
        cut = max(1, int(h * min(max(ratio, 0.05), 0.95)))
        return im.crop((0, 0, w, cut))
    except Exception:
        return im


def _crop_right_half(im: Any, start_ratio: float = 0.45) -> Any:
    """이미지 오른쪽만 자르되 왼쪽을 조금 겹쳐 둠."""
    try:
        if not isinstance(im, Image.Image):
            return im
        w, h = im.width, im.height
        if w <= 0 or h <= 0:
            return im
        x0 = int(w * min(max(start_ratio, 0.0), 0.95))
        x0 = max(0, min(x0, w - 1))
        return im.crop((x0, 0, w, h))
    except Exception:
        return im


def _extract_table_caption_from_pdf_text(pdf_path: str, page_num: int) -> Optional[str]:
    """가능하면 PDF 텍스트 레이어에서 표 캡션/제목을 뽑아냅니다.

    이유:
    - 페이지 순회 모드에서 비전 모델이 제목의 대괄호 수식어를 누락할 때가 있음
      (예: "서울시장 후보 선호도 [보수진영]" -> "서울시장 후보 선호도")
    - News1/YTN 계열처럼 캡션 라인("[표6] ...")이 텍스트 레이어에 있는 PDF가 많음

    보수적으로 동작:
    - 페이지 텍스트의 앞쪽 10여 줄만 확인
    - 캡션 패턴이 명확할 때만 값을 반환
    """

    if not pdf_path or not page_num or page_num < 1:
        return None

    try:
        import pdfplumber  # type: ignore

        with pdfplumber.open(pdf_path) as pdf:
            if page_num > len(pdf.pages):
                return None
            txt = pdf.pages[int(page_num) - 1].extract_text() or ""
    except Exception:
        return None

    if not txt:
        return None

    lines = [ln.strip() for ln in txt.replace("\x00", "").splitlines() if ln.strip()]
    if not lines:
        return None

    head_lines = lines[:12]
    head = "\n".join(head_lines)

    # 패턴 1: "[표6] 서울시장 후보 선호도 [보수진영]"
    m = re.search(r"\[\s*표\s*[^\]]+\]\s*(.+)", head)
    if m:
        title = m.group(1).strip()

        # 캡션이 줄바꿈된 경우 다음 줄에 "[보수진영]" 같은 괄호 접미어가 올 수 있어 붙여 줌
        if title and ("[" not in title) and ("]" not in title):
            for ln in head_lines[1:6]:
                if re.match(r"^\[[^\]]+\]$", ln) and ("표" not in ln):
                    title = f"{title} {ln}".strip()
                    break

        title = re.sub(r"\s+", " ", title).strip()
        return title or None

    # 패턴 2(폴백): "표6. 서울시장 후보 선호도 ..." (드묾)
    m2 = re.search(r"(?m)^표\s*\d+\s*[\)\.]?\s*(.+)$", head)
    if m2:
        title = re.sub(r"\s+", " ", m2.group(1)).strip()
        return title or None

    return None


def _should_use_psr_columns_values(org: OrgConfig, item_name: str) -> bool:
    """기관이 정당지지도용 특수 PSR 프롬프트/검증을 설정했으면 True."""
    if not org or not isinstance(org.raw, dict):
        return False

    # 후보 지지도/적합도 류는 PSR 엄격 검증 대상 아님 (정당 지지도와 혼동 방지)
    name_compact = re.sub(r"\s+", "", str(item_name or ""))
    if "후보" in name_compact:
        return False

    try:
        sp = org.raw.get("특수프롬프트") or org.raw.get("special_prompts") or {}
        sp_psr = (sp.get("PSR") or sp.get("psr") or {}) if isinstance(sp, dict) else {}
        mode = sp_psr.get("정당지지도") if isinstance(sp_psr, dict) else None
    except Exception:
        mode = None
    if mode != "columns_values_v1":
        return False
    name = item_name or ""
    return ("정당" in name) and ("지지도" in name)


def _issues_missing_items_only(issues: List[str]) -> bool:
    """ISSUE 폴백 판단: 문제 목록이 '응답항목/데이터 누락'뿐이면 True."""
    if not issues:
        return False
    allow = {"missing 응답항목", "missing 데이터"}
    return all(any(tok in msg for tok in allow) for msg in issues)


def _run_psr_strict_validation_if_configured(org: OrgConfig, item_name: str, cand: Dict[str, Any]) -> List[str]:
    """기관 설정이 opt-in인 경우에만 PSR 엄격 검증을 적용."""
    if not _should_use_psr_columns_values(org, item_name):
        return []

    # YAML의 추가 검증 설정 읽기(선택)
    tol = 5.0
    required_cols: List[str] = []
    try:
        sv = org.raw.get("특수검증") or org.raw.get("special_validation") or {}
        sv_psr = (sv.get("PSR") or sv.get("psr") or {}) if isinstance(sv, dict) else {}
        rule = sv_psr.get("정당지지도") if isinstance(sv_psr, dict) else None
        if isinstance(rule, dict):
            if rule.get("행합계허용오차") is not None:
                tol = float(rule.get("행합계허용오차"))
            req = rule.get("필수열")
            if isinstance(req, list):
                required_cols = [str(x) for x in req if str(x).strip()]
    except Exception:
        pass

    return validate_psr_strict(cand, required_cols=required_cols, tol=tol)


def save_json(data: Any, output_path: str) -> None:
    """JSON을 원자적으로 기록(최대한 안전하게).

    이유:
    - 중간에 크래시/취소되면 불완전한 파일이 남을 수 있음
    - 이후 실행이 존재하는 파일을 캐시로 오인할 수 있음

    패턴: *.tmp에 쓰기 -> fsync -> os.replace(tmp, 최종)
    """

    p = Path(output_path)
    try:
        p.parent.mkdir(parents=True, exist_ok=True)
    except Exception:
        pass

    tmp = p.with_suffix(p.suffix + ".tmp")

    try:
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
            try:
                f.flush()
                os.fsync(f.fileno())
            except Exception:
                # 일부 FS에서 fsync 실패 가능: 무시
                pass
        os.replace(tmp, p)
    finally:
        # 중간 실패 시 임시 파일 정리 시도
        try:
            if tmp.exists():
                tmp.unlink()
        except Exception:
            pass




def _is_kri_org(org: OrgConfig) -> bool:
    """휴리스틱: 코리아리서치인터내셔널(KRI) 여부."""
    try:
        blob = f"{org.id} {org.name} {' '.join(org.aliases or [])}"
    except Exception:
        blob = f"{getattr(org, 'id', '')} {getattr(org, 'name', '')}"
    blob_l = blob.lower()
    return ("코리아리서치" in blob) or ("kri" in blob_l)




def _is_kri_style_org(org: OrgConfig) -> bool:
    """휴리스틱: KRI 스타일 통계표 계열.

    - 코리아리서치인터내셔널(KRI)
    - 한국리서치(Hankook Research)

    이 계열은 '조사 개요/설계' 페이지의 텍스트가 비교적 구조적이라
    pdfplumber 기반 폴백 파서가 잘 동작하는 편입니다.
    """
    try:
        blob = f"{org.id} {org.name} {' '.join(org.aliases or [])}"
    except Exception:
        blob = f"{getattr(org, 'id', '')} {getattr(org, 'name', '')}"
    blob_l = blob.lower()

    return (
        _is_kri_org(org)
        or ("한국리서치" in blob)
        or ("hankook" in blob_l)
    )


def _parse_metadata_from_text_kri(text: str) -> Dict[str, Any]:
    """KRI 계열 '조사 설계' 페이지의 pdfplumber 텍스트에서 메타데이터 파싱.

    목표: 결정적이고 보수적으로 동작(질문 리스트를 과하게 잡지 않음).
    """
    if not text:
        return {}

    t = text.replace("\x00", "")

    def _line_after(label: str) -> str:
        m = re.search(rf"{re.escape(label)}\s+(.+)", t)
        return m.group(1).strip() if m else ""

    def _block_after(label: str, stop_labels: List[str]) -> str:
        # 다음 구분 라벨이 줄 맨 앞에 나오기 전까지 캡처
        stop_re = "|".join(re.escape(x) for x in stop_labels)
        m = re.search(
            rf"{re.escape(label)}\s*(.*?)(?=\n(?:{stop_re})\b)",
            t,
            flags=re.S,
        )
        return (m.group(1).strip() if m else "")

    out: Dict[str, Any] = {}

    # 표본크기
    s = _line_after("표본 크기") or _line_after("표본크기")
    if s:
        m = re.search(r"([0-9][0-9,]*)\s*명", s)
        if m:
            try:
                out["표본크기"] = int(m.group(1).replace(",", ""))
            except Exception:
                out["표본크기"] = s
        else:
            out["표본크기"] = s

    # 표본오차 (단일 줄)
    s = _line_after("표본오차")
    if s:
        out["표본오차"] = re.sub(r"\s+", " ", s).strip()

    # 조사기관 (단일 줄)
    s = _line_after("조사기관")
    if s:
        out["조사기관"] = re.sub(r"\s+", " ", s).strip()

    # 조사기간 (인라인 또는 멀티라인)
    s = _line_after("조사기간")
    if s:
        out["조사기간"] = re.sub(r"\s+", " ", s).strip()
    else:
        blk = _block_after(
            "조사기간",
            stop_labels=["조사대상", "조사방법", "표본 크기", "표본크기", "표본오차", "응답률", "접촉률", "피조사자"],
        )
        if blk:
            parts = []
            for ln in blk.splitlines():
                ln = ln.strip().lstrip("-").strip()
                if ln:
                    parts.append(ln)
            out["조사기간"] = " / ".join(parts) if parts else ""

    # 조사방법 (멀티라인, '표본 크기' 전에 끝나는 게 일반적)
    blk = _block_after(
        "조사방법",
        stop_labels=["표본 크기", "표본크기", "표본오차", "응답률", "접촉률", "피조사자", "가중치값", "질문 내용"],
    )
    # 대부분 PDF는 '조사방법 · ...'처럼 한 줄로 표기
    method_prefix = _line_after("조사방법")
    m = re.search(r"(?m)^(.{0,120}(?:ARS|CATI|면접).{0,120})\n조사방법\b", t)
    if m:
        method_prefix = m.group(1).strip()
    if blk or method_prefix:
        parts = []
        if method_prefix:
            parts.append(method_prefix)
        if blk:
            for ln in blk.splitlines():
                ln = ln.strip().lstrip("-").strip()
                if ln:
                    parts.append(ln)
        out["조사방법"] = re.sub(r"\s+", " ", " ".join(parts)).strip()

    # 응답률
    s = _line_after("응답률")
    if s:
        m = re.search(r"([0-9]+(?:\.[0-9]+)?)\s*%", s)
        if m:
            out["응답률"] = f"{m.group(1)}%"
        else:
            out["응답률"] = re.sub(r"\s+", " ", s).strip()

    return {k: v for k, v in out.items() if v not in (None, "", [])}


def _fill_metadata_with_text_fallback(
    metadata: Dict[str, Any], pdf_path: str, meta_page: int, org: OrgConfig
) -> Dict[str, Any]:
    """메타데이터가 비었을 때 pdfplumber 텍스트로 보충(KRI 폴백).

    보수적 동작:
    - KRI 계열로 보이는 기관에서만 실행
    - LLM 메타데이터가 대부분 비어 있을 때만 실행
    """

    if not pdf_path or meta_page is None or not org:
        return metadata

    if not _is_kri_style_org(org):
        return metadata

    fields = ["표본크기", "표본오차", "조사방법", "조사기간", "조사기관", "응답률"]
    empty = sum(1 for f in fields if not (metadata or {}).get(f))
    # 하나라도 비어 있으면 결정적 텍스트 파서로 채워보기
    if empty <= 0:
        return metadata

    try:
        import pdfplumber  # type: ignore

        with pdfplumber.open(pdf_path) as pdf:
            if meta_page < 1 or meta_page > len(pdf.pages):
                return metadata
            txt = pdf.pages[int(meta_page) - 1].extract_text() or ""
    except Exception:
        return metadata

    parsed = _parse_metadata_from_text_kri(txt)
    if not parsed:
        return metadata

    out = dict(metadata or {})
    for k, v in parsed.items():
        if not out.get(k):
            out[k] = v
    return out


def _normalize_metadata_fields(metadata: Dict[str, Any]) -> Dict[str, Any]:
    """메타데이터 주요 필드를 정규화(응답률 등)."""
    if not isinstance(metadata, dict):
        return metadata

    out = dict(metadata)

    # 응답률: 첫 퍼센트 숫자만 보존(예: "11.8% (총 ...)" -> "11.8%")
    val = out.get("응답률")
    if isinstance(val, str):
        m = re.search(r"([0-9]+(?:\.[0-9]+)?)\s*%", val)
        if m:
            out["응답률"] = f"{m.group(1)}%"

    return out



def _parse_region_info_from_text_kri_style(text: str) -> Dict[str, Any]:
    """pdfplumber 텍스트에서 '권역 구분' 매핑을 파싱.

    예상 패턴(KBS-한국리서치 자료에서 흔함):
      - '권 역 권역 구분' 이후에
        '<권역명> <구/군/시 ...>' 형태의 라인이 이어짐.

    반환:
      {'권역': [{'name': '...', 'areas': '...'}, ...]}
    """
    if not text:
        return {}

    lines = [ln.strip() for ln in (text or "").replace("\x00", "").splitlines()]

    start_idx = None
    for i, ln in enumerate(lines):
        # '권 역'처럼 띄어쓰기가 섞인 경우에도 대응
        compact = re.sub(r"\s+", "", ln)
        if "권역구분" in compact:
            start_idx = i + 1
            break

    if start_idx is None:
        return {}

    regions: List[Dict[str, str]] = []
    for ln in lines[start_idx:]:
        if not ln:
            continue
        # 권역 매핑 뒤에 '※'로 시작하는 주석 블록이 이어지기도 함
        if ln.startswith("※") and regions:
            break
        if ln.startswith("[") and regions:
            break

        m = re.match(r"^(\S+)\s+(.+)$", ln)
        if not m:
            continue
        name = m.group(1).strip()
        areas = m.group(2).strip()
        if name and areas:
            regions.append({"name": name, "areas": areas})

    return {"권역": regions} if regions else {}


def _fill_region_with_text_fallback(
    region_info: Any, pdf_path: str, region_page: Any, org: OrgConfig
) -> Any:
    """권역 정보가 없을 때 pdfplumber 텍스트로 보충.

    보수적 동작:
    - KRI 계열로 보이는 기관에서만 실행
    - region_info가 비어 있을 때만 실행
    """

    # region_page는 기관/자료에 따라 없을 수 있음.
    # 일부 YAML은 권역설명 전용 페이지를 의도적으로 생략(표 행 라벨 우선 정책).
    # 이런 경우 region_page는 None일 수 있으니 크래시 금지.
    if not pdf_path or region_page is None or not org:
        return region_info

    # str/int 모두 수용
    try:
        region_page_i = int(region_page)
    except Exception:
        return region_info

    if not _is_kri_style_org(org):
        return region_info

    # 권역 항목이 1개 이상 있으면 그대로 유지
    try:
        if isinstance(region_info, dict) and isinstance(region_info.get("권역"), list) and region_info.get("권역"):
            return region_info
        if isinstance(region_info, list) and region_info:
            return region_info
    except Exception:
        pass

    try:
        import pdfplumber  # type: ignore

        with pdfplumber.open(pdf_path) as pdf:
            if region_page_i < 1 or region_page_i > len(pdf.pages):
                return region_info
            txt = pdf.pages[region_page_i - 1].extract_text() or ""
    except Exception:
        return region_info

    parsed = _parse_region_info_from_text_kri_style(txt)
    return parsed if parsed else region_info


def _compact_text_for_match(text: str) -> str:
    """키워드 매칭 강화를 위해 텍스트를 압축.

    pdfplumber 결과에 ￾ 같은 이상한 구분 기호가 섞일 수 있어
    숫자/알파벳/한글만 남겨 "조사￾ 설계"도 "조사설계"로 매칭되게 함.
    """
    return re.sub(r"[^0-9A-Za-z가-힣]+", "", text or "")


def _score_metadata_page_text(text: str) -> int:
    """'조사 설계/조사 개요' 페이지일 가능성을 점수화."""
    if not text:
        return 0

    c = _compact_text_for_match(text)
    score = 0

    # 강한 신호
    if "조사설계" in c:
        score += 5
    if "조사개요" in c:
        score += 4

    # 공통 메타데이터 필드
    for kw in [
        "조사의뢰자",
        "조사기관",
        "조사지역",
        "조사기간",
        "조사대상",
        "조사방법",
        "표본크기",
        "표본오차",
        "응답률",
        "접촉률",
        "가중치",
    ]:
        if kw in c:
            score += 1

    return score


def _resolve_metadata_page(
    pdf_path: str,
    yaml_meta_page: Optional[int],
    toc_meta_page: Optional[int],
    org: OrgConfig,
) -> Optional[int]:
    """가장 신뢰할 메타데이터 페이지 선택.

    이유:
    - 어떤 목차는 "조사 설계"가 아예 없어 모델이 페이지를 추정할 수 있음.
    - toc_meta_page를 그대로 믿으면 엉뚱한 페이지에서 추출해 빈 메타데이터.json을 캐시할 수 있음.

    전략:
    - PDF 텍스트를 읽을 수 있으면 앞쪽 페이지만 훑어 메타데이터 키워드 신호가 가장 강한 페이지 선택
    - 불가하면 YAML(최우선) -> TOC 순으로 폴백

    메모: 의도적으로 보수적이고 가벼운 접근.
    """

    _ = org  # org는 시그니처 호환성을 위해 유지

    if yaml_meta_page is None and toc_meta_page is None:
        return None

    # PDF 경로가 없으면 YAML을 우선 신뢰
    if not pdf_path:
        return yaml_meta_page if yaml_meta_page is not None else toc_meta_page

    try:
        import pdfplumber  # type: ignore

        with pdfplumber.open(pdf_path) as pdf:
            n = len(pdf.pages)
            if n <= 0:
                return yaml_meta_page if yaml_meta_page is not None else toc_meta_page

            # 메타데이터는 보통 앞부분에 있으므로 초반 페이지만 스캔
            cands = [p for p in [yaml_meta_page, toc_meta_page] if isinstance(p, int) and p > 0]
            max_cand = max(cands) if cands else 10
            scan_upto = min(n, max(10, min(max_cand + 2, 30)))

            best_page: Optional[int] = None
            best_score = -1

            for p in range(1, scan_upto + 1):
                try:
                    txt = pdf.pages[p - 1].extract_text() or ""
                except Exception:
                    txt = ""
                sc = _score_metadata_page_text(txt)
                if sc > best_score:
                    best_score = sc
                    best_page = p

            # 덮어쓰려면 최소 신뢰도 요구
            if best_page is not None and best_score >= 6:
                return best_page

    except Exception:
        # pdfplumber 실패(또는 텍스트 레이어 없음) 시 폴백
        pass

    return yaml_meta_page if yaml_meta_page is not None else toc_meta_page

def _guess_item_type_from_name(name: str) -> str:
    """TOC 항목명으로 항목 유형을 추정하는 휴리스틱.

    보수적으로:
    - PSR: '정당( )지지도' 계열만
    - GE: '국정'+(수행/운영/평가) 또는 '대통령'+(직무/수행/운영/국정)
    - 그 외: ISSUE
    """
    s = str(name or "")
    s_compact = re.sub(r"\s+", "", s)

    # 후보 지지도/적합도 류는 ISSUE로 분류 (PSR/GE 템플릿 적용 방지)
    if "후보" in s:
        return "ISSUE"

    # 이유형 문항(예: "국정운영 긍정 평가 이유")은 GE/PSR 분포표가 아니므로 ISSUE로 처리해 응답 항목을 자유롭게 둠
    if "이유" in s:
        return "ISSUE"

    # 미래 전망/향후 인식 류도 GE/PSR 분포표가 아니므로 ISSUE 처리
    if ("향후" in s) or ("전망" in s):
        return "ISSUE"

    # "국정운영 성과 분야" 류는 직무평가(잘함/못함) 분포표가 아니라
    # '어떤 분야에서 성과가 크냐' 같은 항목 선택(정책/분야 리스트) 문항이므로 ISSUE.
    # - '성과' 단독은 다른 문장(예: 성과 평가)에서도 등장할 수 있어 너무 광범위하므로
    #   '성과' + '분야' 조합(또는 '성과분야' 축약)만 예외 처리한다.
    if ("성과" in s and "분야" in s) or ("성과분야" in s_compact):
        return "ISSUE"

    # PSR: 정당지지도(공백 허용)만
    if "정당지지도" in s_compact or re.search(r"정당\s*지지도", s):
        return "PSR"

    # GE: 국정운영/직무수행 계열만 (대통령+평가 같은 일반 '평가'는 제외)
    ge_by_gukjeong = ("국정" in s) and any(k in s for k in ["수행", "운영", "평가"])
    ge_by_president = ("대통령" in s) and any(k in s for k in ["직무", "수행", "운영", "국정"])

    if ge_by_gukjeong or ge_by_president:
        return "GE"

    return "ISSUE"


def _postprocess_auto_type(item_type: str, question: str | None, response_items: list | None) -> str:
    """자주 틀리는 분류를 줄이기 위한 일반 후처리.

    - 대통령/국정운영 직무평가가 ISSUE로 떨어지는 경우 -> GE
    - 지방 단체장 직무평가가 GE로 잘못 올라가는 경우 -> ISSUE

    기관 특화 규칙이 아닌 경험칙 기반.
    """

    itype = (item_type or "ISSUE").upper()
    q = re.sub(r"\s+", "", str(question or ""))
    resp = response_items or []
    resp_norm = [re.sub(r"\s+", "", str(x)) for x in resp]

    # 미래 전망/향후 인식 류는 GE 분포표가 아니므로 ISSUE로 강등
    if ("향후" in q) or ("전망" in q):
        return "ISSUE"

    has_ge_resp = (
        any(x in resp_norm for x in ["잘함", "잘하고있다", "잘하고있음"]) and
        any(x in resp_norm for x in ["잘못함", "잘못하고있다", "잘못하고있음"])
    )

    local_office_keywords = ["구청장", "군수", "도지사", "교육감", "청장", "시장"]
    is_local_office = any(k in q for k in local_office_keywords)

    # "대통령"이 들어가도 GE가 아닌 경우(전 대통령/해외 대통령/일반 평가형 이슈)를 구분
    is_ex_president = ("전대통령" in q) or ("전직대통령" in q) or (("전직" in q) and ("대통령" in q))
    is_foreign_president = any(k in q for k in ["트럼프", "바이든", "푸틴", "시진핑", "김정은"]) or ("미국대통령" in q)

    if (
        itype != "GE" and (not is_ex_president) and (not is_foreign_president) and (
            (("대통령" in q) and any(k in q for k in ["직무평가", "직무수행", "국정운영", "국정수행"])) or
            (has_ge_resp and any(k in q for k in ["직무", "국정", "운영"]) and not is_local_office)
        )
    ):
        itype = "GE"

    if (
        itype == "GE" and is_local_office and ("대통령" not in q) and ("국정" not in q)
    ):
        itype = "ISSUE"

    # 형태는 GE 같지만 실제 국정/대통령 직무평가가 아니면 ISSUE로 내림
    if itype == "GE":
        is_true_ge = (
            (("국정" in q) and any(k in q for k in ["운영", "수행", "평가"])) or
            (("대통령" in q) and any(k in q for k in ["직무평가", "직무수행", "국정운영", "국정수행", "직무"]))
        )
        if (is_ex_president or is_foreign_president) or (not is_true_ge):
            itype = "ISSUE"

        # "국정운영 성과 분야" 같은 '분야 선택' 문항은 GE가 아니라 ISSUE로 보는 것이 안전.
        # - GE 프롬프트(3열 고정)를 쓰면 응답항목(분야/정책 리스트)이 유실될 수 있음.
        if ("성과" in q and "분야" in q) or ("성과분야" in q):
            # 질문 자체가 '성과 분야'(분야 선택형)인 경우는 GE가 아니라 ISSUE로 보는 것이 안전.
            # (이미 GE 프롬프트로 추출한 경우 응답항목이 '잘하고/못하고'로 오염될 수 있어서
            #  resp 패턴을 신뢰하지 않는다.)
            itype = "ISSUE"


    # "국정운영/직무수행 평가 이유" 류는 GE(직무평가) 분포표가 아니라 이유(다항목) 문항이므로 ISSUE로 처리
    if itype == "GE" and ("이유" in q):
        itype = "ISSUE"

    return itype


# -----------------------------------------------------------------------------
# 후처리 헬퍼
# -----------------------------------------------------------------------------
_CANDIDATE_ITEM_HINTS = [
    "적합도", "지지도", "선호도", "대결", "가상대결", "후보", "단일화", "양자대결", "다자대결",
]

def _is_candidate_issue_item(obj: Dict[str, Any]) -> bool:
    """휴리스틱: 후보 적합도/지지도 등 후보 관련 ISSUE 테이블 여부."""
    try:
        name = str(obj.get("항목명") or "")
        q = str(obj.get("질문") or "")
        blob = re.sub(r"\s+", "", f"{name} {q}")
        return any(k in blob for k in _CANDIDATE_ITEM_HINTS)
    except Exception:
        return False


def _enrich_issue_candidate_titles_across_items(result_dir: Path) -> None:
    """같은 PDF 실행 내 ISSUE 항목들 사이에서 축약된 후보 헤더를 보강.

    배경:
    - 후보 테이블 헤더가 여러 줄(이름+정당/직책)일 때 비전 모델이 어떤 표에서는 이름만 내고
      같은 PDF의 다른 표에서는 완전한 제목을 내보낼 수 있음.
    - 같은 실행의 다른 ISSUE 항목들로부터 (기본 이름 -> 가장 풍부한 라벨) 맵을 만들고
      축약된 항목을 보완할 수 있음.

    보수적인 최선 시도로 동작.
    """

    issue_paths = sorted([p for p in result_dir.glob("p*_ISSUE_*.json") if p.is_file()])
    if len(issue_paths) < 2:
        return

    issue_items: List[tuple[Path, Dict[str, Any]]] = []
    for p in issue_paths:
        try:
            obj = json.loads(p.read_text(encoding="utf-8"))
        except Exception:
            continue
        if not isinstance(obj, dict):
            continue
        if not _is_candidate_issue_item(obj):
            continue
        issue_items.append((p, obj))

    if len(issue_items) < 2:
        return

    # 모든 후보 관련 ISSUE 항목에서 가장 풍부한 라벨 맵 구축
    rich: Dict[str, str] = {}

    def _is_non_candidate_choice(s: str) -> bool:
        s2 = re.sub(r"\s+", "", s)
        return s2 in {
            "적합한사람이없다",
            "모름.무응답",
            "모름/무응답",
            "잘모름",
            "없음",
        }

    for _, obj in issue_items:
        resp = obj.get("응답항목") or []
        if not isinstance(resp, list):
            continue
        for label in resp:
            if not isinstance(label, str):
                continue
            s = label.strip()
            if not s:
                continue

            # 명확한 띄어쓰기 변형 보정
            if s == "다른사람":
                s = "다른 사람"

            if _is_non_candidate_choice(s):
                continue

            base = s.split()[0]
            if not base:
                continue

            # 기본 이름 외에 설명이 붙어 있어야 "풍부한" 라벨로 인정
            # (공백이 있고 기본 토큰보다 충분히 길 때)
            if (" " not in s) or (len(s) <= len(base) + 1):
                continue

            cur = rich.get(base)
            if cur is None or len(s) > len(cur):
                rich[base] = s

    if not rich:
        return

    # 축약된 라벨 보정
    for path, obj in issue_items:
        resp = obj.get("응답항목") or []
        data = obj.get("데이터") or {}
        if not isinstance(resp, list) or not isinstance(data, dict):
            continue

        repl: Dict[str, str] = {}

        for label in resp:
            if not isinstance(label, str):
                continue
            s = label.strip()
            if not s:
                continue

            # 띄어쓰기 보정
            if s == "다른사람":
                repl[label] = "다른 사람"
                continue

            # 기본 이름만 있을 때 다른 표에서 확보한 풍부한 라벨로 교체
            if (" " not in s) and (s in rich) and (len(rich[s]) > len(s)):
                repl[label] = rich[s]
                continue

            # 흔한 붙여쓰기 오류: "조국현" == "조국 현 ..."
            if (" " not in s) and s.endswith("현") and (s[:-1] in rich):
                repl[label] = rich[s[:-1]]
                continue

        if not repl:
            continue

        obj["응답항목"] = [repl.get(x, x) if isinstance(x, str) else x for x in resp]

        new_data: Dict[str, Any] = {}
        for row_k, row_v in data.items():
            if not isinstance(row_v, dict):
                new_data[row_k] = row_v
                continue
            new_row: Dict[str, Any] = {}
            for col_k, col_v in row_v.items():
                if isinstance(col_k, str) and col_k in repl:
                    new_row[repl[col_k]] = col_v
                else:
                    new_row[col_k] = col_v
            new_data[row_k] = new_row
        obj["데이터"] = new_data

        # 변경 내용을 저장
        try:
            save_json(obj, str(path))
        except Exception:
            pass





def _cleanup_duplicate_m_outputs(result_dir: Path) -> None:
    """결과 폴더 안에 중복된 '*_m.json' 산출물을 정리.

    일부 레거시 파이프라인/외부 후처리가 동일 항목 JSON을 '_m' 접미사로 하나 더 만들기도 함.
    이는 항목이 중복되어 이후 단계(예: Excel 작성기)를 혼란시킬 수 있음.

    안전 정책:
    - 기본 파일과 *_m 둘 다 있고 내용이 같으면 *_m 삭제
    - 기본 파일이 없으면 *_m을 기본 이름으로 변경
    - 그 외에는 둘 다 유지(데이터 보존 우선)
    """
    for m_path in sorted(result_dir.glob("p*_*_m.json")):
        if not m_path.is_file():
            continue
        base_name = m_path.name.replace("_m.json", ".json")
        base_path = m_path.with_name(base_name)

        try:
            m_bytes = m_path.read_bytes()
        except Exception:
            continue

        if base_path.exists():
            try:
                if base_path.read_bytes() == m_bytes:
                    m_path.unlink()
            except Exception:
                pass
        else:
            try:
                m_path.rename(base_path)
            except Exception:
                pass


def _norm_label_key(s: str) -> str:
    """구분자/공백 변형까지 정리해 라벨 매칭을 견고하게 정규화."""
    return re.sub(r"[^0-9A-Za-z가-힣]+", "", str(s or ""))


_BUDONG_NORM = _norm_label_key("부동층")

# '부동층'을 구성하는 잔여/비선택 카테고리 모음
_BUDONG_COMPONENT_NORMS = {
    _norm_label_key("없음"),
    _norm_label_key("없다"),
    _norm_label_key("지지정당없음"),
    _norm_label_key("지지정당없다"),
    _norm_label_key("적합한사람이없다"),
    _norm_label_key("적합한사람없다"),
    _norm_label_key("적합한사람이없음"),
    _norm_label_key("적합한사람없음"),
    _norm_label_key("잘모름"),
    _norm_label_key("모름"),
    _norm_label_key("무응답"),
    _norm_label_key("모름무응답"),
    _norm_label_key("모름/무응답"),
    _norm_label_key("모름.무응답"),
    _norm_label_key("잘모름무응답"),
    _norm_label_key("잘모름/무응답"),
    _norm_label_key("잘모름.무응답"),
}


def _to_float(v: Any) -> Optional[float]:
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, str):
        s = v.strip()
        if not s:
            return None
        try:
            return float(s.replace(",", ""))
        except Exception:
            return None
    return None


def _pick_total_row_key(data: Dict[str, Any]) -> Optional[str]:
    if not isinstance(data, dict) or not data:
        return None
    # 가능하면 '전체'와 정확히 일치하는 행 우선
    for k in data.keys():
        if isinstance(k, str) and k.strip() == "전체":
            return k
    # 없으면 '전체'로 시작하는 행
    for k in data.keys():
        if isinstance(k, str) and k.strip().startswith("전체"):
            return k
    # 최종 폴백: 첫 번째 키
    for k in data.keys():
        if isinstance(k, str):
            return k
    return None


def _remove_derived_budongcheung_in_obj(obj: Dict[str, Any]) -> bool:
    """명백히 파생된 '부동층'(예: 없음+잘모름 합산)일 때 제거.

    보수적으로 동작:
    - 응답항목에 '부동층'이 있을 때만
    - 전체 행에서 부동층 값이 구성 요소 합과 거의 같고,
      이를 제거하면 합계가 100에 가까워질 때만
    """
    if not isinstance(obj, dict):
        return False

    resp = obj.get("응답항목")
    data = obj.get("데이터")

    if not isinstance(resp, list) or not isinstance(data, dict) or not resp or not data:
        return False

    # 응답항목에서 실제 사용된 '부동층' 라벨 찾기
    budong_label = None
    for c in resp:
        if isinstance(c, str) and _norm_label_key(c) == _BUDONG_NORM:
            budong_label = c
            break
    if not budong_label:
        return False

    total_key = _pick_total_row_key(data)
    if not total_key:
        return False

    row = data.get(total_key)
    if not isinstance(row, dict):
        return False

    budong_val = _to_float(row.get(budong_label))
    if budong_val is None or budong_val <= 0:
        return False

    # 같은 행에서 구성 요소 합산
    comp_sum = 0.0
    comp_used = 0
    for k, v in row.items():
        if not isinstance(k, str):
            continue
        if _norm_label_key(k) in _BUDONG_COMPONENT_NORMS:
            fv = _to_float(v)
            if fv is not None:
                comp_sum += fv
                comp_used += 1

    # 최소 1개(보통 없음+잘모름 두 개) 구성 요소 필요
    if comp_used <= 0:
        return False

    # 부동층 값이 구성 요소 합과 거의 같은지 확인
    if abs(budong_val - comp_sum) > 1.0:
        return False

    # 전체 행 합계(before/after)를 응답항목 리스트 기준으로 계산
    before = 0.0
    for c in resp:
        if not isinstance(c, str):
            continue
        fv = _to_float(row.get(c))
        if fv is not None:
            before += fv

    after = before - budong_val

    # 합계가 명확히 100을 넘고 제거가 유의미하게 고치는 경우에만 실행
    if before < 110.0:
        return False
    if abs(after - 100.0) > 5.0:
        return False
    if abs(after - 100.0) >= abs(before - 100.0):
        return False

    # 실제 제거 적용: 응답항목/모든 행에서 삭제
    obj["응답항목"] = [c for c in resp if not (isinstance(c, str) and _norm_label_key(c) == _BUDONG_NORM)]

    for rk, rv in list(data.items()):
        if isinstance(rv, dict) and budong_label in rv:
            try:
                rv.pop(budong_label, None)
            except Exception:
                pass

    obj["데이터"] = data
    return True


def _drop_derived_budongcheung_outputs(result_dir: Path) -> None:
    """저장된 항목 JSON에서 파생 '부동층' 열을 최선 시도로 제거."""
    for path in sorted(result_dir.glob("p*_*.json")):
        if not path.is_file():
            continue
        if path.name in ("메타데이터.json", "권역설명.json"):
            continue
        if path.name.endswith("_m.json"):
            # *_m 정리 단계가 먼저 처리하도록 건너뜀
            continue

        try:
            obj = json.loads(path.read_text(encoding="utf-8"))
        except Exception:
            continue
        if not isinstance(obj, dict):
            continue

        changed = False
        try:
            changed = _remove_derived_budongcheung_in_obj(obj)
        except Exception:
            changed = False

        if changed:
            try:
                save_json(obj, str(path))
            except Exception:
                pass



def _sim_ratio(a: str, b: str) -> float:
    """라벨 유사도 간단 계산(0~1)."""
    try:
        from difflib import SequenceMatcher
        return SequenceMatcher(None, a or "", b or "").ratio()
    except Exception:
        return 0.0


def _remove_near_duplicate_response_items_in_obj(obj: Dict[str, Any]) -> bool:
    """명백한 중복 응답항목을 병합/제거(보수적).

    페이지 순회 출력에서 발견된 실패 패턴:
    - 아주 작은 OCR/LLM 오타로 동일 선택지가 두 번 등장
      (예: "편의이" vs "편익이" 등 한 글자 차이)
    - 한 변형이 일부 교차표(또는 전체 행)에만 나타남
    - 둘 다 합산되어 총합이 100을 초과

    보수적 조건:
    - 라벨 유사도가 충분히 높아야 함
    - 전체 행에서 두 값이 존재하고 거의 동일해야 함
    - 어느 행에서든 값이 충돌하면 제외
    - 제거 후 전체 행 합계가 100에 더 가까워질 때만 적용
    """
    if not isinstance(obj, dict):
        return False

    resp = obj.get("응답항목")
    data = obj.get("데이터")
    if not isinstance(resp, list) or not isinstance(data, dict) or len(resp) < 2:
        return False

    total_key = _pick_total_row_key(data)
    if not total_key:
        return False

    total_row = data.get(total_key)
    if not isinstance(total_row, dict):
        return False

    def _sum_row(row: Dict[str, Any], cols: List[str]) -> float:
        s = 0.0
        for c in cols:
            fv = _to_float(row.get(c))
            if fv is not None:
                s += fv
        return s

    resp_now = [c for c in resp if isinstance(c, str)]
    before = _sum_row(total_row, resp_now)
    if before < 110.0:
        return False

    changed = False

    # 한 번에 하나씩 중복을 제거하면서 반복
    while True:
        resp_now = [c for c in obj.get("응답항목", []) if isinstance(c, str)]
        total_row = data.get(total_key) if isinstance(data.get(total_key), dict) else {}
        if not isinstance(total_row, dict):
            break

        before = _sum_row(total_row, resp_now)
        if before < 110.0:
            break

        found = False
        for i, a in enumerate(resp_now):
            for b in resp_now[i + 1:]:

                na = _norm_label_key(a)
                nb = _norm_label_key(b)
                if not na or not nb:
                    continue

                if _sim_ratio(na, nb) < 0.88:
                    continue

                ta = _to_float(total_row.get(a))
                tb = _to_float(total_row.get(b))
                if ta is None or tb is None:
                    continue
                if abs(ta - tb) > 0.1:
                    continue

                # 어느 행에서든 두 값이 충돌하면 제외
                conflict = False
                for _, rv in data.items():
                    if not isinstance(rv, dict):
                        continue
                    xa = _to_float(rv.get(a))
                    xb = _to_float(rv.get(b))
                    if xa is not None and xb is not None and abs(xa - xb) > 0.1:
                        conflict = True
                        break
                if conflict:
                    continue

                # 더 많은 행에 등장한 쪽을 기준 라벨로 선택
                count_a = 0
                count_b = 0
                for _, rv in data.items():
                    if not isinstance(rv, dict):
                        continue
                    if _to_float(rv.get(a)) is not None:
                        count_a += 1
                    if _to_float(rv.get(b)) is not None:
                        count_b += 1

                keep = a
                drop = b
                if count_b > count_a:
                    keep, drop = b, a

                drop_val = _to_float(total_row.get(drop)) or 0.0
                after = before - drop_val

                # 적용 시 전체 행 합계를 100 쪽으로 확실히 개선해야 함
                if abs(after - 100.0) > 5.0:
                    continue
                if abs(after - 100.0) >= abs(before - 100.0):
                    continue

                # 값 병합: keep이 비어 있고 drop만 있으면 복사해 채움
                for _, rv in data.items():
                    if not isinstance(rv, dict):
                        continue
                    v_keep = _to_float(rv.get(keep))
                    v_drop = _to_float(rv.get(drop))
                    if v_keep is None and v_drop is not None:
                        rv[keep] = v_drop
                    rv.pop(drop, None)

                obj["응답항목"] = [c for c in resp_now if c != drop]
                obj["데이터"] = data

                changed = True
                found = True
                break
            if found:
                break

        if not found:
            break

    return changed

def _dedupe_near_duplicate_response_items_outputs(result_dir: Path) -> None:
    """저장된 항목 JSON의 유사 중복 응답항목을 최선 시도로 정리."""
    for path in sorted(result_dir.glob("p*_*.json")):
        if not path.is_file():
            continue
        if path.name in ("메타데이터.json", "권역설명.json"):
            continue
        if path.name.endswith("_m.json"):
            continue

        try:
            obj = json.loads(path.read_text(encoding="utf-8"))
        except Exception:
            continue
        if not isinstance(obj, dict):
            continue

        changed = False
        try:
            changed = _remove_near_duplicate_response_items_in_obj(obj)
        except Exception:
            changed = False

        if changed:
            try:
                save_json(obj, str(path))
            except Exception:
                pass

def extract_from_pdf(
    pdf_path: str,
    org_name: str,
    api_key: str,
    model: str,
    poppler_path: str | None,
    dpi: int = 300,
    output_dir: str = "output",
    extract_types: list | None = None,
    resume: bool = True,
    page_whitelist: list[int] | None = None,
    skip_meta_region: bool | None = None,
    progress_callback: Callable[[Dict[str, Any]], None] | None = None,
    should_cancel: Callable[[], bool] | None = None,
) -> dict:
    """여론조사 PDF에서 표를 추출해 JSON 파일로 저장."""

    from .. import prompts as prompt_builder
    from ..util.cancel import UserCancelled

    def _progress(stage: str, *, current: int | None = None, total: int | None = None, message: str | None = None):
        """최선 시도의 진행률 보고.

        progress_callback은 절대 예외를 던져 파이프라인을 깨면 안 됨.
        """
        if progress_callback is None:
            return
        try:
            progress_callback({
                "stage": stage,
                "current": current,
                "total": total,
                "message": message,
            })
        except Exception:
            pass

    def _cancel_check() -> None:
        """취소 요청이 있으면 UserCancelled 발생."""
        try:
            if should_cancel is not None and should_cancel():
                raise UserCancelled("cancelled")
        except UserCancelled:
            raise
        except Exception:
            # should_cancel 콜백이 추출을 중단시키면 안 됨
            return

    if extract_types is None:
        extract_types = ["PSR", "GE", "ISSUE"]

    # 설정 로드 및 검증
    org = load_org_config(org_name, project_root=_project_root())
    problems = validate_org_config(org)
    if problems:
        raise ValueError("\n".join(problems))

    survey_type = detect_survey_type(pdf_path)
    side = org.national if survey_type == "전국" else org.local

    # ------------------------------------------------------------------
    # 최종 결과 vs 중간 산출물
    # ------------------------------------------------------------------
    # 사용자 지정 출력 폴더에는 최종 Excel만 있어야 하므로
    # 모든 중간 JSON(메타데이터/권역/항목)은 OS 캐시 디렉터리에 저장
    pdf_name = decode_hashu(Path(pdf_path).stem)
    result_dir = json_output_dir(pdf_path)

    # ------------------------------------------------------------------
    # 캐시 정책
    # ------------------------------------------------------------------
    # 중요: 캐시는 사용자 출력 폴더에 쓰지 않으며 OS 캐시에만 둔다.
    # 캐시 베이스는 "<pdf_name>..." 구조(사람이 읽기 좋게)와 충돌 방지용 짧은 해시를 함께 유지
    cache_base = pdf_cache_base_dir(pdf_path)
    # 캐시 용량 초과 시 오래된 PDF 캐시부터 삭제, 현재 작업 캐시는 제외
    try:
        enforce_cache_quota(exclude=[cache_base])
    except Exception:
        pass

    renderer = PDFPageRenderer(
        pdf_path=pdf_path,
        poppler_path=poppler_path,
        cache_dir=default_cache_dir(pdf_path, dpi),
        options=RenderOptions(dpi=dpi),
    )

    table_cropper = PDFTableCropper(
        pdf_path=pdf_path,
        dpi=dpi,
        margin_percent=renderer.options.margin_percent,
    )

    client = OpenAIChatClient(api_key=api_key)
    oa_opts = OpenAIOptions(model=model)

    # 산출물 배치
    #  - 사용자용: 메타데이터/권역설명+항목 JSON (result_dir 바로 아래)
    #  - 캐시/디버그: <cache_base>/_cache/ (OS 캐시)
    cache_dir = run_cache_dir(pdf_path)

    # ------------------------------------------------------------------
    # 안전 재개 마커("dirty" marker: 미완료 표시)
    # ------------------------------------------------------------------
    # 추출이 끝나기 전까지 캐시는 불완전 상태로 간주
    # 중단/크래시 시 마커가 남아 다음 실행은 재개를 시도하지 않음

    dirty_marker = cache_dir / "__INCOMPLETE__"
    forced_fresh = bool(resume and dirty_marker.exists())
    if forced_fresh:
        # 중단된 실행의 캐시를 신뢰하지 않음
        resume = False
        _progress(
            "resume",
            message="불완전 캐시 감지: 안전을 위해 재개(resume)를 끄고 새로 추출합니다.",
        )
        # 구 캐시 JSON을 최대한 정리해 신구 데이터 혼합 방지
        try:
            for fp in result_dir.glob("*.json"):
                try:
                    fp.unlink()
                except Exception:
                    pass
        except Exception:
            pass

    # 완료 전까지는 실행 상태를 'incomplete'로 표시
    try:
        save_json(
            {
                "status": "incomplete",
                "started_at": datetime.now().isoformat(timespec="seconds"),
            },
            str(dirty_marker),
        )
    except Exception:
        pass

    meta_path = result_dir / "메타데이터.json"
    region_path = result_dir / "권역설명.json"
    toc_path = cache_dir / "_toc.json"

    whitelist_set = {int(p) for p in (page_whitelist or [])}

    items: List[Dict[str, Any]] = []
    metadata: Dict[str, Any] = {}
    region_info: Any = None
    side_meta_page: Optional[int] = None
    side_region_page: Optional[int] = None

    force_skip_meta = skip_meta_region
    skip_meta_region = bool(force_skip_meta) if force_skip_meta is not None else False
    if whitelist_set and force_skip_meta is None:
        skip_meta_region = True

    # -------------------------
    # 1) 계획/목차/초기 스캔
    # -------------------------
    _cancel_check()
    _progress("plan", message="계획 수립/목차 분석 중...")
    fixed_items = side.fixed_items

    if side.initial_scan and not side.has_toc:
        # 초기 스캔 모드
        scan_start = int(side.initial_scan.get("시작", 2))
        scan_end = int(side.initial_scan.get("끝", 8))

        # 메타/목차 계획 캐시가 있으면 재개 가능
        scan_plan_path = cache_dir / "_initial_scan.json"
        scan_data = None
        if resume and scan_plan_path.exists():
            try:
                scan_data = json.loads(scan_plan_path.read_text(encoding="utf-8"))
            except Exception:
                scan_data = None

        if scan_data is None:
            print(f"\n[1/4] 초기 스캔... (PDF {scan_start}~{scan_end}페이지)")
            _cancel_check()
            _progress("plan", message=f"초기 스캔 중... (PDF {scan_start}~{scan_end}페이지)")
            scan_images = renderer.render_range(scan_start, scan_end)
            prompt_initial = prompt_builder.build_prompt_initial_scan(org.raw, survey_type, scan_start, scan_end)
            _cancel_check()
            scan_text = client.chat(prompt_initial, scan_images, oa_opts)
            scan_data = parse_json_response(scan_text)
            save_json(scan_data, str(scan_plan_path))

        metadata = scan_data.get("메타데이터", {}) or {}
        if metadata and (not meta_path.exists() or not resume):
            save_json(metadata, str(meta_path))

        region_info = scan_data.get("권역설명")
        if survey_type == "지방" and region_info:
            # 레거시 초기 스캔 결과가 리스트일 수 있어 {'권역': [...]} 형태로 정규화
            if isinstance(region_info, list):
                region_info = {"권역": region_info}
            region_info = cleanup_region_info(region_info)
            if not region_path.exists() or not resume:
                save_json(region_info if isinstance(region_info, dict) else {"권역": region_info}, str(region_path))

        crosstab_start_page = int(scan_data.get("교차분석시작페이지", 7))
        # 페이지 순회 설정으로 변환
        side_page_walk = {"시작": crosstab_start_page, "끝": -1}
        skip_meta_region = True
        page_walk = side_page_walk

    elif not side.has_toc and fixed_items:
        print("\n[1/4] 목차 없음 - 고정 항목 사용")
        for name, page in fixed_items.items():
            items.append({"name": name, "page": page})

        page_walk = None

    elif side.has_toc:
        # 목차 모드
        if resume and toc_path.exists():
            try:
                toc_data = json.loads(toc_path.read_text(encoding="utf-8"))
            except Exception:
                toc_data = None
        else:
            toc_data = None

        if toc_data is None:
            if side.toc_page is None:
                raise ValueError(f"[{org.name}] {survey_type}: 목차 페이지가 없습니다")
            toc_page = int(side.toc_page)
            print(f"\n[1/4] 목차 분석... (PDF {toc_page}페이지)")
            _cancel_check()
            _progress("plan", message=f"목차 분석 중... (PDF {toc_page}페이지)")
            toc_images = renderer.render_range(toc_page, toc_page)
            prompt_toc = prompt_builder.build_prompt_toc(org.raw)
            _cancel_check()
            toc_text = client.chat(prompt_toc, toc_images, oa_opts)
            toc_data = parse_json_response(toc_text)
            save_json(toc_data, str(toc_path))

        items = toc_data.get("items", []) or []

        # 선택: 목차 정보로 메타/권역 페이지를 덮어쓰기
        # 참고: 목차에 적힌 페이지 번호는 '문서 내 페이지(인쇄 번호)'인 경우가 많아
        #       실제 PDF 페이지로 변환하기 위해 page_offset을 적용한다.
        #       (items 추출에서 이미 page_offset을 적용하고 있으며, 메타/권역도 동일 규칙을 따른다.)
        toc_offset = int(side.page_offset or 0)

        if side.metadata_page is not None:
            side_meta_page = side.metadata_page
        elif toc_data.get("조사설계페이지"):
            side_meta_page = int(toc_data["조사설계페이지"]) + toc_offset
        else:
            side_meta_page = None

        if side.region_page is not None:
            side_region_page = side.region_page
        elif toc_data.get("응답자특성페이지"):
            side_region_page = int(toc_data["응답자특성페이지"]) + toc_offset
        else:
            side_region_page = None

        page_walk = None

    elif side.page_walk:
        page_walk = side.page_walk
    else:
        raise ValueError(f"[{org.name}] {survey_type}: 목차/고정항목/페이지순회 설정이 없습니다")

    # (목차 덮어쓰기 반영) 실제 메타/권역 페이지 결정
    if side.has_toc:
        meta_page = side_meta_page
        region_page = side_region_page
    else:
        meta_page = side.metadata_page
        region_page = side.region_page

    # 목차가 YAML을 덮어쓸 때 모델 추정이 빗나갈 수 있으므로 메타 페이지를 보강 결정
    if side.has_toc and meta_page != side.metadata_page:
        meta_page = _resolve_metadata_page(str(pdf_path), side.metadata_page, meta_page, org)

    # -------------------------
    # 2) 메타데이터
    # -------------------------
    if not skip_meta_region:
        if resume and meta_path.exists():
            try:
                metadata = json.loads(meta_path.read_text(encoding="utf-8"))
            except Exception:
                metadata = {}

            # KRI 메타데이터가 빈 채로 캐시된 경우가 많아서, 로드 후에도 텍스트 폴백을 한번 더 시도
            if meta_page is not None:
                metadata2 = _fill_metadata_with_text_fallback(metadata, pdf_path, int(meta_page), org)
                if metadata2 != metadata:
                    metadata = metadata2
                metadata_norm = _normalize_metadata_fields(metadata)
                if metadata_norm != metadata:
                    metadata = metadata_norm
                save_json(metadata, str(meta_path))
        else:
            if meta_page is None:
                metadata = {}
            else:
                print(f"\n[2/4] 메타데이터... (PDF {meta_page}페이지)")
                _cancel_check()
                _progress("meta", message=f"메타데이터 추출 중... (PDF {meta_page}페이지)")
                meta_images = renderer.render_range(int(meta_page), int(meta_page))
                prompt_meta = prompt_builder.build_prompt_metadata(org.raw, survey_type)
                _cancel_check()
                meta_text = client.chat(prompt_meta, meta_images, oa_opts)
                metadata = parse_json_response(meta_text) or {}
                metadata = _fill_metadata_with_text_fallback(metadata, pdf_path, int(meta_page), org)
                metadata = _normalize_metadata_fields(metadata)
                save_json(metadata, str(meta_path))
    else:
        if resume and meta_path.exists():
            try:
                metadata = json.loads(meta_path.read_text(encoding="utf-8"))
            except Exception:
                pass

    # -------------------------
    # 3) 권역설명 (지방 전용)
    # -------------------------
    if survey_type == "지방":
        if not skip_meta_region:
            if resume and region_path.exists():
                try:
                    region_info = json.loads(region_path.read_text(encoding="utf-8"))
                except Exception:
                    region_info = None
                region_info = cleanup_region_info(region_info)
                # LLM 결과가 비면 pdfplumber 텍스트의 '권역 구분' 블록으로 보충
                region_info2 = _fill_region_with_text_fallback(region_info, pdf_path, region_page, org)
                if region_info2 is not region_info:
                    region_info = cleanup_region_info(region_info2)
            else:
                if region_page is None:
                    region_info = None
                else:
                    print(f"\n[3/4] 권역설명... (PDF {region_page}페이지)")
                    _cancel_check()
                    _progress("region", message=f"권역설명 추출 중... (PDF {region_page}페이지)")

                    if side.region_with_metadata and meta_page is not None and int(region_page) == int(meta_page):
                        region_images = renderer.render_range(int(meta_page), int(meta_page))
                    else:
                        region_images = renderer.render_range(int(region_page), int(region_page))

                    prompt_region = prompt_builder.build_prompt_region(org.raw)
                    _cancel_check()
                    region_text = client.chat(prompt_region, region_images, oa_opts)
                    region_info = parse_json_response(region_text)
                    if isinstance(region_info, list):
                        region_info = {"권역": region_info}
                    region_info = cleanup_region_info(region_info)
                # LLM 결과가 비면 pdfplumber 텍스트의 '권역 구분' 블록으로 보충
                region_info2 = _fill_region_with_text_fallback(region_info, pdf_path, region_page, org)
                if region_info2 is not region_info:
                    region_info = cleanup_region_info(region_info2)
                    save_json(region_info, str(region_path))
        else:
            if resume and region_path.exists():
                try:
                    region_info = json.loads(region_path.read_text(encoding="utf-8"))
                except Exception:
                    region_info = region_info
                region_info = cleanup_region_info(region_info)
                # LLM 결과가 비면 pdfplumber 텍스트의 '권역 구분' 블록으로 보충
                region_info2 = _fill_region_with_text_fallback(region_info, pdf_path, region_page, org)
                if region_info2 is not region_info:
                    region_info = cleanup_region_info(region_info2)

    # -------------------------
    # 4) 표 추출
    # -------------------------
    # 참고: '항목'에는 성공한 항목만 담습니다.
    #       GUI에서 빠진 항목을 쉽게 보도록 실패 항목을 별도로 기록합니다.
    results = {
        "메타데이터": metadata,
        "권역설명": region_info,
        "항목": [],
        "실패": [],
    }

    if page_walk:
        # 페이지 순회 모드
        start_page = int(page_walk.get("시작", 4))
        end_page = int(page_walk.get("끝", -1))

        # 총 페이지 수가 필요할 때만 계산
        if end_page == -1:
            from pdf2image.pdf2image import pdfinfo_from_path

            info = pdfinfo_from_path(pdf_path, userpw=None, poppler_path=poppler_path)
            end_page = int(info.get("Pages", start_page))

        print(f"\n[4/4] 표 추출... (페이지 순회: {start_page}~{end_page})")

        auto_prompt = prompt_builder.build_prompt_table_auto_type(org.raw, survey_type)

        skip_keywords = []
        if side.page_skip and isinstance(side.page_skip.get("키워드"), list):
            skip_keywords = [str(x) for x in side.page_skip.get("키워드")]

        # 페이지 순회 자동 캐시:
        #  - (구버전) _cache/pXXXX_auto.json (페이지별 파일)
        #  - (현재)   _cache/_page_walk_cache.json (단일 파일)
        #
        # 이 캐시를 두면 모델 재호출 없이 재개할 수 있으면서도 _cache에 수백 개의 JSON이 생기는 것을 방지
        page_cache_path = cache_dir / "_page_walk_cache.json"
        page_cache: Dict[str, Any] = {}

        def _load_page_walk_cache() -> Dict[str, Any]:
            # 단일 캐시 파일을 우선 사용
            if resume and page_cache_path.exists():
                try:
                    raw_cache = json.loads(page_cache_path.read_text(encoding="utf-8"))
                    if isinstance(raw_cache, dict):
                        if isinstance(raw_cache.get("pages"), dict):
                            return raw_cache["pages"]
                        # 구버전 페이지별 캐시 파일을 발견하면 병합
                        return raw_cache
                except Exception:
                    pass

            # 구버전 페이지별 캐시 파일이 있으면 병합 후 단일 캐시로 저장
            pages: Dict[str, Any] = {}
            if resume:
                for fp in sorted(cache_dir.glob("p*_auto.json")):
                    m = re.match(r"p(\d{4})_auto\.json$", fp.name)
                    if not m:
                        continue
                    try:
                        pnum = str(int(m.group(1)))
                        pages[pnum] = json.loads(fp.read_text(encoding="utf-8"))
                    except Exception:
                        continue

                if pages:
                    try:
                        save_json({"schema": 1, "pages": pages}, str(page_cache_path))
                        # 구버전 파일을 지워 캐시를 깔끔하게 유지
                        for fp in cache_dir.glob("p*_auto.json"):
                            try:
                                fp.unlink()
                            except Exception:
                                pass
                    except Exception:
                        pass

            return pages

        def _save_page_walk_cache(pages: Dict[str, Any]) -> None:
            try:
                # 캐시가 비면 "중복 JSON" 착시를 막기 위해 캐시 파일 삭제
                if not pages:
                    try:
                        if page_cache_path.exists():
                            page_cache_path.unlink()
                    except Exception:
                        pass
                    return
                save_json({"schema": 1, "pages": pages}, str(page_cache_path))
            except Exception:
                # 캐시 저장 실패가 추출을 중단시키면 안 됨
                pass

        def _build_table_retry_plan(
            item_type: str,
            images: List[Any],
            page_num: int,
            missing_only: bool,
            prompt_table: str,
            prompt_strict: str,
        ) -> List[Tuple[str, List[Any], Any]]:
            # 크롭/세로분할 변형 생성(촘촘한 ISSUE/PSR에 도움)
            cropped_imgs: List[Any] = []
            try:
                cropped_once = table_cropper.crop_images(images, [page_num])
                cropped_imgs = list(cropped_once)
                if item_type in {"PSR", "ISSUE"} and cropped_imgs:
                    split_imgs: List[Any] = []
                    for cim in cropped_imgs:
                        split_imgs.extend(_split_vertical_image(cim))
                    if split_imgs:
                        cropped_imgs = split_imgs
            except Exception:
                cropped_imgs = []

            # 크롭이 없을 때를 대비한 풀페이지 세로분할
            split_full_imgs: List[Any] = []
            if not cropped_imgs and item_type in {"PSR", "ISSUE"} and images:
                try:
                    split_full_imgs = _split_vertical_image(images[0])
                except Exception:
                    split_full_imgs = []

            # 상단만 다시 자르는 변형(헤더 유실 보완)
            top_crop_imgs: List[Any] = []
            if missing_only and images:
                try:
                    top_crop_imgs = [_crop_top_portion(img, 0.5) for img in images if img is not None]
                    split_top: List[Any] = []
                    for tim in top_crop_imgs:
                        split_top.extend(_split_vertical_image(tim))
                    if split_top:
                        top_crop_imgs = split_top
                except Exception:
                    top_crop_imgs = []

            attempt_plan: List[Tuple[str, List[Any], Any]] = []
            if cropped_imgs:
                attempt_plan.extend(
                    [
                        ("cropped", cropped_imgs, prompt_table),
                        ("cropped_strict", cropped_imgs, prompt_strict),
                    ]
                )
            if top_crop_imgs:
                attempt_plan.extend(
                    [
                        ("top_crop", top_crop_imgs, prompt_table),
                        ("top_crop_strict", top_crop_imgs, prompt_strict),
                    ]
                )
            if split_full_imgs:
                attempt_plan.extend(
                    [
                        ("split_full", split_full_imgs, prompt_table),
                        ("split_full_strict", split_full_imgs, prompt_strict),
                    ]
                )
            attempt_plan.extend(
                [
                    ("full", images, prompt_table),
                    ("full_strict", images, prompt_strict),
                ]
            )
            return attempt_plan

        def _execute_table_retry_plan(
            attempt_plan: List[Tuple[str, List[Any], Any]],
            *,
            item_type: str,
            question_keep: Any,
            include_cfg: Any,
            region_info: Any,
            item_name: str,
        ) -> Any:
            best: Any = None
            for _tag, _imgs, _pr in attempt_plan:
                _cancel_check()
                table_text2 = client.chat(_pr, _imgs, oa_opts)
                cand = parse_json_response(table_text2)
                try:
                    cand = coerce_extracted_table_schema(cand)
                except Exception:
                    pass

                cand_issues: List[str] = []
                try:
                    cand_issues.extend(validate_table_data(item_type, cand))
                except Exception:
                    pass
                if item_type == "ISSUE" and _issues_missing_items_only(cand_issues or []):
                    cand_issues = []
                try:
                    cand_issues.extend(validate_expected_groups(cand, include=include_cfg, region_info=region_info))
                except Exception:
                    pass
                try:
                    if item_type == "PSR":
                        cand_issues.extend(_run_psr_strict_validation_if_configured(org, item_name, cand))
                except Exception:
                    pass

                if not cand_issues and isinstance(cand, dict) and isinstance(cand.get("데이터"), dict):
                    if (not cand.get("질문")) and question_keep:
                        cand["질문"] = question_keep
                    best = cand
                    break
            return best

        page_cache = _load_page_walk_cache()

        # 중요: 사용자용 페이지별 JSON과 _page_walk_cache.json에 같은 데이터를 중복 저장하지 않기
        #
        # 어떤 페이지에 결과 JSON이 이미 있으면 그 파일로 재개하고,
        # 캐시된 raw 엔트리는 제거해 중복을 피함
        if resume and page_cache:
            cleaned = False
            for k in list(page_cache.keys()):
                try:
                    pnum = int(str(k))
                except Exception:
                    continue
                if list(result_dir.glob(f"p{pnum:04d}_*.json")):
                    page_cache.pop(str(k), None)
                    cleaned = True
            if cleaned:
                _save_page_walk_cache(page_cache)

        page_list = list(range(start_page, end_page + 1))
        if whitelist_set:
            page_list = [p for p in page_list if p in whitelist_set]

        total_pages = len(page_list) if page_list else 0
        if total_pages:
            _progress("tables", current=0, total=total_pages, message=f"표 추출 중... (0/{total_pages})")
        else:
            _progress("tables", message="표 추출 중...")

        for idx_page, page_num in enumerate(page_list, start=1):
            _cancel_check()
            if total_pages:
                _progress(
                    "tables",
                    current=idx_page,
                    total=total_pages,
                    message=f"표 추출 중... ({idx_page}/{total_pages})  p{page_num}",
                )
            print(f"\n  - PDF {page_num}페이지")

            # 빠른 재개: 해당 페이지 사용자용 JSON이 있으면 그대로 신뢰하고 모델을 다시 부르지 않음
            if resume:
                existing = list(result_dir.glob(f"p{page_num:04d}_*.json"))
                if existing:
                    # 가장 최근 파일 선택(보통 1개)
                    chosen = max(existing, key=lambda p: p.stat().st_mtime)
                    try:
                        obj = json.loads(chosen.read_text(encoding="utf-8"))
                        existing_type = (obj.get("항목유형") or "").upper()
                        existing_name = obj.get("항목명") or obj.get("name") or chosen.stem
                    except Exception:
                        m = re.match(r"p\d{4}_([A-Z]+)_", chosen.name)
                        existing_type = (m.group(1) if m else "").upper()
                        existing_name = chosen.stem

                    if existing_type in extract_types:
                        results["항목"].append(
                            {"name": existing_name, "type": existing_type, "file": chosen.name}
                        )
                    continue

            images = renderer.render_range(page_num, page_num)
            if not images:
                continue

            # 키워드 기반 빠른 건너뛰기: OCR은 없지만 모델이 'SKIP'을 내면 건너뜀

            # 통합 재개 캐시: _page_walk_cache.json의 페이지별 엔트리
            page_key = str(page_num)
            table_data = page_cache.get(page_key) if resume else None

            if table_data is None:
                imgs_for_prompt = images
                _cancel_check()
                table_text = client.chat(auto_prompt, imgs_for_prompt, oa_opts)
                table_data = parse_json_response(table_text)
                # 모델 출력의 흔한 스키마/키 변형 보정
                # (예: '헤더'/'columns' -> '응답항목', 'data'/'rows' -> '데이터')
                try:
                    table_data = coerce_extracted_table_schema(table_data)
                except Exception:
                    pass
                page_cache[page_key] = table_data
                _save_page_walk_cache(page_cache)

            raw_item_type = (table_data.get("항목유형") or "ISSUE").upper()
            if raw_item_type == "SKIP":
                continue

            question = table_data.get("질문") or table_data.get("제목") or table_data.get("항목명") or table_data.get("title")
            item_type = _postprocess_auto_type(raw_item_type, question, table_data.get("응답항목"))

            # 후처리 결과 유형이 바뀌면(예: '성과 분야' 문항: GE -> ISSUE)
            # 첫 추출이 GE 고정 스키마(3열)를 썼을 수 있음.
            # 이 경우 올바른 유형 프롬프트로 한 번 더 돌려 응답항목 유실을 방지
            if item_type != raw_item_type:
                try:
                    fix_item_name = (
                        "국정운영평가" if item_type == "GE" else
                        ("정당지지도" if item_type == "PSR" else (question or f"문{page_num}"))
                    )
                    fix_prompt = prompt_builder.build_prompt_table(org.raw, item_type, fix_item_name, survey_type)
                    _cancel_check()
                    fix_text = client.chat(fix_prompt, images, oa_opts)
                    fix_data = parse_json_response(fix_text)
                    try:
                        fix_data = coerce_extracted_table_schema(fix_data)
                    except Exception:
                        pass
                    if isinstance(fix_data, dict) and isinstance(fix_data.get("데이터"), dict):
                        fix_data["항목유형"] = item_type
                        if fix_data.get("질문") in (None, ""):
                            fix_data["질문"] = question
                        table_data = fix_data
                        page_cache[page_key] = table_data
                        _save_page_walk_cache(page_cache)
                        question = table_data.get("질문") or question
                except Exception:
                    # 재시도 실패 시 원본 auto-type 결과 유지
                    pass

            if item_type not in extract_types:
                continue

            if item_type == "GE":
                item_name = "국정운영평가"
            elif item_type == "PSR":
                item_name = "정당지지도"
            else:
                item_name = question or f"문{page_num}"

            # -----------------------------------------------------------------
            # 리서치뷰 등: auto-type 단계에서는 제목/캡션 유실을 피하려고 크롭을 끄지만,
            # PSR/GE는 풀페이지 이미지에서 숫자 오독/누락이 자주 발생합니다.
            # => 유형이 확정된 뒤, 표 본문만 크롭해서 한 번 더 추출(재시도)합니다.
            #    (PSR은 촘촘한 표라 세로로 2~3분할해 가독성을 올립니다.)
            # -----------------------------------------------------------------
            if item_type in {"PSR", "GE", "ISSUE"}:
                try:
                    cfg_side = org.raw.get("전국" if survey_type == "전국" else "지방", {}) or {}
                    include_cfg = cfg_side.get("포함") or cfg_side.get("include")
                except Exception:
                    include_cfg = None

                issues: List[str] = []
                try:
                    issues.extend(validate_table_data(item_type, table_data))
                except Exception:
                    pass
                missing_only = _issues_missing_items_only(issues or [])
                try:
                    issues.extend(validate_expected_groups(table_data, include=include_cfg, region_info=region_info))
                except Exception:
                    pass
                try:
                    if item_type == "PSR":
                        issues.extend(_run_psr_strict_validation_if_configured(org, item_name, table_data))
                except Exception:
                    pass

                # PSR/ISSUE는 풀페이지 이미지에서 하단 블록(지역/직업/이념 등) 누락이 잦아,
                # 이 단계에서 한 번 더 크롭 기반 추출을 시도합니다.
                # (GE는 문제가 감지된 경우에만 재시도)
                should_retry = bool(issues) or missing_only or (item_type == "PSR")

                if should_retry:
                    try:
                        prompt_table = prompt_builder.build_prompt_table(org.raw, item_type, item_name, survey_type)
                        prompt_strict = prompt_table + """

[추가 규칙]
- 보이는 값만 추출하고, 보이지 않는 값은 추정하지 마세요.
- 숫자는 퍼센트 그대로(예: .9는 0.9).
- JSON 외 텍스트를 절대 출력하지 마세요.
"""
                        attempt_plan = _build_table_retry_plan(
                            item_type, images, page_num, missing_only, prompt_table, prompt_strict
                        )
                        best = _execute_table_retry_plan(
                            attempt_plan,
                            item_type=item_type,
                            question_keep=question,
                            include_cfg=include_cfg,
                            region_info=region_info,
                            item_name=item_name,
                        )

                        if isinstance(best, dict) and isinstance(best.get("데이터"), dict):
                            # 원래 질문(자동 타입에서 얻은 캡션)을 유지하되 비었을 때만 대체
                            if not question and isinstance(best.get("질문"), str):
                                question = best.get("질문")
                            table_data = best
                            # 개선된 추출 결과를 재개 시에도 쓰도록 페이지별 캐시 갱신
                            page_cache[page_key] = table_data
                            _save_page_walk_cache(page_cache)
                    except Exception:
                        # 재시도 실패 시 원본 추출 결과 유지
                        pass

            # -----------------------------------------------------------------
            # ISSUE 제목 보정 (텍스트 레이어 기반)
            #
            # 페이지 순회(auto-type)에서 LLM이 제목의 대괄호 수식어를
            # 누락하는 케이스가 관측됨:
            #   "서울시장 후보 선호도 [보수진영]" -> "서울시장 후보 선호도"
            #
            # News1/YTN 통계표 계열은 보통 '[표X] ...' 캡션이 텍스트로 들어가므로
            # pdfplumber로 캡션을 뽑아, 괄호 수식어가 있는 경우 우선 적용합니다.
            # -----------------------------------------------------------------
            if item_type == "ISSUE":
                cap = _extract_table_caption_from_pdf_text(pdf_path, page_num)
                if isinstance(cap, str) and cap.strip():
                    cap_s = cap.strip()
                    q_s = (question or "").strip()
                    cap_comp = re.sub(r"\s+", "", cap_s)
                    q_comp = re.sub(r"\s+", "", q_s)

                    # 모델이 누락한 대괄호 수식어가 캡션에 있으면 캡션을 우선 사용
                    if (
                        (not q_s)
                        or (("[" in cap_s and "]" in cap_s) and ("[" not in q_s) and cap_comp.startswith(q_comp))
                    ):
                        question = cap_s
                        item_name = cap_s

            # 정규화된 항목 JSON 저장(작성기가 기대하는 형태)
            final_name = f"p{page_num:04d}_{item_type}_{_safe_filename(item_name, 50)}.json"
            final_path = result_dir / final_name
            if not (resume and final_path.exists()):
                # 새 산출물을 쓸 때 과거 실행의 구버전 파일은 정리
                # (예: 패치 후 GE/ISSUE 유형이 바뀐 경우)
                page_prefix = f"p{page_num:04d}_"
                for old in result_dir.glob(f"{page_prefix}*.json"):
                    if old.name != final_name:
                        try:
                            old.unlink()
                        except Exception:
                            pass

                item_result = {
                    "항목명": item_name,
                    "항목유형": item_type,
                    "페이지": [page_num],
                    "메타데이터": metadata,
                    "질문": question,
                    **{k: v for k, v in table_data.items() if k not in ["항목유형", "질문"]},
                }
                if region_info:
                    item_result["권역설명"] = region_info
                # 저장 전 한 번 정규화해 JSON 결과도 결정적이도록 유지(동의어 병합, 라벨 표준화 등)
                # 엑셀 작성기도 정규화하지만 JSON을 깨끗하게 두면 디버깅이 쉬움
                try:
                    item_result, _norm_report = normalize_table_json(
                        item_result, org, survey_type, region_info=region_info
                    )
                except Exception:
                    # 정규화 실패는 치명적이지 않으므로 추출을 중단하지 않음
                    pass
                save_json(item_result, str(final_path))

                # 사용자용 JSON을 썼으면 중복을 막기 위해 해당 페이지 캐시 원본 제거
                if resume and page_key in page_cache:
                    try:
                        page_cache.pop(page_key, None)
                        _save_page_walk_cache(page_cache)
                    except Exception:
                        pass

            results["항목"].append({"name": item_name, "type": item_type, "file": final_path.name})

    else:
        # 목차/고정항목 모드
        print(f"\n[4/4] 표 추출... ({len(items)}개 항목)")

        if side.page_offset is None:
            raise ValueError(
                f"[{org.name}] {survey_type}: '페이지오프셋'이 없습니다. "
                f"(실제 PDF 페이지 번호를 쓰면 0)"
            )
        page_offset = int(side.page_offset)

        # TOC 모드에서도 '원본(정규화 전) 추출 결과'를 _cache에 남겨두면
        # 테스트/디버깅 시 모델 출력과 정규화 결과를 비교하기 쉽습니다.
        # (페이지순회 모드의 _page_walk_cache.json 과 유사한 목적)
        toc_table_cache_path = cache_dir / "_toc_table_cache.json"
        toc_table_cache: Dict[str, Any] = {}

        if resume and toc_table_cache_path.exists():
            try:
                raw_cache = json.loads(toc_table_cache_path.read_text(encoding="utf-8"))
                if isinstance(raw_cache, dict):
                    if isinstance(raw_cache.get("items"), dict):
                        toc_table_cache = raw_cache["items"]
                    elif isinstance(raw_cache.get("tables"), dict):
                        toc_table_cache = raw_cache["tables"]
                    else:
                        # 구버전 직접 dict 구조도 허용
                        toc_table_cache = raw_cache
            except Exception:
                toc_table_cache = {}

        def _save_toc_table_cache() -> None:
            try:
                save_json({"schema": 1, "items": toc_table_cache}, str(toc_table_cache_path))
            except Exception:
                pass


        # page_whitelist가 있으면 해당 PDF 페이지에 매핑되는 항목만 선택
        filtered_items = []
        if whitelist_set:
            for item in items:
                try:
                    page_num_int = int(item.get("page") or item.get("페이지"))
                except Exception:
                    continue
                if page_num_int is not None and (page_num_int + int(side.page_offset or 0)) in whitelist_set:
                    filtered_items.append(item)
        else:
            filtered_items = items

        # 진행률을 정확히 보여주도록 작업 리스트를 미리 산출(선택된 유형만)
        todo: List[Dict[str, Any]] = []
        for it in filtered_items:
            item_name = it.get("name") or it.get("항목명") or ""
            page_num = it.get("page") or it.get("페이지")
            try:
                page_num_int = int(page_num)
            except Exception:
                continue

            # 주의: _toc.json 안 it["type"] 같은 LLM 추가 필드는 신뢰하지 않음.
            # 항목명을 기반으로만 추정해 GE 오검출(전 대통령, 트럼프 대통령 등)을 방지
            item_type = _guess_item_type_from_name(item_name).upper()
            if item_type not in extract_types:
                continue

            todo.append({"item": it, "name": item_name, "type": item_type, "page_num": page_num_int})

        total_items = len(todo)
        if total_items:
            _progress("tables", current=0, total=total_items, message=f"표 추출 중... (0/{total_items})")
        else:
            _progress("tables", message="표 추출 중...")

        for idx, t in enumerate(todo, start=1):
            _cancel_check()
            if total_items:
                _progress(
                    "tables",
                    current=idx,
                    total=total_items,
                    message=f"표 추출 중... ({idx}/{total_items})",
                )

            item = t["item"]
            item_name = t["name"]
            item_type = t["type"]
            page_num_int = int(t["page_num"])

            pdf_first_page = page_num_int + page_offset
            pdf_last_page = pdf_first_page + int(side.pages_per_item) - 1

            print(f"\n  [{idx}/{total_items or len(todo)}] {item_name} ({item_type})")
            print(f"    PDF {pdf_first_page}~{pdf_last_page}페이지")

            out_name = f"p{pdf_first_page:04d}_{item_type}_{_safe_filename(item_name, 50)}.json"
            out_path = result_dir / out_name
            if resume and out_path.exists():
                results["항목"].append({"name": item_name, "type": item_type, "file": out_path.name})
                continue

            pages = list(range(pdf_first_page, pdf_last_page + 1))
            images_full = renderer.render_pages(pages)
            if not images_full:
                continue

            prompt_table = prompt_builder.build_prompt_table(org.raw, item_type, item_name, survey_type)
            prompt_strict = prompt_table + """

[추가 규칙]
- 보이는 값만 추출하고, 보이지 않는 값은 추정하지 마세요.
- 숫자는 퍼센트 그대로(예: .9는 0.9).
- 후보 선호도/적합도 표의 **후보명/정당명**은 표에 적힌 그대로 추출하세요.
- '후보1', '후보2' 같은 임의 치환/축약은 절대 사용하지 마세요.
- JSON 외 텍스트를 절대 출력하지 마세요.
"""

            # 크롭 -> 풀페이지 -> strict 변형 순서로 시도(크롭 우선)
            table_data = None
            chosen_tag = None
            last_err = None
            question_keep = item_name

            use_right_half = bool(getattr(side, "toc_right_half_crop", False))
            right_half_imgs: List[Any] = []
            if use_right_half:
                try:
                    right_half_imgs = [_crop_right_half(img, 0.45) for img in images_full if img is not None]
                    split_right: List[Any] = []
                    for rim in right_half_imgs:
                        split_right.extend(_split_vertical_image(rim))
                    if split_right:
                        right_half_imgs = split_right
                except Exception:
                    right_half_imgs = []

            # 크롭/세로분할 변형 생성(촘촘한 ISSUE/PSR에 도움)
            cropped_imgs: List[Any] = []
            try:
                cropped_once = table_cropper.crop_images(images_full, pages)
                cropped_imgs = list(cropped_once)
                if item_type in {"PSR", "ISSUE"} and cropped_imgs:
                    split_imgs: List[Any] = []
                    for cim in cropped_imgs:
                        split_imgs.extend(_split_vertical_image(cim))
                    if split_imgs:
                        cropped_imgs = split_imgs
            except Exception:
                cropped_imgs = []

            # 크롭 결과가 없으면 풀페이지를 세로분할해 재시도
            split_full_imgs: List[Any] = []
            if not cropped_imgs and item_type in {"PSR", "ISSUE"} and images_full:
                try:
                    split_full_imgs = _split_vertical_image(images_full[0])
                except Exception:
                    split_full_imgs = []

            top_crop_imgs: List[Any] = []
            # 검증 결과가 응답/데이터 누락 위주면 헤더 중심 상단 크롭도 시도
            if item_type == "ISSUE":
                try:
                    # 앞선 strict 루프 메시지에서 missing_only 플래그를 재사용하기 어려우니
                    # 백업 시도로 ISSUE에 상단 크롭을 미리 추가
                    top_crop_imgs = [_crop_top_portion(img, 0.5) for img in images_full if img is not None]
                    split_top: List[Any] = []
                    for tim in top_crop_imgs:
                        split_top.extend(_split_vertical_image(tim))
                    if split_top:
                        top_crop_imgs = split_top
                except Exception:
                    top_crop_imgs = []

            attempt_plan: List[tuple[str, List[Any], str]] = []
            if use_right_half and right_half_imgs:
                attempt_plan.extend(
                    [
                        ("right_half", right_half_imgs, prompt_table),
                        ("right_half_strict", right_half_imgs, prompt_strict),
                    ]
                )
            if cropped_imgs:
                attempt_plan.extend(
                    [
                        ("cropped", cropped_imgs, prompt_table),
                        ("cropped_strict", cropped_imgs, prompt_strict),
                    ]
                )
            if top_crop_imgs:
                attempt_plan.extend(
                    [
                        ("top_crop", top_crop_imgs, prompt_table),
                        ("top_crop_strict", top_crop_imgs, prompt_strict),
                    ]
                )
            if split_full_imgs:
                attempt_plan.extend(
                    [
                        ("split_full", split_full_imgs, prompt_table),
                        ("split_full_strict", split_full_imgs, prompt_strict),
                    ]
                )
            attempt_plan.extend(
                [
                    ("full", images_full, prompt_table),
                    ("full_strict", images_full, prompt_strict),
                ]
            )
            for tag, imgs, pr in attempt_plan:
                _cancel_check()
                table_text = client.chat(pr, imgs, oa_opts)
                try:
                    cand = parse_json_response(table_text)
                except Exception as e:
                    last_err = e
                    continue

                # 모델 출력의 흔한 스키마/키 변형 보정
                # (예: '헤더'/'응답지'/'columns' -> '응답항목', 'data'/'rows' -> '데이터')
                try:
                    cand = coerce_extracted_table_schema(cand)
                except Exception:
                    # 최선 시도일 뿐이며 실패 시 원본 cand 유지
                    pass

                # KRI: PDF 텍스트 기반 보정 (누락 행/병합 행 자동 복구)
                if _is_kri_org(org):
                    try:
                        from .kri_text_repair import try_repair_kri_table_data
                        cand = try_repair_kri_table_data(pdf_path, pages, cand, side)
                    except Exception:
                        # 보정 실패는 치명적이지 않으므로 원본 cand 유지
                        pass


                issues = validate_table_data(item_type, cand)
                missing_only = _issues_missing_items_only(issues or [])
                if item_type == "ISSUE" and missing_only:
                    # 다음 시도(예: 상단 크롭)가 돌도록 실패로 처리
                    last_err = ValueError(f"quality check failed ({tag}): missing 응답항목/데이터")
                    continue

                # 추가 완전성 검증: 기관 설정에 기대되는 주요 그룹(지역/연령/성별 등)이 비었는지 확인
                # 보수적으로, 다른 주요 그룹이 하나라도 있을 때만 실행
                if not issues:
                    try:
                        cfg_side = org.raw.get("전국" if survey_type == "전국" else "지방", {}) or {}
                        include_cfg = cfg_side.get("포함") or cfg_side.get("include")
                    except Exception:
                        include_cfg = None
                    more_issues = validate_expected_groups(cand, include=include_cfg, region_info=region_info)
                    if more_issues:
                        issues = (issues or []) + more_issues

                if issues:
                    last_err = ValueError(f"quality check failed ({tag}): " + "; ".join(issues))
                    continue

                # 촘촘한 PSR 테이블에 대한 선택적 엄격 검증
                extra = _run_psr_strict_validation_if_configured(org, item_name, cand)
                if extra:
                    last_err = ValueError(f"quality check failed ({tag}, strict): " + "; ".join(extra))
                    continue

                if question_keep and not cand.get("질문"):
                    cand["질문"] = question_keep

                table_data = cand
                chosen_tag = tag
                break

            if table_data is None:
                print(f"    ⚠️ 추출 실패: {last_err}")
                try:
                    results.setdefault("실패", []).append(
                        {
                            "name": item_name,
                            "type": item_type,
                            "pages": list(range(pdf_first_page, pdf_last_page + 1)),
                            "error": str(last_err) if last_err is not None else "unknown",
                        }
                    )
                except Exception:
                    pass
                continue

            # 캐시(정규화 전 원본) 저장: TOC 모드에서도 디버깅 가능하도록
            try:
                toc_table_cache[out_name] = {
                    "item_name": item_name,
                    "item_type": item_type,
                    "pages": pages,
                    "chosen_attempt": chosen_tag,
                    "raw": table_data,
                }
                _save_toc_table_cache()
            except Exception:
                pass

            item_result: Dict[str, Any] = {
                "항목명": item_name,
                "항목유형": item_type,
                "페이지": list(range(pdf_first_page, pdf_last_page + 1)),
                "메타데이터": metadata,
                **table_data,
            }
            if region_info:
                item_result["권역설명"] = region_info

            # 저장 전 한 번 정규화해 JSON 결과도 결정적이도록 유지(동의어 병합, 라벨 표준화 등)
            # 엑셀 작성기도 정규화하지만 JSON을 깨끗하게 두면 디버깅이 쉬움
            try:
                item_result, _norm_report = normalize_table_json(
                    item_result, org, survey_type, region_info=region_info
                )
            except Exception:
                # 정규화 실패는 치명적이지 않으므로 추출을 중단하지 않음
                pass

            save_json(item_result, str(out_path))
            results["항목"].append({"name": item_name, "type": item_type, "file": out_path.name})


    # 후처리: ISSUE 항목 전반에서 후보 헤더 보강(최선 시도)
    try:
        _enrich_issue_candidate_titles_across_items(result_dir)
    except Exception:
        pass

    # 후처리: 중복 *_m.json 산출물 정리(안전)
    try:
        _cleanup_duplicate_m_outputs(result_dir)
    except Exception:
        pass

    # 후처리: 파생 '부동층' 열이 잔여 카테고리와 중복일 때 제거(안전)
    try:
        _drop_derived_budongcheung_outputs(result_dir)
    except Exception:
        pass

    # 후처리: 유사 중복 응답항목 정리(안전)
    try:
        _dedupe_near_duplicate_response_items_outputs(result_dir)
    except Exception:
        pass

    # 실행 후 캐시 용량 정리(오래된 캐시부터 삭제)
    # 캐시 정리 실패가 파이프라인을 멈추게 해서는 안 됨
    try:
        enforce_cache_quota()
    except Exception:
        pass

    # GUI가 사용자 출력 폴더를 건드리지 않고도 Excel 변환할 수 있도록 JSON 디렉터리 전달
    try:
        results["json_dir"] = str(result_dir)
        results["pdf_name"] = pdf_name
        results["forced_fresh"] = bool(forced_fresh)
    except Exception:
        pass

    # 완료 표시: "incomplete" 마커 제거
    # 사용자가 취소하거나 예외로 중단되면 여기까지 못 오므로 마커가 남아 다음 실행을 안전 재실행으로 강제
    try:
        if dirty_marker.exists():
            dirty_marker.unlink()
    except Exception:
        pass

    print(f"\n완료! JSON 결과 폴더(캐시): {result_dir}")
    return results
