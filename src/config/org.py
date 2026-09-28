"""기관 설정 로드/정규화.

목표:
- 설정은 YAML 기반(코드에 기관별 하드코딩 없음)
- 기존 프로젝트의 레거시 YAML 형태도 지원
- '#Uacb0#Uacfc...' 같은 유니코드 깨짐 파일명도 처리

공개 API는 의도적으로 최소화:
- list_org_configs()
- load_org_config(selector)
- detect_org_from_pdf(pdf_path)

추출기/작성기는 설정 저장 방식에 신경 쓰지 않아도 된다.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple

import re

import yaml

from ..util.unicode_name import decode_hashu, normalize_identifier


DEFAULT_CONFIG_DIR_CANDIDATES = [
    # 레거시
    "config",
    # 향후/옵션
    "configs/orgs",
]


@dataclass(frozen=True)
class SurveySideConfig:
    """조사 유형(전국/지방)별 정규화된 설정."""

    # 목차 섹션
    has_toc: bool
    toc_page: Optional[int]
    toc_right_half_crop: bool
    page_offset: Optional[int]
    section_filter: Optional[str]
    fixed_items: Dict[str, int]

    # 기타 페이지
    metadata_page: Optional[int]
    region_page: Optional[int]
    region_with_metadata: bool
    fixed_regions: List[str]

    # 추출 동작
    pages_per_item: int
    include_groups: List[str]
    exclude_groups: List[str]

    # 선택 전략
    page_walk: Optional[Dict[str, Any]]
    initial_scan: Optional[Dict[str, Any]]
    page_skip: Optional[Dict[str, Any]]
    type_rules: Optional[Dict[str, Any]]


@dataclass(frozen=True)
class OrgConfig:
    """기관 설정(정규화된 형태)."""

    id: str
    name: str
    aliases: List[str]
    supported_items: List[str]

    national: SurveySideConfig
    local: SurveySideConfig

    # 정규화
    term_mapping: Dict[str, Dict[str, str]]
    region_merges: List[Dict[str, Any]]
    excel_patches: Dict[str, Any]

    raw: Dict[str, Any]
    source_path: Path


def _read_yaml(path: Path) -> Dict[str, Any]:
    with open(path, "r", encoding="utf-8") as f:
        data = yaml.safe_load(f) or {}
    if not isinstance(data, dict):
        raise ValueError(f"YAML root must be a mapping: {path}")
    return data


def _norm_supported_items(raw: Dict[str, Any]) -> List[str]:
    items = raw.get("지원항목") or raw.get("supported_items") or raw.get("items")
    if items is None:
        # 안전 기본값: 모두 지원
        return ["PSR", "GE", "ISSUE"]
    if isinstance(items, str):
        return [items]
    if isinstance(items, list):
        out = []
        for x in items:
            if x is None:
                continue
            s = str(x).strip().upper()
            if not s:
                continue
            out.append(s)
        return out
    return ["PSR", "GE", "ISSUE"]


def _get_side(raw: Dict[str, Any], side_key: str) -> Dict[str, Any]:
    """레거시 형태를 포함해 전국/지방 설정 dict를 반환."""
    # 레거시 구조: 전국:, 지방:
    side = raw.get(side_key) or {}
    return side if isinstance(side, dict) else {}


def _extract_toc_side(raw: Dict[str, Any], side_key: str) -> Dict[str, Any]:
    """전국/지방에 대한 목차 설정을 반환(분리/공통 모두 지원).
    - 목차: {전국:{...}, 지방:{...}}
    - 목차: {...} (공통 적용)
    """
    toc = raw.get("목차") or {}
    if not isinstance(toc, dict):
        return {}

    if side_key in toc:
        side = toc.get(side_key) or {}
        return side if isinstance(side, dict) else {}

    # 분리 설정이 없으면 공통 적용
    return toc


def _extract_page_block(raw: Dict[str, Any], block_key: str, side_key: str) -> Dict[str, Any]:
    """조사설계/권역설명 페이지 설정을 분리/공통 모두 지원해 추출."""
    block = raw.get(block_key) or {}
    if not isinstance(block, dict):
        return {}
    if side_key in block:
        side = block.get(side_key) or {}
        return side if isinstance(side, dict) else {}
    return block


def _parse_int(v: Any) -> Optional[int]:
    if v is None:
        return None
    try:
        return int(v)
    except Exception:
        return None


def _normalize_side_config(raw: Dict[str, Any], side_key: str) -> SurveySideConfig:
    toc_side = _extract_toc_side(raw, side_key)

    has_toc = bool(toc_side.get("있음", toc_side.get("has_toc", True)))

    toc_page = _parse_int(toc_side.get("페이지", toc_side.get("page")))
    toc_right_half_crop = bool(
        toc_side.get("right_half_crop")
        or toc_side.get("toc_right_half_crop")
        or toc_side.get("오른쪽크롭")
    )

    # 중요: page_offset을 추측하지 않고, 명시가 없으면 None으로 둬서 검증 단계에서 판단하게 함
    page_offset = toc_side.get("페이지오프셋", toc_side.get("page_offset"))
    page_offset = _parse_int(page_offset)

    section_filter = toc_side.get("섹션필터") or toc_side.get("section_filter")

    fixed_items = toc_side.get("고정항목") or toc_side.get("fixed_items") or {}
    if not isinstance(fixed_items, dict):
        fixed_items = {}
    # 페이지 값을 int로 정규화
    fixed_items_norm: Dict[str, int] = {}
    for k, v in fixed_items.items():
        p = _parse_int(v)
        if k is None or p is None:
            continue
        fixed_items_norm[str(k).strip()] = p

    metadata_block = _extract_page_block(raw, "조사설계", side_key)
    metadata_page = _parse_int(metadata_block.get("페이지", metadata_block.get("page")))

    region_block = _extract_page_block(raw, "권역설명", side_key)
    region_page = _parse_int(region_block.get("페이지", region_block.get("page")))
    region_with_metadata = bool(region_block.get("조사설계와함께", region_block.get("with_metadata", False)))
    fixed_regions = region_block.get("고정권역") or region_block.get("권역목록") or []
    if isinstance(fixed_regions, str):
        fixed_regions = [fixed_regions]
    fixed_regions = [str(x).strip() for x in fixed_regions if str(x).strip()] if isinstance(fixed_regions, list) else []

    side = _get_side(raw, side_key)
    pages_per_item = _parse_int(side.get("페이지수", side.get("pages_per_item"))) or 1

    include_groups = side.get("포함", side.get("include")) or []
    exclude_groups = side.get("제외", side.get("exclude")) or []
    include_groups = [str(x) for x in include_groups] if isinstance(include_groups, list) else []
    exclude_groups = [str(x) for x in exclude_groups] if isinstance(exclude_groups, list) else []

    page_walk = (raw.get("페이지순회") or {}).get(side_key) if isinstance(raw.get("페이지순회"), dict) else raw.get("페이지순회")
    if page_walk is not None and not isinstance(page_walk, dict):
        page_walk = None

    initial_scan = raw.get("초기스캔")
    if initial_scan is not None and not isinstance(initial_scan, dict):
        initial_scan = None

    page_skip = raw.get("페이지스킵")
    if page_skip is not None and not isinstance(page_skip, dict):
        page_skip = None

    type_rules = raw.get("유형판단")
    if type_rules is not None and not isinstance(type_rules, dict):
        type_rules = None

    return SurveySideConfig(
        has_toc=has_toc,
        toc_page=toc_page,
        toc_right_half_crop=toc_right_half_crop,
        page_offset=page_offset,
        section_filter=section_filter,
        fixed_items=fixed_items_norm,
        metadata_page=metadata_page,
        region_page=region_page,
        region_with_metadata=region_with_metadata,
        fixed_regions=fixed_regions,
        pages_per_item=pages_per_item,
        include_groups=include_groups,
        exclude_groups=exclude_groups,
        page_walk=page_walk,
        initial_scan=initial_scan,
        page_skip=page_skip,
        type_rules=type_rules,
    )


def _normalize_term_mapping(raw: Dict[str, Any]) -> Dict[str, Dict[str, str]]:
    mapping = raw.get("용어매핑") or raw.get("term_mapping") or {}
    if not isinstance(mapping, dict):
        return {}

    # 허용 형태:
    # - {분류:{...}, 응답:{...}}
    # - 평평한 dict
    if any(isinstance(v, dict) for v in mapping.values()):
        out: Dict[str, Dict[str, str]] = {}
        for k, v in mapping.items():
            if isinstance(v, dict):
                out[str(k)] = {str(kk): (None if vv is None else str(vv)) for kk, vv in v.items()}
        return out

    # 평평한 dict
    flat = {str(k): (None if v is None else str(v)) for k, v in mapping.items()}
    return {"all": flat}


def _org_id_from(raw: Dict[str, Any], path: Path) -> Tuple[str, str, List[str]]:
    name = raw.get("기관명") or raw.get("name") or decode_hashu(path.stem)
    name = str(name).strip() if name is not None else decode_hashu(path.stem)

    # id 우선순위: YAML 지정값 > 파일명(디코딩)
    explicit_id = raw.get("id") or raw.get("org_id")
    if explicit_id:
        org_id = str(explicit_id).strip()
    else:
        # 파일명(디코딩)에서 파생: 가능하면 ASCII, 아니면 기관명 사용
        org_id = decode_hashu(path.stem).strip()

    aliases = raw.get("별칭") or raw.get("aliases") or []
    if isinstance(aliases, str):
        aliases = [aliases]
    if not isinstance(aliases, list):
        aliases = []
    aliases = [str(a).strip() for a in aliases if a is not None and str(a).strip()]

    # 매칭 편의를 위해 파일명(디코딩)과 기관명을 별칭에 포함
    stem_dec = decode_hashu(path.stem).strip()
    base_aliases = [stem_dec, name]

    # 중복 제거(대소문자 무시)
    seen = set()
    merged: List[str] = []
    for a in base_aliases + aliases:
        key = normalize_identifier(a)
        if not key or key in seen:
            continue
        seen.add(key)
        merged.append(a)

    return org_id, name, merged


def _normalize_excel_patches(raw: Dict[str, Any]) -> Dict[str, Any]:
    patches = raw.get("엑셀패치") or raw.get("excel_patches") or {}
    return patches if isinstance(patches, dict) else {}


def _normalize_region_merges(raw: Dict[str, Any]) -> List[Dict[str, Any]]:
    merges = raw.get("지역합산") or raw.get("region_merges") or []
    if not isinstance(merges, list):
        return []
    out: List[Dict[str, Any]] = []
    for m in merges:
        if isinstance(m, dict):
            out.append(m)
    return out


def load_org_config(selector: str, project_root: Path | None = None) -> OrgConfig:
    """selector로 단일 기관 설정을 로드.

    selector 허용 형태:
    - 기관 id(yaml의 'id')
    - 기관명
    - yaml 파일명(디코딩)
    - 별칭
    """
    project_root = project_root or Path(__file__).resolve().parents[2]

    candidates = list_org_configs(project_root)
    if not candidates:
        raise FileNotFoundError("No org YAML configs found")

    sel_key = normalize_identifier(selector)

    best: Optional[Tuple[Path, Dict[str, Any], str]] = None

    for path in candidates:
        raw = _read_yaml(path)
        org_id, name, aliases = _org_id_from(raw, path)

        keys = [org_id, name] + aliases
        norm_keys = [normalize_identifier(k) for k in keys]

        if sel_key in norm_keys:
            best = (path, raw, org_id)
            break

    if best is None:
        # 부분 일치 허용: selector가 어느 키에든 포함되면 매칭
        for path in candidates:
            raw = _read_yaml(path)
            org_id, name, aliases = _org_id_from(raw, path)
            keys = [org_id, name] + aliases
            norm_keys = [normalize_identifier(k) for k in keys]
            if any(sel_key and sel_key in k for k in norm_keys):
                best = (path, raw, org_id)
                break

    if best is None:
        avail = ", ".join(sorted({decode_hashu(p.stem) for p in candidates}))
        raise ValueError(f"Unknown org '{selector}'. Available: {avail}")

    path, raw, org_id = best
    org_id, name, aliases = _org_id_from(raw, path)

    supported_items = _norm_supported_items(raw)

    national = _normalize_side_config(raw, "전국")
    local = _normalize_side_config(raw, "지방")

    term_mapping = _normalize_term_mapping(raw)
    region_merges = _normalize_region_merges(raw)
    excel_patches = _normalize_excel_patches(raw)

    return OrgConfig(
        id=org_id,
        name=name,
        aliases=aliases,
        supported_items=supported_items,
        national=national,
        local=local,
        term_mapping=term_mapping,
        region_merges=region_merges,
        excel_patches=excel_patches,
        raw=raw,
        source_path=path,
    )


def list_org_configs(project_root: Path | None = None) -> List[Path]:
    project_root = project_root or Path(__file__).resolve().parents[2]

    paths: List[Path] = []
    for rel in DEFAULT_CONFIG_DIR_CANDIDATES:
        d = project_root / rel
        if not d.exists():
            continue
        paths.extend(sorted(d.glob("*.yaml")))

    # 실제 경로 기준 중복 제거
    uniq = []
    seen = set()
    for p in paths:
        rp = str(p.resolve())
        if rp in seen:
            continue
        seen.add(rp)
        uniq.append(p)

    return uniq


def get_available_orgs(project_root: Path | None = None) -> List[str]:
    """사용자용 기관 목록 반환(디코딩된 파일명)."""
    return [decode_hashu(p.stem) for p in list_org_configs(project_root)]


def detect_org_from_pdf(pdf_path: str, project_root: Path | None = None) -> Optional[str]:
    """기관 selector를 추정.

    1) 빠른 경로: PDF 파일명 부분 문자열 매칭
    2) 폴백(파일명에 기관이 없을 때): 앞쪽 페이지 텍스트에서 기관 토큰 스캔
    """
    project_root = project_root or Path(__file__).resolve().parents[2]

    pdf_stem = decode_hashu(Path(pdf_path).stem)
    filename_key = normalize_identifier(pdf_stem)

    for path in list_org_configs(project_root):
        raw = _read_yaml(path)
        org_id, name, aliases = _org_id_from(raw, path)
        for token in [org_id, name] + aliases:
            tkey = normalize_identifier(token)
            if tkey and tkey in filename_key:
                return org_id

    # 폴백: 내용 기반 매칭(파일명이 'MBC_통계표...'처럼 기관이 없지만 표지에 기관 표기가 있는 경우 등)
    try:
        import pdfplumber  # 이미 핵심 의존성(표 크롭에서 사용)

        text_chunks: List[str] = []
        with pdfplumber.open(pdf_path) as pdf:
            for p in pdf.pages[:3]:
                t = p.extract_text() or ""
                if t:
                    text_chunks.append(t)
        content_key = re.sub(r"\s+", "", normalize_identifier("\n".join(text_chunks)))
        if content_key:
            scored: List[Tuple[int, str]] = []
            for path in list_org_configs(project_root):
                raw = _read_yaml(path)
                org_id, name, aliases = _org_id_from(raw, path)
                tokens = [org_id, name] + aliases
                # 공백 제거된 키로 매칭
                score = 0
                for tok in tokens:
                    tkey = re.sub(r"\s+", "", normalize_identifier(tok))
                    if tkey and len(tkey) >= 2 and tkey in content_key:
                        score += 1
                if score > 0:
                    scored.append((score, org_id))

            if scored:
                scored.sort(key=lambda x: x[0], reverse=True)
                best_score = scored[0][0]
                best_ids = [oid for sc, oid in scored if sc == best_score]
                if len(best_ids) == 1:
                    return best_ids[0]
    except Exception:
        # 실패하면 수동 선택으로 폴백
        pass

    return None


def validate_org_config(org: OrgConfig) -> List[str]:
    """사람이 읽기 쉬운 문제 목록 반환(유효하면 빈 리스트)."""
    problems: List[str] = []

    def _check_side(side_name: str, s: SurveySideConfig):
        # 목차 모드 검사
        if s.has_toc:
            if s.toc_page is None:
                problems.append(f"[{org.name}] {side_name}: 목차 있음인데 '페이지'가 없습니다")
            if s.page_offset is None:
                problems.append(f"[{org.name}] {side_name}: 목차 있음인데 '페이지오프셋'이 없습니다")
        else:
            # fixed_items가 있을 때 page_offset이 필요한가? 유형이 둘로 나뉘므로 명시 필요:
            # - 목차 번호(오프셋 필요)
            # - 실제 PDF 페이지(오프셋 0)
            # 숨은 +1 버그를 막으려면 fixed_items 사용 시 page_offset을 반드시 요구
            if s.fixed_items and s.page_offset is None:
                problems.append(
                    f"[{org.name}] {side_name}: 고정항목 사용인데 '페이지오프셋'이 없습니다 (권장: 실제 PDF면 0)"
                )

    _check_side("전국", org.national)
    _check_side("지방", org.local)

    return problems
