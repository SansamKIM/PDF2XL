"""캐시 관리 유틸리티.

목표
----
모든 캐시는 사용자가 지정한 출력 폴더가 아닌 OS 캐시 디렉터리에 저장한다.

캐시 위치(OS가 관리하는 영역):
  - macOS  : ~/Library/Caches/<app>/...
  - Windows: %LOCALAPPDATA%\\<app>\\Cache\\...
  - Linux  : ~/.cache/<app>/...

캐시 구조
---------
기존 디스크 구조를 유지하되 루트만 OS 캐시 디렉터리로 옮긴다.

  <cache_root>/<pdf_key>/_cache/
      pages_dpi300/p0001.png
      _toc.json
      _initial_scan.json
      _page_walk_cache.json
      ...

<pdf_key>는 읽기 쉬운 PDF 파일명에 짧은 해시를 붙여 충돌과 버전 혼선을 막는다.
"""

from __future__ import annotations

import hashlib
import os
import shutil
import time
from pathlib import Path
from typing import Iterable, List, Optional, Tuple

try:
    from platformdirs import user_cache_dir
except Exception:  # pragma: no cover
    # platformdirs가 없어도 동작하도록 최소한의 폴백 제공
    def user_cache_dir(appname: str) -> str:  # type: ignore
        # 리눅스 스타일의 기본 경로 사용
        return str(Path.home() / ".cache" / appname)

from .. import __app_name__
from .unicode_name import decode_hashu


# 캐시 최대 용량: 4 GiB(바이너리 기준)
DEFAULT_MAX_CACHE_BYTES = 4 * 1024 * 1024 * 1024


def _sanitize_folder_name(name: str, max_len: int = 120) -> str:
    """경로 세그먼트로 안전하게 쓸 수 있는 문자열로 정리."""

    s = (name or "").strip()
    # 경로 이동/구분자를 제거
    s = s.replace("/", "_").replace("\\", "_")
    if os.altsep:
        s = s.replace(os.altsep, "_")

    # Windows에서 금지된 문자(다른 OS에서도 안전)
    for ch in (":", "*", "?", '"', "<", ">", "|"):
        s = s.replace(ch, "_")

    # 공백을 하나로 축소
    s = " ".join(s.split())
    if not s:
        s = "pdf"
    return s[:max_len]


def cache_root() -> Path:
    """캐시 루트 폴더 반환(없으면 생성)."""

    # 고급 사용자/디버깅용 수동 지정 환경변수
    env = os.environ.get("PDF2XLAI_CACHE_DIR")
    if env:
        root = Path(env).expanduser()
    else:
        # 캐시를 앱 소유 폴더 아래로 모음
        # (platformdirs가 OS별 기본 캐시 경로는 이미 선택함)
        root = Path(user_cache_dir(__app_name__)) / "pdf_cache"

    root.mkdir(parents=True, exist_ok=True)
    return root


def _pdf_identity_hash(pdf_path: Path) -> str:
    """캐시를 구분하기 위한 짧은 식별자.

    실경로+크기+mtime을 섞어 동일 이름의 PDF라도 폴더/수정 시점을 구분한다.
    """

    try:
        p = pdf_path.resolve()
    except Exception:
        p = pdf_path

    try:
        st = p.stat()
        payload = f"{p}|{st.st_size}|{int(st.st_mtime)}"
    except Exception:
        payload = str(p)

    return hashlib.sha1(payload.encode("utf-8", "ignore")).hexdigest()[:10]


def pdf_cache_key(pdf_path: str | Path) -> str:
    p = Path(pdf_path)
    display = decode_hashu(p.stem) if p.stem else "pdf"
    display = _sanitize_folder_name(display)
    h = _pdf_identity_hash(p)
    return f"{display}_{h}"


def _touch(path: Path) -> None:
    """가능하면 mtime을 지금 시각으로 갱신."""

    try:
        now = time.time()
        os.utime(path, (now, now))
    except Exception:
        pass


def pdf_cache_base_dir(pdf_path: str | Path) -> Path:
    """PDF의 캐시 베이스 디렉터리 반환(없으면 생성)."""

    base = cache_root() / pdf_cache_key(pdf_path)
    base.mkdir(parents=True, exist_ok=True)
    _touch(base)
    return base


def run_cache_dir(pdf_path: str | Path) -> Path:
    """PDF용 '_cache' 디렉터리 반환(생성 포함)."""

    d = pdf_cache_base_dir(pdf_path) / "_cache"
    d.mkdir(parents=True, exist_ok=True)
    _touch(d)
    return d


def pages_cache_dir(pdf_path: str | Path, dpi: int) -> Path:
    """PDF와 DPI별 렌더 이미지 캐시 폴더 반환."""

    d = run_cache_dir(pdf_path) / f"pages_dpi{int(dpi)}"
    d.mkdir(parents=True, exist_ok=True)
    _touch(d)
    return d


def json_output_dir(pdf_path: str | Path) -> Path:
    """추출된 JSON 출력(메타데이터/권역/항목)용 디렉터리.

    정책: 사용자가 선택한 출력 폴더에는 최종 Excel 파일만 있어야 한다.
    따라서 중간 산출물(JSON)은 모두 OS 캐시 디렉터리 아래에 저장한다.
    """

    d = pdf_cache_base_dir(pdf_path) / "json"
    d.mkdir(parents=True, exist_ok=True)
    _touch(d)
    return d


def _dir_size_bytes(path: Path) -> int:
    total = 0
    try:
        for root, _, files in os.walk(path):
            for fn in files:
                fp = os.path.join(root, fn)
                try:
                    total += os.path.getsize(fp)
                except OSError:
                    continue
    except Exception:
        return 0
    return total


def cache_total_size_bytes() -> int:
    """캐시 전체 크기(바이트)."""

    root = cache_root()
    return _dir_size_bytes(root) if root.exists() else 0


def clear_all_cache() -> None:
    """모든 캐시 삭제(최대한)."""

    root = cache_root()
    try:
        if root.exists():
            shutil.rmtree(root)
    finally:
        # 이후 로직 단순화를 위해 폴더는 다시 생성
        root.mkdir(parents=True, exist_ok=True)


def enforce_cache_quota(
    max_bytes: int = DEFAULT_MAX_CACHE_BYTES,
    *,
    exclude: Optional[Iterable[Path]] = None,
) -> dict:
    """캐시 총량을 max_bytes 이하로 유지.

    전략: PDF별 캐시 디렉터리를 mtime 기준 오래된 것부터 통째로 삭제.

    로그용 소형 통계 dict를 반환한다.
    """

    root = cache_root()
    if not root.exists():
        return {"total_before": 0, "total_after": 0, "deleted": []}

    exclude_set = set()
    for p in exclude or []:
        try:
            exclude_set.add(Path(p).resolve())
        except Exception:
            continue

    entries: List[Tuple[Path, int, float]] = []
    for child in root.iterdir():
        if not child.is_dir():
            continue
        try:
            size = _dir_size_bytes(child)
            mtime = child.stat().st_mtime
            entries.append((child, size, mtime))
        except Exception:
            continue

    total_before = sum(sz for _, sz, _ in entries)
    if total_before <= max_bytes:
        return {"total_before": total_before, "total_after": total_before, "deleted": []}

    # 오래된 순으로 정렬
    entries.sort(key=lambda t: t[2])

    deleted: List[str] = []
    total_after = total_before

    for path, size, _ in entries:
        if total_after <= max_bytes:
            break

        try:
            if path.resolve() in exclude_set:
                continue
        except Exception:
            # resolve에 실패하면 일단 삭제 시도
            pass

        try:
            shutil.rmtree(path)
            deleted.append(path.name)
            total_after -= size
        except Exception:
            # 삭제에 실패해도 계속 진행(최대한)
            continue

    if total_after < 0:
        total_after = 0

    return {"total_before": total_before, "total_after": total_after, "deleted": deleted}
