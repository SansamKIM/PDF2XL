"""PDF -> PIL 이미지 렌더링(디스크 캐시 포함).

개발 중 가장 시간을 잡아먹는 작업이 PDF 페이지 반복 렌더링이라 (pdf, dpi, page)별로 캐시해
재실행 시 poppler 호출을 건너뛴다.

캐시 구조(시스템 캐시):
  <system-cache>/<pdf_key>/_cache/pages_dpi300/p0001.png
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import List, Optional, Sequence, Tuple

from pdf2image import convert_from_path

# -----------------------------------------------------------------------------
# Windows: Poppler 실행 시 콘솔 창이 번쩍 뜨는 문제 방지.
#
# pdf2image는 subprocess.Popen으로 poppler 바이너리(pdfinfo/pdftoppm)를 호출하며,
# GUI 앱(PyInstaller --noconsole/pythonw)에서는 플래그를 주지 않으면 콘솔이 잠깐 열린다.
#
# Windows에서 pdf2image.pdf2image.Popen을 패치해 모든 poppler 호출을
# CREATE_NO_WINDOW/숨김 startupinfo로 실행한다. 최선 시도이며 파이프라인에 영향 없어야 한다.
# -----------------------------------------------------------------------------
try:
    import sys
    if sys.platform == "win32":
        import subprocess
        import pdf2image.pdf2image as _p2i

        _orig_popen = getattr(_p2i, "Popen", None)
        if _orig_popen and getattr(_orig_popen, "__name__", "") != "_popen_no_window":
            def _popen_no_window(*args, **kwargs):
                # 콘솔 창 숨기기(Windows 전용)
                try:
                    si_cls = getattr(subprocess, "STARTUPINFO", None)
                    if ("startupinfo" not in kwargs) or (kwargs.get("startupinfo") is None):
                        if si_cls is not None:
                            si = si_cls()
                            si.dwFlags |= getattr(subprocess, "STARTF_USESHOWWINDOW", 0)
                            kwargs["startupinfo"] = si

                    cnw = getattr(subprocess, "CREATE_NO_WINDOW", 0)
                    if cnw:
                        kwargs["creationflags"] = kwargs.get("creationflags", 0) | cnw
                except Exception:
                    # 외형 설정 때문에 실패하지 않도록 무시
                    pass

                return _orig_popen(*args, **kwargs)

            _p2i.Popen = _popen_no_window
except Exception:
    # 패치 시도 때문에 import가 실패하지 않도록 무시
    pass
from PIL import Image

from ..util.cache import pages_cache_dir


@dataclass
class RenderOptions:
    """PDF 페이지 렌더링 시 사용할 전처리/해상도 옵션."""
    dpi: int = 300
    preprocess: bool = True
    margin_percent: int = 3
    contrast: float = 1.5
    sharpness: float = 2.0


def _crop_margins(image: Image.Image, margin_percent: int = 3) -> Image.Image:
    width, height = image.size
    left = int(width * margin_percent / 100)
    top = int(height * margin_percent / 100)
    right = int(width * (100 - margin_percent) / 100)
    bottom = int(height * (100 - margin_percent) / 100)
    return image.crop((left, top, right, bottom))


def _enhance_image(image: Image.Image, contrast: float = 1.5, sharpness: float = 2.0) -> Image.Image:
    from PIL import ImageEnhance

    im = image
    if contrast and contrast != 1.0:
        im = ImageEnhance.Contrast(im).enhance(contrast)
    if sharpness and sharpness != 1.0:
        im = ImageEnhance.Sharpness(im).enhance(sharpness)
    return im


class PDFPageRenderer:
    """PDF 페이지를 PIL 이미지로 렌더링하고 디스크 캐시에 저장하는 래퍼."""
    def __init__(
        self,
        pdf_path: str,
        poppler_path: str | None,
        cache_dir: str | Path,
        options: RenderOptions | None = None,
    ) -> None:
        self.pdf_path = str(pdf_path)
        self.poppler_path = poppler_path
        self.cache_dir = Path(cache_dir)
        self.options = options or RenderOptions()

        self.cache_dir.mkdir(parents=True, exist_ok=True)

    def _cache_path(self, page: int) -> Path:
        # 정렬이 안정되도록 0으로 패딩
        return self.cache_dir / f"p{page:04d}.png"

    def _load_cached(self, page: int) -> Optional[Image.Image]:
        p = self._cache_path(page)
        if not p.exists():
            return None
        try:
            return Image.open(p)
        except Exception:
            return None

    def _save_cached(self, page: int, image: Image.Image) -> None:
        p = self._cache_path(page)
        try:
            image.save(p, format="PNG")
        except Exception:
            # 캐시 저장 실패가 파이프라인을 깨면 안 됨
            pass

    def _preprocess(self, image: Image.Image) -> Image.Image:
        if not self.options.preprocess:
            return image
        im = _crop_margins(image, margin_percent=self.options.margin_percent)
        im = _enhance_image(im, contrast=self.options.contrast, sharpness=self.options.sharpness)
        return im

    def render_page(self, page: int) -> Optional[Image.Image]:
        """단일 PDF 페이지를 렌더링하고 캐시를 활용."""
        cached = self._load_cached(page)
        if cached is not None:
            return cached

        imgs = convert_from_path(
            self.pdf_path,
            dpi=self.options.dpi,
            poppler_path=self.poppler_path,
            first_page=page,
            last_page=page,
        )
        if not imgs:
            return None
        im = self._preprocess(imgs[0])
        self._save_cached(page, im)
        return im

    def render_pages(self, pages: Sequence[int]) -> List[Image.Image]:
        """임의 페이지 목록을 렌더링.

        호출자 순서를 유지하되, 내부적으로는 연속된 누락 페이지를 묶어
        poppler 시작 비용을 줄입니다.
        """

        if not pages:
            return []

        pages_int = [int(p) for p in pages]
        # 정렬된 뷰에서 배치 렌더링 후 호출 순서로 재정렬
        sorted_pages = sorted(set(pages_int))

        # 캐시된 페이지를 우선 로드
        rendered: dict[int, Image.Image] = {}
        missing: List[int] = []
        for p in sorted_pages:
            im = self._load_cached(p)
            if im is None:
                missing.append(p)
            else:
                rendered[p] = im

        # 누락된 페이지를 연속 구간별로 배치 렌더링
        if missing:
            ranges: List[Tuple[int, int]] = []
            start = prev = missing[0]
            for p in missing[1:]:
                if p == prev + 1:
                    prev = p
                    continue
                ranges.append((start, prev))
                start = prev = p
            ranges.append((start, prev))

            for start, end in ranges:
                imgs = convert_from_path(
                    self.pdf_path,
                    dpi=self.options.dpi,
                    poppler_path=self.poppler_path,
                    first_page=start,
                    last_page=end,
                )
                # 반환된 이미지를 각 페이지에 매핑(길이 안전하게 처리)
                expected_pages = list(range(start, end + 1))
                for idx, p in enumerate(expected_pages):
                    if idx >= len(imgs):
                        break
                    im = self._preprocess(imgs[idx])
                    self._save_cached(p, im)
                    rendered[p] = im

        # 렌더링 실패 페이지를 건너뛰고 호출자 순서로 반환
        out: List[Image.Image] = []
        for p in pages_int:
            im = rendered.get(p)
            if im is not None:
                out.append(im)
        return out

    def render_range(self, first_page: int, last_page: int) -> List[Image.Image]:
        """지정 구간(first~last)을 렌더링."""
        if first_page > last_page:
            return []
        return self.render_pages(list(range(first_page, last_page + 1)))


def default_cache_dir(pdf_path: str | Path, dpi: int) -> Path:
    """렌더된 페이지 PNG를 위한 기본 캐시 디렉터리를 반환."""

    return pages_cache_dir(pdf_path, dpi)
