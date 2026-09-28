"""PDF 렌더 이미지에서 표 영역을 자동 크롭.

동기
- 촘촘한 교차표(특히 PSR)는 페이지 전체를 이미지로 주면 LLM 숫자 인식 오류가 잦음.
- pdfplumber는 텍스트가 흐트러져도( CID 폰트/임베드 이미지) 룰라인으로 표 bbox를 잡을 수 있음.

전략: 페이지마다 대표 표를 pdfplumber로 찾고, 이미 렌더된 PIL 이미지를 그 bbox로 크롭해 가독성을 높이고 프롬프트를 짧게 유지.

설계
- 일반적: 기관별 하드코딩 없음.
- 안전: 탐지 실패 시 원본 이미지 반환.
- 캐시: 페이지별 bbox 캐싱.

좌표 메모
- pdfplumber bbox는 (x0, top, x1, bottom) 형태로 좌상단 기준 PDF 좌표계.
- pdf2image 렌더는 dpi 기준(PDF 포인트 72/inch => px = pt * dpi/72).
- 렌더러가 대칭 여백을 자를 수 있어 margin_percent>0이면 이를 보정.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, Iterable, List, Optional, Sequence, Tuple

import pdfplumber
from PIL import Image


@dataclass
class TableCropOptions:
    padding_px: int = 10
    prefer_largest: bool = True
    min_area_ratio: float = 0.06  # 각주 등 매우 작은 "표"는 무시


class PDFTableCropper:
    def __init__(
        self,
        pdf_path: str,
        dpi: int,
        margin_percent: int = 0,
        options: Optional[TableCropOptions] = None,
    ) -> None:
        self.pdf_path = str(pdf_path)
        self.dpi = int(dpi)
        self.margin_percent = int(margin_percent or 0)
        self.options = options or TableCropOptions()

        # page_num(1부터 시작) -> (bbox, page_w_pt, page_h_pt)
        self._bbox_cache: Dict[int, Tuple[Tuple[float, float, float, float], float, float]] = {}

    @property
    def _scale(self) -> float:
        return float(self.dpi) / 72.0

    def _margin_offsets_px(self, cropped_image: Image.Image) -> Tuple[float, float]:
        """대칭 여백 크롭으로 잘려나간 (left_px, top_px) 복원."""

        m = float(self.margin_percent or 0) / 100.0
        if m <= 0:
            return 0.0, 0.0

        # 크롭 결과 = 원본 * (1 - 2m)
        factor = 1.0 / max(1e-6, (1.0 - 2.0 * m))
        orig_w = cropped_image.width * factor
        orig_h = cropped_image.height * factor
        return orig_w * m, orig_h * m

    def _choose_bbox(self, tables) -> Optional[Tuple[float, float, float, float]]:
        if not tables:
            return None

        # 가장 큰 영역의 표를 우선 선택
        best = None
        best_area = -1.0
        for t in tables:
            try:
                x0, top, x1, bottom = t.bbox
                area = max(0.0, x1 - x0) * max(0.0, bottom - top)
            except Exception:
                continue
            if area > best_area:
                best_area = area
                best = (x0, top, x1, bottom)

        return best

    def _bbox_from_chars(self, page) -> Optional[Tuple[float, float, float, float]]:
        """텍스트 문자 기반 폴백 bbox.

        일부 PDF는 표 테두리를 '-' 문자 반복으로 그려 find_tables() 결과가 없을 수 있다.
        page.chars의 합집합 bbox를 사용하면 여백을 줄이고 LLM 가독성을 높이는 안전한 크롭을 얻을 수 있다.
        """

        try:
            chars = getattr(page, "chars", None) or []
            if not chars:
                return None

            x0 = min(float(c.get("x0", 0.0)) for c in chars)
            x1 = max(float(c.get("x1", 0.0)) for c in chars)
            top = min(float(c.get("top", 0.0)) for c in chars)
            bottom = max(float(c.get("bottom", 0.0)) for c in chars)

            # 간단한 안전검사
            if x1 <= x0 + 1 or bottom <= top + 1:
                return None
            return (x0, top, x1, bottom)
        except Exception:
            return None

    def get_bbox(self, page_num: int) -> Optional[Tuple[float, float, float, float]]:
        """1-based 페이지 번호에 대한 표(메인 테이블) bbox 캐시를 반환."""

        p = int(page_num)
        if p in self._bbox_cache:
            return self._bbox_cache[p][0]

        try:
            with pdfplumber.open(self.pdf_path) as pdf:
                idx = p - 1
                if idx < 0 or idx >= len(pdf.pages):
                    return None
                page = pdf.pages[idx]
                page_w, page_h = float(page.width), float(page.height)
                tables = page.find_tables()  # 룰라인 기반 탐지(CID 폰트에서도 비교적 견고)
                bbox = self._choose_bbox(tables)
                if bbox is None:
                    # 텍스트로 그린 표(대시 테두리 등) 폴백
                    bbox = self._bbox_from_chars(page)
                if bbox is None:
                    return None

                # 너무 작은 bbox 필터링
                x0, top, x1, bottom = bbox
                area = max(0.0, x1 - x0) * max(0.0, bottom - top)
                if page_w * page_h > 0:
                    ratio = area / (page_w * page_h)
                    if ratio < float(self.options.min_area_ratio or 0):
                        return None

                self._bbox_cache[p] = (bbox, page_w, page_h)
                return bbox
        except Exception:
            return None

    def warmup(self, pages: Iterable[int]) -> None:
        """지정 페이지의 bbox를 미리 계산(최대한)."""
        pages_int = sorted({int(p) for p in pages if p is not None})
        if not pages_int:
            return

        # 캐시에 없는 페이지만 계산
        missing = [p for p in pages_int if p not in self._bbox_cache]
        if not missing:
            return

        try:
            with pdfplumber.open(self.pdf_path) as pdf:
                for p in missing:
                    idx = p - 1
                    if idx < 0 or idx >= len(pdf.pages):
                        continue
                    page = pdf.pages[idx]
                    page_w, page_h = float(page.width), float(page.height)
                    tables = page.find_tables()
                    bbox = self._choose_bbox(tables)
                    if bbox is None:
                        bbox = self._bbox_from_chars(page)
                    if bbox is None:
                        continue
                    x0, top, x1, bottom = bbox
                    area = max(0.0, x1 - x0) * max(0.0, bottom - top)
                    if page_w * page_h > 0:
                        ratio = area / (page_w * page_h)
                        if ratio < float(self.options.min_area_ratio or 0):
                            continue
                    self._bbox_cache[p] = (bbox, page_w, page_h)
        except Exception:
            # 최선 시도 후 실패는 무시
            return

    def crop_image(self, image: Image.Image, page_num: int) -> Image.Image:
        """렌더된(여백 크롭될 수 있는) 페이지 이미지를 감지한 표 영역으로 크롭."""

        bbox = self.get_bbox(page_num)
        if bbox is None:
            return image

        try:
            x0, top, x1, bottom = bbox
            scale = self._scale
            x0_px = x0 * scale
            x1_px = x1 * scale
            y0_px = top * scale
            y1_px = bottom * scale

            # 렌더러의 대칭 여백 크롭 보정
            left_off, top_off = self._margin_offsets_px(image)
            x0_px -= left_off
            x1_px -= left_off
            y0_px -= top_off
            y1_px -= top_off

            pad = int(self.options.padding_px or 0)
            x0_px = max(0, int(x0_px) - pad)
            y0_px = max(0, int(y0_px) - pad)
            x1_px = min(image.width, int(x1_px) + pad)
            y1_px = min(image.height, int(y1_px) + pad)

            # 안전 검사
            if x1_px <= x0_px + 10 or y1_px <= y0_px + 10:
                return image

            return image.crop((x0_px, y0_px, x1_px, y1_px))
        except Exception:
            return image

    def crop_images(self, images: Sequence[Image.Image], pages: Sequence[int]) -> List[Image.Image]:
        """페이지 순서에 맞춰 이미지들을 크롭."""
        if not images or not pages:
            return list(images)

        pages_int = [int(p) for p in pages]
        self.warmup(pages_int)

        out: List[Image.Image] = []
        for im, p in zip(images, pages_int):
            out.append(self.crop_image(im, p))
        # 페이지 수보다 이미지가 많으면 남은 이미지는 그대로 유지
        if len(images) > len(pages_int):
            out.extend(images[len(pages_int):])
        return out
