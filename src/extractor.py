"""레거시 엔트리포인트 래퍼.

프로젝트 초기에는 src/extractor.py에서 추출 함수를 노출했다.
이 파일은 import 경로 호환성을 유지하기 위해 남겨두었고,
실제 구현은 src.extract.extractor 모듈에 있다.
"""

from __future__ import annotations

from .extract.extractor import extract_from_pdf, detect_survey_type, save_json

__all__ = [
    "extract_from_pdf",
    "detect_survey_type",
    "save_json",
]
