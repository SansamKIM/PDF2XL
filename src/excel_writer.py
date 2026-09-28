"""레거시 진입점 래퍼.

초기에는 src/excel_writer.py에서 작성 함수를 노출했으나,
지금은 구현이 src.write.excel_writer에 있고 이 파일은 호환 경로만 유지한다.
"""

from __future__ import annotations

from .write.excel_writer import json_to_excel, process_all_json, get_unique_filepath

__all__ = ["json_to_excel", "process_all_json", "get_unique_filepath"]
