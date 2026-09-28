"""모델 출력에서 JSON을 안정적으로 파싱.

모델이 markdown 펜스나 앞뒤 텍스트를 붙이는 경우가 있어 공통 래퍼를 제거한 뒤 json.loads한다.
"""

from __future__ import annotations

import json


def parse_json_response(text: str) -> dict:
    """모델 응답 텍스트에서 JSON 본문만 추출해 파싱."""
    t = (text or "").strip()

    # markdown 펜스 제거
    if "```json" in t:
        start = t.find("```json") + 7
        end = t.rfind("```")
        if end > start:
            t = t[start:end].strip()
    elif "```" in t:
        start = t.find("```") + 3
        end = t.rfind("```")
        if end > start:
            t = t[start:end].strip()
            if t.startswith("json"):
                t = t[4:].strip()

    # 첫 JSON 괄호 위치 찾기
    json_start = -1
    for i, ch in enumerate(t):
        if ch in "{[":
            json_start = i
            break
    if json_start > 0:
        t = t[json_start:]

    # 마지막 JSON 괄호 위치 찾기
    json_end = -1
    for i in range(len(t) - 1, -1, -1):
        if t[i] in "}]":
            json_end = i + 1
            break
    if json_end > 0:
        t = t[:json_end]

    try:
        return json.loads(t)
    except json.JSONDecodeError as e:
        preview = t[:200].replace("\n", "\\n")
        raise json.JSONDecodeError(f"{e.msg} (preview: {preview})", e.doc, e.pos) from e
