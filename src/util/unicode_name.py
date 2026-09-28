"""'#UXXXX' 이스케이프 유니코드 파일명을 다루기 위한 유틸.

환경을 오가며 zip을 만들거나 풀 때 비ASCII 파일명이 '#Uacb0#Uacfc...'처럼 표현될 수 있어,
기관명과 PDF 파일명을 안정적으로 매칭할 수 있도록 디코딩을 지원한다.
"""

from __future__ import annotations

import re

_HASHU_RE = re.compile(r"#U([0-9a-fA-F]{4})")


def decode_hashu(text: str) -> str:
    """'#UXXXX' 시퀀스를 유니코드 문자로 디코드.

    패턴이 잘못되면 그대로 둔다.
    """

    if not isinstance(text, str) or "#U" not in text:
        return text

    def _repl(m: re.Match[str]) -> str:
        try:
            return chr(int(m.group(1), 16))
        except Exception:
            return m.group(0)

    return _HASHU_RE.sub(_repl, text)


def maybe_decode_hashu(text: str) -> str:
    """'#U' 시퀀스가 있을 때만 디코드."""
    return decode_hashu(text) if isinstance(text, str) and "#U" in text else text


def normalize_identifier(text: str) -> str:
    """매칭용 식별자 정규화.

    - #U 이스케이프 디코드
    - 소문자화
    - 앞뒤 공백 제거
    """
    if not isinstance(text, str):
        return ""
    t = decode_hashu(text)
    return t.strip().lower()
