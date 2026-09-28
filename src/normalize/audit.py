"""정규화 감사(audit) 유틸리티.

이 모듈은 라벨 정규화 계층에 대한 **안전한** 간이 점검(sanity check)을 제공한다.
기관 YAML이 없어도 실행할 수 있다.

현재 점검 항목
- 멱등성(idempotence): normalize_label(normalize_label(x)) == normalize_label(x)
  (정규화가 여러 단계에서 여러 번 적용될 수 있으므로 중요)
- 기본 키 위생: 키/값에 앞뒤 공백이 있으면 안 된다.

사용법
    python -m src.normalize.audit
"""

from __future__ import annotations

from typing import List, Tuple

from .labels import LABEL_ALIASES, normalize_label


def check_idempotence() -> List[Tuple[str, str, str]]:
    """[(입력, 1회 결과, 2회 결과), ...] 형태로, 2회 결과 != 1회 결과인 항목을 반환."""
    bad: List[Tuple[str, str, str]] = []
    samples = set(LABEL_ALIASES.keys()) | set(LABEL_ALIASES.values())
    for s in sorted(samples):
        once = normalize_label(s)
        twice = normalize_label(once)
        if twice != once:
            bad.append((s, once, twice))
    return bad


def check_whitespace_hygiene() -> List[str]:
    """앞/뒤 공백이 포함된 라벨(키/값)을 반환."""
    bad = []
    for s in list(LABEL_ALIASES.keys()) + list(LABEL_ALIASES.values()):
        if s != s.strip():
            bad.append(s)
    return bad


def main() -> int:
    idemp = check_idempotence()
    ws = check_whitespace_hygiene()

    if not idemp and not ws:
        print("OK: 정규화 매핑이 안정적으로 보입니다")
        return 0

    if ws:
        print("\n[공백 문제]")
        for s in ws:
            print(f"  - {s!r}")

    if idemp:
        print("\n[멱등성 문제] (2회 정규화 결과 != 1회 정규화 결과)")
        for inp, once, twice in idemp:
            print(f"  - {inp!r} -> once={once!r} -> twice={twice!r}")

    return 1


if __name__ == "__main__":
    raise SystemExit(main())
