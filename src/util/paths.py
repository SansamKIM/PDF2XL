"""런타임 안전 경로 헬퍼.

일반 파이썬 패키지로 개발하지만 Windows .exe(PyInstaller) 배포도 염두에 둔다.
frozen 환경에서는 상대경로나 __file__ 기반 루트 탐지가 자주 깨지므로,
아래 헬퍼로 경로 기준을 일원화한다:
- project root(templates/, config/, third_party/가 있는 곳)
- output root(기본 출력 폴더)
"""

from __future__ import annotations

import os
import sys
from pathlib import Path

from .. import __app_name__


def is_frozen() -> bool:
    return bool(getattr(sys, "frozen", False))


def runtime_base_dir() -> Path:
    """리소스 탐색의 기준 폴더.

    - PyInstaller onefile: sys._MEIPASS
    - PyInstaller onedir: 실행 파일이 있는 폴더
    - 소스 실행: 저장소 루트(src/의 두 단계 위)
    """

    if is_frozen():
        meipass = getattr(sys, "_MEIPASS", None)
        if meipass:
            return Path(meipass)
        exe_dir = Path(sys.executable).resolve().parent

        # macOS .app 구조: <App>.app/Contents/MacOS/<exe>
        # 리소스는 <App>.app/Contents/Resources/
        if sys.platform == "darwin" and exe_dir.name == "MacOS":
            res_dir = exe_dir.parent / "Resources"
            if res_dir.exists():
                return res_dir

        return exe_dir

    # 소스 실행: .../<repo>/src/util/paths.py -> parents[2] == <repo>
    return Path(__file__).resolve().parents[2]


def project_root() -> Path:
    """templates/, config/, third_party/가 있는 폴더 반환.

    우선 runtime_base_dir()을 시도하고, 없으면 현 파일 기준으로 위로 탐색.
    """

    base = runtime_base_dir()
    if (base / "templates").exists() and (base / "config").exists():
        return base

    # 방어적으로 현재 파일에서 위로 탐색
    for parent in Path(__file__).resolve().parents:
        if (parent / "templates").exists() and (parent / "config").exists():
            return parent

    return base


def default_output_dir() -> Path:
    """기본 출력 디렉터리.

    .exe 실행 시 cwd가 System32일 수 있어, OUTPUT_DIR가 없으면 실행 파일 옆의 안정된 폴더를 사용.
    """

    env = os.environ.get("OUTPUT_DIR")
    if env:
        return Path(env)

    # macOS: .app 번들/임시 경로 내부 쓰기를 피함
    if sys.platform == "darwin":
        return Path.home() / "Documents" / __app_name__ / "output"

    # Windows 배포본(frozen): 실행 파일 옆의 안정된 폴더 사용
    if sys.platform == "win32":
        return runtime_base_dir() / "output"

    # Linux/기타
    return Path.home() / __app_name__ / "output"
