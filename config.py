# config.py - 설정 파일
"""프로젝트 전역 설정.

환경변수 설정 방법:

[Windows CMD]
  set OPENAI_API_KEY=sk-xxxxxxxxxxxx
  set POPPLER_PATH=C:\\poppler\\Library\\bin

[Windows PowerShell]
  $env:OPENAI_API_KEY="sk-xxxxxxxxxxxx"
  $env:POPPLER_PATH="C:\\poppler\\Library\\bin"

[Mac/Linux]
  export OPENAI_API_KEY=sk-xxxxxxxxxxxx
  export POPPLER_PATH=/usr/local/bin

Mac 배포(.app)에서는 Poppler를 앱 안에 번들링하는 전제를 사용합니다.
- 번들 Poppler가 존재하면 자동으로 그 경로를 사용합니다.
- POPPLER_PATH 환경변수로 언제든 override 할 수 있습니다.
"""

from __future__ import annotations

import os
import sys
from pathlib import Path

try:
    from src import __app_name__ as APP_NAME, __version__ as APP_VERSION
except Exception:
    # 예외적인 경우를 위한 최후 폴백(정상 실행에서는 사용되지 않음)
    APP_NAME = "pdf2xlAI"
    APP_VERSION = "1.0.1"


def _runtime_base_dir() -> Path:
    """번들 리소스를 찾기 위한 기준 폴더 반환.

    - 소스 실행: 이 파일이 있는 폴더(프로젝트 루트)
    - PyInstaller onefile: sys._MEIPASS 임시 폴더
    - PyInstaller onedir: 실행 파일이 있는 폴더
    - macOS .app: 감지되면 .../Contents/Resources 우선
    """

    if getattr(sys, "frozen", False):
        meipass = getattr(sys, "_MEIPASS", None)
        if meipass:
            return Path(meipass)

        exe_dir = Path(sys.executable).resolve().parent

        # macOS .app 구조: <App>.app/Contents/MacOS/<exe>
        # 리소스는 <App>.app/Contents/Resources/ 아래에 위치
        if sys.platform == "darwin" and exe_dir.name == "MacOS":
            res_dir = exe_dir.parent / "Resources"
            if res_dir.exists():
                return res_dir

        return exe_dir

    return Path(__file__).resolve().parent


_BASE_DIR = _runtime_base_dir()


def _prepend_env_path(var: str, path: str) -> None:
    """환경변수(PATH 계열) 앞에 경로를 하나 추가."""

    if not path:
        return

    current = os.environ.get(var, "")
    if not current:
        os.environ[var] = path
        return

    parts = current.split(os.pathsep)
    if path in parts:
        return

    os.environ[var] = path + os.pathsep + current


def _is_under(child: Path, parent: Path) -> bool:
    try:
        child.resolve().relative_to(parent.resolve())
        return True
    except Exception:
        return False


# =============================================================================
# OpenAI API 설정
# =============================================================================
# 방법 1: 환경변수 (권장)
# 방법 2: 아래에 직접 입력
OPENAI_API_KEY = os.environ.get("OPENAI_API_KEY", "")

# 모델 선택 (환경변수 OPENAI_MODEL 우선)
#  - gpt-5.4-mini: 빠르고 효율적인 기본 모델
#  - gpt-5-mini: 기존 저비용 모델
#  - gpt-5: 더 정확(비용/시간 증가)
OPENAI_MODEL = os.environ.get("OPENAI_MODEL", "gpt-5.4-mini")

# =============================================================================
# Poppler 경로 (PDF → 이미지 변환용)
# =============================================================================
# Windows: https://github.com/oschwartz10612/poppler-windows/releases
# Mac:    (배포용) third_party/poppler/macos/ 아래에 번들
#         (개발용) brew install poppler
# Linux:  sudo apt install poppler-utils
#
# 우선순위:
# 1) 환경변수 POPPLER_PATH
# 2) 프로젝트(또는 번들) 내 내장 poppler
# 3) OS별 흔한 기본 경로(homebrew 등)

_EMBEDDED_POPPLER_WIN = _BASE_DIR / "third_party" / "poppler" / "poppler-24.08.0" / "Library" / "bin"
_EMBEDDED_POPPLER_MAC = _BASE_DIR / "third_party" / "poppler" / "macos" / "bin"


def _poppler_bin_is_valid(bin_dir: Path) -> bool:
    """해당 디렉터리가 사용 가능한 poppler bin 폴더이면 True."""

    if not bin_dir.exists():
        return False

    if sys.platform == "win32":
        return (bin_dir / "pdftoppm.exe").exists() and (bin_dir / "pdfinfo.exe").exists()

    # macOS/Linux
    return (bin_dir / "pdftoppm").exists() and (bin_dir / "pdfinfo").exists()


def _detect_poppler_path() -> str | None:
    env = os.environ.get("POPPLER_PATH")
    if env:
        return env

    if sys.platform == "win32":
        if _poppler_bin_is_valid(_EMBEDDED_POPPLER_WIN):
            return str(_EMBEDDED_POPPLER_WIN)
        # 개발 환경용 기본 경로
        return r"C:\\poppler\\poppler-24.08.0\\Library\\bin"

    if sys.platform == "darwin":
        # 번들된 poppler를 우선 사용
        if _poppler_bin_is_valid(_EMBEDDED_POPPLER_MAC):
            return str(_EMBEDDED_POPPLER_MAC)

        # 개발용 폴백(Homebrew)
        for p in (Path("/opt/homebrew/bin"), Path("/usr/local/bin"), Path("/usr/bin")):
            if _poppler_bin_is_valid(p):
                return str(p)
        return None

    # Linux 기본 경로(가능한 경우)
    for p in (Path("/usr/bin"), Path("/usr/local/bin")):
        if _poppler_bin_is_valid(p):
            return str(p)
    return None


POPPLER_PATH: str | None = _detect_poppler_path()

# poppler bin을 PATH에 자동 추가
if POPPLER_PATH:
    _prepend_env_path("PATH", POPPLER_PATH)

    # macOS: 번들 poppler의 dylib를 찾을 수 있도록 lib 경로도 함께 설정
    if sys.platform == "darwin":
        poppler_bin = Path(POPPLER_PATH)
        poppler_root = poppler_bin.parent
        lib_dir = poppler_root / "lib"
        # 번들된 poppler일 때만 환경변수 주입(사용자 시스템 오염 최소화)
        if _is_under(poppler_root, _BASE_DIR):
            if lib_dir.exists():
                _prepend_env_path("DYLD_LIBRARY_PATH", str(lib_dir))
                _prepend_env_path("DYLD_FALLBACK_LIBRARY_PATH", str(lib_dir))

            # Poppler 데이터(encoding/CMap 등)가 필요한 경우를 대비해 같이 번들링할 수 있음
            data_dir = poppler_root / "share" / "poppler"
            if data_dir.exists():
                # poppler에서 이 환경변수를 참조하는 구현이 있는 버전들을 위해 best-effort로 세팅
                os.environ.setdefault("POPPLER_DATADIR", str(data_dir))
                os.environ.setdefault("POPPLER_DATA_DIR", str(data_dir))

# =============================================================================
# PDF 처리 설정
# =============================================================================
PDF_DPI = int(os.environ.get("PDF_DPI", "300"))  # 높을수록 정확하지만 느림 (200~400 권장)

# =============================================================================
# 출력 경로
# =============================================================================
# Windows exe는 cwd가 System32로 잡히는 경우가 있어 프로그램 폴더 옆 output을 기본으로.
# macOS .app은 앱 번들 내부에 쓰기 권한 문제가 잦아서 기본 출력은 Documents로.


def _default_output_dir() -> Path:
    env = os.environ.get("OUTPUT_DIR")
    if env:
        return Path(env).expanduser()

    if sys.platform == "darwin":
        return Path.home() / "Documents" / APP_NAME / "output"

    if sys.platform == "win32":
        return _BASE_DIR / "output"

    # Linux/others
    return Path.home() / APP_NAME / "output"


OUTPUT_DIR = str(_default_output_dir())
