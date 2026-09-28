# -*- mode: python ; coding: utf-8 -*-

"""PyInstaller spec (macOS 대응).

- macOS 배포는 .app(BUNDLE)로 생성
- Poppler를 프로젝트에 번들한 경우(third_party/poppler/macos/) 해당 디렉터리를 포함합니다.
  (번들하지 않는 경우: 시스템 Poppler(pdfinfo/pdftoppm)를 사용)

참고
- PyInstaller의 권장 패턴에 맞춰, COLLECT 결과(coll)를 BUNDLE로 감싸 .app을 생성합니다.
"""

import re
import sys
from pathlib import Path

# PyInstaller는 spec 파일을 exec()로 실행하며, 항상 __file__을 정의하지 않을 수 있습니다.
# 대신 SPEC(이 spec 파일의 경로)를 제공합니다.
_spec = globals().get("SPEC")
project_dir = Path(_spec).resolve().parent if _spec else Path.cwd()


def _read_app_version(default: str = "1.0.2") -> str:
    """src/__init__.py에서 __version__을 문자열로만 추출(임포트/실행 없이)."""
    init_py = project_dir / "src" / "__init__.py"
    try:
        text = init_py.read_text(encoding="utf-8")
    except Exception:
        return default

    m = re.search(r"__version__\s*=\s*['\"]([^'\"]+)['\"]", text)
    return m.group(1).strip() if m else default


APP_VERSION = _read_app_version()

# --- 데이터 파일(templates/config 등) ---
datas = [
    (str(project_dir / "templates"), "templates"),
    (str(project_dir / "config"), "config"),
]

# --- Poppler 번들(플랫폼별) ---
if sys.platform == "darwin":
    poppler_dir = project_dir / "third_party" / "poppler" / "macos"
    if poppler_dir.exists():
        datas.append((str(poppler_dir), "third_party/poppler/macos"))
elif sys.platform == "win32":
    poppler_dir = project_dir / "third_party" / "poppler" / "poppler-24.08.0"
    if poppler_dir.exists():
        datas.append((str(poppler_dir), "third_party/poppler/poppler-24.08.0"))


a = Analysis(
    ["gui.py"],
    pathex=[str(project_dir)],
    binaries=[],
    datas=datas,
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="pdf2xlAI",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=True if sys.platform == "darwin" else False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name="pdf2xlAI",
)

# macOS: .app 번들 생성
if sys.platform == "darwin":
    app = BUNDLE(
        coll,
        name="pdf2xlAI.app",
        icon=None,
        bundle_identifier="com.internal.pdf2xlai",
        # CFBundleShortVersionString(버전 문자열)
        version=str(APP_VERSION),
        # 추가 Info.plist 키
        info_plist={
            # Retina 디스플레이에서 Tk UI를 선명하게 표시
            "NSHighResolutionCapable": True,
            "CFBundleDisplayName": "pdf2xlAI",
        },
    )
