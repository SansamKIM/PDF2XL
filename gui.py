# gui.py - PDF → Excel 변환 GUI
"""
GUI 실행: python gui.py
"""

import os
import sys
import json
import threading
import queue
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from pathlib import Path
from datetime import datetime

try:
    from platformdirs import user_config_dir
except ImportError:
    # platformdirs가 없으면 홈 아래의 기본 설정 폴더를 사용
    def user_config_dir(appname: str = "pdf2xlAI") -> str:
        return str(Path.home() / ".config" / appname)

# Windows DPI 스케일링 해결
try:
    import ctypes
    ctypes.windll.shcore.SetProcessDpiAwareness(1)  # 시스템 DPI 인식
except:
    pass

# 프로젝트 경로 설정
sys.path.insert(0, str(Path(__file__).parent))

from config import APP_NAME, APP_VERSION, OPENAI_API_KEY, OPENAI_MODEL, POPPLER_PATH, PDF_DPI, OUTPUT_DIR
from src.util.unicode_name import decode_hashu
from src.config.org import detect_org_from_pdf, get_available_orgs, load_org_config
from src.util.paths import project_root as get_project_root
from src.util.cache import cache_root, cache_total_size_bytes, clear_all_cache


CONFIG_PATH = Path(user_config_dir("pdf2xlAI")) / "config.json"


class PDFConverterGUI:
    def __init__(self, root):
        self.root = root
        self.root.title(f"{APP_NAME} v{APP_VERSION} - 여론조사 PDF → Excel 변환")
        self.root.geometry("800x800")
        self.root.resizable(True, True)

        # 상태 변수
        self.pdf_path = tk.StringVar()
        self.output_path = tk.StringVar(value=os.path.abspath(OUTPUT_DIR))
        self.org_var = tk.StringVar()
        self.type_var = tk.StringVar(value="자동감지")
        self.dpi_var = tk.IntVar(value=PDF_DPI)
        self.model_var = tk.StringVar(value=OPENAI_MODEL)
        self.api_key_var = tk.StringVar(value=self._load_initial_api_key())
        self.remember_api_key = tk.BooleanVar(value=bool(self.api_key_var.get()) and not bool(OPENAI_API_KEY))
        self.current_api_key = self.api_key_var.get()
        self.is_running = False

        # 마지막 확정 진행률을 저장해 100%를 일찍 표시하지 않음
        # 단계가 실제 끝난 뒤에만 100%를 찍도록 사용
        self._last_progress_stage: str | None = None
        self._last_progress_total: int | None = None

        # 현재 실행을 중단시키는 플래그(소프트 스톱)
        self.cancel_event = threading.Event()

        # 워커→UI 스레드로 전달하는 스레드 안전 로그 큐
        self._log_queue: "queue.Queue[tuple[str, str]]" = queue.Queue()
        self._log_pump_started = False

        # 추출 항목 체크 변수
        self.extract_psr = tk.BooleanVar(value=True)   # 정당지지도
        self.extract_ge = tk.BooleanVar(value=True)    # 국정운영평가
        self.extract_issue = tk.BooleanVar(value=True) # 이슈

        # 기관별 설정 캐시
        self.org_configs = {}

        self.create_widgets()
        self._start_log_pump()
        self.load_organizations()
        self.check_api_key()

        # 시작 시 캐시 용량을 가볍게 정리(할당량 유지)
        # 본격적인 정리는 추출 중에도 한 번 더 실행됨
        try:
            from src.util.cache import enforce_cache_quota

            enforce_cache_quota()
        except Exception:
            pass

    def create_widgets(self):
        """위젯 생성"""
        # 메인 프레임
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # === 1. 파일 선택 ===
        file_frame = ttk.LabelFrame(main_frame, text="1. PDF 파일 선택", padding="10")
        file_frame.pack(fill=tk.X, pady=(0, 10))

        self.file_entry = ttk.Entry(file_frame, textvariable=self.pdf_path, width=70)
        self.file_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))

        browse_btn = ttk.Button(file_frame, text="찾아보기", command=self.browse_file)
        browse_btn.pack(side=tk.RIGHT)

        # === 2. 출력 폴더 선택 ===
        output_frame = ttk.LabelFrame(main_frame, text="2. 출력 폴더", padding="10")
        output_frame.pack(fill=tk.X, pady=(0, 10))

        self.output_entry = ttk.Entry(output_frame, textvariable=self.output_path, width=70)
        self.output_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))

        output_browse_btn = ttk.Button(output_frame, text="변경", command=self.browse_output_folder)
        output_browse_btn.pack(side=tk.RIGHT)

        # === 3. 기관 선택 ===
        org_frame = ttk.LabelFrame(main_frame, text="3. 기관 선택", padding="10")
        org_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(org_frame, text="기관:").pack(side=tk.LEFT, padx=(0, 10))
        self.org_combo = ttk.Combobox(org_frame, textvariable=self.org_var, 
                                       state="readonly", width=20)
        self.org_combo.pack(side=tk.LEFT, padx=(0, 20))
        self.org_combo.bind("<<ComboboxSelected>>", self.on_org_change)

        ttk.Label(org_frame, text="조사유형:").pack(side=tk.LEFT, padx=(0, 10))
        self.type_combo = ttk.Combobox(org_frame, textvariable=self.type_var,
                                        values=["자동감지", "전국", "지방"],
                                        state="readonly", width=12)
        self.type_combo.pack(side=tk.LEFT)
        self.type_warn_var = tk.StringVar(value="")
        ttk.Label(org_frame, textvariable=self.type_warn_var, foreground="red").pack(side=tk.LEFT, padx=(10, 0))

        # === 4. 추출 항목 선택 ===
        extract_frame = ttk.LabelFrame(main_frame, text="4. 추출 항목 선택", padding="10")
        extract_frame.pack(fill=tk.X, pady=(0, 10))

        self.psr_check = ttk.Checkbutton(extract_frame, text="정당지지도 (PSR)",
                                          variable=self.extract_psr)
        self.psr_check.pack(side=tk.LEFT, padx=(0, 20))

        self.ge_check = ttk.Checkbutton(extract_frame, text="국정운영평가 (GE)",
                                         variable=self.extract_ge)
        self.ge_check.pack(side=tk.LEFT, padx=(0, 20))

        self.issue_check = ttk.Checkbutton(extract_frame, text="이슈 여론 (ISSUE)",
                                            variable=self.extract_issue)
        self.issue_check.pack(side=tk.LEFT)

        # === 5. 설정 ===
        settings_frame = ttk.LabelFrame(main_frame, text="5. 설정", padding="10")
        settings_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(settings_frame, text="DPI (해상도):").pack(side=tk.LEFT, padx=(0, 10))
        dpi_spin = ttk.Spinbox(settings_frame, from_=150, to=500, increment=50,
                               textvariable=self.dpi_var, width=8)
        dpi_spin.pack(side=tk.LEFT, padx=(0, 20))

        ttk.Label(settings_frame, text="모델:").pack(side=tk.LEFT, padx=(0, 10))
        self.model_combo = ttk.Combobox(
            settings_frame,
            textvariable=self.model_var,
            values=["gpt-5.4-mini", "gpt-5", "gpt-5-mini"],
            state="readonly",
            width=14,
        )
        self.model_combo.pack(side=tk.LEFT)
        ttk.Button(settings_frame, text="API 키 설정", command=self.open_api_key_dialog).pack(side=tk.LEFT, padx=(20, 0))

        # === 6. 변환 버튼 ===
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=(0, 10))

        self.convert_btn = ttk.Button(btn_frame, text="🚀 변환 시작",
                                       command=self.start_conversion)
        self.convert_btn.pack(side=tk.LEFT, padx=(0, 10))

        self.stop_btn = ttk.Button(btn_frame, text="⏹ 중지",
                                    command=self.stop_conversion, state=tk.DISABLED)
        self.stop_btn.pack(side=tk.LEFT, padx=(0, 10))

        self.clear_cache_btn = ttk.Button(btn_frame, text="🗑 캐시 삭제", command=self.clear_cache)
        self.clear_cache_btn.pack(side=tk.RIGHT, padx=(0, 10))

        self.open_folder_btn = ttk.Button(btn_frame, text="📁 출력 폴더 열기",
                                           command=self.open_output_folder)
        self.open_folder_btn.pack(side=tk.RIGHT)

        # === 7. 진행 상태 ===
        progress_frame = ttk.LabelFrame(main_frame, text="진행 상태", padding="10")
        progress_frame.pack(fill=tk.X, pady=(0, 10))

        self.progress_var = tk.StringVar(value="대기 중...")
        self.progress_label = ttk.Label(progress_frame, textvariable=self.progress_var)
        self.progress_label.pack(anchor=tk.W)

        # 진행 표시 (B): 결정형 막대 + 텍스트
        # 총량을 모를 때 '준비' 단계에서는 잠시 비결정형으로 전환
        self.progress_bar = ttk.Progressbar(progress_frame, mode='determinate', maximum=100)
        try:
            self.progress_bar["value"] = 0
        except Exception:
            pass
        self.progress_bar.pack(fill=tk.X, pady=(5, 0))

        # === 8. 로그 ===
        log_frame = ttk.LabelFrame(main_frame, text="로그", padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True)

        # macOS: Consolas가 기본 폰트가 아니라서 Menlo로 자동 전환
        # (없으면 TkFixedFont를 사용)
        mono_font = "Consolas"
        if sys.platform == "darwin":
            mono_font = "Menlo"
        elif sys.platform.startswith("linux"):
            mono_font = "DejaVu Sans Mono"

        try:
            self.log_text = scrolledtext.ScrolledText(
                log_frame,
                height=12,
                state=tk.DISABLED,
                font=(mono_font, 9),
            )
        except Exception:
            import tkinter.font as tkfont

            self.log_text = scrolledtext.ScrolledText(
                log_frame,
                height=12,
                state=tk.DISABLED,
                font=tkfont.nametofont("TkFixedFont"),
            )
        self.log_text.pack(fill=tk.BOTH, expand=True)

        # === API 키 상태 ===
        self.api_status = ttk.Label(main_frame, text="", foreground="red")
        self.api_status.pack(anchor=tk.W)
        self.api_key_var.trace_add("write", lambda *args: self.check_api_key())

    def load_organizations(self):
        """기관 목록 로드 (YAML 파일명/기관명/별칭 포함)

        - 표시/선택 값은 org_id(기본: YAML 파일명 디코딩) 기준
        - supported_items 등은 load_org_config로 정규화된 값을 사용
        """
        project_root = get_project_root()

        orgs = ["자동감지"]
        self.org_configs = {}

        try:
            available = get_available_orgs(project_root=project_root)
        except Exception as e:
            available = []
            print(f"기관 목록 로드 실패: {e}")

        for org_id in available:
            orgs.append(org_id)
            try:
                self.org_configs[org_id] = load_org_config(org_id, project_root=project_root)
            except Exception as e:
                # GUI는 계속 동작해야 하므로 실패는 로그만
                print(f"설정 로드 실패 ({org_id}): {e}")

        self.org_combo["values"] = orgs
        self.org_combo.set("자동감지")

    def _load_config_file(self):
        """config.json 로드 (깨져있으면 무시)."""
        try:
            if CONFIG_PATH.exists():
                with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                    return json.load(f) or {}
        except Exception:
            pass
        return {}

    def _load_initial_api_key(self) -> str:
        """환경변수 > config.json(openai_api_key) > 빈값 순으로 로드."""
        if OPENAI_API_KEY:
            return OPENAI_API_KEY
        cfg = self._load_config_file()
        key = cfg.get("openai_api_key")
        return str(key).strip() if key else ""

    def _save_api_key(self, key: str, remember: bool):
        """API 키 저장/삭제."""
        if not remember:
            if CONFIG_PATH.exists():
                try:
                    CONFIG_PATH.unlink()
                except Exception:
                    pass
            return
        try:
            CONFIG_PATH.parent.mkdir(parents=True, exist_ok=True)
            with open(CONFIG_PATH, "w", encoding="utf-8") as f:
                json.dump({"openai_api_key": key}, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def _delete_saved_api_key(self):
        try:
            if CONFIG_PATH.exists():
                CONFIG_PATH.unlink()
        except Exception:
            pass

    def get_supported_types(self, org_name: str) -> list:
        """기관에서 지원하는 추출 항목 반환"""
        if org_name == "자동감지":
            return ["PSR", "GE", "ISSUE"]

        cfg = self.org_configs.get(org_name)
        if cfg is None:
            return ["PSR", "GE", "ISSUE"]

        return list(cfg.supported_items or ["PSR", "GE", "ISSUE"])

    def on_org_change(self, event=None):
        """기관 선택 변경 시 - 지원 항목에 따라 체크박스 활성화/비활성화

        IMPORTANT: 사용자가 수동으로 체크 해제한 값을 '자동으로 다시 켜지' 않도록,
        지원하는 항목은 현재 체크 상태를 유지합니다.
        (지원하지 않는 항목만 강제로 OFF + 비활성화)
        """
        org = self.org_var.get()
        supported = self.get_supported_types(org)

        # PSR
        if "PSR" in supported:
            self.psr_check.config(state=tk.NORMAL)
            # 현재 선택 상태 유지
        else:
            self.psr_check.config(state=tk.DISABLED)
            self.extract_psr.set(False)

        # GE
        if "GE" in supported:
            self.ge_check.config(state=tk.NORMAL)
            # 현재 선택 상태 유지
        else:
            self.ge_check.config(state=tk.DISABLED)
            self.extract_ge.set(False)

        # ISSUE
        if "ISSUE" in supported:
            self.issue_check.config(state=tk.NORMAL)
            # 현재 선택 상태 유지
        else:
            self.issue_check.config(state=tk.DISABLED)
            self.extract_issue.set(False)

        # 기타기관: 조사유형 자동감지 지원 중단(전국/지방만 선택 가능)
        if org == "기타기관":
            allowed_types = ["전국", "지방"]
            self.type_combo.config(values=allowed_types)
            if self.type_var.get() not in allowed_types:
                self.type_var.set("전국")
            self.type_warn_var.set("조사유형을 꼭 확인해주세요.")
        else:
            allowed_types = ["자동감지", "전국", "지방"]
            self.type_combo.config(values=allowed_types)
            if self.type_var.get() not in allowed_types:
                self.type_var.set("자동감지")
            self.type_warn_var.set("")

        self.log(f"기관 선택: {org} (지원: {', '.join(supported)})")

    def check_api_key(self):
        """API 키 확인"""
        key = (self.api_key_var.get() or "").strip()
        if key:
            env_set = bool(OPENAI_API_KEY)
            masked = key if len(key) <= 8 else f"{key[:4]}...{key[-4:]}"
            self.api_status.config(
                text=f"✅ API 키 설정됨 ({'env' if env_set else 'local'}: {masked})",
                foreground="green"
            )
            if not self.is_running:
                self.convert_btn.config(state=tk.NORMAL)
        else:
            self.api_status.config(
                text="⚠️ API 키 미설정: 입력 또는 환경변수 OPENAI_API_KEY 설정 필요",
                foreground="red"
            )
            self.convert_btn.config(state=tk.DISABLED)

    def browse_file(self):
        """파일 선택 다이얼로그"""
        filename = filedialog.askopenfilename(
            title="PDF 파일 선택",
            filetypes=[("PDF 파일", "*.pdf"), ("모든 파일", "*.*")]
        )
        if filename:
            self.pdf_path.set(filename)
            self.auto_detect_org(filename)

            # 출력 폴더 기본값: PDF와 같은 폴더의 output
            pdf_dir = os.path.dirname(filename)
            default_output = os.path.join(pdf_dir, "output")
            self.output_path.set(default_output)

    def browse_output_folder(self):
        """출력 폴더 선택 다이얼로그"""
        folder = filedialog.askdirectory(
            title="출력 폴더 선택",
            initialdir=self.output_path.get()
        )
        if folder:
            self.output_path.set(folder)
            self.log(f"출력 폴더 변경: {folder}")

    def auto_detect_org(self, pdf_path):
        """파일명에서 기관 자동 감지 (실패 시 자동감지 유지)"""
        project_root = get_project_root()
        detected = detect_org_from_pdf(pdf_path, project_root=project_root)
        if detected:
            self.org_combo.set(detected)
        else:
            self.org_combo.set("자동감지")
        self.on_org_change()

    def prompt_org_selection(self, pdf_path: str):
        """기관 자동판별 실패 시 사용자 선택 다이얼로그."""
        choices = [x for x in (self.org_combo["values"] or []) if x and x != "자동감지"]
        if not choices:
            return None

        dialog = tk.Toplevel(self.root)
        dialog.title("기관 선택")
        dialog.transient(self.root)
        dialog.grab_set()

        fname = Path(pdf_path).name
        ttk.Label(dialog, text="기관을 자동으로 판별할 수 없습니다.\n아래에서 조사기관을 선택하세요.").pack(
            padx=12, pady=(12, 6), anchor=tk.W
        )
        ttk.Label(dialog, text=f"PDF: {fname}", foreground="gray").pack(
            padx=12, pady=(0, 10), anchor=tk.W
        )

        sel = tk.StringVar(value=choices[0])
        combo = ttk.Combobox(dialog, textvariable=sel, values=choices, state="readonly", width=28)
        combo.pack(padx=12, pady=(0, 12), fill=tk.X)
        combo.focus_set()

        btns = ttk.Frame(dialog)
        btns.pack(padx=12, pady=(0, 12), fill=tk.X)

        result = {"value": None}

        def on_ok():
            result["value"] = sel.get().strip() if sel.get() else None
            dialog.destroy()

        def on_cancel():
            result["value"] = None
            dialog.destroy()

        ttk.Button(btns, text="확인", command=on_ok).pack(side=tk.LEFT)
        ttk.Button(btns, text="취소", command=on_cancel).pack(side=tk.LEFT, padx=(8, 0))

        self.root.wait_window(dialog)
        return result["value"]

    def _start_log_pump(self):
        """워커 스레드 로그를 UI에 반영하는 주기 작업을 시작."""
        if self._log_pump_started:
            return
        self._log_pump_started = True
        self.root.after(80, self._drain_log_queue)

    def _drain_log_queue(self):
        """UI 스레드에서 로그 큐 메시지를 비우기(최선 시도)."""
        try:
            if not self.root.winfo_exists():
                return
        except Exception:
            return

        flushed = False
        try:
            while True:
                ts, msg = self._log_queue.get_nowait()
                if msg is None:
                    continue
                if not flushed:
                    self.log_text.config(state=tk.NORMAL)
                    flushed = True
                self.log_text.insert(tk.END, f"[{ts}] {msg}\n")
        except queue.Empty:
            pass
        except Exception:
            # 로그 문제로 GUI가 멈추지 않도록 무시
            pass
        finally:
            if flushed:
                try:
                    self.log_text.see(tk.END)
                    self.log_text.config(state=tk.DISABLED)
                except Exception:
                    pass

        # 다음 주기로 재호출
        try:
            if self.root.winfo_exists():
                self.root.after(80, self._drain_log_queue)
        except Exception:
            pass

    def log(self, message):
        """워커 스레드에서도 안전한 로그 큐잉."""
        try:
            ts = datetime.now().strftime("%H:%M:%S")
        except Exception:
            ts = "--:--:--"
        try:
            self._log_queue.put_nowait((ts, str(message)))
        except Exception:
            pass

    def _fmt_bytes(self, n: int) -> str:
        try:
            n = int(n)
        except Exception:
            return "0 B"
        units = ["B", "KB", "MB", "GB", "TB"]
        size = float(n)
        for u in units:
            if size < 1024 or u == units[-1]:
                if u == "B":
                    return f"{int(size)} {u}"
                return f"{size:.2f} {u}"
            size /= 1024.0

    def clear_cache(self):
        """OS 캐시 디렉터리에 저장된 모든 캐시 파일을 삭제."""

        if self.is_running:
            messagebox.showwarning("캐시 삭제", "변환 중에는 캐시를 삭제할 수 없습니다.\n중지 후 다시 시도하세요.")
            return

        root = cache_root()
        size = cache_total_size_bytes()

        if size <= 0:
            messagebox.showinfo("캐시 삭제", f"삭제할 캐시가 없습니다.\n\n위치: {root}")
            return

        msg = (
            "캐시를 모두 삭제할까요?\n\n"
            f"위치: {root}\n"
            f"현재 크기: {self._fmt_bytes(size)}\n\n"
            "삭제하면 다음 실행에서 PDF 렌더링/재시도 캐시가 다시 생성되어\n"
            "처음 실행이 더 느릴 수 있습니다."
        )

        if not messagebox.askyesno("캐시 삭제", msg):
            return

        try:
            clear_all_cache()
            self.log(f"캐시 삭제 완료: {root}")
            messagebox.showinfo("캐시 삭제", "캐시를 삭제했습니다.")
        except Exception as e:
            self.log(f"캐시 삭제 실패: {e}")
            messagebox.showerror("캐시 삭제", f"캐시 삭제 중 오류가 발생했습니다.\n\n{e}")

    def open_api_key_dialog(self):
        """API 키 설정/저장 다이얼로그."""
        dialog = tk.Toplevel(self.root)
        dialog.title("API 키 설정")
        dialog.transient(self.root)
        dialog.grab_set()

        ttk.Label(dialog, text="API 키 (env > 파일 > 입력 순 적용)").pack(padx=12, pady=(12, 4), anchor=tk.W)
        entry = ttk.Entry(dialog, textvariable=self.api_key_var, width=40, show="*")
        entry.pack(padx=12, fill=tk.X)
        entry.focus_set()

        remember_chk = ttk.Checkbutton(dialog, text="API 키 기억하기", variable=self.remember_api_key)
        remember_chk.pack(padx=12, pady=(8, 12), anchor=tk.W)

        btns = ttk.Frame(dialog)
        btns.pack(padx=12, pady=(0, 12), fill=tk.X)

        def on_save():
            key = (self.api_key_var.get() or "").strip()
            self.current_api_key = key
            self._save_api_key(key, self.remember_api_key.get())
            self.check_api_key()
            dialog.destroy()

        def on_reset():
            self.api_key_var.set("")
            self.remember_api_key.set(False)
            self._delete_saved_api_key()
            self.check_api_key()

        ttk.Button(btns, text="저장", command=on_save).pack(side=tk.LEFT)
        ttk.Button(btns, text="초기화(삭제)", command=on_reset).pack(side=tk.LEFT, padx=(8, 0))
        ttk.Button(btns, text="닫기", command=dialog.destroy).pack(side=tk.RIGHT)

        self.root.wait_window(dialog)

    def start_conversion(self):
        """변환 시작"""
        if self.is_running:
            return
        # 입력 검증
        pdf_path = self.pdf_path.get()
        if not pdf_path or not os.path.exists(pdf_path):
            messagebox.showerror("오류", "PDF 파일을 선택하세요.")
            return

        # API 키는 환경변수를 우선 사용
        api_key = (OPENAI_API_KEY or self.api_key_var.get() or "").strip()
        if not api_key:
            messagebox.showerror("오류", "API 키가 설정되지 않았습니다.\nAPI 키를 입력하거나 환경변수 OPENAI_API_KEY를 설정해주세요.")
            return

        # 추출 항목 확인
        if not any([self.extract_psr.get(), self.extract_ge.get(), self.extract_issue.get()]):
            messagebox.showerror("오류", "최소 하나의 추출 항목을 선택하세요.")
            return

        # 기관 자동감지(파일명 기반) + 실패 시 선택 팝업
        if self.org_var.get() == "자동감지":
            detected = self.detect_org_name(pdf_path)
            if detected:
                self.org_combo.set(detected)
                self.on_org_change()
                self.log(f"기관 자동 감지: {detected}")
            else:
                messagebox.showwarning(
                    "기관 판별 실패",
                    "파일명에서 기관을 찾지 못했습니다.\n조사기관을 선택해 주세요.",
                )
                chosen = self.prompt_org_selection(pdf_path)
                if not chosen:
                    return
                self.org_combo.set(chosen)
                self.on_org_change()
                self.log(f"기관 수동 선택: {chosen}")

        # UI 상태 변경
        self.is_running = True
        self.cancel_event = threading.Event()
        self.current_api_key = api_key
        self.convert_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)
        try:
            self.clear_cache_btn.config(state=tk.DISABLED)
        except Exception:
            pass

        # 워커 스레드에서 Tk 변수를 건드리지 않도록 추출 항목을 미리 복사
        extract_types: list[str] = []
        if self.extract_psr.get():
            extract_types.append("PSR")
        if self.extract_ge.get():
            extract_types.append("GE")
        if self.extract_issue.get():
            extract_types.append("ISSUE")

        settings = {
            "pdf_path": pdf_path,
            "output_dir": self.output_path.get(),
            "org_name": self.org_var.get(),
            "survey_type": self.type_var.get(),
            "dpi": int(self.dpi_var.get()),
            "model": (self.model_var.get() or OPENAI_MODEL).strip(),
            "api_key": api_key,
            "extract_types": extract_types,
        }

        # 총량을 모르는 준비 단계는 비결정형으로 표시
        self._set_progress_ui("준비 중...", mode="indeterminate", value=0, maximum=100)

        # 별도 스레드에서 실행
        thread = threading.Thread(target=self.run_conversion, args=(settings,), daemon=True)
        thread.start()

    def stop_conversion(self):
        """변환 중지"""
        if not self.is_running:
            return

        # 소프트 스톱: 취소 플래그만 세우고 워커가 안전 지점에서 멈추게 함
        try:
            self.cancel_event.set()
        except Exception:
            pass

        try:
            self.stop_btn.config(state=tk.DISABLED)
        except Exception:
            pass

        self.log("⏹ 중지 요청됨... (현재 단계 완료 후 중지됩니다)")
        self._set_progress_ui("⏹ 중지 요청됨...", mode="indeterminate")

    def reset_ui(self):
        """UI 상태 초기화"""
        self.is_running = False
        try:
            self.stop_btn.config(state=tk.DISABLED)
        except Exception:
            pass
        try:
            self.clear_cache_btn.config(state=tk.NORMAL)
        except Exception:
            pass

        # 진행 표시 초기화
        self._set_progress_ui("대기 중...", mode="determinate", value=0, maximum=100)

        # API 키 조건 다시 반영
        try:
            self.check_api_key()
        except Exception:
            # 실행 중이 아니면 버튼 다시 활성화(예외 대비)
            try:
                self.convert_btn.config(state=tk.NORMAL)
            except Exception:
                pass

    def _ask_yesno_threadsafe(self, title: str, message: str) -> bool:
        """UI 스레드에서 예/아니요를 물어보고 워커 스레드에 전달."""
        if self.cancel_event.is_set():
            return False

        result: dict = {"value": False}
        ev = threading.Event()

        def _ask():
            try:
                result["value"] = messagebox.askyesno(title, message)
            except Exception:
                result["value"] = False
            finally:
                try:
                    ev.set()
                except Exception:
                    pass

        try:
            self.root.after(0, _ask)
        except Exception:
            return False

        # 대기하면서 취소 요청은 계속 확인
        while True:
            if ev.wait(timeout=0.1):
                break
            if self.cancel_event.is_set():
                return False
            try:
                if not self.root.winfo_exists():
                    return False
            except Exception:
                return False

        return bool(result.get("value"))

    def run_conversion(self, settings: dict):
        """변환 실행 (별도 스레드)."""
        from src.util.cancel import UserCancelled

        try:
            from src.extractor import extract_from_pdf, detect_survey_type
            from src.excel_writer import process_all_json

            pdf_path = str(settings.get("pdf_path") or "").strip()
            output_dir = str(settings.get("output_dir") or "").strip()
            org_name = str(settings.get("org_name") or "").strip()
            survey_type = str(settings.get("survey_type") or "자동감지").strip()
            dpi = int(settings.get("dpi") or PDF_DPI)
            model = str(settings.get("model") or OPENAI_MODEL).strip()
            api_key = str(settings.get("api_key") or OPENAI_API_KEY or "").strip()
            extract_types = settings.get("extract_types") or ["PSR", "GE", "ISSUE"]

            if not api_key:
                self.log("API 키가 없어 변환을 중단합니다.")
                self.root.after(0, lambda: messagebox.showerror("오류", "API 키를 입력하세요."))
                return

            if not pdf_path or not os.path.exists(pdf_path):
                self.root.after(0, lambda: messagebox.showerror("오류", "PDF 파일 경로가 올바르지 않습니다."))
                return

            os.makedirs(output_dir, exist_ok=True)

            def _should_cancel() -> bool:
                try:
                    return self.cancel_event.is_set()
                except Exception:
                    return False

            # 필요하면 조사유형 자동 감지
            if (survey_type == "자동감지") and (org_name == "기타기관"):
                survey_type = "전국"  # 기타기관은 자동감지 지원 안 함
                self.log("기타기관은 조사유형 자동감지를 사용하지 않습니다. 기본값 '전국'으로 설정합니다.")
            elif survey_type == "자동감지":
                survey_type = detect_survey_type(pdf_path)
                self.log(f"조사유형 자동 감지: {survey_type}")

            self.log(f"추출 항목: {', '.join(extract_types)}")

            # 1) PDF -> JSON
            self._set_progress_ui("PDF → JSON 추출 준비...", mode="indeterminate")
            self.log("=" * 50)
            self.log(f"PDF: {Path(pdf_path).name}")
            self.log(f"기관: {org_name}")
            self.log(f"조사유형: {survey_type}")
            self.log(f"DPI: {dpi}")
            self.log(f"모델: {model}")
            self.log(f"출력: {output_dir}")
            self.log("=" * 50)

            result = extract_from_pdf(
                pdf_path=pdf_path,
                org_name=org_name,
                api_key=api_key,
                model=model,
                poppler_path=POPPLER_PATH,
                dpi=dpi,
                output_dir=output_dir,
                extract_types=extract_types,
                resume=True,
                progress_callback=self._on_progress_event,
                should_cancel=_should_cancel,
            )

            if _should_cancel():
                raise UserCancelled("cancelled")

            # 실제 추출이 끝난 뒤에야 결정형 진행률을 완료로 표시
            # 마지막 항목 처리 중 100%가 먼저 뜨는 문제를 막음
            try:
                last_total = getattr(self, "_last_progress_total", None)
                if isinstance(last_total, int) and last_total > 0:
                    self._set_progress_ui(
                        "✅ PDF → JSON 추출 완료",
                        mode="determinate",
                        value=int(last_total),
                        maximum=int(last_total),
                    )
            except Exception:
                pass

            if result.get("forced_fresh"):
                self.log("ℹ️ 이전 실행이 중단된 흔적이 있어 안전 모드로 새로 추출했습니다. (resume 비활성화)")

            ok_items = result.get("항목", []) or []
            fail_items = result.get("실패", []) or []

            self.log(f"✅ JSON 추출 완료: {len(ok_items)}개 항목")

            if fail_items and not _should_cancel():
                self.log(f"⚠️ JSON 추출 실패: {len(fail_items)}개 항목")
                for f in fail_items:
                    name = f.get("name")
                    itype = f.get("type")
                    pages = f.get("pages")
                    err = f.get("error")
                    self.log(f"   - {name} ({itype}) p{pages}: {err}")

                retry_pages = sorted({int(p) for f in fail_items for p in (f.get("pages") or []) if p is not None})
                if retry_pages and not _should_cancel():
                    do_retry = self._ask_yesno_threadsafe(
                        "재시도",
                        f"JSON 추출 실패 {len(fail_items)}개 항목이 있습니다.\n"
                        f"해당 페이지만 다시 시도할까요?\n\n재시도 페이지: {retry_pages}",
                    )
                    if do_retry and not _should_cancel():
                        self.log(f"재시도: 페이지 {retry_pages}")
                        retry_result = extract_from_pdf(
                            pdf_path=pdf_path,
                            org_name=org_name,
                            api_key=api_key,
                            model=model,
                            poppler_path=POPPLER_PATH,
                            dpi=dpi,
                            output_dir=output_dir,
                            extract_types=extract_types,
                            resume=False,
                            page_whitelist=retry_pages,
                            skip_meta_region=True,
                            progress_callback=self._on_progress_event,
                            should_cancel=_should_cancel,
                        )
                        if _should_cancel():
                            raise UserCancelled("cancelled")

                        # 재시도까지 끝난 뒤 결정형 진행률도 완료 처리
                        try:
                            last_total = getattr(self, "_last_progress_total", None)
                            if isinstance(last_total, int) and last_total > 0:
                                self._set_progress_ui(
                                    "✅ 재시도 추출 완료",
                                    mode="determinate",
                                    value=int(last_total),
                                    maximum=int(last_total),
                                )
                        except Exception:
                            pass

                        ok_items.extend(retry_result.get("항목", []) or [])
                        fail_items = retry_result.get("실패", []) or []
                        if fail_items:
                            self.log(f"⚠️ 재시도 후에도 실패: {len(fail_items)}개")
                            for f in fail_items:
                                name = f.get("name")
                                itype = f.get("type")
                                pages = f.get("pages")
                                err = f.get("error")
                                self.log(f"   - {name} ({itype}) p{pages}: {err}")
                        else:
                            self.log("재시도 성공: 모든 실패 항목 복구")

            for item in ok_items:
                try:
                    self.log(f"   - {item['name']} ({item['type']})")
                except Exception:
                    pass

            if _should_cancel():
                raise UserCancelled("cancelled")

            # 2) JSON -> Excel
            self._set_progress_ui("JSON → Excel 변환 준비...", mode="indeterminate")

            pdf_name = result.get("pdf_name") or decode_hashu(Path(pdf_path).stem)
            json_dir = result.get("json_dir") or os.path.join(output_dir, pdf_name)
            excel_dir = os.path.join(output_dir, pdf_name)

            excel_results = process_all_json(
                json_dir,
                excel_dir,
                survey_type,
                org_name,
                extract_types=extract_types,
                progress_callback=self._on_progress_event,
                should_cancel=_should_cancel,
            )

            if _should_cancel():
                raise UserCancelled("cancelled")

            success_count = sum(1 for r in excel_results if r.get("status") == "success")
            fail_count = len(excel_results) - success_count

            self.log("=" * 50)
            self.log("✅ Excel 변환 완료!")
            self.log(f"   성공: {success_count}개")
            if fail_count > 0:
                self.log(f"   실패: {fail_count}개")
            self.log(f"   출력: {excel_dir}")
            self.log("=" * 50)

            # 진행률 완료 표시
            self._set_progress_ui(f"✅ 완료! (성공 {success_count}개)", mode="determinate", value=1, maximum=1)

            json_fail_count = len(fail_items)
            json_fail_msg = ""
            if json_fail_count:
                json_fail_msg = f"\n\n⚠️ JSON 추출 실패: {json_fail_count}개 (자세한 내용은 로그 확인)"

            self.root.after(
                0,
                lambda: messagebox.showinfo(
                    "완료",
                    f"변환이 완료되었습니다!\n\n"
                    f"성공: {success_count}개\n"
                    f"실패: {fail_count}개\n\n"
                    f"출력 폴더: {excel_dir}" + json_fail_msg,
                ),
            )

        except UserCancelled:
            self.log("⏹ 중지 완료")
            self._set_progress_ui("⏹ 중지됨", mode="determinate", value=0, maximum=1)

        except Exception as e:
            self.log(f"❌ 오류: {e}")
            import traceback

            self.log(traceback.format_exc())
            error_msg = str(e)
            self.root.after(0, lambda msg=error_msg: messagebox.showerror("오류", msg))

        finally:
            self.root.after(0, self.reset_ui)

    def detect_org_name(self, pdf_path):
        """기관명 자동 감지"""
        project_root = get_project_root()
        return detect_org_from_pdf(pdf_path, project_root=project_root)

    def update_progress(self, message):
        """진행 상태 텍스트 업데이트 (UI 스레드)."""
        self._set_progress_ui(message)

    def _set_progress_ui(self, message: str, *, mode: str | None = None, value: int | None = None, maximum: int | None = None):
        """Tk 메인 스레드에서 진행 텍스트/막대를 안전하게 갱신.

        mode: 'determinate' | 'indeterminate' | None (keep)
        """

        def _do():
            try:
                if mode in {"determinate", "indeterminate"}:
                    try:
                        self.progress_bar.config(mode=mode)
                    except Exception:
                        pass
                    if mode == "indeterminate":
                        try:
                            self.progress_bar.start(10)
                        except Exception:
                            pass
                    else:
                        try:
                            self.progress_bar.stop()
                        except Exception:
                            pass

                if maximum is not None:
                    try:
                        self.progress_bar.config(maximum=max(1, int(maximum)))
                    except Exception:
                        pass

                if value is not None:
                    try:
                        self.progress_bar["value"] = max(0, int(value))
                    except Exception:
                        pass

                if message is not None:
                    self.progress_var.set(str(message))
            except Exception:
                pass

        try:
            self.root.after(0, _do)
        except Exception:
            pass

    def _on_progress_event(self, evt: dict):
        """추출기/작성기 진행 콜백을 처리(스레드 안전)."""
        try:
            stage = (evt.get("stage") or "").strip()
            cur = evt.get("current")
            total = evt.get("total")
            msg = evt.get("message")
        except Exception:
            stage, cur, total, msg = "", None, None, None

        # 결정형/비결정형 모드 결정
        determinate = isinstance(total, int) and total > 0 and cur is not None
        mode = "determinate" if determinate else "indeterminate"

        # 표시용 문구 구성
        text = str(msg) if msg else (stage or "진행 중...")

        # UX 보정: current는 '처리 중인 항목' 값이라 마지막 항목 시작에 100%가 찍히는 문제를 방지
        # 완료 건수 = current-1로 계산해 단계가 끝났을 때만 100%를 표시
        maximum = int(total) if determinate else None
        display_cur = None
        completed = None
        if maximum is not None:
            try:
                display_cur = int(cur)
            except Exception:
                display_cur = None
            if display_cur is not None:
                # 메시지에 '완료'가 있으면 완료로 간주, 아니면 진행 중으로 간주
                is_done_msg = False
                try:
                    if isinstance(msg, str) and any(k in msg for k in ("완료", "done", "complete", "finished")):
                        is_done_msg = True
                except Exception:
                    is_done_msg = False

                completed = display_cur if is_done_msg else (display_cur - 1)
                try:
                    completed = max(0, min(int(completed), int(maximum)))
                except Exception:
                    completed = None

        value = completed if (maximum is not None and completed is not None) else None

        if maximum is not None and value is not None:
            # 단계 종료 시 100% 처리를 위해 마지막 결정형 진행률 저장
            try:
                self._last_progress_stage = stage or None
                self._last_progress_total = int(maximum)
            except Exception:
                pass

            # 완료 건수 기준으로 백분율 계산
            try:
                pct = int((value / maximum) * 100)
            except Exception:
                pct = None

            if pct is not None:
                # 모델이 준 current/total 표시를 살리되 완료 기준 퍼센트를 덧붙임
                marker = None
                try:
                    if display_cur is not None:
                        marker = f"{display_cur}/{maximum}"
                except Exception:
                    marker = None

                if marker and marker in text:
                    text = f"{text} ({pct}%)"
                elif marker:
                    text = f"{text} ({marker}, {pct}%)"
                else:
                    text = f"{text} ({pct}%)"

        self._set_progress_ui(text, mode=mode, value=value, maximum=maximum)

    def open_output_folder(self):
        """출력 폴더 열기"""
        output_path = Path(self.output_path.get())

        # PDF가 선택되어 있으면 해당 PDF 출력 폴더 열기
        pdf_path = self.pdf_path.get()
        if pdf_path and os.path.exists(pdf_path):
            pdf_name = decode_hashu(Path(pdf_path).stem)
            # 현재는 Excel을 <output>/<pdf_name>/ 바로 아래에 저장
            specific_output = output_path / pdf_name
            if specific_output.exists():
                output_path = specific_output

        # 폴더가 없으면 기본 출력 폴더
        if not output_path.exists():
            output_path = Path(self.output_path.get())
            output_path.mkdir(parents=True, exist_ok=True)

        # 폴더 열기
        if sys.platform == "win32":
            os.startfile(output_path)
        elif sys.platform == "darwin":
            os.system(f'open "{output_path}"')
        else:
            os.system(f'xdg-open "{output_path}"')


def main():
    """간단한 GUI 앱 엔트리 포인트."""
    root = tk.Tk()

    # 스타일 설정
    style = ttk.Style()
    try:
        style.theme_use("clam")
    except:
        pass

    app = PDFConverterGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
