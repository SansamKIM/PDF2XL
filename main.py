# main.py - PDF → JSON → Excel 변환 CLI
"""
사용법:
  python main.py --pdf "파일경로.pdf"
  python main.py --pdf "파일경로.pdf" --org "리얼미터"
  python main.py --pdf "파일경로.pdf" --type "지방"
  
환경변수:
  OPENAI_API_KEY: OpenAI API 키 (필수)
  POPPLER_PATH: Poppler 설치 경로 (선택)
"""

import os
import sys
import json
import argparse
from pathlib import Path
from datetime import datetime

sys.path.insert(0, str(Path(__file__).parent))

from config import APP_NAME, APP_VERSION, OPENAI_API_KEY, OPENAI_MODEL, POPPLER_PATH, PDF_DPI, OUTPUT_DIR
from src.extractor import extract_from_pdf, detect_survey_type
from src.excel_writer import process_all_json
from src.config.org import detect_org_from_pdf, get_available_orgs as list_available_orgs
from src.util.unicode_name import decode_hashu


def detect_org_name(pdf_path: str) -> str:
    """PDF 파일명에서 기관명 자동 판별 (config/org loader 기반)"""
    return detect_org_from_pdf(pdf_path)


def get_available_orgs() -> list:
    """사용 가능한 기관 목록 (표시용, 디코딩 포함)"""
    return list_available_orgs()


def validate_json_data(json_path: str) -> dict:
    """JSON 데이터 검증 (누락률, 합계 체크)"""
    with open(json_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    stats = {
        "total_categories": 0,
        "total_values": 0,
        "null_values": 0,
        "sum_issues": []
    }

    table_data = data.get("데이터", {})
    response_items = data.get("응답항목", [])

    for category, values in table_data.items():
        stats["total_categories"] += 1

        # 응답항목 합계 체크 (100% ± 5% 허용)
        total = 0
        null_count = 0

        for item in response_items:
            val = values.get(item)
            stats["total_values"] += 1

            if val is None:
                null_count += 1
                stats["null_values"] += 1
            else:
                total += val

        # 합계가 95~105 범위 밖이면 경고
        if total > 0 and (total < 95 or total > 105):
            stats["sum_issues"].append({
                "category": category,
                "sum": round(total, 1),
                "expected": "95~105%"
            })

    # 누락률 계산
    if stats["total_values"] > 0:
        stats["null_rate"] = round(stats["null_values"] / stats["total_values"] * 100, 1)
    else:
        stats["null_rate"] = 0

    return stats


def save_run_report(report: dict, output_dir: str):
    """실행 리포트 저장"""
    report_path = os.path.join(output_dir, "run_report.json")
    with open(report_path, "w", encoding="utf-8") as f:
        json.dump(report, f, ensure_ascii=False, indent=2)
    return report_path


def main():
    # CLI 파서 설정
    parser = argparse.ArgumentParser(
        description=f"{APP_NAME} v{APP_VERSION} - PDF 여론조사 → Excel 변환",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
예시:
  python main.py --pdf "조사결과.pdf"
  python main.py --pdf "조사결과.pdf" --org "리얼미터"
  python main.py --pdf "조사결과.pdf" --type "지방"
  
환경변수:
  OPENAI_API_KEY: OpenAI API 키 (필수)
  POPPLER_PATH: Poppler 설치 경로 (선택)
        """
    )

    parser.add_argument("--pdf", required=True, help="PDF 파일 경로")
    parser.add_argument("--org", help=f"기관명 (자동감지 또는: {', '.join(get_available_orgs())})")
    parser.add_argument("--type", choices=["전국", "지방"], help="조사유형 (자동감지 또는 직접지정)")
    parser.add_argument("--output", default=OUTPUT_DIR, help="출력 폴더 (기본: OUTPUT_DIR)")
    parser.add_argument("--dpi", type=int, default=PDF_DPI, help="PDF 변환 DPI (기본: PDF_DPI)")
    parser.add_argument("--skip-validation", action="store_true", help="검증 단계 건너뛰기")

    args = parser.parse_args()

    # === 실행 리포트 초기화 ===
    report = {
        "시작시간": datetime.now().isoformat(),
        "입력": {
            "pdf": args.pdf,
            "org": None,
            "type": None,
            "dpi": args.dpi
        },
        "단계별결과": [],
        "검증": [],
        "최종결과": None
    }

    # === 입력 검증 ===
    if not os.path.exists(args.pdf):
        print(f"❌ PDF 파일 없음: {args.pdf}")
        sys.exit(1)

    if not OPENAI_API_KEY:
        print("❌ OPENAI_API_KEY 환경변수를 설정하세요.")
        print("   Windows: set OPENAI_API_KEY=sk-...")
        print("   Mac/Linux: export OPENAI_API_KEY=sk-...")
        sys.exit(1)

    # === 기관명 판별 ===
    org_name = args.org or detect_org_name(args.pdf)
    if org_name is None:
        print("❌ 기관 판별 실패: 파일명에서 기관명을 찾을 수 없음")
        print(f"   파일명: {Path(args.pdf).stem}")
        print(f"   지원 기관: {', '.join(get_available_orgs())}")
        print("   --org 옵션으로 직접 지정하세요.")
        sys.exit(1)

    # === 조사유형 판별 ===
    survey_type = args.type or detect_survey_type(args.pdf)

    report["입력"]["org"] = org_name
    report["입력"]["type"] = survey_type

    # === 실행 ===
    print("=" * 60)
    print(f"{APP_NAME} v{APP_VERSION}")
    print("PDF → JSON → Excel 변환")
    print("=" * 60)
    print(f"PDF: {args.pdf}")
    print(f"기관: {org_name}")
    print(f"조사유형: {survey_type}")
    print(f"모델: {OPENAI_MODEL}")
    print(f"DPI: {args.dpi}")
    print("-" * 60)

    pdf_name = decode_hashu(Path(args.pdf).stem)
    json_dir = os.path.join(args.output, pdf_name)
    excel_dir = os.path.join(args.output, pdf_name, "excel")

    try:
        # === 1단계: PDF → JSON ===
        print("\n[1/3] PDF → JSON 추출 중...")

        result = extract_from_pdf(
            pdf_path=args.pdf,
            org_name=org_name,
            api_key=OPENAI_API_KEY,
            model=OPENAI_MODEL,
            poppler_path=POPPLER_PATH,
            dpi=args.dpi,
            output_dir=args.output
        )

        extraction_result = {
            "단계": "PDF→JSON",
            "상태": "성공",
            "메타데이터필드": len(result.get('메타데이터', {})),
            "추출항목수": len(result.get('항목', [])),
            "항목": [{"name": i["name"], "type": i["type"], "file": i["file"]} 
                    for i in result.get('항목', [])]
        }
        report["단계별결과"].append(extraction_result)

        print(f"  ✅ 메타데이터: {len(result.get('메타데이터', {}))}개 필드")
        print(f"  ✅ 추출 항목: {len(result.get('항목', []))}개")

        # === 2단계: 검증 ===
        if not args.skip_validation:
            print("\n[2/3] 데이터 검증 중...")

            for item in result.get('항목', []):
                json_path = os.path.join(json_dir, item['file'])
                if os.path.exists(json_path):
                    stats = validate_json_data(json_path)

                    validation_result = {
                        "항목": item['name'],
                        "분류수": stats["total_categories"],
                        "누락률": f"{stats['null_rate']}%",
                        "합계이상": len(stats["sum_issues"])
                    }
                    report["검증"].append(validation_result)

                    # 경고 출력
                    if stats["null_rate"] > 10:
                        print(f"  ⚠️ {item['name']}: 누락률 {stats['null_rate']}%")
                    if stats["sum_issues"]:
                        print(f"  ⚠️ {item['name']}: 합계 이상 {len(stats['sum_issues'])}건")
                        for issue in stats["sum_issues"][:3]:  # 최대 3개만 출력
                            print(f"      - {issue['category']}: {issue['sum']}%")

            print("  ✅ 검증 완료")
        else:
            print("\n[2/3] 검증 건너뜀 (--skip-validation)")

        # === 3단계: JSON → Excel ===
        print("\n[3/3] JSON → Excel 변환 중...")

        excel_results = process_all_json(json_dir, excel_dir, survey_type, org_name)

        success_count = sum(1 for r in excel_results if r["status"] == "success")
        fail_count = len(excel_results) - success_count

        excel_result = {
            "단계": "JSON→Excel",
            "상태": "성공" if fail_count == 0 else "일부실패",
            "성공": success_count,
            "실패": fail_count,
            "파일": [r["excel"] for r in excel_results if r["status"] == "success"]
        }
        report["단계별결과"].append(excel_result)

        # === 최종 결과 ===
        report["종료시간"] = datetime.now().isoformat()
        report["최종결과"] = {
            "상태": "성공" if fail_count == 0 else "일부실패",
            "JSON폴더": json_dir,
            "Excel폴더": excel_dir,
            "성공": success_count,
            "실패": fail_count
        }

        # 리포트 저장
        report_path = save_run_report(report, json_dir)

        print("\n" + "=" * 60)
        print("✅ 변환 완료!")
        print("=" * 60)
        print(f"JSON 폴더: {json_dir}")
        print(f"Excel 폴더: {excel_dir}")
        print(f"실행 리포트: {report_path}")
        print(f"결과: 성공 {success_count}개 / 실패 {fail_count}개")

        if fail_count > 0:
            print("\n실패 항목:")
            for r in excel_results:
                if r["status"] != "success":
                    print(f"  ❌ {r['json']}: {r.get('error', '알 수 없음')}")

    except Exception as e:
        report["종료시간"] = datetime.now().isoformat()
        report["최종결과"] = {"상태": "실패", "오류": str(e)}

        # 리포트 저장 시도
        try:
            os.makedirs(json_dir, exist_ok=True)
            save_run_report(report, json_dir)
        except:
            pass

        print(f"\n❌ 오류 발생: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
