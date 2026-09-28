"""Excel 템플릿 후처리: 동적 테두리/병합/포맷 확장.

값을 미리 스타일된 템플릿에 쓸 때, 작성 범위가 템플릿의 사전 그려진 영역을 넘는 경우가 있다:

- ISSUE: 응답 열이 동적이라 오른쪽을 넘어갈 수 있음
- 지방: 지역 행이 동적이라 아래를 넘어갈 수 있음

이 모듈은 **스타일**(테두리/행 높이/열 너비/숫자 포맷)을 확장해
동적 데이터에도 템플릿 룩을 유지하도록 한다.

설계 목표
------------
1) 확장만 하고 템플릿을 줄이지 않는다.
2) 표 밖의 '권역설명' 같은 블록은 건드리지 않는다.
3) 지방 '지역별' 병합 라벨 블록(case C)을 처리한다.
4) 숫자 포맷 적용:
   - '조사완료'(사례수) 열: 정수 포맷(템플릿 기반)
   - 응답 열: 소수 첫째 자리 포맷(템플릿 기반)
"""

from __future__ import annotations

from copy import copy
from dataclasses import dataclass
import re
from typing import Any, Dict, Iterable, List, Optional, Tuple

from openpyxl.styles import Border
from openpyxl.styles.borders import Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet


@dataclass(frozen=True)
class BaseBounds:
    header_row: int
    data_start_row: int
    left_col: int
    sample_col: int
    base_last_row: int
    base_last_col: int


def _cell_has_any_border(ws: Worksheet, row: int, col: int) -> bool:
    b = ws.cell(row=row, column=col).border
    try:
        return any(
            getattr(b, side).style is not None
            for side in ("left", "right", "top", "bottom")
        )
    except Exception:
        return False


def _border_with(b: Border, **updates: Any) -> Border:
    """선택된 면을 교체한 새 Border를 반환."""
    return Border(
        left=updates.get("left", b.left),
        right=updates.get("right", b.right),
        top=updates.get("top", b.top),
        bottom=updates.get("bottom", b.bottom),
        diagonal=updates.get("diagonal", b.diagonal),
        vertical=updates.get("vertical", b.vertical),
        horizontal=updates.get("horizontal", b.horizontal),
        diagonalUp=b.diagonalUp,
        diagonalDown=b.diagonalDown,
        outline=b.outline,
        start=b.start,
        end=b.end,
    )


def _copy_style(src, dst) -> None:
    """셀 스타일 속성만 복사(값은 복사하지 않음)."""
    # _style은 전체 스타일 인덱스를 담고 있어 가장 빠르지만, 명시적 필드도 함께 복사
    try:
        dst._style = copy(src._style)
    except Exception:
        pass
    try:
        dst.font = copy(src.font)
    except Exception:
        pass
    try:
        dst.fill = copy(src.fill)
    except Exception:
        pass
    try:
        dst.border = copy(src.border)
    except Exception:
        pass
    try:
        dst.alignment = copy(src.alignment)
    except Exception:
        pass
    try:
        dst.protection = copy(src.protection)
    except Exception:
        pass
    try:
        dst.number_format = src.number_format
    except Exception:
        pass


def _find_header_col(ws: Worksheet, header_row: int, label: str, start_col: int = 1, max_cols: int = 60) -> Optional[int]:
    want = str(label).strip()
    for c in range(start_col, start_col + max_cols):
        v = ws.cell(row=header_row, column=c).value
        if v is None:
            continue
        if str(v).strip() == want:
            return c
    return None


def _detect_base_bounds(
    ws: Worksheet,
    header_row: int = 4,
    data_start_row: int = 5,
    left_col: int = 2,  # B열
    max_scan_cols: int = 60,
    max_scan_rows: int = 400,
) -> BaseBounds:
    # '조사완료' 열을 기준점으로 사용
    sample_col = _find_header_col(ws, header_row, "조사완료", start_col=left_col, max_cols=max_scan_cols) or 4

    # 기본 마지막 열: 헤더/데이터 영역 중 경계선이 존재하는 가장 오른쪽 열
    scan_rows = [header_row, data_start_row, data_start_row + 1]
    base_last_col = left_col
    for c in range(left_col, left_col + max_scan_cols):
        if any(_cell_has_any_border(ws, r, c) for r in scan_rows if r >= 1):
            base_last_col = max(base_last_col, c)

    # 기본 마지막 행: [left_col..base_last_col] 범위에서 경계선이 있는 가장 아래 행
    base_last_row = header_row
    for r in range(header_row, max_scan_rows + 1):
        if any(_cell_has_any_border(ws, r, c) for c in range(left_col, base_last_col + 1)):
            base_last_row = r

    return BaseBounds(
        header_row=header_row,
        data_start_row=data_start_row,
        left_col=left_col,
        sample_col=sample_col,
        base_last_row=base_last_row,
        base_last_col=base_last_col,
    )


def _find_region_label_row(ws: Worksheet, left_col: int = 2, max_scan_rows: int = 120) -> Optional[int]:
    """B열에 '지역별'이 있는 행 찾기(지방 템플릿 기본)."""
    for r in range(1, max_scan_rows + 1):
        v = ws.cell(row=r, column=left_col).value
        if v is None:
            continue
        if str(v).strip() == "지역별":
            return r
    return None


def _merged_range_containing(ws: Worksheet, row: int, col: int) -> Optional[str]:
    coord = ws.cell(row=row, column=col).coordinate
    for mr in ws.merged_cells.ranges:
        try:
            if coord in mr:
                return str(mr)
        except Exception:
            continue
    return None


def _extend_columns(ws: Worksheet, bounds: BaseBounds, desired_last_col: int) -> None:
    if desired_last_col <= bounds.base_last_col:
        return

    table_rows = range(bounds.header_row, bounds.base_last_row + 1)
    old_last = bounds.base_last_col
    proto_outer = old_last
    proto_internal = max(bounds.sample_col + 1, old_last - 1)

    # 열 단위로 스타일 복사
    for c in range(old_last + 1, desired_last_col + 1):
        src_col = proto_outer if c == desired_last_col else proto_internal
        for r in table_rows:
            _copy_style(ws.cell(row=r, column=src_col), ws.cell(row=r, column=c))

        # 열 너비 복사
        src_letter = get_column_letter(src_col)
        dst_letter = get_column_letter(c)
        try:
            ws.column_dimensions[dst_letter].width = ws.column_dimensions[src_letter].width
        except Exception:
            pass

    # 기존 마지막 열은 내부 열이 되므로 경계선 조정
    for r in table_rows:
        cell_old_last = ws.cell(row=r, column=old_last)
        cell_proto_int = ws.cell(row=r, column=proto_internal)
        b = cell_old_last.border
        if b is None:
            continue
        new_right: Side = cell_proto_int.border.right
        cell_old_last.border = _border_with(b, right=new_right)


def _extend_rows(ws: Worksheet, bounds: BaseBounds, desired_last_row: int, desired_last_col: int) -> None:
    if desired_last_row <= bounds.base_last_row:
        return

    old_last = bounds.base_last_row
    proto_row = max(bounds.data_start_row, old_last - 1)
    cols = range(bounds.left_col, desired_last_col + 1)

    # 기존 마지막 행의 바깥쪽 하단 테두리를 미리 저장
    outer_bottom: Dict[int, Side] = {
        c: ws.cell(row=old_last, column=c).border.bottom for c in cols
    }

    # 새 행에 스타일 복사
    for r in range(old_last + 1, desired_last_row + 1):
        for c in cols:
            _copy_style(ws.cell(row=proto_row, column=c), ws.cell(row=r, column=c))

        # 행 높이 복사
        try:
            ws.row_dimensions[r].height = ws.row_dimensions[proto_row].height
        except Exception:
            pass

    # 기존 마지막 행은 내부 행이 되므로 프로토 행의 하단 테두리로 교체
    for c in cols:
        cell_old_last = ws.cell(row=old_last, column=c)
        proto_cell = ws.cell(row=proto_row, column=c)
        b = cell_old_last.border
        if b is None:
            continue
        cell_old_last.border = _border_with(b, bottom=proto_cell.border.bottom)

    # 새 마지막 행은 바깥쪽 하단 테두리(템플릿에서는 중간 굵기)를 받음
    for c in cols:
        cell_new_last = ws.cell(row=desired_last_row, column=c)
        b = cell_new_last.border
        if b is None:
            continue
        cell_new_last.border = _border_with(b, bottom=outer_bottom.get(c, b.bottom))


def _extend_region_merge_if_needed(ws: Worksheet, region_label_row: int, region_needed_last_row: int, left_col: int = 2) -> None:
    """케이스 C: B열의 병합된 '지역별' 블록을 필요한 만큼 확장.

    확장만 하고 축소는 하지 않음.
    """
    rng = _merged_range_containing(ws, region_label_row, left_col)
    if not rng:
        return
    try:
        # 예: 'B14:B17'
        start, end = rng.split(":")
        end_row = int("".join(ch for ch in end if ch.isdigit()))
    except Exception:
        return

    if region_needed_last_row <= end_row:
        return

    # 병합 범위를 아래로 확장
    try:
        ws.unmerge_cells(rng)
    except Exception:
        pass
    col_letter = get_column_letter(left_col)
    ws.merge_cells(f"{col_letter}{region_label_row}:{col_letter}{region_needed_last_row}")


def _extend_vertical_double_separator(ws: Worksheet, bounds: BaseBounds, desired_last_row: int) -> None:
    """B열에 이중 세로선이 있는 템플릿이면 이를 연장."""
    # 기본 테이블에서 B.right가 double인 첫 행 찾기
    c = bounds.left_col
    first: Optional[int] = None
    for r in range(bounds.header_row, bounds.base_last_row + 1):
        try:
            if ws.cell(row=r, column=c).border.right.style == "double":
                first = r
                break
        except Exception:
            continue
    if first is None:
        return

    # 템플릿의 Side 객체를 그대로 사용해 색상을 유지
    template_side = ws.cell(row=first, column=c).border.right

    for r in range(first, desired_last_row + 1):
        cell = ws.cell(row=r, column=c)
        b = cell.border
        if b is None:
            continue
        cell.border = _border_with(b, right=template_side)


def _apply_number_formats(ws: Worksheet, bounds: BaseBounds, desired_last_row: int, desired_last_col: int) -> None:
    """테이블 데이터 영역에 숫자 서식 적용.

    요구사항: '조사완료' 열 제외, 응답 열에는 소수 1자리 형식 적용.
    실제 서식 문자열은 템플릿(D5, E5 등)에서 가져옴.
    """

    sample_col = bounds.sample_col
    data_r1 = bounds.data_start_row
    data_r2 = desired_last_row

    # 템플릿 셀에서 서식 추출
    try:
        sample_fmt = ws.cell(row=data_r1, column=sample_col).number_format or "#,##0"
    except Exception:
        sample_fmt = "#,##0"
    try:
        percent_fmt = ws.cell(row=data_r1, column=sample_col + 1).number_format or "0.0"
    except Exception:
        percent_fmt = "0.0"
    # General 서식을 명시적 숫자 서식으로 변환
    if str(sample_fmt).strip().lower() == "general":
        sample_fmt = "#,##0"
    if str(percent_fmt).strip().lower() == "general":
        percent_fmt = "0.0"

    for r in range(data_r1, data_r2 + 1):
        # 조사완료(사례수)
        try:
            ws.cell(row=r, column=sample_col).number_format = sample_fmt
        except Exception:
            pass

        # 응답(%)
        for c in range(sample_col + 1, desired_last_col + 1):
            try:
                ws.cell(row=r, column=c).number_format = percent_fmt
            except Exception:
                continue


def _cell_has_value(cell) -> bool:
    v = cell.value
    if v is None:
        return False
    if isinstance(v, str):
        return bool(v.strip())
    return True


def _find_sample_anchor(ws: Worksheet, start_row: int = 3, max_row: int = 30, max_col: int = 80) -> Tuple[Optional[int], Optional[int]]:
    for r in range(start_row, max_row + 1):
        for c in range(1, max_col + 1):
            try:
                v = ws.cell(row=r, column=c).value
            except Exception:
                continue
            if isinstance(v, str) and v.strip() == "조사완료":
                return r, c
    return None, None


def _find_last_data_row(ws: Worksheet, left_col: int, right_col: int, start_row: int, end_row: int) -> Optional[int]:
    last = None
    for r in range(start_row, end_row + 1):
        for c in range(left_col, right_col + 1):
            try:
                if _cell_has_value(ws.cell(row=r, column=c)):
                    last = r
                    break
            except Exception:
                continue
    return last


def _find_last_value_col(ws: Worksheet, left_col: int, right_col: int, start_row: int, end_row: int) -> Optional[int]:
    last_col = None
    for r in range(start_row, end_row + 1):
        for c in range(left_col, right_col + 1):
            try:
                if _cell_has_value(ws.cell(row=r, column=c)):
                    last_col = c if last_col is None else max(last_col, c)
            except Exception:
                continue
    return last_col


def _shrink_to_last_response(
    ws: Worksheet,
    *,
    bounds: BaseBounds,
    desired_last_col: int,
    desired_last_row: int,
) -> Optional[Dict[str, Any]]:
    header_row, sample_col = _find_sample_anchor(ws)
    if header_row is None or sample_col is None:
        return None
    if header_row < 3:
        return None

    old_last_col = desired_last_col
    data_start_row = max(bounds.data_start_row, 3)

    last_header_col = None
    for c in range(bounds.left_col, old_last_col + 1):
        try:
            if _cell_has_value(ws.cell(row=header_row, column=c)):
                last_header_col = c
        except Exception:
            continue

    last_data_row = _find_last_data_row(ws, bounds.left_col, old_last_col, data_start_row, desired_last_row)
    if last_data_row is None:
        return None

    last_value_col = _find_last_value_col(ws, bounds.left_col, old_last_col, data_start_row, last_data_row)

    new_last_col = max(
        sample_col,
        last_header_col or sample_col,
        last_value_col or sample_col,
    )
    if new_last_col >= old_last_col:
        return {
            "skipped": False,
            "reason": "no_shrink_needed",
            "header_row": header_row,
            "sample_col": sample_col,
            "old_last_col": old_last_col,
            "new_last_col": old_last_col,
            "last_data_row": last_data_row,
        }

    try:
        template_right = ws.cell(row=header_row, column=old_last_col).border.right
    except Exception:
        template_right = None
    try:
        template_bottom = ws.cell(row=last_data_row, column=old_last_col).border.bottom
    except Exception:
        template_bottom = None

    clear_side = Side(style=None, color=None)

    for r in range(header_row, last_data_row + 1):
        cell_new = ws.cell(row=r, column=new_last_col)
        b = cell_new.border or Border()
        cell_new.border = _border_with(b, right=template_right or b.right)

    for r in range(header_row, last_data_row + 1):
        for c in range(new_last_col + 1, old_last_col + 1):
            cell = ws.cell(row=r, column=c)
            b = cell.border or Border()
            cell.border = _border_with(b, left=clear_side, right=clear_side, top=clear_side, bottom=clear_side)

    bottom_cell = ws.cell(row=last_data_row, column=new_last_col)
    b = bottom_cell.border or Border()
    bottom_cell.border = _border_with(b, bottom=template_bottom or b.bottom, right=template_right or b.right)

    return {
        "skipped": False,
        "header_row": header_row,
        "sample_col": sample_col,
        "old_last_col": old_last_col,
        "new_last_col": new_last_col,
        "last_data_row": last_data_row,
        "last_header_col": last_header_col,
        "last_value_col": last_value_col,
    }


def finalize_table_styles(
    ws: Worksheet,
    *,
    survey_type: str,
    item_type: str,
    response_items: List[str] | None,
    region_list: List[str] | None,
    shrink_to_last_response: bool = False,
) -> Dict[str, Any]:
    """동적 데이터에 맞게 스타일/병합/서식을 확장.

    디버그용 작은 dict를 반환(무시해도 안전).
    """

    bounds = _detect_base_bounds(ws)

    # 목표 마지막 열: 응답항목 수 기반(D + len(responses))
    resp_n = len(response_items or [])
    desired_last_col = max(bounds.base_last_col, bounds.sample_col + resp_n)

    # 목표 마지막 행: 지방 권역 행이 있을 때만 확장
    desired_last_row = bounds.base_last_row
    region_label_row = None
    region_needed_last_row = None
    if survey_type == "지방":
        region_label_row = _find_region_label_row(ws, left_col=bounds.left_col)
        if region_label_row is not None and region_list:
            region_needed_last_row = region_label_row + len(region_list) - 1
            # 확장만 하고 축소는 하지 않음
            if region_needed_last_row > bounds.base_last_row:
                desired_last_row = region_needed_last_row

    # 케이스 A: 열을 먼저 확장(행 확장 시 새 열 스타일을 복사하기 위함)
    _extend_columns(ws, bounds, desired_last_col)

    # 케이스 A: 행 확장(지역 행 확장도 같은 방식 사용)
    _extend_rows(ws, bounds, desired_last_row, desired_last_col)

    # 케이스 C: 필요할 때만 '지역별' 병합 블록 확장
    if region_label_row is not None and region_needed_last_row is not None:
        _extend_region_merge_if_needed(ws, region_label_row, region_needed_last_row, left_col=bounds.left_col)

    # 템플릿에 있는 이중 세로선이 확장된 행에서도 이어지도록 보장
    _extend_vertical_double_separator(ws, bounds, desired_last_row)

    # 퍼센트 서식 적용(조사완료 제외)
    _apply_number_formats(ws, bounds, desired_last_row, desired_last_col)

    shrink_info = None
    if shrink_to_last_response:
        shrink_info = _shrink_to_last_response(
            ws,
            bounds=bounds,
            desired_last_col=desired_last_col,
            desired_last_row=desired_last_row,
        )

    return {
        "base_last_row": bounds.base_last_row,
        "base_last_col": bounds.base_last_col,
        "desired_last_row": desired_last_row,
        "desired_last_col": desired_last_col,
        "sample_col": bounds.sample_col,
        "region_label_row": region_label_row,
        "region_needed_last_row": region_needed_last_row,
        "shrink": shrink_info,
    }
