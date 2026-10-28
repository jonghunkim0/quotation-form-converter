"""
Excel 견적서 파서
업체별 Excel 견적서를 읽어 항목 목록으로 추출
"""

import pandas as pd
import re
from io import BytesIO
from typing import List, Dict, Any, Optional

# 소계/합계/총액 등 요약 행을 식별하는 키워드
_SUBTOTAL_KEYWORDS = ("소계", "합계", "총액", "총 금액", "총금액", "subtotal", "total")


def parse_excel(file) -> List[Dict[str, Any]]:
    """
    Excel 파일을 파싱하여 견적 항목 목록을 반환.
    
    Args:
        file: 업로드된 파일 객체 (BytesIO 또는 파일 핸들)
        
    Returns:
        [{"raw_data": {...}, "sheet_name": str, "row_index": int}, ...]
    """
    items = []
    
    # 파일 읽기
    if hasattr(file, "read"):
        data = file.read()
        file = BytesIO(data)
    
    # 모든 시트 읽기
    try:
        excel_file = pd.ExcelFile(file)
    except Exception as e:
        raise ValueError(f"Excel 파일을 읽을 수 없습니다: {e}")
    
    for sheet_name in excel_file.sheet_names:
        try:
            df = pd.read_excel(excel_file, sheet_name=sheet_name, header=None)
        except Exception:
            continue
        
        # 헤더 행 찾기 (컬럼명이 있는 행)
        header_row, headers = _find_header_row(df)
        if header_row is None:
            continue

        # 데이터 행 추출
        data_start = header_row + 1
        for idx in range(data_start, len(df)):
            row = df.iloc[idx]
            # 빈 행 스킵
            if _is_empty_row(row):
                continue
            
            raw_data = {}
            for col_idx, header in enumerate(headers):
                if col_idx < len(row):
                    header_str = str(header).strip()
                    # 빈 헤더 컬럼 제외 (이름 없는 컬럼)
                    if not header_str or header_str.startswith("Unnamed:"):
                        continue
                    value = row.iloc[col_idx]
                    if pd.notna(value):
                        raw_data[header_str] = value

            # 소계/합계/총액 행 제외
            if _is_subtotal_row(raw_data):
                continue

            # 2차 헤더 행 제외: 값이 컬럼명과 동일한 행 (예: 수량*=수량*, 단위=단위)
            if _is_duplicate_header(raw_data, headers):
                continue

            if raw_data:
                items.append({
                    "raw_data": raw_data,
                    "sheet_name": sheet_name,
                    "row_index": idx,
                })
    
    return items


def _find_header_row(df: pd.DataFrame) -> tuple:
    """
    헤더 행을 자동 감지.
    여러 컬럼명이 연속으로出现的 행을 헤더로 판단.
    
    Returns:
        (header_row_index, [col1, col2, ...]) or (None, None)
    """
    best_row = None
    best_count = 0
    
    # 첫 10행 중 헤더 후보 탐색
    search_range = min(10, len(df))
    
    for i in range(search_range):
        row = df.iloc[i]
        # 문자열 값이 3개 이상인 행을 헤더 후보로 봄
        str_count = sum(
            1 for v in row 
            if isinstance(v, str) and v.strip() and not _is_number(v)
        )
        if str_count > best_count:
            best_count = str_count
            best_row = i
    
    if best_row is None or best_count < 2:
        return None, None
    
    # 헤더 값 추출
    headers = []
    for v in df.iloc[best_row]:
        if pd.notna(v) and str(v).strip():
            headers.append(str(v).strip())
        else:
            headers.append("")
    
    return best_row, headers


def _is_empty_row(row) -> bool:
    """행이 빈 행인지 확인"""
    return all(pd.isna(v) or str(v).strip() == "" for v in row)


def _is_subtotal_row(raw_data: Dict[str, Any]) -> bool:
    """소계/합계/총액 등 요약 행인지 확인"""
    for value in raw_data.values():
        text = str(value).lower().replace(" ", "")
        for kw in _SUBTOTAL_KEYWORDS:
            if kw.lower().replace(" ", "") in text:
                return True
    return False


def _is_section_header(raw_data: Dict[str, Any]) -> bool:
    """
    섹션 헤더 행인지 확인.
    No 컬럼이 "(N) 이름" 또는 "N)" 형식이면 섹션 헤더로 판단.
    예: "(1) 기구부", "1)", "(2) 전장부"
    """
    no_val = str(raw_data.get("No", "")).strip()
    if not no_val:
        return False
    # "(N) 이름" 형식
    if no_val.startswith("(") and ")" in no_val:
        return True
    # "N)" 형식 (숫자 + 괄호)
    import re
    if re.match(r"^\d+\)$", no_val):
        return True
    return False


def _has_meaningful_data(raw_data: Dict[str, Any]) -> bool:
    """
    실질적 데이터가 있는지 확인.
    자재명/품명/규격/Maker/단가/금액 중 의미 있는 값이 하나 이상 있어야 함.
    """
    data_cols = ("자재명*", "자재명", "품명", "규격*", "규격", "Maker", "단가*", "단가", "금액*", "금액")
    for col in data_cols:
        val = raw_data.get(col)
        if val is None:
            continue
        val_str = str(val).strip()
        # 빈 값, 대시, 0만인 경우 제외
        if val_str and val_str != "-" and val_str != "0" and val_str != "0.0":
            return True
    return False


def _is_duplicate_header(raw_data: Dict[str, Any], headers: List[str]) -> bool:
    """
    2차 헤더 행인지 확인.
    raw_data의 값이 headers와 동일한 경우 (예: {'수량*': '수량*', '단위': '단위'}).
    """
    header_set = {str(h).strip() for h in headers if str(h).strip()}
    match_count = 0
    for key, val in raw_data.items():
        if str(val).strip() == str(key).strip() and str(key).strip() in header_set:
            match_count += 1
    # 2개 이상 컬럼명이 값과 동일하면 2차 헤더
    return match_count >= 2


def _is_number(value) -> bool:
    """값이 숫자인지 확인"""
    try:
        float(str(value).replace(",", "").replace("₩", "").replace("$", "").strip())
        return True
    except (ValueError, TypeError):
        return False
