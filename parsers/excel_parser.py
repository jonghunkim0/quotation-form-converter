"""
Excel 견적서 파서
업체별 Excel 견적서를 읽어 항목 목록으로 추출
"""

import pandas as pd
import re
from io import BytesIO
from typing import List, Dict, Any, Optional


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
                    value = row.iloc[col_idx]
                    if pd.notna(value):
                        raw_data[str(header).strip()] = value
            
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


def _is_number(value) -> bool:
    """값이 숫자인지 확인"""
    try:
        float(str(value).replace(",", "").replace("₩", "").replace("$", "").strip())
        return True
    except (ValueError, TypeError):
        return False
