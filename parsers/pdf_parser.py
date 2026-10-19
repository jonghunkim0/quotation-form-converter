"""
PDF 견적서 파서
업체별 PDF 견적서를 읽어 항목 목록으로 추출
"""

import pdfplumber
import re
from io import BytesIO
from typing import List, Dict, Any


def parse_pdf(file) -> List[Dict[str, Any]]:
    """
    PDF 파일을 파싱하여 견적 항목 목록을 반환.
    
    Args:
        file: 업로드된 파일 객체 (BytesIO 또는 파일 핸들)
        
    Returns:
        [{"raw_data": {...}, "page": int, "row_index": int}, ...]
    """
    items = []
    
    # 파일 읽기
    if hasattr(file, "read"):
        data = file.read()
        file = BytesIO(data)
    
    with pdfplumber.open(file) as pdf:
        for page_num, page in enumerate(pdf.pages, 1):
            # 테이블 추출 시도
            tables = page.extract_tables()
            
            for table in tables:
                if not table or len(table) < 2:
                    continue
                
                # 헤더 행 찾기
                header_row, headers = _find_pdf_header(table)
                if header_row is None:
                    continue
                
                # 데이터 행 추출
                for row_idx in range(header_row + 1, len(table)):
                    row = table[row_idx]
                    if _is_empty_pdf_row(row):
                        continue
                    
                    raw_data = {}
                    for col_idx, header in enumerate(headers):
                        if col_idx < len(row) and row[col_idx]:
                            raw_data[header.strip()] = row[col_idx].strip()
                    
                    if raw_data:
                        items.append({
                            "raw_data": raw_data,
                            "page": page_num,
                            "row_index": row_idx,
                        })
            
            # 테이블이 없으면 텍스트 기반 파싱
            if not tables:
                text_items = _parse_pdf_text(page.extract_text() or "")
                for item in text_items:
                    item["page"] = page_num
                    items.append(item)
    
    return items


def _find_pdf_header(table: list) -> tuple:
    """PDF 테이블에서 헤더 행 감지"""
    best_row = None
    best_count = 0
    
    search_range = min(5, len(table))
    
    for i in range(search_range):
        row = table[i]
        str_count = sum(
            1 for v in row 
            if v and isinstance(v, str) and v.strip() and not _is_number(v)
        )
        if str_count > best_count:
            best_count = str_count
            best_row = i
    
    if best_row is None or best_count < 2:
        return None, None
    
    headers = []
    for v in table[best_row]:
        headers.append(v.strip() if v else "")
    
    return best_row, headers


def _is_empty_pdf_row(row: list) -> bool:
    """PDF 테이블 행이 빈 행인지 확인"""
    return all(v is None or str(v).strip() == "" for v in row)


def _is_number(value) -> bool:
    """값이 숫자인지 확인"""
    try:
        float(str(value).replace(",", "").replace("₩", "").replace("$", "").strip())
        return True
    except (ValueError, TypeError):
        return False


def _parse_pdf_text(text: str) -> List[Dict[str, Any]]:
    """
    텍스트 기반 PDF 파싱 (테이블이 없는 경우)
    간단한 패턴 매칭으로 항목 추출
    """
    items = []
    lines = text.split("\n")
    
    for line in lines:
        line = line.strip()
        if not line:
            continue
        
        # 숫자 패턴이 포함된 행 (품번 + 품명 + 수량 + 단가 + 금액 형태)
        # 예: "ABC-001  볼트 M6  100  500  50000"
        match = re.match(
            r"([A-Za-z0-9\-]+)\s+(.+?)\s+(\d+)\s+([\d,]+)\s+([\d,]+)",
            line
        )
        if match:
            items.append({
                "raw_data": {
                    "품번": match.group(1),
                    "품명": match.group(2),
                    "수량": match.group(3),
                    "단가": match.group(4),
                    "금액": match.group(5),
                },
                "row_index": 0,
            })
    
    return items
