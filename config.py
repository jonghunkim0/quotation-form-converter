"""
견적 표준 양식 변환기 - 설정 파일
"""

import os
from pathlib import Path

# ─── LLM (Qwen3.8 / DashScope) ─────────────────────────────────
DASHSCOPE_API_KEY = os.environ.get(
    "DASHSCOPE_API_KEY",
    "sk-9370684cc9794f03b2b951b02bab8d98",
)

# ─── 표준 출력 컬럼 ─────────────────────────────────────────────
STANDARD_COLUMNS = [
    "업체명",
    "견적일",
    "품목코드",
    "품명",
    "규격",
    "Maker",
    "수량",
    "단위",
    "단가",
    "금액",
    "납기",
    "비고",
    "검증결과",
]

# 원본명 표시 옵션 (기본: False)
SHOW_ORIGINAL_NAME = False

# ─── 컬럼 매핑 규칙 ─────────────────────────────────────────────
# 원본 컬럼명 → 표준 컬럼명
COLUMN_MAPPING_RULES = {
    # 품목코드
    "품번": "품목코드",
    "part no": "품목코드",
    "part no.": "품목코드",
    "item code": "품목코드",
    "item no": "품목코드",
    "item no.": "품목코드",
    "item number": "품목코드",
    "품목번호": "품목코드",
    "코드": "품목코드",
    "no": "품목코드",
    "no.": "품목코드",
    "#": "품목코드",
    # 품명 (D열 자재명) - Raw 견적에서 품명/부품명/자재명 등 다양한 표현 대응
    "품명": "품명",
    "부품명": "품명",
    "자재명": "품명",
    "재질명": "품명",
    "description": "품명",
    "item name": "품명",
    "item": "품명",
    "item description": "품명",
    "part name": "품명",
    "product name": "품명",
    "product": "품명",
    "명칭": "품명",
    "품목명": "품명",
    "제품명": "품명",
    "내용": "품명",
    # 규격
    "규격": "규격",
    "spec": "규격",
    "specification": "규격",
    "spec.": "규격",
    "size": "규격",
    "dimension": "규격",
    "dimensions": "규격",
    "치수": "규격",
    "재질": "규격",
    "material": "규격",
    # Maker (F열) - 협력사/제조사/공급사/Maker 등
    "maker": "Maker",
    "maker name": "Maker",
    "협력사": "Maker",
    "협력업체": "Maker",
    "제조사": "Maker",
    "공급사": "Maker",
    "공급업체": "Maker",
    "manufacturer": "Maker",
    "brand": "Maker",
    # 수량
    "ea": "수량",
    "qty": "수량",
    "quantity": "수량",
    "수량": "수량",
    "count": "수량",
    "개수": "수량",
    # 단위
    "단위": "단위",
    "unit": "단위",
    "uom": "단위",
    "unit of measure": "단위",
    # 단가
    "단가": "단가",
    "unit price": "단가",
    "price": "단가",
    "unit cost": "단가",
    "단가(원)": "단가",
    "단가($)": "단가",
    # 금액
    "amount": "금액",
    "total price": "금액",
    "금액": "금액",
    "total": "금액",
    "total amount": "금액",
    "합계": "금액",
    "price (total)": "금액",
    "금액(원)": "금액",
    "금액($)": "금액",
    # 납기
    "lead time": "납기",
    "납기": "납기",
    "delivery": "납기",
    "delivery time": "납기",
    "交期": "납기",
    "납기일": "납기",
    # 비고
    "비고": "비고",
    "remark": "비고",
    "remarks": "비고",
    "note": "비고",
    "notes": "비고",
    "comment": "비고",
    "comments": "비고",
    "memo": "비고",
    "특기사항": "비고",
    # 업체명 (전체 견적 업체 - 요약 시트용)
    "업체명": "업체명",
    "vendor": "업체명",
    "supplier": "업체명",
    "company": "업체명",
    "회사명": "업체명",
    # 견적일
    "견적일": "견적일",
    "date": "견적일",
    "quote date": "견적일",
    "견적일자": "견적일",
    "일 자": "견적일",
    "일자": "견적일",
}

# ─── 단위 정규화 규칙 ───────────────────────────────────────────
UNIT_NORMALIZATION = {
    "pcs": "EA",
    "pc": "EA",
    "ea": "EA",
    "개": "EA",
    "개/box": "EA",
    "mm": "mm",
    "MM": "mm",
    "Mm": "mm",
    "m": "m",
    "cm": "cm",
    "kg": "kg",
    "KG": "kg",
    "g": "g",
    "set": "SET",
    "set.": "SET",
    "세트": "SET",
    "lot": "LOT",
    "roll": "ROLL",
    "m²": "m²",
    "sqm": "m²",
}

# ─── 금액 파싱 규칙 ─────────────────────────────────────────────
CURRENCY_SYMBOLS = ["₩", "¥", "$", "€", "£"]
CURRENCY_CODES = {
    "KRW": "₩",
    "USD": "$",
    "EUR": "€",
    "JPY": "¥",
    "GBP": "£",
}

# ─── 파일 경로 ──────────────────────────────────────────────────
BASE_DIR = Path(__file__).parent
DATA_DIR = BASE_DIR / "data"
TEMPLATE_DIR = BASE_DIR / "templates"
OUTPUT_DIR = DATA_DIR / "output"

# 표준 양식 템플릿 파일 (고객사 표준 양식)
DEFAULT_TEMPLATE_FILE = TEMPLATE_DIR / "고객사제출_표준양식.xlsx"

# 매핑 규칙 파일
RULES_FILE = DATA_DIR / "mapping_rules.json"

# ─── 신뢰도 임계값 ──────────────────────────────────────────────
CONFIDENCE_THRESHOLD = 0.8  # 80% 미만 → "사용자 확인 필요"
