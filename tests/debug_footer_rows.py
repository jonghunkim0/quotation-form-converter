# -*- coding: utf-8 -*-
"""원본 파일 푸터(마지막 행) 구조 비교 — 기업이윤/최종금액 행의 컬럼 분포"""
import sys
sys.path.insert(0, r"C:\AI_Agent\quotation-form-converter")
import pandas as pd
from parsers.excel_parser import parse_excel
from io import BytesIO

files = {
    "풍산_AZS(이전 정상)": r"D:\견적분석\풍산 견적_AZS 각형 VER1.0(260602)_1.xlsx",
    "U56A(현재 문제)": (r"D:\견적분석\견적 샘플\ESMI1\10. U56A 형교환"
                        r"\(나인테크) Quotation_ESMI1동 LAMI U56A 형개발 건_261001_예산용 가견적_111.xlsx"),
}

for label, fpath in files.items():
    print(f"=== {label} ===")
    with open(fpath, "rb") as f:
        data = f.read()
    items = parse_excel(BytesIO(data))
    print(f"  파싱 항목 수: {len(items)}")
    # 푸터 후보: 품명/규격/단가가 비어있는 항목
    for it in items:
        rd = it["raw_data"]
        name = rd.get("품 명") or rd.get("품명") or rd.get("자재명") or ""
        spec = rd.get("규 격") or rd.get("규격") or ""
        price = rd.get("단가")
        if not str(name).strip() and not str(spec).strip() and price in (None, ""):
            print(f"  [푸터 후보] row={it['row_index']} raw={rd}")
    print()
