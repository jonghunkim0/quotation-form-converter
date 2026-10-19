"""원본 (5)전장배선 Unit 행(row 2163)과 주변 확인."""
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from parsers import parse_excel

RAW = r"D:\견적분석\풍산 견적_AZS 각형 VER1.0(260602).xlsx"
with open(RAW, "rb") as f:
    items = parse_excel(f)
for it in items:
    r = it.get("row_index")
    if r is not None and 2155 <= r <= 2175:
        print(f"row={r}: {it.get('raw_data')}")
