"""원본 파일의 빈 Unit 행(row=740) 주변 확인."""
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from parsers import parse_excel
RAW = r"D:\견적분석\풍산 견적_AZS 각형 VER1.0(260602).xlsx"
with open(RAW, "rb") as f:
    items = parse_excel(f)
# row=740 주변
print("=== 원본 row 735~745 ===")
for it in items:
    r = it.get("_행") or it.get("row") or it.get("source_row")
    if r is None:
        continue
    try:
        r = int(r)
    except (ValueError, TypeError):
        continue
    if 735 <= r <= 745:
        print(f"  row={r} 합산={it.get('_합산')} 품명={it.get('품명')!r} 수량={it.get('수량')!r} 단가={it.get('단가')!r} 금액={it.get('금액')!r}")
