"""Short Tester/Light Curtain 원본 위치 확인 (R771 Unit 추적)."""
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from parsers import parse_excel

RAW = r"D:\견적분석\풍산 견적_AZS 각형 VER1.0(260602).xlsx"
with open(RAW, "rb") as f:
    items = parse_excel(f)
for it in items:
    rd = it.get("raw_data", {})
    name = str(rd.get("품 명", "")) + str(rd.get("품명", ""))
    if "Short Tester" in name or "Light Curtain" in name or "Ionizer" in name:
        print(f"row={it.get('row_index')}: {rd}")
