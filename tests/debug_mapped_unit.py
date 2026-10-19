"""매핑된 Unit 항목의 _unit_name 확인."""
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from parsers import parse_excel
from converters import RuleMapper

RAW = r"D:\견적분석\풍산 견적_AZS 각형 VER1.0(260602).xlsx"
with open(RAW, "rb") as f:
    parsed = parse_excel(f)
mapper = RuleMapper({})
mapped = mapper.map_items(parsed)

print("=== 품명이 비어있는 Unit 항목 (매핑 후) ===")
for it in mapped:
    if it.get("_합산") and not str(it.get("품명", "")).strip():
        print(f"  row={it.get('_행')} 품목코드={it.get('품목코드')!r} _unit_name={it.get('_unit_name')!r} 품명={it.get('품명')!r} 단위={it.get('단위')!r}")
