"""품명이 비어있는 Unit(합산) 항목 확인."""
import json
from pathlib import Path
ROOT = Path(__file__).resolve().parent.parent
d = json.load(open(ROOT / "data" / "_debug_export.json", encoding="utf-8"))
print("=== 품명이 비어있는 Unit(합산) 항목 ===")
for it in d["items"]:
    if it["unit"] and not it["name"].strip():
        print(f"  [{it['section']}] row={it['row']} name={it['name']!r} bunryu={it['bunryu']!r}")
print()
print("=== 기구_구매품 섹션의 Unit 항목 (마지막 5개) ===")
units = [it for it in d["items"] if it["unit"] and it["section"] == "기구_구매품"]
for it in units[-5:]:
    print(f"  row={it['row']} name={it['name']!r} bunryu={it['bunryu']!r}")
