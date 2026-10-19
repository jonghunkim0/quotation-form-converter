"""인건비 섹션에 들어온 Unit(합산) 항목 확인."""
import json
from pathlib import Path
ROOT = Path(__file__).resolve().parent.parent
d = json.load(open(ROOT / "data" / "_debug_export.json", encoding="utf-8"))
print("section_counts:")
for k, v in d["section_counts"].items():
    print(f"  {k}: {v}")
print()
print("=== 인건비/경비 계열 섹션의 Unit(합산) 항목 ===")
LABOR_SECS = ("설계", "조립", "셋업", "양산대응", "특수제어", "경비", "포장", "기타항목")
for it in d["items"]:
    if it["unit"] and it["section"] in LABOR_SECS:
        print(f"  [{it['section']}] row={it['row']} name={it['name']!r} bunryu={it['bunryu']!r}")
