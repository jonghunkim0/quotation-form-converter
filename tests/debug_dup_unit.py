"""전장_가공품 섹션의 Unit 항목 확인 (중복 Unit 원인)."""
import json
from pathlib import Path
ROOT = Path(__file__).resolve().parent.parent
d = json.load(open(ROOT / "data" / "_debug_export.json", encoding="utf-8"))
print("=== 전장_가공품 섹션 항목 (앞 25개) ===")
cnt = 0
for it in d["items"]:
    if it["section"] == "전장_가공품":
        tag = "UNIT" if it["unit"] else "    "
        print(f"  {tag} row={it['row']} name={it['name']!r}")
        cnt += 1
        if cnt > 25:
            break
