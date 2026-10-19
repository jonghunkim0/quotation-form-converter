"""포장/국내 운송 Unit 하위 항목 분류 확인"""
import json
d = json.load(open("data/_debug_export.json", encoding="utf-8"))
items = d["items"]
# row 2167 (포장/국내 운송 Unit) 이후 다음 Unit까지
start = None
for i, it in enumerate(items):
    if it["row"] == 2167:
        start = i
        break
if start is not None:
    for it in items[start:start + 12]:
        print(f"row={it['row']} unit={it['unit']} name={it['name']!r} bunryu={it['bunryu']!r} section={it['section']!r}")
print()
print("section_counts:", json.dumps(d["section_counts"], ensure_ascii=False))
