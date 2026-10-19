import json

d = json.load(open("data/_debug_export.json", "r", encoding="utf-8"))
print("section_counts:", json.dumps(d["section_counts"], ensure_ascii=False, indent=1))
units = [i for i in d["items"] if i["unit"]]
print(f"unit rows: {len(units)}")
for u in units:
    print(f"  row={u['row']} name={u['unit_name']} section={u['section']}")

# 각 Unit의 하위 항목 섹션 분포 확인
items = d["items"]
for idx, it in enumerate(items):
    if not it["unit"]:
        continue
    subs = []
    for j in range(idx + 1, len(items)):
        if items[j]["unit"]:
            break
        subs.append(items[j]["section"])
    from collections import Counter
    print(f"unit row={it['row']} ({it['unit_name']}): 하위 {len(subs)}개, 분포={dict(Counter(subs))}")
