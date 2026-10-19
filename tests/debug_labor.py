"""인건비/경비 항목 디버그"""
import json

d = json.load(open("data/_debug_export.json", "r", encoding="utf-8"))
LABOR = ("설계", "조립", "셋업", "양산대응", "특수제어", "경비", "포장", "기타항목")
print("=== 인건비/경비 항목 ===")
for i in d["items"]:
    if i["bunryu"] in LABOR:
        print(f"  row={i['row']} name={i['name']!r} bunryu={i['bunryu']!r} section={i['section']!r}")
