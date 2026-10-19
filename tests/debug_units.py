"""Unit 행 구조 확인"""
import json
d = json.load(open("data/_debug_export.json", encoding="utf-8"))
for i in d["items"]:
    if i["unit"]:
        print(f"row={i['row']} unit_name={i['unit_name']!r} bunryu={i['bunryu']!r} section={i['section']!r}")
