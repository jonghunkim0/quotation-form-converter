"""row 2167-2170 원본 데이터 확인"""
import csv
rows = list(csv.reader(open("data/analyzer_export.csv", encoding="utf-8-sig")))
hdr = rows[0]
print("HDR:", hdr)
for r in rows[1:]:
    if len(r) > 14 and r[3] in ("2166", "2167", "2168", "2169", "2170"):
        print(f"row={r[3]} name={r[4]!r} qty={r[6]!r} price={r[7]!r} amt={r[8]!r} auto={r[11]!r} final={r[14]!r}")
