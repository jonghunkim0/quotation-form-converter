"""병합 겹침(overlap) + drawing 참조 정합성 진단."""
import zipfile
import re
from pathlib import Path
from openpyxl import load_workbook
from openpyxl.utils import range_boundaries

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "data" / "output" / "_labor_check.xlsx"
TPL = ROOT / "templates" / "고객사제출_표준양식.xlsx"

wb = load_workbook(OUT)
print("=== 1) 병합 겹침(overlap) 검사 ===")
total_overlap = 0
for ws in wb.worksheets:
    ranges = []
    for m in ws.merged_cells.ranges:
        c1, r1, c2, r2 = range_boundaries(str(m))
        ranges.append((r1, c1, r2, c2, str(m)))
    # 겹침 검사 (O(n^2) — 시트당 병합 수 적음)
    for i in range(len(ranges)):
        for j in range(i + 1, len(ranges)):
            r1a, c1a, r2a, c2a, sa = ranges[i]
            r1b, c1b, r2b, c2b, sb = ranges[j]
            # 겹침: 행 범위 겹치고 열 범위 겹침
            if not (r2a < r1b or r2b < r1a or c2a < c1b or c2b < c1a):
                print(f"  [OVERLAP] {ws.title}: {sa} <-> {sb}")
                total_overlap += 1
print(f"총 겹침: {total_overlap}")

print("\n=== 2) drawing 참조 정합성 ===")
with zipfile.ZipFile(OUT) as z:
    names = z.namelist()
    # 시트 rels에서 drawing 참조
    for n in names:
        if re.match(r"xl/worksheets/_rels/.*\.rels", n):
            content = z.read(n).decode("utf-8")
            for m in re.finditer(r'Target="([^"]*drawing[^"]*)"', content):
                target = m.group(1)
                # 상대 경로 → 절대
                base = n.replace("_rels/", "").replace(".rels", "")
                abs_target = base + target
                # normalize ../
                parts = abs_target.split("/")
                out = []
                for p in parts:
                    if p == "..":
                        out.pop()
                    elif p != ".":
                        out.append(p)
                abs_target = "/".join(out)
                exists = abs_target in names
                print(f"  {n} -> {target} => {abs_target} 존재={exists}")
    # drawing 파일 존재
    print("  drawing 파일:", [n for n in names if "drawing" in n])

print("\n=== 3) 템플릿 시트 rels (drawing 참조 원본) ===")
with zipfile.ZipFile(TPL) as z:
    for n in z.namelist():
        if re.match(r"xl/worksheets/_rels/.*\.rels", n):
            content = z.read(n).decode("utf-8")
            if "drawing" in content:
                print(f"  {n}:")
                for m in re.finditer(r'<Relationship[^>]*>', content):
                    print(f"    {m.group(0)}")
