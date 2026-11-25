# -*- coding: utf-8 -*-
"""빈 행(No.9/No.10) 원인 추적: 전체 파이프라인 실행"""
import sys
sys.path.insert(0, r"C:\AI_Agent\quotation-form-converter")
from parsers.excel_parser import parse_excel
from converters.rule_mapper import RuleMapper
from classification import classify_items
from io import BytesIO

fpath = (r"D:\견적분석\견적 샘플\ESMI1\10. U56A 형교환"
         r"\(나인테크) Quotation_ESMI1동 LAMI U56A 형개발 건_261001_예산용 가견적_111.xlsx")

with open(fpath, "rb") as f:
    data = f.read()

# Step 1: 파싱
items = parse_excel(BytesIO(data))
print(f"[파싱] 총 항목: {len(items)}")

# Step 2: 매핑
mapper = RuleMapper({})
mapped = mapper.map_items(items)
print(f"[매핑] 총 항목: {len(mapped)}")
for i, it in enumerate(mapped):
    print(f"  {i+1:2d}. _행={it.get('_행')} 품명={it.get('품명')!r} "
          f"규격={it.get('규격')!r} 수량={it.get('수량')!r} "
          f"단가={it.get('단가')!r} 금액={it.get('금액')!r} "
          f"_합산={it.get('_합산')} _소계={it.get('_소계')} "
          f"_분류={it.get('_분류')!r}")

# Step 3: 분류
mapped = classify_items(mapped)
print(f"\n[분류 후]")
for i, it in enumerate(mapped):
    print(f"  {i+1:2d}. 품명={str(it.get('품명',''))[:30]:30s} "
          f"_분류={str(it.get('_분류','')):20s} "
          f"_합산={it.get('_합산')} _소계={it.get('_소계')}")

# Step 4: exporter의 _is_meaningful 필터 시뮬레이션
print(f"\n[exporter 필터 시뮬레이션]")
for i, it in enumerate(mapped):
    if it.get("_소계"):
        print(f"  {i+1:2d}. 품명={str(it.get('품명',''))[:30]:30s} → 소계 제외")
        continue
    meaningful = False
    for k in ("품명", "규격", "Maker", "수량", "단가", "금액"):
        v = it.get(k)
        if v not in (None, "", 0, 0.0):
            meaningful = True
            break
    status = "유지" if meaningful else "제외"
    print(f"  {i+1:2d}. 품명={str(it.get('품명',''))[:30]:30s} → {status}")
