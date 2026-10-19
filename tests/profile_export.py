"""export 내부 단계별 시간 측정 (instrumented)."""
import sys, time
from pathlib import Path
ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import converters.standard_exporter as se

# 원본 함수 래핑으로 타이밍 측정
_orig_insert = se._insert_rows_with_remap
_orig_delete = se._delete_row_with_remap
_orig_fill_labor = se._fill_labor_rows
stats = {"insert": 0.0, "insert_n": 0, "delete": 0.0, "delete_n": 0, "labor": 0.0}

def _ti(ws, ip, count, style_ref_row=None):
    t = time.time()
    r = _orig_insert(ws, ip, count, style_ref_row)
    stats["insert"] += time.time() - t
    stats["insert_n"] += 1
    return r

def _td(ws, row):
    t = time.time()
    r = _orig_delete(ws, row)
    stats["delete"] += time.time() - t
    stats["delete_n"] += 1
    return r

def _tl(ws, labor_data, shift):
    t = time.time()
    r = _orig_fill_labor(ws, labor_data, shift)
    stats["labor"] += time.time() - t
    return r

se._insert_rows_with_remap = _ti
se._delete_row_with_remap = _td
se._fill_labor_rows = _tl

from parsers import parse_excel
from converters import RuleMapper, export_standard_excel
from classification import classify_items

RAW = r"D:\견적분석\풍산 견적_AZS 각형 VER1.0(260602).xlsx"
with open(RAW, "rb") as _f:
    parsed_items = parse_excel(_f)
mapped_items = RuleMapper({}).map_items(parsed_items)
mapped_items = classify_items(mapped_items)

t0 = time.time()
out = export_standard_excel(
    mapped_items, company_name="풍산", quote_no="TEST",
    template_file=str(ROOT / "templates" / "고객사제출_표준양식.xlsx"),
)
t1 = time.time()
print(f"export total: {t1-t0:.1f}s")
print(f"  insert: {stats['insert']:.1f}s ({stats['insert_n']} calls)")
print(f"  delete: {stats['delete']:.1f}s ({stats['delete_n']} calls)")
print(f"  fill_labor: {stats['labor']:.1f}s")
print(f"  other: {t1-t0-stats['insert']-stats['delete']-stats['labor']:.1f}s")
