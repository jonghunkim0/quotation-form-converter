"""파이프라인 단계별 시간 측정."""
import sys, time
from pathlib import Path
ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from parsers import parse_excel
from converters import RuleMapper, export_standard_excel
from classification import classify_items

RAW = r"D:\견적분석\풍산 견적_AZS 각형 VER1.0(260602).xlsx"

t0 = time.time()
with open(RAW, "rb") as _f:
    parsed_items = parse_excel(_f)
mapper = RuleMapper({})
mapped_items = mapper.map_items(parsed_items)
t1 = time.time()
print(f"parse+map: {t1-t0:.1f}s")

from quotation_analyzer import ExcelParser, AutoClassifier
a_parser = ExcelParser()
sheets = a_parser.list_sheets(RAW)
analyzer_qf = a_parser.parse_sheet(RAW, sheets[0], Path(RAW).name, "견적가")
a_classifier = AutoClassifier()
a_classifier.classify_all(analyzer_qf.items)
for it in analyzer_qf.items:
    it.final_category = it.auto_category
t2 = time.time()
print(f"analyzer: {t2-t1:.1f}s")

mapped_items = classify_items(mapped_items)
t3 = time.time()
print(f"classify: {t3-t2:.1f}s")

out = export_standard_excel(
    mapped_items, company_name="풍산", quote_no="TEST",
    template_file=str(ROOT / "templates" / "고객사제출_표준양식.xlsx"),
)
t4 = time.time()
print(f"export: {t4-t3:.1f}s")

out_path = ROOT / "data" / "output" / "_profile_check.xlsx"
out_path.write_bytes(out)
t5 = time.time()
print(f"write: {t5-t4:.1f}s")
print(f"TOTAL: {t5-t0:.1f}s")
