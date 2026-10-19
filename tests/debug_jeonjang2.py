"""(5)전장배선 항목의 분류/배치 추적."""
import sys, ast, json
from pathlib import Path
ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

src = (ROOT / "app.py").read_text(encoding="utf-8")
tree = ast.parse(src)
WANT = {"_ANALYZER_TO_SECTION", "_apply_analyzer_classification",
        "_adjust_family_by_gubun", "_reinherit_unit_classification",
        "_collect_labor_data"}
ns = {"__name__": "app_extract"}
for node in tree.body:
    if isinstance(node, ast.Assign) and any(
            isinstance(t, ast.Name) and t.id in WANT for t in node.targets):
        exec(compile(ast.Module([node], []), "app.py", "exec"), ns)
    elif isinstance(node, ast.FunctionDef) and node.name in WANT:
        exec(compile(ast.Module([node], []), "app.py", "exec"), ns)

from parsers import parse_excel
from converters import RuleMapper
from classification import classify_items
from quotation_analyzer import ExcelParser, AutoClassifier

RAW = r"D:\견적분석\풍산 견적_AZS 각형 VER1.0(260602).xlsx"
with open(RAW, "rb") as _f:
    parsed_items = parse_excel(_f)
mapped_items = RuleMapper({}).map_items(parsed_items)

a_parser = ExcelParser()
sheets = a_parser.list_sheets(RAW)
analyzer_qf = a_parser.parse_sheet(RAW, sheets[0], Path(RAW).name, "견적가")
a_classifier = AutoClassifier()
a_classifier.classify_all(analyzer_qf.items)
for it in analyzer_qf.items:
    it.final_category = it.auto_category

mapped_items = classify_items(mapped_items)
ns["_apply_analyzer_classification"](mapped_items, analyzer_qf)
ns["_adjust_family_by_gubun"](mapped_items)
ns["_reinherit_unit_classification"](mapped_items)

print("=== '전장배선' 포함 매핑 항목 ===")
for it in mapped_items:
    code = str(it.get("품목코드", ""))
    name = str(it.get("품명", ""))
    if "전장배선" in code or "전장배선" in name:
        print(f"  _합산={it.get('_합산')} _분류={it.get('_분류')!r} "
              f"_분류_이유={it.get('_분류_이유')!r} 품목코드={code!r} "
              f"단가={it.get('단가')!r} 금액={it.get('금액')!r}")

print("\n=== analyzer 항목 중 전장배선 ===")
for it in analyzer_qf.items:
    nm = str(getattr(it, "name", ""))
    if "전장배선" in nm:
        print(f"  name={nm!r} auto_category={it.auto_category!r} "
              f"final={it.final_category!r} amount={it.amount!r}")
