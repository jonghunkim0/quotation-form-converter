"""포장/국내 운송 Unit의 실제 배치 섹션 추적"""
import sys, ast
from pathlib import Path
ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

# app.py 함수 추출
src = (ROOT / "app.py").read_text(encoding="utf-8")
tree = ast.parse(src)
WANT = {"_ANALYZER_TO_SECTION", "_apply_analyzer_classification",
        "_adjust_family_by_gubun", "_reinherit_unit_classification"}
ns = {"__name__": "app_extract"}
for node in tree.body:
    if isinstance(node, ast.Assign) and any(isinstance(t, ast.Name) and t.id in WANT for t in node.targets):
        exec(compile(ast.Module([node], []), "app.py", "exec"), ns)
    elif isinstance(node, ast.FunctionDef) and node.name in WANT:
        exec(compile(ast.Module([node], []), "app.py", "exec"), ns)

from parsers import parse_excel
from converters import RuleMapper
from classification import classify_items
from quotation_analyzer import ExcelParser, AutoClassifier

RAW = r"D:\견적분석\풍산 견적_AZS 각형 VER1.0(260602).xlsx"
with open(RAW, "rb") as f:
    parsed_items = parse_excel(f)
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

# 포장/국내 운송 Unit 및 하위 항목
print("=== 포장/국내 운송 Unit 및 하위 ===")
for i, m in enumerate(mapped_items):
    nm = str(m.get("품명", ""))
    if "포장" in nm or "운송" in nm:
        print(f"[UNIT] _행={m.get('_행')} 품명={nm!r} _합산={m.get('_합산')} _분류={m.get('_분류')!r} _unit_name={m.get('_unit_name')!r}")
        # 하위 항목
        for j in range(i + 1, min(i + 8, len(mapped_items))):
            m2 = mapped_items[j]
            if m2.get("_합산"):
                break
            print(f"   sub _행={m2.get('_행')} 품명={str(m2.get('품명',''))[:20]!r} _분류={m2.get('_분류')!r} _구분={m2.get('_구분')!r}")
        break
