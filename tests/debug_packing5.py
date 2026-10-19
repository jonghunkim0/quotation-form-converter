"""포장/국내 운송 Unit의 Pass 2 분할 로직 추적"""
import sys, ast
from pathlib import Path
ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

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
from classification import classify_items, TEMPLATE_SECTIONS
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

# Pass 2 로직 재현
known_sections = {s[0] for s in TEMPLATE_SECTIONS}
filtered_items = [item for item in mapped_items if not item.get("_소계")]

# 포장/국내 운송 Unit 찾기
unit_idx = None
for i, item in enumerate(filtered_items):
    if item.get("_합산") and "포장" in str(item.get("품명", "")):
        unit_idx = i
        break

if unit_idx is None:
    # _unit_name으로 찾기
    for i, item in enumerate(filtered_items):
        if item.get("_합산") and "포장" in str(item.get("_unit_name", "")):
            unit_idx = i
            break

if unit_idx is None:
    print("포장 Unit not found!")
    # 모든 Unit 출력
    for i, item in enumerate(filtered_items):
        if item.get("_합산"):
            print(f"  idx={i} _행={item.get('_행')} 품명={str(item.get('품명',''))[:30]!r} _unit_name={item.get('_unit_name','')!r} _분류={item.get('_분류','')!r}")
else:
    unit = filtered_items[unit_idx]
    print(f"Unit found at idx={unit_idx}: _행={unit.get('_행')} 품명={str(unit.get('품명',''))[:30]!r} _분류={unit.get('_분류','')!r}")
    print(f"  known_sections = {known_sections}")
    print(f"  Unit _분류 in known_sections: {unit.get('_분류','') in known_sections}")
    
    # 하위 항목 추적
    first_idx = {}
    for j in range(unit_idx + 1, len(filtered_items)):
        sub = filtered_items[j]
        if sub.get("_합산"):
            print(f"  [STOP] next Unit at idx={j}: 품명={str(sub.get('품명',''))[:30]!r}")
            break
        sec = sub.get("_분류", "")
        in_known = sec in known_sections
        if in_known and sec not in first_idx:
            first_idx[sec] = j
        print(f"  sub idx={j} _행={sub.get('_행')} 품명={str(sub.get('품명',''))[:25]!r} _분류={sec!r} in_known={in_known}")
    
    print(f"\n  first_idx = {first_idx}")
    if first_idx:
        print(f"  → Unit 분할 대상 섹션: {list(first_idx.keys())}")
    else:
        sec = unit.get("_분류", "")
        fallback = sec if sec in known_sections else "기구_구매품"
        print(f"  → 폴백: _분류={sec!r} → {fallback}")
