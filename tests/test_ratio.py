"""견적 조정 모듈 테스트: 실제 파일 파이프라인으로 RatioModel 구성 검증."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from parsers import parse_excel
from converters import RuleMapper
from classification import classify_items
from app import _apply_analyzer_classification, _adjust_family_by_gubun, _reinherit_unit_classification, _collect_labor_data
from converters.ratio_adjuster import build_ratio_model, apply_adjustment, SECTION_ORDER

SRC = r"D:\견적분석\풍산 견적_AZS 각형 VER1.0(260602).xlsx"

# ── 파이프라인 실행 ──
import io
buf = io.BytesIO(Path(SRC).read_bytes())
buf.name = Path(SRC).name
parsed = parse_excel(buf)
rules = {}
mapped = RuleMapper(rules).map_items(parsed)

# analyzer
from quotation_analyzer import ExcelParser, AutoClassifier
import tempfile, os
tmp = os.path.join(tempfile.gettempdir(), "_ratio_test_src.xlsx")
Path(tmp).write_bytes(Path(SRC).read_bytes())
a_parser = ExcelParser()
sheets = a_parser.list_sheets(tmp)
qf = a_parser.parse_sheet(tmp, sheets[0], Path(SRC).name, "견적가")
AutoClassifier().classify_all(qf.items)
for it in qf.items:
    it.final_category = it.auto_category
os.unlink(tmp)

mapped = classify_items(mapped)
_apply_analyzer_classification(mapped, qf)
_adjust_family_by_gubun(mapped)
_reinherit_unit_classification(mapped)
labor_data = _collect_labor_data(qf)

# ── 모델 구성 ──
model = build_ratio_model(labor_data, mapped)

print("=== 인건비 행 ===")
for r in model.labor_rows:
    print(f"  {r.section} {r.job}: rate={r.rate:,.0f} cur=({r.cur_people}명×{r.cur_days}일) "
          f"MD={r.cur_md:,.0f} 금액={r.cur_amount:,.0f}")

print("\n=== 재료비 ===")
print(f"  가공품: {model.mat_made} 합={model.mat_total('made'):,.0f}")
print(f"  구매품: {model.mat_buy} 합={model.mat_total('buy'):,.0f}")
print(f"  경비: {model.expense:,.0f}")

print("\n=== 섹션 합계 (현견적) ===")
for s in SECTION_ORDER:
    print(f"  {s}: {model.section_sum(s, 'cur'):,.0f}")
print(f"  인건비 합계: {model.labor_total('cur'):,.0f}")
print(f"  총 견적: {model.grand_total('cur'):,.0f}")

# ── 조정 적용 테스트 1: 가공품 비율(%) ──
print("\n=== 조정 테스트 1: 가공품 기구부 80% (20% 절감) ===")
base_total = model.grand_total("cur")
model.made_pct["기구부"] = 80.0
adj_made = model.made_adj("기구부")
print(f"  기구부 가공품: cur={model.mat_made['기구부']:,.0f} adj={adj_made:,.0f} "
      f"증감={adj_made - model.mat_made['기구부']:+,.0f}")
assert abs(adj_made - model.mat_made["기구부"] * 0.8) < 1e-6
assert model.grand_total("adj") < base_total, "80% 적용 시 총액 감소해야 함"

# ── 조정 적용 테스트 2: 경비 비율(%) ──
print("\n=== 조정 테스트 2: 경비 120% ===")
model.expense_pct = 120.0
assert abs(model.expense_adj() - model.expense * 1.2) < 1e-6

# ── 조정 적용 테스트 3: 인건비 인원 2배 ──
print("\n=== 조정 테스트 3: (1) 설계 모든 행 인원 2배 ===")
model.made_pct["기구부"] = 100.0  # 리셋
model.expense_pct = 100.0
inputs = {}
for r in model.labor_rows:
    if r.section == "(1) 설계":
        inputs[f"{r.section}|{r.job}|people"] = r.adj_people * 2
apply_adjustment(model, inputs)
for s in SECTION_ORDER:
    print(f"  {s}: cur={model.section_sum(s, 'cur'):,.0f} adj={model.section_sum(s, 'adj'):,.0f}")
print(f"  총액: cur={model.grand_total('cur'):,.0f} adj={model.grand_total('adj'):,.0f} "
      f"증감={model.grand_total('adj')-model.grand_total('cur'):,.0f}")

# ── 검증 ──
assert model.labor_rows, "인건비 행이 없음"
# MD = 인원×기간, 금액 = MD×단가
for r in model.labor_rows:
    assert abs(r.cur_md - r.cur_people * r.cur_days) < 1e-9
    assert abs(r.cur_amount - r.cur_md * r.rate) < 1e-9
# 조정 후 총액이 증가해야 함 (인원 2배)
assert model.grand_total("adj") > model.grand_total("cur"), "조정 후 총액이 증가해야 함"
print("\nALL PASS")
