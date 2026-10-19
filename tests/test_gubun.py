"""Raw 구분 컬럼 기반 계열 재조정 테스트"""
import sys
sys.path.insert(0, ".")
from converters import RuleMapper
from classification import classify_items
import app as app_module

# Raw 견적 시뮬레이션: 구분 컬럼(병합 셀 → forward-fill 필요)
test_items = [
    # 기구부 섹션
    {'raw_data': {'구분': '기구부', 'No': '1)', '품 명': 'MAIN_FRAME', '규격': 'S1', 'Maker': 'PSS', '수량': '1', '단위': 'EA', '단가': '100000', '금액': '100000'}},
    {'raw_data': {'No': '2)', '품 명': 'FOOT_SPACER', '규격': 'S2', 'Maker': 'PSS', '수량': '2', '단위': 'EA', '단가': '50000', '금액': '100000'}},
    # 전장/제어 섹션 (구분은 첫 행에만 - 병합 셀)
    {'raw_data': {'구분': '전장/제어', 'No': '1)', '품 명': 'PLC', '규격': 'AJ65VBCE32-16DT', 'Maker': 'Mitsubishi', '수량': '10', '단위': 'EA', '단가': '0', '금액': '0', '비고': '- 사급'}},
    {'raw_data': {'No': '2)', '품 명': 'Inverter', '규격': 'FR-E820-3.7K', 'Maker': 'Mitsubishi', '수량': '2', '단위': 'EA', '단가': '0', '금액': '0', '비고': '- 사급'}},
    {'raw_data': {'No': '3)', '품 명': 'Touch Panel', '규격': 'GT2712-STBA', 'Maker': 'Mitsubishi', '수량': '1', '단위': 'EA', '단가': '0', '금액': '0', '비고': '- 사급'}},
    {'raw_data': {'No': '4)', '품 명': 'Motor', '규격': 'S250DSR-QX00', 'Maker': 'NIPPON', '수량': '4', '단위': 'EA', '단가': '0', '금액': '0', '비고': '- 사급'}},
    {'raw_data': {'No': '5)', '품 명': 'CABLE', '규격': 'C1', 'Maker': 'MISUMI', '수량': '1', '단위': 'EA', '단가': '0', '금액': '0', '비고': '- 사급'}},
]

mapped = RuleMapper().map_items(test_items)
print("=== 매핑 후 (_구분 forward-fill) ===")
for m in mapped:
    print(f"  품명={m.get('품명')!r:20} _구분={m.get('_구분')!r}")

mapped = classify_items(mapped)
print("\n=== 분류 후 (analyzer 없이) ===")
for m in mapped:
    print(f"  품명={m.get('품명')!r:20} _분류={m.get('_분류')!r}")

# analyzer가 사급품으로 분류했다고 가정 → _분류_이유 확인용 시뮬레이션
# (실제 app.py에서는 _apply_analyzer_classification이 사급품→기구_구매품으로 설정)
for m in mapped:
    if '사급' in str(m.get('비고', '')):
        m['_분류'] = '기구_구매품'  # analyzer 사급품 → 기구_구매품 시뮬레이션

print("\n=== analyzer 사급품 적용 시뮬레이션 ===")
for m in mapped:
    print(f"  품명={m.get('품명')!r:20} _분류={m.get('_분류')!r}")

app_module._adjust_family_by_gubun(mapped)
print("\n=== _adjust_family_by_gubun 적용 후 (구분=전장/제어) ===")
for m in mapped:
    print(f"  품명={m.get('품명')!r:20} _분류={m.get('_분류')!r}  ({m.get('_분류_이유', '')})")

# 검증
ok = True
for m in mapped:
    name = m.get('품명', '')
    sec = m.get('_분류', '')
    if name in ('PLC', 'Inverter', 'Touch Panel'):
        if sec != '제어_구매품':
            print(f"FAIL: {name} → {sec} (제어_구매품 기대)")
            ok = False
    elif name in ('Motor', 'CABLE'):
        if sec != '전장_구매품':
            print(f"FAIL: {name} → {sec} (전장_구매품 기대)")
            ok = False
    elif name in ('MAIN_FRAME', 'FOOT_SPACER'):
        if sec == '제어_구매품' or sec == '전장_구매품':
            print(f"FAIL: {name} → {sec} (기구 계열 유지 기대)")
            ok = False

# ── 시나리오 2: Raw 구분 컬럼이 "기구부"로 되어 있는 경우 (실제 풍산 AZS) ──
print("\n=== 시나리오 2: 구분='기구부' (실제 풍산 AZS) ===")
test_items2 = [
    {'raw_data': {'구분': '기구부', 'No': '1)', '품 명': 'MAIN_FRAME', '규격': 'S1', 'Maker': 'PSS', '수량': '1', '단위': 'EA', '단가': '100000', '금액': '100000'}},
    {'raw_data': {'No': '2)', '품 명': 'PLC', '규격': 'AJ65VBCE32-16DT', 'Maker': 'Mitsubishi', '수량': '10', '단위': 'EA', '단가': '0', '금액': '0', '비고': '- 사급'}},
    {'raw_data': {'No': '3)', '품 명': 'Inverter', '규격': 'FR-E820-3.7K', 'Maker': 'Mitsubishi', '수량': '2', '단위': 'EA', '단가': '0', '금액': '0', '비고': '- 사급'}},
    {'raw_data': {'No': '4)', '품 명': 'Touch Panel', '규격': 'GT2712-STBA', 'Maker': 'Mitsubishi', '수량': '1', '단위': 'EA', '단가': '0', '금액': '0', '비고': '- 사급'}},
    {'raw_data': {'No': '5)', '품 명': 'SAFETY PLC', '규격': 'G9SP-N20S', 'Maker': 'OMRON', '수량': '1', '단위': 'EA', '단가': '0', '금액': '0', '비고': '- 사급'}},
    {'raw_data': {'No': '6)', '품 명': '초소형 감시 SENSOR', '규격': 'E3T-FD11 2M', 'Maker': 'OMRON', '수량': '4', '단위': 'EA', '단가': '27250', '금액': '109000'}},
]
mapped2 = RuleMapper().map_items(test_items2)
mapped2 = classify_items(mapped2)
# analyzer 사급품 → 기구_구매품 시뮬레이션
for m in mapped2:
    if '사급' in str(m.get('비고', '')):
        m['_분류'] = '기구_구매품'
app_module._adjust_family_by_gubun(mapped2)
for m in mapped2:
    print(f"  품명={m.get('품명')!r:20} _구분={m.get('_구분')!r:8} _분류={m.get('_분류')!r}  ({m.get('_분류_이유', '')})")
for m in mapped2:
    name = m.get('품명', '')
    sec = m.get('_분류', '')
    if name in ('PLC', 'Inverter', 'Touch Panel', 'SAFETY PLC'):
        if sec != '제어_구매품':
            print(f"FAIL: {name} → {sec} (제어_구매품 기대)")
            ok = False
    elif name == 'MAIN_FRAME':
        if sec in ('제어_구매품', '전장_구매품'):
            print(f"FAIL: {name} → {sec} (기구 계열 유지 기대)")
            ok = False
    elif name == '초소형 감시 SENSOR':
        # sensor는 약한 키워드 → 구분=기구부이므로 그대로 유지 (오분류 방지)
        if sec == '제어_구매품':
            print(f"FAIL: {name} → {sec} (기구 계열 유지 기대, 약한 키워드 오분류)")
            ok = False

print("\nPASS" if ok else "\nFAIL")
