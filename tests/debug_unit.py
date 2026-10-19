"""Unit 행 유실 원인 디버그"""
import sys
sys.path.insert(0, ".")
from converters import RuleMapper
from classification import classify_items

# 원본 구조 시뮬레이션 (실제 파일과 동일)
test_items = [
    {'raw_data': {'No': '(1) Unit명 : BASE_FRAME', '수량': '1', '단위': '식', '단가': '31886000', '금액': '31886000'}},
    {'raw_data': {'No': '1)', '품 명': 'MAIN_FRAME', '규격': 'BT2420-L1-1001M0001-CTQ-00', 'Maker': 'PSS', '수량': '1', '단위': 'EA', '단가': '27073800', '금액': '27073800'}},
    {'raw_data': {'No': '2)', '품 명': 'FOOT_SPACER-01', '규격': 'BT2420-L1-1001M0002-00', 'Maker': 'PSS', '수량': '20', '단위': 'EA', '단가': '28100', '금액': '562000'}},
    {'raw_data': {'No': '가공품 소계', '금액': '30676000'}},
    {'raw_data': {'No': '(2) Unit명 : MAGAZINE_UNIT', '수량': '1', '단위': '식', '단가': '4624040', '금액': '4624040'}},
    {'raw_data': {'No': '1)', '품 명': 'AIR CYLINDER', '규격': '25A-MGPM25-50Z-M9NVL', 'Maker': 'SMC', '수량': '4', '단위': 'EA', '단가': '173910', '금액': '695640'}},
    {'raw_data': {'No': '2)', '품 명': 'MANIFOLD', '규격': '25A-SSSY5-20-02', 'Maker': 'SMC', '수량': '1', '단위': 'EA', '단가': '9470', '금액': '37880'}},
]
mapper = RuleMapper()
mapped = mapper.map_items(test_items)
print("=== 매핑 후 ===")
for i, m in enumerate(mapped):
    print(f"[{i}] _합산={m.get('_합산')} _소계={m.get('_소계')} _unit_name={m.get('_unit_name')!r} 품목코드={m.get('품목코드')!r} 품명={m.get('품명')!r} 단위={m.get('단위')!r}")

mapped = classify_items(mapped)
print("\n=== 분류 후 ===")
for i, m in enumerate(mapped):
    print(f"[{i}] _합산={m.get('_합산')} _소계={m.get('_소계')} _분류={m.get('_분류')} 품명={m.get('품명')!r}")
