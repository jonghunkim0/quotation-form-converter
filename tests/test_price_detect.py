"""
'견적가' 헤더 파일의 단가 매칭 테스트
- config 직접 매핑 (견적가→단가, 총액→금액)
- 데이터 기반 자동 감지 (_detect_price_column)
"""
import sys
sys.path.insert(0, ".")
from converters import RuleMapper

PASS = 0
FAIL = 0

def check(name, cond, detail=""):
    global PASS, FAIL
    if cond:
        PASS += 1
        print(f"  PASS {name} {detail}")
    else:
        FAIL += 1
        print(f"  FAIL {name} {detail}")

# ── Case 1: 헤더 '견적가' + '총액' (직접 매핑) ──
print("=== Case 1: 헤더 '견적가' + '총액' ===")
items = [
    {"raw_data": {"No": "1)", "품 명": "MAIN_FRAME", "규격": "BT2420", "Maker": "PSS",
                  "수량": "1", "단위": "EA", "견적가": "27073800", "총액": "27073800"},
     "sheet_name": "세부", "row_index": 4},
    {"raw_data": {"No": "2)", "품 명": "FOOT_SPACER-01", "규격": "BT2420", "Maker": "PSS",
                  "수량": "20", "단위": "EA", "견적가": "28100", "총액": "562000"},
     "sheet_name": "세부", "row_index": 5},
    {"raw_data": {"No": "3)", "품 명": "FOOT_SPACER-02", "규격": "BT2420", "Maker": "PSS",
                  "수량": "20", "단위": "EA", "견적가": "4600", "총액": "92000"},
     "sheet_name": "세부", "row_index": 6},
]
mapper = RuleMapper({})
mapped = mapper.map_items(items)
check("견적가→단가 매칭", mapped[0]["단가"] == 27073800.0, f"(단가={mapped[0]['단가']})")
check("총액→금액 매칭", mapped[0]["금액"] == 27073800.0, f"(금액={mapped[0]['금액']})")
check("금액 검증 OK", "금액 검증 오류" not in mapped[1]["검증결과"], f"(검증={mapped[1]['검증결과']})")

# ── Case 2: 헤더 '견적단가' (직접 매칭 안 됨 → 데이터 기반 감지) ──
print("=== Case 2: 헤더 '견적단가' (부분일치 실패 가정) + '총액' ===")
# '견적단가'는 '단가' 부분일치로 매칭될 수 있음 → '견적가격'처럼 단가 키워드 없는 이름 사용
items2 = [
    {"raw_data": {"No": "1)", "품 명": "A", "규격": "S1", "Maker": "M",
                  "수량": "2", "단위": "EA", "견적가격": "1000", "총액": "2000"},
     "sheet_name": "세부", "row_index": 4},
    {"raw_data": {"No": "2)", "품 명": "B", "규격": "S2", "Maker": "M",
                  "수량": "5", "단위": "EA", "견적가격": "300", "총액": "1500"},
     "sheet_name": "세부", "row_index": 5},
    {"raw_data": {"No": "3)", "품 명": "C", "규격": "S3", "Maker": "M",
                  "수량": "10", "단위": "EA", "견적가격": "50", "총액": "500"},
     "sheet_name": "세부", "row_index": 6},
]
mapper2 = RuleMapper({})
mapped2 = mapper2.map_items(items2)
check("견적가격→단가 자동 감지", mapped2[0]["단가"] == 1000.0, f"(단가={mapped2[0]['단가']})")
check("견적가격→단가 (2행)", mapped2[1]["단가"] == 300.0, f"(단가={mapped2[1]['단가']})")

# ── Case 3: 기존 '단가' 헤더 파일 (회귀 방지) ──
print("=== Case 3: 기존 '단가' 헤더 (회귀 방지) ===")
items3 = [
    {"raw_data": {"No": "1)", "품 명": "A", "규격": "S1", "Maker": "M",
                  "수량": "2", "단위": "EA", "단가": "1000", "금액": "2000"},
     "sheet_name": "세부", "row_index": 4},
]
mapper3 = RuleMapper({})
mapped3 = mapper3.map_items(items3)
check("단가→단가 유지", mapped3[0]["단가"] == 1000.0, f"(단가={mapped3[0]['단가']})")
check("금액→금액 유지", mapped3[0]["금액"] == 2000.0, f"(금액={mapped3[0]['금액']})")

print()
print(f"=== 결과: PASS {PASS}, FAIL {FAIL} ===")
sys.exit(1 if FAIL else 0)
