"""
데이터 내용 기반 컬럼 유추 테스트
- 헤더가 규칙과 전혀 다른 표기일 때 데이터 값/레이아웃으로 유추
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

# ── Case 1: 헤더가 전부 다른 표기 (품목/세부/제조사/개수/단위명/가격/합계) ──
print("=== Case 1: 비표준 헤더 (품목|세부|제조사|개수|단위명|가격|합계) ===")
items = [
    {"raw_data": {"품목": "MAIN_FRAME", "세부": "BT2420-L1", "제조사": "PSS",
                  "개수": "1", "단위명": "EA", "가격": "27073800", "합계": "27073800"},
     "sheet_name": "s", "row_index": 1},
    {"raw_data": {"품목": "FOOT_SPACER-01", "세부": "BT2420-L1", "제조사": "PSS",
                  "개수": "20", "단위명": "EA", "가격": "28100", "합계": "562000"},
     "sheet_name": "s", "row_index": 2},
    {"raw_data": {"품목": "FOOT_SPACER-02", "세부": "BT2420-L1", "제조사": "PSS",
                  "개수": "20", "단위명": "EA", "가격": "4600", "합계": "92000"},
     "sheet_name": "s", "row_index": 3},
    {"raw_data": {"품목": "PLATE-1", "세부": "BT2420-L1", "제조사": "PSS",
                  "개수": "1", "단위명": "EA", "가격": "132500", "합계": "132500"},
     "sheet_name": "s", "row_index": 4},
]
mapper = RuleMapper({})
mapped = mapper.map_items(items)
check("품목→품명", mapped[0]["품명"] == "MAIN_FRAME", f"(품명={mapped[0]['품명']})")
check("세부→규격", mapped[0]["규격"] == "BT2420-L1", f"(규격={mapped[0]['규격']})")
check("제조사→Maker", mapped[0]["Maker"] == "PSS", f"(Maker={mapped[0]['Maker']})")
check("개수→수량", mapped[0]["수량"] == 1, f"(수량={mapped[0]['수량']})")
check("단위명→단위", mapped[0]["단위"] == "EA", f"(단위={mapped[0]['단위']})")
check("가격→단가", mapped[0]["단가"] == 27073800.0, f"(단가={mapped[0]['단가']})")
check("합계→금액", mapped[0]["금액"] == 27073800.0, f"(금액={mapped[0]['금액']})")

# ── Case 2: 헤더가 아예 의미 없는 이름 (A, B, C, D, E, F, G) ──
print("=== Case 2: 의미 없는 헤더 (A~G) 레이아웃 유추 ===")
items2 = [
    {"raw_data": {"A": "1)", "B": "MAIN_FRAME", "C": "BT2420", "D": "PSS",
                  "E": "1", "F": "EA", "G": "27073800", "H": "27073800"},
     "sheet_name": "s", "row_index": 1},
    {"raw_data": {"A": "2)", "B": "FOOT_SPACER-01", "C": "BT2420", "D": "PSS",
                  "E": "20", "F": "EA", "G": "28100", "H": "562000"},
     "sheet_name": "s", "row_index": 2},
    {"raw_data": {"A": "3)", "B": "FOOT_SPACER-02", "C": "BT2420", "D": "PSS",
                  "E": "20", "F": "EA", "G": "4600", "H": "92000"},
     "sheet_name": "s", "row_index": 3},
    {"raw_data": {"A": "4)", "B": "PLATE-1", "C": "BT2420", "D": "PSS",
                  "E": "1", "F": "EA", "G": "132500", "H": "132500"},
     "sheet_name": "s", "row_index": 4},
]
mapper2 = RuleMapper({})
mapped2 = mapper2.map_items(items2)
check("A→품목코드", mapped2[0]["품목코드"] == "1)", f"(품목코드={mapped2[0]['품목코드']})")
check("B→품명", mapped2[0]["품명"] == "MAIN_FRAME", f"(품명={mapped2[0]['품명']})")
check("C→규격", mapped2[0]["규격"] == "BT2420", f"(규격={mapped2[0]['규격']})")
check("D→Maker", mapped2[0]["Maker"] == "PSS", f"(Maker={mapped2[0]['Maker']})")
check("E→수량", mapped2[0]["수량"] == 1, f"(수량={mapped2[0]['수량']})")
check("F→단위", mapped2[0]["단위"] == "EA", f"(단위={mapped2[0]['단위']})")
check("G→단가", mapped2[0]["단가"] == 27073800.0, f"(단가={mapped2[0]['단가']})")
check("H→금액", mapped2[0]["금액"] == 27073800.0, f"(금액={mapped2[0]['금액']})")

# ── Case 3: 기존 표준 헤더 (회귀 방지) ──
print("=== Case 3: 표준 헤더 (회귀 방지) ===")
items3 = [
    {"raw_data": {"No": "1)", "품 명": "A", "규격": "S1", "Maker": "M",
                  "수량": "2", "단위": "EA", "단가": "1000", "금액": "2000"},
     "sheet_name": "s", "row_index": 1},
]
mapper3 = RuleMapper({})
mapped3 = mapper3.map_items(items3)
check("표준 매칭 유지", mapped3[0]["품명"] == "A" and mapped3[0]["단가"] == 1000.0,
      f"(품명={mapped3[0]['품명']}, 단가={mapped3[0]['단가']})")

print()
print(f"=== 결과: PASS {PASS}, FAIL {FAIL} ===")
sys.exit(1 if FAIL else 0)
