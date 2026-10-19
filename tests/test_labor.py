"""인건비 분류 규칙 테스트: 특급/고급 MD 분배 + 일액 정수 + Engineering 배수"""
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from openpyxl import load_workbook
from converters.standard_exporter import _fill_detail_sheet, _fill_labor_rows

TPL = Path(__file__).resolve().parent.parent / "templates" / "고객사제출_표준양식.xlsx"

# 실제 풍산 파일 인건비 데이터 (analyzer_export.csv 기준)
LABOR = {
    "인건비-기구 설계":       {"md": 30,  "rate": 400000, "amount": 12000000, "remark": ""},
    "인건비-전장 설계":       {"md": 20,  "rate": 450000, "amount": 9000000,  "remark": ""},
    "인건비-제어 설계":       {"md": 45,  "rate": 450000, "amount": 20250000, "remark": "1名x45日(기 자료 고융 기준)"},
    "인건비-기구 조립":       {"md": 140, "rate": 400000, "amount": 56000000, "remark": "4名x20日/2名x30日"},
    "인건비-전장/제어 셋업":  {"md": 1,   "rate": 40000000, "amount": 40000000, "remark": ""},
    "인건비-기구 셋업":       {"md": 0,   "rate": 400000, "amount": 0, "remark": "해외 셋업 제외"},
    "인건비-전장/제어 양산대응": {"md": 0, "rate": 450000, "amount": 0, "remark": "해외 셋업 제외"},
}

def get(ws, r, c):
    return ws.cell(r, c).value

def check(label, actual, expected):
    ok = actual == expected
    print(f"  {'PASS' if ok else 'FAIL'} {label}: {actual!r} (expected {expected!r})")
    return ok

def main():
    wb = load_workbook(TPL)
    ws = wb["2.상세 견적"]
    # 재료비 항목 없음 → shift=0
    notes = _fill_labor_rows(ws, LABOR, 0)
    for r, text in notes:
        ws.cell(r, 13).value = text  # M열(검증결과)

    all_ok = True
    print("=== 인건비 특급/고급 분배 (인원×일수) ===")
    # 기구 설계: MD=30, 비고 없음 → 1명×30일 → 특급 1명×30일=30 / 고급 0명×30일=0
    all_ok &= check("기구설계자(특급) G", get(ws, 36, 7), 30)
    all_ok &= check("기구설계자(특급) I", get(ws, 36, 9), 400000)
    all_ok &= check("기구설계자(고급) G", get(ws, 37, 7), 0)
    all_ok &= check("기구설계자(고급) I", get(ws, 37, 9), 400000)

    # 전장 설계: MD=20, 비고 없음 → 7일 후보: 2명×7일+1명×7일=21 → 보정 특급13/고급7
    all_ok &= check("전장설계자(특급) G", get(ws, 40, 7), 13)
    all_ok &= check("전장설계자(특급) I", get(ws, 40, 9), 450000)
    all_ok &= check("전장설계자(고급) G", get(ws, 41, 7), 7)
    all_ok &= check("전장설계자(고급) I", get(ws, 41, 9), 450000)

    # 제어 설계: MD=45, 비고 '1名x45日' → 단일 표현: 1명×45일 → 특급 1명×45일=45 / 고급 0
    all_ok &= check("SW설계자(특급) G", get(ws, 44, 7), 45)
    all_ok &= check("SW설계자(특급) I", get(ws, 44, 9), 450000)
    all_ok &= check("SW설계자(특급) K", get(ws, 44, 11), "1名x45日(기 자료 고융 기준)")
    all_ok &= check("SW설계자(고급) G", get(ws, 45, 7), 0)
    all_ok &= check("SW설계자(고급) I", get(ws, 45, 9), 450000)
    all_ok &= check("SW설계자(고급) K", get(ws, 45, 11), "1名x45日(기 자료 고융 기준)")

    # 기구 조립: MD=140, 비고 '4名x20日/2名x30日' → 평균일수 round(140/6)=23 → 3명×23일+3명×23일=138 → 보정 특급71/고급69
    all_ok &= check("기구조립(특급) G", get(ws, 49, 7), 71)
    all_ok &= check("기구조립(특급) I", get(ws, 49, 9), 400000)
    all_ok &= check("기구조립(고급) G", get(ws, 50, 7), 69)
    all_ok &= check("기구조립(고급) I", get(ws, 50, 9), 400000)
    all_ok &= check("기구조립(특급) K", get(ws, 49, 11), "4名x20日/2名x30日")

    # 전장/제어 셋업: MD=1, 비고 없음 → 1명×7일 → 보정 특급1/고급0
    all_ok &= check("셋업제어(특급) G", get(ws, 62, 7), 1)
    all_ok &= check("셋업제어(특급) I", get(ws, 62, 9), 40000000)
    all_ok &= check("셋업제어(고급) G", get(ws, 63, 7), 0)
    all_ok &= check("셋업제어(고급) I", get(ws, 63, 9), 40000000)

    # 기구 셋업: MD=0, 금액=0 → 템플릿 유지 (G=0)
    all_ok &= check("셋업기구(특급) G (0유지)", get(ws, 58, 7), 0)

    # Engineering: 총 MD=30+20+45+140+1+0+0=236 → 21 초과 → 30의 배수: round(236/30)=8 → 240
    all_ok &= check("Engineering G", get(ws, 35, 7), 240)

    # Engineering 21일 미만 → 7(주)의 배수: MD=15 → round(15/7)=2 → 14
    _fill_labor_rows(ws, {"인건비-기구 설계": {"md": 15, "rate": 400000, "amount": 6000000, "remark": ""}}, 0)
    all_ok &= check("Engineering G (21미만→7배수)", get(ws, 35, 7), 14)

    # J열 수식 유지
    all_ok &= check("기구설계자(특급) J 수식", get(ws, 36, 10), "=G36*I36")

    print("=== M열(검증결과) 인원×일수 표기 ===")
    all_ok &= check("M36 (기구설계)", get(ws, 36, 13), "MD=30 (1명×30일+0명×30일)")
    all_ok &= check("M44 (제어설계)", get(ws, 44, 13), "MD=45 (1명×45일+0명×45일)")
    all_ok &= check("M49 (기구조립)", get(ws, 49, 13), "MD=140 (3명×23일+3명×23일)")

    print()
    print("ALL PASS" if all_ok else "SOME FAILED")
    return 0 if all_ok else 1

if __name__ == "__main__":
    sys.exit(main())
