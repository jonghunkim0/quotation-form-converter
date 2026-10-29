"""수식 유지 검증 테스트"""
import sys
sys.path.insert(0, ".")
import openpyxl

wb = openpyxl.load_workbook('data/test_output.xlsx', data_only=False)

# 1.표준 견적서(갑지) 수식 확인
ws1 = wb['1.표준 견적서(갑지)']
print('=== 1.표준 견적서(갑지) - 수식 확인 ===')
print(f'  C10 = {ws1["C10"].value}')
print(f'  F19 = {ws1["F19"].value}')
print(f'  H19 = {ws1["H19"].value}')
print(f'  C23 = {ws1["C23"].value}')
print(f'  G23 = {ws1["G23"].value}')
print(f'  H23 = {ws1["H23"].value}')
print(f'  H24 = {ws1["H24"].value}')
print(f'  B8  = {ws1["B8"].value}')
print(f'  B19 = {ws1["B19"].value}')
print()

# 2.상세 견적 수식 확인
ws2 = wb['2.상세 견적']
print('=== 2.상세 견적 - 수식 확인 ===')
print(f'  J9   = {ws2["J9"].value}')
print(f'  J14  = {ws2["J14"].value}')
print(f'  J71  = {ws2["J71"].value}')
print(f'  J98  = {ws2["J98"].value}')
print(f'  J101 = {ws2["J101"].value}')
print()

# 수식 검증
checks = [
    (ws1["C10"].value, "=H19"),
    (ws1["F19"].value, "=H24"),
    (ws1["H19"].value, "=ROUNDDOWN(D19*F19, -5)"),
    (ws1["C23"].value, "=B19"),
    (ws1["G23"].value, "='2.상세 견적'!J101"),
    (ws1["H23"].value, "=G23*E23"),
    (ws1["H24"].value, "=SUM(H23)"),
    (ws2["J9"].value, "=SUM(J6:J8)"),
    (ws2["J14"].value, "=G14*I14"),
    (ws2["J71"].value, "=SUM(J31,J42,J51,J60,J69)"),
    (ws2["J98"].value, "=SUM(J97,J71,J13,J30)"),
    (ws2["J101"].value, "=SUM(J98:J100)"),
]

all_pass = True
for actual, expected in checks:
    if actual != expected:
        print(f'  ❌ FAIL: {actual} != {expected}')
        all_pass = False

if all_pass:
    print('✅ 모든 수식 유지 확인 완료!')
else:
    print('❌ 일부 수식이 변경되었습니다!')
