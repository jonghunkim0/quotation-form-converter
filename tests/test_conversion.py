"""변환 파이프라인 테스트"""
import sys
sys.path.insert(0, ".")

from converters import RuleMapper, export_standard_excel

# 테스트 데이터 (가상 견적 항목)
test_items = [
    {"raw_data": {"품번": "ABC-001", "품명": "볼트 M6x20", "규격": "SS304", "수량": "100", "단위": "EA", "단가": "500", "금액": "50000", "비고": ""}},
    {"raw_data": {"품번": "ABC-002", "품명": "너트 M6", "규격": "SS304", "수량": "100", "단위": "EA", "단가": "200", "금액": "20000", "비고": ""}},
    {"raw_data": {"품번": "ABC-003", "품명": "기구설계자(고급)", "규격": "10~14년차", "수량": "2", "단위": "M", "단가": "420000", "금액": "840000", "비고": ""}},
]

mapper = RuleMapper()
mapped = mapper.map_items(test_items)

print("=== 매핑 결과 ===")
for item in mapped:
    code = item.get("품목코드", "")
    name = item.get("품명", "")
    qty = item.get("수량", "")
    price = item.get("단가", "")
    amount = item.get("금액", "")
    verify = item.get("검증결과", "")
    print(f"  {code} | {name} | 수량={qty} | 단가={price} | 금액={amount} | 검증={verify}")

# Excel 생성
output = export_standard_excel(mapped, company_name="테스트업체", quote_no="QT-2026-001")
print(f"\nExcel 생성 성공: {len(output)} bytes")

# 파일 저장
with open("data/test_output.xlsx", "wb") as f:
    f.write(output)
print("data/test_output.xlsx 저장 완료")
