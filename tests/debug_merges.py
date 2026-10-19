"""Unit 분할 시나리오에서 삭제 전 병합 상태 디버그"""
import sys
sys.path.insert(0, ".")
from converters import RuleMapper
from classification import classify_items
from converters.standard_exporter import _fill_detail_sheet
from openpyxl import load_workbook
from config import DEFAULT_TEMPLATE_FILE

test_items3 = [
    {'raw_data': {'No': '(1) Unit명 : BASE_FRAME', '수량': '1', '단위': '식', '단가': '31886000', '금액': '31886000'}},
    {'raw_data': {'No': '1)', '품 명': 'MAIN_FRAME', '규격': 'S1', 'Maker': 'PSS', '수량': '1', '단위': 'EA', '단가': '27073800', '금액': '27073800'}},
    {'raw_data': {'No': '2)', '품 명': 'FOOT_SPACER-01', '규격': 'S2', 'Maker': 'PSS', '수량': '20', '단위': 'EA', '단가': '28100', '금액': '562000'}},
    {'raw_data': {'No': '3)', '품 명': 'FOOT_SPACER-02', '규격': 'S3', 'Maker': 'PSS', '수량': '20', '단위': 'EA', '단가': '4600', '금액': '92000'}},
    {'raw_data': {'No': '4)', '품 명': 'HI MOUNT', '규격': 'JK15-02', 'Maker': '진건산업', '수량': '20', '단위': 'EA', '단가': '60500', '금액': '1210000'}},
    {'raw_data': {'No': '5)', '품 명': 'SPEED CONTROLLER', '규격': 'AS1002FS-06', 'Maker': 'SMC', '수량': '48', '단위': 'EA', '단가': '6490', '금액': '311520'}},
]
mapped3 = RuleMapper().map_items(test_items3)
mapped3 = classify_items(mapped3)
for m in mapped3:
    if m.get('_합산'):
        continue
    if m.get('품명', '').startswith(('MAIN_FRAME', 'FOOT_SPACER')):
        m['_분류'] = '기구_가공품'
    else:
        m['_분류'] = '기구_구매품'

wb = load_workbook(str(DEFAULT_TEMPLATE_FILE))
ws = wb['2.상세 견적']

# _fill_detail_sheet를 monkeypatch로 삭제 직전 병합 상태 캡처
import converters.standard_exporter as se
orig_delete = se._delete_row_with_remap
def patched_delete(ws_, row):
    print(f'--- delete row {row} 전 병합 상태 ---')
    for m in sorted(ws_.merged_cells.ranges, key=lambda x: (x.min_row, x.min_col)):
        print(' ', m)
    orig_delete(ws_, row)
se._delete_row_with_remap = patched_delete

try:
    se._fill_detail_sheet(ws, mapped3, False)
except Exception as e:
    print(f'ERROR: {type(e).__name__}: {e}')
