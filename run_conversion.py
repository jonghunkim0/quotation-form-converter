"""
run_conversion.py - 8단계 견적 표준 양식 변환 파이프라인

1. Raw 견적 파일의 모든 데이터 행 식별
2. quotation-analyzer 분석 결과 참조 (analyzer_export.csv)
3. 최종분류 기준으로 품목 분류
4. 원본 품목 정보 분리 (제품명/규격/메이커/수량/단위/단가/비고)
5. 검증 (누락/중복/오류/신뢰도)
6. 검증 결과를 포함한 상세 파일용 데이터 생성
7. 표준 견적서 Excel 컬럼에 매핑
8. 표준 포멧에 누락없이 반영하여 내보내기
"""

import pandas as pd
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
from pathlib import Path
from collections import defaultdict
from datetime import datetime

# ─── 경로 설정 ────────────────────────────────────────────────
RAW_FILE = r"D:\견적분석\풍산 견적_AZS 각형 VER1.0(260602).xlsx"
ANALYZER_CSV = r"c:\AI_Agent\quotation-form-converter\data\analyzer_export.csv"
TEMPLATE = r"c:\AI_Agent\quotation-form-converter\templates\고객사제출_표준양식.xlsx"
OUTPUT_DIR = Path(r"c:\AI_Agent\quotation-form-converter\data\output")
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

# ─── Step 1: Raw 파일에서 Maker/단위/섹션 추출 ────────────────
def extract_raw_metadata(raw_file: str) -> dict:
    """Raw 파일에서 행번호별 Maker(E), 단위(G), 섹션(A열) 추출"""
    wb = openpyxl.load_workbook(raw_file, data_only=True)
    ws = wb["세부"]
    metadata = {}
    current_section = ""

    for r in range(1, ws.max_row + 1):
        a_val = ws.cell(row=r, column=1).value
        if a_val:
            a_str = str(a_val).replace("\n", "").replace(" ", "").strip()
            if a_str.startswith("[1]"):
                current_section = "기구"
            elif a_str.startswith("[2]"):
                current_section = "전장/제어"
            elif a_str.startswith("[3]"):
                current_section = "인건비"
            elif a_str.startswith("[4]"):
                current_section = "경비"
            elif a_str.startswith("[5]"):
                current_section = "일반관리"

        maker = ws.cell(row=r, column=5).value   # E: Maker
        unit = ws.cell(row=r, column=7).value    # G: 단위

        metadata[r] = {
            "maker": str(maker).strip() if maker and str(maker).strip() not in ("-", "") else "",
            "unit": str(unit).strip() if unit and str(unit).strip() not in ("-", "") else "",
            "section": current_section,
        }

    wb.close()
    return metadata


# ─── Step 2: 최종분류 → 표준 섹션 매핑 ────────────────────────
def map_category_to_section(category: str, raw_section: str) -> str:
    """analyzer 최종분류 → 표준 템플릿 섹션 키"""
    cat = category.strip()

    if cat == "기구 구매품":
        return "기구_구매품"
    elif cat == "기구 가공품":
        return "기구_가공품"
    elif cat in ("전장 구매품", "제어 구매품"):
        return "전장제어_구매품"
    elif cat == "전장 가공품":
        return "전장제어_가공품"
    elif cat == "사급품":
        # 원본 섹션에 따라 구매품으로 매칭
        return "기구_구매품" if raw_section == "기구" else "전장제어_구매품"
    elif "설계" in cat:
        return "인건비_설계"
    elif "조립" in cat:
        return "인건비_조립"
    elif "셋업" in cat:
        return "인건비_셋업"
    elif "양산" in cat:
        return "인건비_양산대응"
    elif "특수" in cat:
        return "인건비_특수제어"
    elif "포장" in cat:
        return "경비_포장"
    elif "숙박" in cat or "교통" in cat or "항공" in cat or "식비" in cat:
        return "경비_경비"
    elif "경비" in cat:
        return "경비_경비"
    else:
        return "미분류"


# ─── Step 3: 검증 ─────────────────────────────────────────────
def validate_item(item: dict) -> list:
    """단일 항목 검증: 누락/오류/신뢰도"""
    issues = []

    # 누락 검증
    if not item.get("제품명"):
        issues.append("품명 누락")
    qty = item.get("수량", 0)
    if qty is None or qty == 0:
        issues.append("수량 0/누락")
    price = item.get("단가", 0)
    if price is None or price == 0:
        if item.get("금액", 0) != 0:
            issues.append("단가 누락(금액 있음)")

    # 오류 검증
    if qty is not None and qty < 0:
        issues.append("수량 음수")
    if price is not None and price < 0:
        issues.append("단가 음수")

    # 신뢰도
    conf = item.get("신뢰도", 0)
    if conf < 80:
        issues.append(f"저신뢰도({conf:.0f}%)")

    return issues


def detect_duplicates(items: list) -> set:
    """품명+규격+단가 기준으로 중복 검출"""
    seen = {}
    dups = set()
    for i, item in enumerate(items):
        key = (item.get("제품명", ""), item.get("규격", ""), item.get("단가", 0))
        if key in seen:
            dups.add(i)
            dups.add(seen[key])
        else:
            seen[key] = i
    return dups


# ─── Step 4: 인건비 MD 3등분 (특급/고급/중급) ─────────────────
def split_labor_md(total_md: int) -> tuple:
    """총 MD를 특급/고급/중급 3등분 (나머지 특급에 합산)"""
    base = total_md // 3
    remainder = total_md % 3
    return (base + remainder, base, base)


# ─── Step 5: 표준 포멧 Excel 생성 ─────────────────────────────
def build_standard_excel(items_by_section: dict, labor_data: dict,
                         company_name: str, output_path: str):
    """표준 양식 Excel 생성 (2.상세 견적 시트 빌드)"""

    # 템플릿 로드 (1.표준 견적서 시트 유지)
    wb = openpyxl.load_workbook(TEMPLATE)

    # 기존 "2.상세 견적" 삭제 후 새로 생성
    if "2.상세 견적" in wb.sheetnames:
        del wb["2.상세 견적"]
    ws = wb.create_sheet("2.상세 견적", 1)

    # ── 스타일 ──
    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )
    header_font = Font(bold=True, size=10)
    data_font = Font(size=10)
    center_align = Alignment(horizontal="center", vertical="center")
    left_align = Alignment(horizontal="left", vertical="center")
    right_align = Alignment(horizontal="right", vertical="center")

    # ── 헤더 행 (R1) ──
    headers = ["구분*", "단가유형*", "No", "자재명*", "규격*", "Maker",
               "수량*", "단위", "단가*", "금액*", "비고"]
    for c, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=c, value=h)
        cell.font = header_font
        cell.border = thin_border
        cell.alignment = center_align

    # 열 너비
    col_widths = [12, 10, 8, 25, 30, 15, 8, 6, 12, 14, 25]
    for c, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(c)].width = w

    row = 4  # 데이터 시작 행 (R2~R3 빈 행)

    # ── 섹션 정의 (순서) ──
    # (섹션키, 구분, 단가유형, No 라벨)
    sections_order = [
        ("기구_구매품",     "기구부",     "구매품", "(1) 기구부"),
        ("기구_가공품",     "기구부",     "가공품", ""),
        ("전장제어_구매품", "전장/제어",  "구매품", "(1) 전장/제어"),
        ("전장제어_가공품", "전장/제어",  "가공품", ""),
        ("인건비_설계",     "인건비",     "인건비", "(1) 설계"),
        ("인건비_조립",     "인건비",     "인건비", "(2) 조립"),
        ("인건비_셋업",     "인건비",     "인건비", "(3) 셋업"),
        ("인건비_양산대응", "인건비",     "인건비", "(4) 양산대응"),
        ("인건비_특수제어", "인건비",     "인건비", "(5) 특수제어"),
        ("경비_포장",       "경비",       "경비",   "(1) 포장"),
        ("경비_기타항목",   "경비",       "경비",   "(2) 기타항목"),
        ("경비_경비",       "경비",       "경비",   "(3) 경비"),
    ]

    section_subtotal_rows = {}  # 섹션 → 소계 행 번호
    section_start_rows = {}     # 섹션 → 시작 행 번호

    for sec_key, gubun, dangayuh, no_label in sections_order:
        items = items_by_section.get(sec_key, [])

        # 인건비 섹션은 labor_data 사용
        if sec_key.startswith("인건비_"):
            labor_items = labor_data.get(sec_key, [])
        else:
            labor_items = None

        # ── 섹션 헤더 행 ──
        ws.cell(row=row, column=1, value=gubun).font = header_font
        ws.cell(row=row, column=2, value=dangayuh).font = header_font
        if no_label:
            ws.cell(row=row, column=3, value=no_label).font = header_font
        ws.cell(row=row, column=7, value=1)  # G=1
        for c in range(1, 12):
            ws.cell(row=row, column=c).border = thin_border
        row += 1

        section_start_rows[sec_key] = row

        # ── 데이터 행 ──
        if labor_items is not None:
            # 인건비: 특급/고급/중급 3행
            for grade, md_val in labor_items:
                ws.cell(row=row, column=3, value=f"{grade}").font = data_font
                ws.cell(row=row, column=4, value=f"{gubun} {dangayuh} ({grade})").font = data_font
                ws.cell(row=row, column=7, value=md_val).font = data_font
                ws.cell(row=row, column=7).alignment = right_align
                ws.cell(row=row, column=8, value="MD").font = data_font
                # 단가(I)는 템플릿 기본값 유지 (빈칸이면 0)
                ws.cell(row=row, column=10, value=f"=G{row}*I{row}").font = data_font
                for c in range(1, 12):
                    ws.cell(row=row, column=c).border = thin_border
                row += 1
        else:
            # 재료비: 전체 항목 입력
            for i, item in enumerate(items):
                ws.cell(row=row, column=3, value=f"{i+1})").font = data_font
                ws.cell(row=row, column=4, value=item.get("제품명", "")).font = data_font
                ws.cell(row=row, column=5, value=item.get("규격", "")).font = data_font
                ws.cell(row=row, column=6, value=item.get("메이커", "")).font = data_font
                ws.cell(row=row, column=7, value=item.get("수량", 0)).font = data_font
                ws.cell(row=row, column=7).alignment = right_align
                ws.cell(row=row, column=8, value=item.get("단위", "")).font = data_font
                ws.cell(row=row, column=9, value=item.get("단가", 0)).font = data_font
                ws.cell(row=row, column=9).alignment = right_align
                ws.cell(row=row, column=10, value=f"=G{row}*I{row}").font = data_font
                ws.cell(row=row, column=10).alignment = right_align
                # 비고: 검증결과 + 원본 비고
                bigo_parts = []
                if item.get("검증결과") and item["검증결과"] != "OK":
                    bigo_parts.append(item["검증결과"])
                if item.get("비고"):
                    bigo_parts.append(item["비고"])
                ws.cell(row=row, column=11, value=" | ".join(bigo_parts)).font = data_font
                for c in range(1, 12):
                    ws.cell(row=row, column=c).border = thin_border
                row += 1

        # ── 소계 행 ──
        data_start = section_start_rows[sec_key]
        data_end = row - 1
        ws.cell(row=row, column=3, value=f"{dangayuh} 소계").font = header_font
        if data_end >= data_start:
            ws.cell(row=row, column=10, value=f"=SUM(J{data_start}:J{data_end})").font = header_font
        for c in range(1, 12):
            ws.cell(row=row, column=c).border = thin_border
        section_subtotal_rows[sec_key] = row
        row += 1

    # ── 기구부 합계 ──
    ws.cell(row=row, column=3, value="기구부 합계").font = header_font
    g1 = section_subtotal_rows.get("기구_구매품", 0)
    g2 = section_subtotal_rows.get("기구_가공품", 0)
    if g1 and g2:
        ws.cell(row=row, column=10, value=f"=J{g1}+J{g2}").font = header_font
    for c in range(1, 12):
        ws.cell(row=row, column=c).border = thin_border
    gubun_total_row = row
    row += 1

    # ── 전장/제어 합계 ──
    ws.cell(row=row, column=3, value="전장/제어부 합계").font = header_font
    t1 = section_subtotal_rows.get("전장제어_구매품", 0)
    t2 = section_subtotal_rows.get("전장제어_가공품", 0)
    if t1 and t2:
        ws.cell(row=row, column=10, value=f"=J{t1}+J{t2}").font = header_font
    for c in range(1, 12):
        ws.cell(row=row, column=c).border = thin_border
    jeongjang_total_row = row
    row += 1

    # ── 인건비 합계 ──
    ws.cell(row=row, column=3, value="인건비 합계").font = header_font
    labor_subs = [section_subtotal_rows.get(k, 0) for k in
                  ["인건비_설계", "인건비_조립", "인건비_셋업",
                   "인건비_양산대응", "인건비_특수제어"]]
    labor_subs = [r for r in labor_subs if r]
    if labor_subs:
        formula = "+".join(f"J{r}" for r in labor_subs)
        ws.cell(row=row, column=10, value=f"={formula}").font = header_font
    for c in range(1, 12):
        ws.cell(row=row, column=c).border = thin_border
    labor_total_row = row
    row += 1

    # ── 경비 합계 ──
    ws.cell(row=row, column=3, value="경비 합계").font = header_font
    exp_subs = [section_subtotal_rows.get(k, 0) for k in
                ["경비_포장", "경비_기타항목", "경비_경비"]]
    exp_subs = [r for r in exp_subs if r]
    if exp_subs:
        formula = "+".join(f"J{r}" for r in exp_subs)
        ws.cell(row=row, column=10, value=f"={formula}").font = header_font
    for c in range(1, 12):
        ws.cell(row=row, column=c).border = thin_border
    expense_total_row = row
    row += 1

    # ── Sub합계 ──
    ws.cell(row=row, column=1, value="Sub합계").font = header_font
    total_refs = [r for r in [gubun_total_row, jeongjang_total_row,
                               labor_total_row, expense_total_row] if r]
    if total_refs:
        formula = "+".join(f"J{r}" for r in total_refs)
        ws.cell(row=row, column=10, value=f"={formula}").font = header_font
    for c in range(1, 12):
        ws.cell(row=row, column=c).border = thin_border
    sub_total_row = row
    row += 1

    # ── 일반관리 (5%) ──
    ws.cell(row=row, column=1, value="[5]일반관리(안전관리비포함)").font = header_font
    ws.cell(row=row, column=10, value=f"=J{sub_total_row}*0.05").font = header_font
    for c in range(1, 12):
        ws.cell(row=row, column=c).border = thin_border
    row += 1

    # ── 기업이윤 (5%) ──
    ws.cell(row=row, column=1, value="[6]기업이윤").font = header_font
    ws.cell(row=row, column=10, value=f"=J{sub_total_row}*0.05").font = header_font
    for c in range(1, 12):
        ws.cell(row=row, column=c).border = thin_border
    row += 1

    # ── TOTAL합계 ──
    ws.cell(row=row, column=1, value="TOTAL합계").font = Font(bold=True, size=11)
    ws.cell(row=row, column=10, value=f"=SUM(J{sub_total_row}:J{row-1})").font = Font(bold=True, size=11)
    for c in range(1, 12):
        ws.cell(row=row, column=c).border = thin_border

    # ── 1.표준 견적서(갑지) 시트: 업체명/견적일 입력 ──
    if "1.표준 견적서(갑지)" in wb.sheetnames:
        ws1 = wb["1.표준 견적서(갑지)"]
        # 업체명, 견적일 등 메타데이터는 템플릿 구조에 따라 입력
        # (템플릿 구조에 따라 셀 위치가 다를 수 있음)

    # 저장
    wb.save(output_path)
    wb.close()
    print(f"  → {output_path}")


# ─── Main ─────────────────────────────────────────────────────
def main():
    print("=" * 60)
    print("견적 표준 양식 변환 파이프라인")
    print("=" * 60)

    # Step 1: Raw 메타데이터
    print("\n[Step 1] Raw 파일 메타데이터 추출...")
    raw_meta = extract_raw_metadata(RAW_FILE)
    print(f"  → {len(raw_meta)}개 행 메타데이터 추출")

    # Step 2: Analyzer CSV
    print("\n[Step 2] Analyzer 내보내기 CSV 로드...")
    df = pd.read_csv(ANALYZER_CSV, encoding="utf-8-sig")
    df.columns = [c.replace(" ", "") for c in df.columns]
    print(f"  → {len(df)}개 항목 로드")

    # Step 3: 항목별 매핑 + 검증
    print("\n[Step 3] 항목별 매핑 + 검증...")
    items_by_section = defaultdict(list)
    all_items = []

    for idx, row in df.iterrows():
        row_num = int(row["행번호"])
        category = str(row["최종분류"]).strip()
        raw_section = raw_meta.get(row_num, {}).get("section", "")
        section = map_category_to_section(category, raw_section)

        maker = raw_meta.get(row_num, {}).get("maker", "")
        unit = raw_meta.get(row_num, {}).get("unit", "")

        item = {
            "행번호": row_num,
            "제품명": str(row["품명"]).strip() if pd.notna(row.get("품명")) else "",
            "규격": str(row["사양"]).strip() if pd.notna(row.get("사양")) else "",
            "메이커": maker,
            "수량": row.get("수량", 0),
            "단위": unit,
            "단가": row.get("단가", 0),
            "금액": row.get("금액", 0),
            "비고": str(row["비고"]).strip() if pd.notna(row.get("비고")) else "",
            "최종분류": category,
            "매핑섹션": section,
            "신뢰도": row.get("신뢰도(%)", 0),
            "분류근거": str(row["분류근거"]).strip() if pd.notna(row.get("분류근거")) else "",
        }

        # 검증
        issues = validate_item(item)
        item["검증결과"] = " | ".join(issues) if issues else "OK"

        items_by_section[section].append(item)
        all_items.append(item)

    # 중복 검출
    dup_indices = detect_duplicates(all_items)
    for i in dup_indices:
        if "중복" not in all_items[i]["검증결과"]:
            all_items[i]["검증결과"] += " | 중복"

    # 검증 요약
    ok_count = sum(1 for it in all_items if it["검증결과"] == "OK")
    warn_count = len(all_items) - ok_count
    print(f"  → OK: {ok_count}, 경고: {warn_count}, 중복: {len(dup_indices)}")

    # Step 4: 상세 검증 파일
    print("\n[Step 4] 상세 검증 파일 생성...")
    val_df = pd.DataFrame(all_items)
    val_path = OUTPUT_DIR / "상세_검증파일.csv"
    val_df.to_csv(val_path, index=False, encoding="utf-8-sig")
    print(f"  → {val_path} ({len(val_df)}개 항목)")

    # Step 5: 인건비 MD 3등분
    print("\n[Step 5] 인건비 MD 3등분...")
    labor_data = {}
    for sec_key in ["인건비_설계", "인건비_조립", "인건비_셋업",
                     "인건비_양산대응", "인건비_특수제어"]:
        items = items_by_section.get(sec_key, [])
        total_md = sum(int(it.get("수량", 0)) for it in items)
        if total_md > 0:
            teuk, go, jung = split_labor_md(total_md)
            labor_data[sec_key] = [
                ("특급", teuk),
                ("고급", go),
                ("중급", jung),
            ]
            print(f"  {sec_key}: 총 {total_md}MD → 특급 {teuk} / 고급 {go} / 중급 {jung}")
        else:
            labor_data[sec_key] = []

    # Step 6: 표준 포멧 Excel
    print("\n[Step 6] 표준 포멧 Excel 생성...")
    # 업체명 추출
    company_name = "㈜풍산시스템"
    out_path = str(OUTPUT_DIR / "표준_견적서_풍산_AZS.xlsx")
    build_standard_excel(items_by_section, labor_data, company_name, out_path)

    # ── 최종 요약 ──
    print("\n" + "=" * 60)
    print("=== 섹션별 항목 수 ===")
    for sec in ["기구_구매품", "기구_가공품", "전장제어_구매품", "전장제어_가공품",
                "인건비_설계", "인건비_조립", "인건비_셋업",
                "인건비_양산대응", "인건비_특수제어",
                "경비_포장", "경비_기타항목", "경비_경비", "미분류"]:
        cnt = len(items_by_section.get(sec, []))
        if cnt > 0:
            print(f"  {sec}: {cnt}")
    print(f"  총: {len(all_items)}")
    print(f"\n출력 파일:")
    print(f"  1. {val_path}")
    print(f"  2. {out_path}")
    print("=" * 60)


if __name__ == "__main__":
    main()
