"""
표준 양식 Excel 출력 로직
고객사 표준 양식(고객사제출_표준양식.xlsx) 구조를 그대로 유지하면서
'2.상세 견적' 시트에 변환된 데이터를 입력.

표준 양식 구조:
- 시트 1: '1.표준 견적서(갑지)' - 견적 요약 (변경 없음)
- 시트 2: '2.상세 견적' - 상세 항목 (여기에 데이터 입력)

'2.상세 견적' 컬럼 구조 (Row 4 헤더):
  A=구분*  B=단가유형*  C=No  D=자재명*  E=규격*  F=Maker
  G=수량*  H=단위  I=단가*  J=금액*  K=비고
  (M~Q: 2번째 견적안 - 동일 구조)

섹션 구조:
  - 기구부 (구매품/가공품)
  - 전장/제어 (구매품/가공품)
  - 인건비 (설계/조립/셋업/양산대응/특수제어)
  - 경비 (포장/기타항목/경비)
  - Sub합계, 일반관리, 기업이윤, TOTAL합계
"""

import copy
from io import BytesIO
from typing import List, Dict, Any, Optional
from pathlib import Path
from openpyxl import load_workbook

from config import DEFAULT_TEMPLATE_FILE


def export_standard_excel(
    mapped_items: List[Dict[str, Any]],
    company_name: str = "",
    quote_no: str = "",
    template_file: Optional[str] = None,
    show_original: bool = False,
    labor_data: Optional[Dict[str, Dict[str, Any]]] = None,
) -> bytes:
    """
    고객사 표준 양식 템플릿에 변환된 데이터를 입력하여 Excel 생성.
    
    ⚠️ openpyxl은 drawing(텍스트 박스, 이미지)을 저장 시 제거하는 문제가 있다.
    해결: openpyxl로 셀 값만 수정 → 저장 → ZIP 레벨에서 템플릿의
    drawing/media 파일을 출력 파일에 복원.
    
    Args:
        mapped_items: RuleMapper.map_items() 결과
        company_name: 업체명
        quote_no: 견적번호
        template_file: 표준 양식 템플릿 경로 (None이면 기본 템플릿 사용)
        show_original: 원본 컬럼명 표시 여부 (비고에 추가)
        labor_data: 인건비 데이터 {analyzer_카테고리: {md, rate, amount, remark}}
        
    Returns:
        Excel 파일 bytes
    """
    import zipfile
    import shutil
    import tempfile
    import os
    
    # 템플릿 경로 결정
    if template_file and Path(template_file).exists():
        tpl_path = str(template_file)
    elif DEFAULT_TEMPLATE_FILE.exists():
        tpl_path = str(DEFAULT_TEMPLATE_FILE)
    else:
        raise FileNotFoundError(
            f"표준 양식 템플릿을 찾을 수 없습니다.\n"
            f"경로: {DEFAULT_TEMPLATE_FILE}\n"
            f"templates/ 폴더에 '고객사제출_표준양식.xlsx'를 배치하세요."
        )
    
    # 1) openpyxl로 셀 값 수정
    #    템플릿에 externalLinks(외부 파일 캐시, 수십 MB)가 있으면 openpyxl이
    #    전부 파싱/재직렬화해서 로딩+저장이 5초 이상 걸림 →
    #    로딩 전에 ZIP 레벨에서 제거한 임시 템플릿으로 교체.
    clean_tpl, clean_dir = _strip_external_links_for_load(tpl_path)
    tmp_dir = tempfile.mkdtemp()
    tmp_output = os.path.join(tmp_dir, "output.xlsx")
    try:
        wb = load_workbook(clean_tpl)
        total_inserted = 0
        if "2.상세 견적" in wb.sheetnames:
            ws = wb["2.상세 견적"]
            total_inserted = _fill_detail_sheet(ws, mapped_items, show_original, labor_data)
        
        if "1.표준 견적서(갑지)" in wb.sheetnames:
            ws1 = wb["1.표준 견적서(갑지)"]
            _fill_summary_sheet(ws1, company_name, quote_no, mapped_items)
            _fix_summary_cross_refs(ws1, total_inserted)
        
        wb.save(tmp_output)
        
        # 2) ZIP 레벨에서 템플릿의 drawing/media 파일을 복원
        _restore_drawing_from_template(tpl_path, tmp_output)
        
        # 3) externalLinks(외부 참조) 제거 - 다른 환경에서 오류 방지
        _remove_external_links(tmp_output)
        
        # 4) bytes로 반환
        with open(tmp_output, "rb") as f:
            result = f.read()
    finally:
        if clean_dir:
            shutil.rmtree(clean_dir, ignore_errors=True)
        shutil.rmtree(tmp_dir, ignore_errors=True)
    
    return result


def _strip_external_links_for_load(template_path: str):
    """
    템플릿의 externalLinks(외부 파일 캐시)를 제거한 임시 템플릿을 생성.
    
    템플릿에 수십 MB의 externalLinks 캐시가 있으면 openpyxl 로딩/저장이
    5초 이상 걸린다. 로딩 전에 ZIP 레벨에서 제거하면 로딩+저장이
    1초 내외로 줄어든다. externalLinks는 출력에서도 제거되므로
    (다른 환경에서 오류 유발) 로딩 시점에 제거해도 결과에 영향 없음.
    
    Returns:
        (임시 템플릿 경로, 임시 디렉터리 경로) — 호출부에서 정리
    """
    import zipfile
    import re
    import tempfile
    import shutil
    import os
    import zipfile
    import re
    import tempfile
    import shutil
    import os
    
    tmp_dir = tempfile.mkdtemp()
    clean_path = os.path.join(tmp_dir, "clean_template.xlsx")
    
    with zipfile.ZipFile(template_path, "r") as z:
        has_links = any("externalLinks" in n for n in z.namelist())
        if not has_links:
            return template_path, None  # 제거 대상 없음 → 원본 사용
        with zipfile.ZipFile(clean_path, "w", zipfile.ZIP_DEFLATED) as new_zip:
            for item in z.infolist():
                if "externalLinks" in item.filename:
                    continue
                data = z.read(item.filename)
                if item.filename == "[Content_Types].xml":
                    content = data.decode("utf-8")
                    content = re.sub(r'<Override[^>]*externalLink[^>]*/>', '', content)
                    data = content.encode("utf-8")
                elif item.filename == "xl/_rels/workbook.xml.rels":
                    content = data.decode("utf-8")
                    content = re.sub(r'<Relationship[^>]*externalLink[^>]*/>', '', content)
                    data = content.encode("utf-8")
                elif item.filename == "xl/workbook.xml":
                    content = data.decode("utf-8")
                    content = re.sub(r'<externalReferences>.*?</externalReferences>', '',
                                     content, flags=re.DOTALL)
                    content = re.sub(r'<definedName[^>]*>[^<]*\[\d+\][^<]*</definedName>',
                                     '', content)
                    if re.search(r'<definedNames>\s*</definedNames>', content):
                        content = re.sub(r'<definedNames>\s*</definedNames>', '', content)
                    data = content.encode("utf-8")
                new_zip.writestr(item, data)
    return clean_path, tmp_dir


def _restore_drawing_from_template(template_path: str, output_path: str):
    """
    템플릿의 drawing(텍스트 박스, 이미지) 파일을 출력 Excel에 복원.
    
    openpyxl이 저장 시 xl/drawings/, xl/media/를 제거하므로
    템플릿 원본에서 해당 파일을 복사하여 출력 파일에 추가한다.
    """
    import zipfile
    import os
    import tempfile
    import shutil
    
    # 템플릿에서 drawing/media 파일 수집
    with zipfile.ZipFile(template_path, "r") as tpl_zip:
        drawing_files = {}
        for name in tpl_zip.namelist():
            if "drawing" in name.lower() or "media" in name.lower():
                drawing_files[name] = tpl_zip.read(name)
    
    if not drawing_files:
        return  # drawing 파일이 없으면 처리 불필요
    
    # 출력 파일에서 기존 drawing/media 제거 후 템플릿 버전으로 교체
    tmp_dir = tempfile.mkdtemp()
    tmp_rebuilt = os.path.join(tmp_dir, "rebuilt.xlsx")
    
    with zipfile.ZipFile(output_path, "r") as out_zip:
        with zipfile.ZipFile(tmp_rebuilt, "w", zipfile.ZIP_DEFLATED) as new_zip:
            for item in out_zip.infolist():
                # 기존 drawing/media 파일은 스킵 (템플릿 버전으로 교체)
                if item.filename in drawing_files:
                    continue
                new_zip.writestr(item, out_zip.read(item.filename))
            
            # 템플릿의 drawing/media 파일 추가
            for name, data in drawing_files.items():
                new_zip.writestr(name, data)
    
    # 출력 파일로 교체
    shutil.move(tmp_rebuilt, output_path)
    shutil.rmtree(tmp_dir, ignore_errors=True)


def _remove_external_links(output_path: str):
    """
    출력 Excel에서 externalLinks(외부 참조)를 제거.
    
    템플릿에 포함된 외부 파일 참조(ENG공급.XLS 등)는
    다른 환경에서 열 때 오류를 유발하므로 제거한다.
    
    제거 대상:
    1. xl/externalLinks/ 폴더 전체
    2. [Content_Types].xml의 externalLink Override
    3. xl/_rels/workbook.xml.rels의 externalLink Relationship
    4. xl/workbook.xml의 <externalReferences> 태그
    """
    import zipfile
    import os
    import tempfile
    import shutil
    import re
    
    tmp_dir = tempfile.mkdtemp()
    tmp_rebuilt = os.path.join(tmp_dir, "rebuilt.xlsx")
    
    with zipfile.ZipFile(output_path, "r") as out_zip:
        with zipfile.ZipFile(tmp_rebuilt, "w", zipfile.ZIP_DEFLATED) as new_zip:
            for item in out_zip.infolist():
                filename = item.filename
                
                # 1) externalLinks 폴더 파일은 제외
                if "externalLinks" in filename:
                    continue
                
                data = out_zip.read(filename)
                
                # 2) [Content_Types].xml에서 externalLink Override 제거
                if filename == "[Content_Types].xml":
                    content = data.decode("utf-8")
                    content = re.sub(
                        r'<Override[^>]*externalLink[^>]*/>',
                        '',
                        content
                    )
                    data = content.encode("utf-8")
                
                # 3) workbook.xml.rels에서 externalLink Relationship 제거
                elif filename == "xl/_rels/workbook.xml.rels":
                    content = data.decode("utf-8")
                    content = re.sub(
                        r'<Relationship[^>]*externalLink[^>]*/>',
                        '',
                        content
                    )
                    data = content.encode("utf-8")
                
                # 4) workbook.xml에서 <externalReferences> 태그 제거
                elif filename == "xl/workbook.xml":
                    content = data.decode("utf-8")
                    content = re.sub(
                        r'<externalReferences>.*?</externalReferences>',
                        '',
                        content,
                        flags=re.DOTALL
                    )
                    # 5) definedNames에서 외부 참조 포함 항목 제거
                    # '[1]#REF'!#REF! 또는 [10]simulation_BSE_0402!전전월 등
                    content = re.sub(
                        r'<definedName[^>]*>[^<]*\[\d+\][^<]*</definedName>',
                        '',
                        content
                    )
                    # definedNames가 비어있으면 태그 자체 제거
                    if re.search(r'<definedNames>\s*</definedNames>', content):
                        content = re.sub(
                            r'<definedNames>\s*</definedNames>',
                            '',
                            content
                        )
                    data = content.encode("utf-8")
                
                new_zip.writestr(item, data)
    
    # 출력 파일로 교체
    shutil.move(tmp_rebuilt, output_path)
    shutil.rmtree(tmp_dir, ignore_errors=True)


def _fill_detail_sheet(ws, items: List[Dict[str, Any]], show_original: bool,
                       labor_data: Optional[Dict[str, Dict[str, Any]]] = None):
    """
    '2.상세 견적' 시트에 매핑된 항목을 입력.
    
    ⚠️ 수식 셀은 절대 변경하지 않는다.
    - J열(금액): =G*I 또는 =SUM(...) 수식으로 자동 계산
    - I열(단가)의 소계 행: =SUM(J35:J47) 등
    
    분류 기반 배치:
    - 각 항목의 _분류 필드를 기준으로 템플릿 섹션에 배치
    - 인건비/경비: labor_data 기반 특급/고급 분배 (MD 2분할, 일액 정수)
    - 재료비(구매품/가공품): 전체 입력
    
    컬럼 매핑 (템플릿 Row 4 기준):
      A(1)=구분  B(2)=단가유형  C(3)=No  D(4)=자재명  E(5)=규격
      F(6)=Maker  G(7)=수량  H(8)=단위  I(9)=단가  J(10)=금액(수식)  K(11)=비고
    """
    import re
    import json
    from collections import Counter
    from openpyxl.cell.cell import MergedCell
    from openpyxl.styles import Alignment
    from classification import TEMPLATE_SECTIONS

    # 표준 컬럼 → 템플릿 컬럼 인덱스
    COL_NO = 3         # C: No
    COL_JARAEM = 4     # D: 자재명
    COL_GYEUK = 5      # E: 규격
    COL_MAKER = 6      # F: Maker
    COL_SANGYANG = 7   # G: 수량
    COL_DANW = 8       # H: 단위
    COL_DANGA = 9      # I: 단가
    COL_JEOMYEOK = 10  # J: 금액(수식)
    COL_BIGO = 11      # K: 비고 (Raw 견적 비고 내용 그대로)
    COL_VERIFY = 13    # M: 검증결과 (K열에서 한 열(L) 띄운 별도 열)

    # 인건비/경비 섹션 키 (단가 변경 금지, 행 추가 없음)
    LABOR_EXPENSE_SECTIONS = {"설계", "조립", "셋업", "양산대응", "특수제어", "포장", "기타항목", "경비"}
    # _fill_labor_rows가 labor_data 기반으로 전부 채우는 섹션.
    #    이 섹션의 원본 항목(Unit 행 포함)은 쓰지 않는다.
    #    (예: '(5)전장배선'은 단위='식'이라 Unit으로 오감지되어 셋업/양산대응에
    #     분할 배치되지만, 금액은 이미 labor_data에 포함되어 _fill_labor_rows가
    #     채우므로 Unit 행을 쓰면 중복된다.)
    LABOR_FILL_SECTIONS = {"설계", "조립", "셋업", "양산대응", "특수제어"}
    # 재료비 섹션 키 (행 추가 대상)
    MATERIAL_SECTIONS = {"기구_구매품", "기구_가공품", "제어_구매품", "제어_가공품", "전장_구매품", "전장_가공품"}
    # 재료비 섹션별 B열(단가유형) 라벨 시작 행
    B_LABEL_ROW = {
        "기구_구매품": 5, "기구_가공품": 10,
        "제어_구매품": 15, "제어_가공품": 21,
        "전장_구매품": 23, "전장_가공품": 30,
    }
    # 섹션 키 → 하위 구분 라벨 (합산 행 D~F 병합 텍스트용)
    SECTION_LABEL = {
        "기구_구매품": "구매품", "기구_가공품": "가공품",
        "제어_구매품": "제어부_구매품", "제어_가공품": "제어부_가공품",
        "전장_구매품": "전장부_구매품", "전장_가공품": "전장부_가공품",
    }

    # 1) 항목을 분류별로 그룹핑
    #    표준모듈 작성방법: "모든 행의 데이터 누락 없도록" →
    #    미분류/알 수 없는 분류의 항목은 재료비 기본 섹션(기구_구매품)으로 폴백 배치.
    items_by_section: Dict[str, List[Dict[str, Any]]] = {}
    known_sections = {s[0] for s in TEMPLATE_SECTIONS}
    
    # 소계/합계 행 제외 후 필터링
    filtered_items = [item for item in items if not item.get("_소계")]
    
    # Pass 1: 각 Unit(합산) 행에 family별 번호 부여
    #          기구부/제어부/전장부 각각 1)~N) (섹션 family마다 1부터 시작)
    _SECTION_TO_FAMILY = {
        "기구_구매품": "기구부", "기구_가공품": "기구부",
        "제어_구매품": "제어부", "제어_가공품": "제어부",
        "전장_구매품": "전장부", "전장_가공품": "전장부",
    }
    unit_no_by_family: Dict[str, int] = {}
    for item in filtered_items:
        if item.get("_합산"):
            sec = item.get("_분류", "")
            family = _SECTION_TO_FAMILY.get(sec, sec or "기타")
            unit_no_by_family[family] = unit_no_by_family.get(family, 0) + 1
            item["_unit_no"] = unit_no_by_family[family]

    # Pass 2: 항목 배정 — Unit 행을 세부항목 섹션별로 분할 생성
    #          (Unit의 세부항목이 구매품/가공품 등 여러 섹션에 걸리면
    #           섹션마다 별도 Unit 행 생성: BASE_FRAME_구매품 / BASE_FRAME_가공품)
    # family(재료비/인건비·경비)별 섹션 — Unit 하위 스캔은 같은 family 안에서만 진행.
    # (인건비/경비 Unit의 스캔이 푸터/메모 행을 지나 재료비 항목까지 닿아
    #  Unit이 재료비 섹션으로 잘못 분할되는 것을 방지)
    _MATERIAL_SECS = {"기구_구매품", "기구_가공품", "제어_구매품", "제어_가공품", "전장_구매품", "전장_가공품"}
    _LABOR_SECS = {"설계", "조립", "셋업", "양산대응", "특수제어", "포장", "기타항목", "경비"}

    def _sec_family(sec: str):
        if sec in _MATERIAL_SECS:
            return "material"
        if sec in _LABOR_SECS:
            return "labor"
        return None

    for idx, item in enumerate(filtered_items):
        if item.get("_합산"):
            unit_family = _sec_family(item.get("_분류", ""))
            # 하위 항목을 섹션별로 그룹핑 (첫 등장 순서)
            first_idx: Dict[str, int] = {}
            for j in range(idx + 1, len(filtered_items)):
                if filtered_items[j].get("_합산"):
                    break
                sec = filtered_items[j].get("_분류", "")
                # Unit과 다른 family의 항목 → Unit 하위 범위 종료
                if unit_family and _sec_family(sec) not in (None, unit_family):
                    break
                if sec in known_sections and sec not in first_idx:
                    first_idx[sec] = j
            if first_idx:
                # 각 섹션의 첫 세부항목에 "Unit 분할 행을 앞에 삽입" 태그
                for sec, j in first_idx.items():
                    split = dict(item)  # Unit 분할 행 (사본)
                    split["_split_section"] = sec
                    filtered_items[j].setdefault("_unit_split_before", []).append(split)
            else:
                # 하위 항목 없음 → 폴백 섹션에 배치
                sec = item.get("_분류", "")
                if sec not in known_sections:
                    sec = "기구_구매품"
                items_by_section.setdefault(sec, []).append(item)
            continue
        sec = item.get("_분류", "미분류")
        if sec not in known_sections:
            sec = "기구_구매품"  # 폴백: 미분류 → 기구부 구매품
        # 직전 Unit 분할 행 삽입 (태그가 있으면)
        for split in item.get("_unit_split_before", []):
            items_by_section.setdefault(sec, []).append(split)
        items_by_section.setdefault(sec, []).append(item)

    # Pass 3: Unit 번호를 "실제 배치된 섹션" 기준으로 재매기.
    #          Pass 1의 번호는 Unit 자체의 _분류 기준이라, Unit이 다른 섹션에
    #          분할 배치된 경우(예: PLC Unit이 _분류=제어_구매품인데 하위 사급품이
    #          기구_구매품이라 기구_구매품에 배치) 번호가 어긋난다.
    #          → 배치된 섹션의 family(기구부/제어부/전장부)별로 1)~N) 연속.
    for sec_key, sec_items in items_by_section.items():
        family = _SECTION_TO_FAMILY.get(sec_key)
        if family is None:
            continue  # 인건비/경비 섹션은 Unit 번호 없음
        unit_no = 0
        for it in sec_items:
            if it.get("_합산"):
                unit_no += 1
                it["_unit_no"] = unit_no

    # ── 디버그 덤프: 섹션 배정 결과 (data/_debug_export.json) ──
    try:
        _id_to_sec = {}
        for _k, _v in items_by_section.items():
            for _it in _v:
                _id_to_sec[id(_it)] = _k
        _dbg = {
            "items": [
                {
                    "row": it.get("_행"),
                    "unit": bool(it.get("_합산")),
                    "unit_name": it.get("_unit_name", ""),
                    "soye": bool(it.get("_소계")),
                    "bunryu": it.get("_분류", ""),
                    "section": _id_to_sec.get(id(it), ""),
                    "name": str(it.get("품명", ""))[:30],
                }
                for it in filtered_items
            ],
            "section_counts": {k: len(v) for k, v in items_by_section.items()},
        }
        _dbg_path = Path(__file__).resolve().parent.parent / "data" / "_debug_export.json"
        _dbg_path.parent.mkdir(parents=True, exist_ok=True)
        _dbg_path.write_text(json.dumps(_dbg, ensure_ascii=False, indent=1), encoding="utf-8")
    except Exception:
        pass

    all_sections = TEMPLATE_SECTIONS  # top-down 순서

    # 2) 재료비 섹션별 필요 삽입 행 수 계산
    #    표준모듈 작성방법: "모든 행의 데이터 누락 없도록, 필요시 행추가를 해서 모두 명기"
    insert_plan: Dict[str, Dict[str, int]] = {}
    for sec_key, gubun, sub_sec, start_row, end_row in all_sections:
        if sec_key not in MATERIAL_SECTIONS:
            continue
        n_items = len(items_by_section.get(sec_key, []))
        available = max(0, end_row - start_row + 1)
        need = max(0, n_items - available)
        if need > 0:
            insert_plan[sec_key] = {
                "insert_point": end_row + 1,  # 소계 행 (여기에 행 삽입)
                "count": need,
                "data_start": start_row,
                "data_end": end_row,
            }

    # 3) bottom-up 행 삽입 (수식/병합 셀 참조 보정 포함)
    #    아래 섹션부터 삽입하면 위 섹션의 행 번호가 변하지 않아 안전.
    for sec_key, gubun, sub_sec, start_row, end_row in reversed(all_sections):
        if sec_key not in insert_plan:
            continue
        plan = insert_plan[sec_key]
        _insert_rows_with_remap(
            ws, plan["insert_point"], plan["count"],
            style_ref_row=plan["data_start"],
        )
        # 소계 수식 확장 (새 데이터 범위까지)
        new_subtotal = plan["insert_point"] + plan["count"]
        new_data_end = plan["data_end"] + plan["count"]
        ws.cell(new_subtotal, COL_JEOMYEOK).value = f"=SUM(J{plan['data_start']}:J{new_data_end})"
        # B열(단가유형) 병합을 새 소계 행까지 확장
        _extend_b_merge(ws, B_LABEL_ROW[sec_key], new_subtotal)

    # 3.5) 템플릿의 기존 아웃라인(레벨 1 섹션 그룹) 제거
    #      Unit(합산) 그룹만 레벨 1로 유지 → 좌측 +/− 버튼이 1단계만 표시
    for _rd in ws.row_dimensions.values():
        _rd.outlineLevel = 0

    # 4) 데이터 입력 (top-down, 최종 행 위치 계산)
    shift = 0
    unit_detail_map: Dict[int, List[int]] = {}  # Unit 행 → 하위 세부항목 행 (그룹핑용, 전역)
    for sec_key, gubun, sub_sec, start_row, end_row in all_sections:
        section_items = items_by_section.get(sec_key, [])
        n_inserted = insert_plan.get(sec_key, {}).get("count", 0)

        if not section_items:
            shift += n_inserted
            continue

        is_labor_expense = sec_key in LABOR_EXPENSE_SECTIONS
        fstart = start_row + shift
        fend = end_row + n_inserted + shift

        item_seq = 0    # 세부항목 No (1, 2, 3 ...) - Unit 행마다 리셋
        unit_seq = 0    # 합산(Unit) 행 No (1), 2), 3) ...) - 섹션(그룹)마다 리셋
        detail_rows = []  # 세부항목(비 Unit) 행 위치 - 소계 수식 재작성용
        for idx, item in enumerate(section_items):
            r = fstart + idx
            if r > fend:
                break  # 안전장치 (이론상 발생 안 함)

            if sec_key in LABOR_FILL_SECTIONS:
                # _fill_labor_rows가 labor_data 기반으로 전부 채우는 섹션.
                # Unit 행 포함 모든 원본 항목을 쓰지 않는다.
                # (예: '(5)전장배선'은 Unit으로 오감지되어 이 섹션에 배치되지만,
                #  쓰면 템플릿 인건비 라벨 행을 덮어쓰고 금액이 중복된다.)
                pass
            elif is_labor_expense and not item.get("_합산"):
                # 인건비/경비 세부항목: labor_data 기반 후처리에서 채우기 (아래 _fill_labor_rows)
                pass
            elif item.get("_합산"):
                # ── 합산(Unit) 행: No '1)' + D~F 병합 + 단가=합산 수식 ──
                unit_seq += 1
                u_no = item.get("_unit_no") or unit_seq  # 글로벌 Unit 번호 (분할 행 동일)
                item_seq = 0  # 새 Unit 시작 → 세부항목 No 리셋
                # No: '1)', '2)', ... (세부항목과 별도 번호 체계)
                no_cell = ws.cell(r, COL_NO)
                if not isinstance(no_cell, MergedCell):
                    no_cell.value = f"{u_no})"
                    no_cell.alignment = Alignment(horizontal="center", vertical="center")
                # D~F 병합: 재료비는 'BASE_FRAME_구매품' (Unit명_섹션라벨,
                #           구매품/가공품 분할 구분용), 인건비·경비는 원본 Unit명 그대로.
                unit_name = item.get("_unit_name") or item.get("품명", "")
                if is_labor_expense:
                    text = unit_name or SECTION_LABEL.get(sec_key, sec_key)
                else:
                    sub_label = SECTION_LABEL.get(sec_key, sec_key)
                    text = f"{unit_name}_{sub_label}" if unit_name else sub_label
                _merge_df_text(ws, r, text)
                # G(수량)=1, H(단위)=식
                _safe_set_cell(ws, r, COL_SANGYANG, 1)
                _safe_set_cell(ws, r, COL_DANW, "식")
                # I(단가): 하위 세부항목이 있으면 합산 수식, 없으면 원본 단가 그대로
                sub_start = r + 1
                sub_end = None
                for j in range(idx + 1, len(section_items)):
                    if section_items[j].get("_합산"):
                        break
                    sub_end = fstart + j
                if sub_end is not None and sub_end >= sub_start:
                    _safe_set_cell(ws, r, COL_DANGA, f"=SUM(J{sub_start}:J{sub_end})")
                else:
                    # 하위 항목이 없는 Unit 행 → 원본 단가 그대로
                    _safe_set_cell(ws, r, COL_DANGA, item.get("단가", ""))
                    detail_rows.append(r)
                # J(금액): =G*I (1 * 합산단가 = 합산금액)
                j_cell = ws.cell(r, COL_JEOMYEOK)
                if not isinstance(j_cell, MergedCell) and j_cell.value in (None, ""):
                    j_cell.value = f"=G{r}*I{r}"
                # 합산 행 스타일: 굵은 글씨 + 금액(J열) 흐린 회색 채움
                _style_sum_row(ws, r)
                # 하위 세부항목 행 기록 (Excel 아웃라인 그룹핑용)
                unit_detail_map[r] = []
            else:
                # 재료비(구매품/가공품) 세부항목: No + 전체 입력
                item_seq += 1
                detail_rows.append(r)
                # No: 문자 중간(수평) 배치
                no_cell = ws.cell(r, COL_NO)
                if not isinstance(no_cell, MergedCell):
                    no_cell.value = item_seq
                    no_cell.alignment = Alignment(horizontal="center", vertical="center")
                # 직전 Unit 행의 하위 항목으로 기록 (그룹핑용)
                if unit_detail_map:
                    unit_detail_map[sorted(unit_detail_map)[-1]].append(r)
                # D열(자재명) = 품명 (Raw에서 품명/부품명 등 다양한 표현 매칭)
                _safe_set_cell(ws, r, COL_JARAEM, item.get("품명", ""))
                _safe_set_cell(ws, r, COL_GYEUK, item.get("규격", ""))
                _safe_set_cell(ws, r, COL_MAKER, item.get("Maker", ""))
                _safe_set_cell(ws, r, COL_SANGYANG, item.get("수량", ""))
                _safe_set_cell(ws, r, COL_DANW, item.get("단위", ""))
                _safe_set_cell(ws, r, COL_DANGA, item.get("단가", ""))
                # J열(금액): =G*I 수식 (빈 셀에만, 기존 수식/값은 건드리지 않음)
                j_cell = ws.cell(r, COL_JEOMYEOK)
                if not isinstance(j_cell, MergedCell) and j_cell.value in (None, ""):
                    j_cell.value = f"=G{r}*I{r}"
                # K열(비고): Raw 견적 비고 내용 그대로
                _safe_set_cell(ws, r, COL_BIGO, str(item.get("비고", "")))
                # M열(검증결과): K열에서 한 열(L) 띄운 별도 열에 검증 문구 명기
                verify_parts = []
                if item.get("검증결과"):
                    verify_parts.append(str(item["검증결과"]))
                if show_original and item.get("_원본명"):
                    verify_parts.append(f"[원본: {item['_원본명']}]")
                if verify_parts:
                    _safe_set_cell(ws, r, COL_VERIFY, " | ".join(verify_parts))
            # 인건비/경비: K열(비고)은 _fill_labor_rows가 특급/고급 행에만 작성

        # 소계 수식 재작성: Unit(합산) 행은 하위 항목 합산이므로
        # 소계가 Unit 행까지 SUM하면 이중계산 → 세부항목만 합산.
        if detail_rows and sec_key in MATERIAL_SECTIONS:
            subtotal_row = fend + 1
            c_val = ws.cell(subtotal_row, COL_NO).value
            if c_val and "소계" in str(c_val):
                ws.cell(subtotal_row, COL_JEOMYEOK).value = _sum_formula(detail_rows)

        shift += n_inserted

    # 4.4) 인건비/경비 행 채우기: 인원×일수 기반 MD 분배 + M열 표기
    labor_md_notes = _fill_labor_rows(ws, labor_data, shift)
    for r, text in labor_md_notes:
        _safe_set_cell(ws, r, COL_VERIFY, text)

    # 4.5) 섹션 구분 행 삭제: 소계 행과 다음 하위 섹션(가공품) 사이의 빈칸 행.
    #      템플릿 구조상 B열에만 하위 섹션 라벨('가공품' 등)이 있고 C~K는 빈칸.
    #      아래쪽부터 삭제하면 위쪽 행 번호가 유지되어 안전.
    sep_rows = []
    shift2 = 0
    for sec_key, gubun, sub_sec, start_row, end_row in all_sections:
        if sec_key in ("기구_가공품", "제어_가공품", "전장_가공품"):
            sep = start_row + shift2 - 1
            # 안전장치: B열에 하위 섹션 라벨('가공품')이 있는 구분 행일 때만 삭제
            if ws.cell(sep, 2).value not in (None, ""):
                sep_rows.append(sep)
        shift2 += insert_plan.get(sec_key, {}).get("count", 0)

    for sep in sorted(sep_rows, reverse=True):
        _delete_row_with_remap(ws, sep)
        # Unit→세부항목 행 위치 리맵 (삭제 행 아래는 1씩 상향)
        unit_detail_map = {
            (u - 1 if u > sep else u): [d - 1 if d > sep else d for d in drows]
            for u, drows in unit_detail_map.items()
        }

    # Unit(합산) 행 → 하위 세부항목 Excel 아웃라인 그룹핑 (+/− 접기/펼치기)
    # 기존 템플릿 그룹은 제거했으므로 Unit 그룹이 레벨 1.
    # summaryBelow=False 이므로 +/− 버튼이 Unit 행(그룹 상단)에 표시됨.
    for u_row, d_rows in unit_detail_map.items():
        if d_rows:
            ws.row_dimensions.group(d_rows[0], d_rows[-1], outline_level=1)

    # 아웃라인 요약 버튼이 Unit(합산) 행(상단)에 표시되도록 설정
    ws.sheet_properties.outlinePr.summaryBelow = False

    # 행 높이 통일: 삽입된 행은 높이가 안 잡혀 크기가 혼재됨 →
    # 데이터 영역(5행~마지막) 전체를 표준 포멧 높이(12.0)로 통일.
    for r in range(5, ws.max_row + 1):
        ws.row_dimensions[r].height = 12.0

    return shift - len(sep_rows)  # 순 삽입 행 수 (요약 시트 교차 참조 교정용)


# ── 인건비: analyzer 카테고리 → 템플릿 행 라벨 (특급, 특급출현순, 고급, 고급출현순) ──
# 템플릿 라벨이 중복되는 행(예: 'SW설계자(무관)' x4, '셋업 제어/전장 엔지니어(특급)' x4)은
# 출현 순서로 특급=1번째, 고급=2번째 행을 지정.
_LABOR_ROW_MAP = {
    "인건비-기구 설계":        ("기구설계자(특급)", 0, "기구설계자(고급)", 0),
    "인건비-전장 설계":        ("전장설계자(특급)", 0, "전장설계자(고급)", 0),
    "인건비-제어 설계":        ("SW설계자(무관)", 0, "SW설계자(무관)", 1),
    "인건비-기구 조립":        ("기구조립(특급)", 0, "기구조립(고급)", 0),
    "인건비-전장/제어 조립":   ("제어조립(특급)", 0, "제어조립(고급)", 0),
    "인건비-기구 셋업":        ("셋업 기구 엔지니어(특급)", 0, "셋업 기구 엔지니어(고급)", 0),
    "인건비-전장/제어 셋업":   ("셋업 제어/전장 엔지니어(특급)", 0, "셋업 제어/전장 엔지니어(특급)", 1),
    "인건비-기구 양산대응":    ("양산 기구 엔지니어(특급)", 0, "양산 기구 엔지니어(고급)", 0),
    "인건비-전장/제어 양산대응": ("양산 제어/전장 엔지니어(특급)", 0, "양산 제어/전장 엔지니어(특급)", 1),
    "인건비-특수제어":         ("제어 엔지니어(무관)", 0, None, 0),
}


def _find_label_row(ws, label: str, start: int, end: int, occurrence: int = 0):
    """D열(4)이 label과 일치하는 행 번호를 반환 (occurrence번째 출현, 0-based). 없으면 None."""
    count = 0
    for r in range(start, end + 1):
        v = ws.cell(r, 4).value
        if v is not None and str(v).strip() == label:
            if count == occurrence:
                return r
            count += 1
    return None


# MD 표기용 기준 일수 후보 (비고에 인원×일수 정보가 없을 때)
_MD_DAY_CANDIDATES = (7, 14, 21, 30, 60)


def _calc_labor_md_split(md: float, remark: str):
    """
    MD를 (대응인원수 × 대응일수) 기준으로 특급/고급 2분할 계산.
    특급/고급은 대응 일수를 동일하게 유지하고 인원만 분배하며, 총 MD는 보존한다.
    
    규칙:
    - 비고에 인원×일수가 하나면 그 일수를 특급/고급 공통으로 사용 (인원만 분배)
    - 비고에 인원×일수가 여러 개면 MD/총인원으로 평균 일수(반올림)를 공통 사용
    - 비고에 정보가 없으면 7/14/21/30/60일 중 총 MD를 맞추는 적절한 인원수를 탐색
    - 인원 2분할(특급=ceil)로 총 MD가 어긋나면 차이를 특급 MD에 보정
    
    Returns:
        (spec_md, adv_md, display) — display: '4명×20일+2명×30일' 형식
    """
    import math
    import re
    md = int(round(md or 0))
    if md <= 0:
        return 0, 0, ""

    # 비고에서 인원×일수 쌍 추출: '4名x20日', '2명×30일', '1名x45日' 등
    pairs = [(int(p), int(d)) for p, d in re.findall(
        r"(\d+)\s*[名名명]\s*[xX×*]\s*(\d+)\s*[日日일]", remark or "")]

    if not pairs:
        # 비고 없음: 기준 일수 후보 × 인원수로 MD 재현 (일수 작은 쪽 우선, 인원 1~10)
        best = None
        for days in _MD_DAY_CANDIDATES:
            for people in range(1, 11):
                if people * days == md:
                    best = (people, days)
                    break
            if best:
                break
        if best is None:
            days = min(_MD_DAY_CANDIDATES, key=lambda d: abs(md - round(md / d) * d))
            best = (max(1, round(md / days)), days)
        p, d = best
        spec_p = math.ceil(p / 2)
        adv_p = p - spec_p
        spec_md, adv_md = spec_p * d, adv_p * d
        spec_md += md - (spec_md + adv_md)  # 총 MD 보존 (차이 → 특급)
        return spec_md, adv_md, f"{spec_p}명×{d}일+{adv_p}명×{d}일"

    if len(pairs) == 1:
        # 단일 표현: 일수 공통, 인원 2분할 (특급=ceil)
        p, d = pairs[0]
        spec_p = math.ceil(p / 2)
        adv_p = p - spec_p
        spec_md, adv_md = spec_p * d, adv_p * d
        spec_md += md - (spec_md + adv_md)  # 총 MD 보존 (차이 → 특급)
        return spec_md, adv_md, f"{spec_p}명×{d}일+{adv_p}명×{d}일"

    # 여러 표현: MD/총인원으로 평균 일수(반올림)를 공통 사용
    total_p = sum(p for p, _ in pairs)
    days = max(1, int(round(md / total_p)))
    spec_p = math.ceil(total_p / 2)
    adv_p = total_p - spec_p
    spec_md, adv_md = spec_p * days, adv_p * days
    spec_md += md - (spec_md + adv_md)  # 총 MD 보존 (차이 → 특급)
    return spec_md, adv_md, f"{spec_p}명×{days}일+{adv_p}명×{days}일"


def _fill_labor_rows(ws, labor_data: Optional[Dict[str, Dict[str, Any]]], shift: int):
    """
    인건비/경비 행을 labor_data 기반으로 채운다.
    
    규칙 (인건비 기본 규칙):
    - MD = 대응인원수 × 대응일수 로 계산 (특급/고급은 일수 동일, 인원만 분배)
    - 일액(I열) 정수로 작성: 고급=int(A/M), 특급=int((A-고급금액)/특급MD) → 나머지는 특급에 추가
    - 동일한 비고(K열)를 특급/고급 양 행에 복제
    - M열(검증결과): 'MD=140 (4명×20일+2명×30일)' 형식으로 인원×일수 표기
    - Engineering 행: 전체 MD 21 미만 → 7(몇 주)의 배수, 21 이상 → 30(몇 달)의 배수
    - J열(금액)은 템플릿의 =G*I 수식이 자동 계산
    
    Returns:
        [(행번호, M열 표기 텍스트), ...] — 호출부에서 M열에 기록
    """
    if not labor_data:
        return []
    import math
    from openpyxl.cell.cell import MergedCell

    # 인건비 영역 (템플릿 R34~R77 + 재료비 삽입 시프트)
    lo, hi = 34 + shift, 77 + shift
    notes: List[tuple] = []

    def set_cell(r, col, val):
        if r is None:
            return
        c = ws.cell(r, col)
        if not isinstance(c, MergedCell):
            c.value = val

    # ── Engineering: 전체 MD 기반 배수 반올림 ──
    # 21일 미만 → 7(몇 주)의 배수, 21일 이상 → 30(몇 달)의 배수에 가장 가까운 값
    total_md = sum(v.get("md", 0) or 0 for v in labor_data.values())
    if total_md > 0:
        eng_row = _find_label_row(ws, "Engineering", lo, hi)
        if total_md < 21:
            eng_md = int(round(total_md / 7.0)) * 7
        else:
            eng_md = int(round(total_md / 30.0)) * 30
        set_cell(eng_row, 7, eng_md)

    # ── 카테고리별 특급/고급 분배 (인원×일수 기준) ──
    for cat, info in labor_data.items():
        spec_label, spec_occ, adv_label, adv_occ = _LABOR_ROW_MAP.get(cat, (None, 0, None, 0))
        if spec_label is None:
            continue
        md = info.get("md", 0) or 0
        amt = info.get("amount", 0) or 0
        remark = info.get("remark", "") or ""
        spec_row = _find_label_row(ws, spec_label, lo, hi, spec_occ)
        if spec_row is None:
            continue
        adv_row = _find_label_row(ws, adv_label, lo, hi, adv_occ) if adv_label else None

        if md <= 0 and amt <= 0:
            continue  # 금액 0 → 템플릿 기본값 유지 (G=0 → J=0)

        spec_md, adv_md, display = _calc_labor_md_split(md, remark)

        if adv_row is None:
            # 단일 행 (특수제어 등) — 단가(I열)는 템플릿 고정값 유지
            if spec_md > 0:
                set_cell(spec_row, 7, spec_md)
            if remark:
                set_cell(spec_row, 11, remark)
            if spec_md > 0:
                notes.append((spec_row, f"MD={spec_md} ({display})"))
            continue

        # 단가(I열)는 템플릿 고정값 유지 — G열(MD)만 채움
        set_cell(spec_row, 7, spec_md)
        if remark:
            set_cell(spec_row, 11, remark)
        set_cell(adv_row, 7, adv_md)
        if remark:
            set_cell(adv_row, 11, remark)
        # M열 표기: 'MD=140 (4명×20일+2명×30일)'
        notes.append((spec_row, f"MD={int(round(md))} ({display})"))

    return notes


def _sum_formula(rows: List[int]) -> str:
    """
    행 목록을 연속 구간으로 묶어 SUM 수식 생성.
    예: [7,8,9,11,12] → "=SUM(J7:J9,J11:J12)"
    """
    if not rows:
        return ""
    rows = sorted(rows)
    ranges = []
    seg_start = rows[0]
    prev = rows[0]
    for r in rows[1:]:
        if r == prev + 1:
            prev = r
        else:
            ranges.append((seg_start, prev))
            seg_start = r
            prev = r
    ranges.append((seg_start, prev))
    parts = [f"J{a}:J{b}" if a != b else f"J{a}" for a, b in ranges]
    return f"=SUM({','.join(parts)})"


def _insert_rows_with_remap(ws, insert_point: int, count: int, style_ref_row: int = None):
    """
    insert_point 행 앞에 count 행을 삽입하고, 아래쪽 수식 참조와 병합 셀을 보정.
    
    openpyxl의 insert_rows는 수식 참조/병합 셀을 자동으로 수정하지 않으므로
    수동으로 리맵한다. 또한 style_ref_row의 스타일(테두리 등)을 새 행에
    복사해 표준 포멧 형태로 표현한다.
    """
    import re
    from copy import copy
    CELL_REF_RE = re.compile(r"(?<![A-Za-z0-9_])([A-Z]{1,3})(\d+)")
    RANGE_RE = re.compile(r"([A-Z]+)(\d+):([A-Z]+)(\d+)")

    def new_row(r: int) -> int:
        return r + (count if r >= insert_point else 0)

    # 1) 삽입 전 수식 수집 (ws._cells 딕셔너리 스캔 — ws.cell() 반복보다 수백 배 빠름)
    formulas = []
    for (r, c), cell in ws._cells.items():
        v = cell.value
        if isinstance(v, str) and v.startswith("="):
            formulas.append((r, c, v))

    # 2) 병합 범위 수집
    merged = [str(m) for m in ws.merged_cells.ranges]

    # 3) 스타일 참조 행 수집 (참조 행은 insert_point 위이므로 시프트 없음)
    ref_styles = None
    if style_ref_row is not None:
        ref_styles = []
        for c in range(1, ws.max_column + 1):
            ref_styles.append(copy(ws.cell(style_ref_row, c)._style))

    # 4) 병합 해제 (insert_rows가 병합을 깨뜨릴 수 있음)
    for m in list(ws.merged_cells.ranges):
        ws.unmerge_cells(str(m))

    # 5) 행 삽입
    ws.insert_rows(insert_point, count)

    # 6) 수식 리맵 (행 참조 시프트)
    def remap(f: str) -> str:
        return CELL_REF_RE.sub(lambda m: f"{m.group(1)}{new_row(int(m.group(2)))}", f)
    for (r, c, t) in formulas:
        ws.cell(new_row(r), c).value = remap(t)

    # 7) 병합 재적용 (리맵)
    for rng in merged:
        m = RANGE_RE.match(rng)
        if not m:
            continue
        c1, r1, c2, r2 = m.group(1), int(m.group(2)), m.group(3), int(m.group(4))
        ws.merge_cells(f"{c1}{new_row(r1)}:{c2}{new_row(r2)}")

    # 8) 새 행에 표준 포멧 스타일(테두리/폰트/채움/숫자형식) 적용
    if ref_styles is not None:
        for i in range(count):
            new_r = insert_point + i
            for c in range(1, len(ref_styles) + 1):
                ws.cell(new_r, c)._style = copy(ref_styles[c - 1])


def _delete_row_with_remap(ws, row: int):
    """
    row 행을 1개 삭제하고, 아래쪽 수식 참조와 병합 셀을 보정.
    
    openpyxl의 delete_rows는 수식 참조/병합 셀을 자동으로 수정하지 않으므로
    수동으로 리맵한다. 삭제 행이 병합의 상단이면 값은 아래 행으로 이동
    (B열 '가공품' 라벨 유지용).
    """
    import re
    CELL_REF_RE = re.compile(r"(?<![A-Za-z0-9_])([A-Z]{1,3})(\d+)")
    RANGE_RE = re.compile(r"([A-Z]+)(\d+):([A-Z]+)(\d+)")

    def new_row(r: int) -> int:
        return r - 1 if r > row else r

    # 1) 삭제 전 수식 수집 — 삭제 행을 참조하는 수식만 리맵 대상.
    #    ⚠️ 삭제 행 "위쪽" 수식도 아래 행을 참조할 수 있음 (예: R5 합계 행의
    #    =SUM(J786,J1799)). 위쪽 수식까지 무시하면 참조가 1행 어긋나
    #    합계 행을 자기 참조 → 순환 참조 오류가 발생한다.
    #    성능: 전체 수식을 수집하되, 삭제 행 이하를 참조하는 수식만 리맵.
    #    ws._cells(딕셔너리)를 스캔하면 ws.cell() 반복보다 수백 배 빠름.
    formulas = []
    for (r, c), cell in ws._cells.items():
        v = cell.value
        if isinstance(v, str) and v.startswith("="):
            formulas.append((r, c, v))

    # 2) 전체 병합 범위 수집
    #    ⚠️ openpyxl의 delete_rows는 병합 셀을 자동으로 조정하지 않는다
    #    (테스트 확인). 삭제 행 "아래"의 병합도 1행씩 위로 이동해야 하므로
    #    모든 병합을 해제 → 삭제 → 리맵 후 재적용해야 한다.
    merged = [str(m) for m in ws.merged_cells.ranges]

    # 3) 전체 병합 해제 (delete_rows가 병합을 깨뜨릴 수 있음)
    #    unmerge_cells가 범위 내 모든 셀을 del하므로, 존재하지 않는 셀이 있으면
    #    KeyError → 먼저 셀을 생성해 둔다.
    for m in list(ws.merged_cells.ranges):
        for r in range(m.min_row, m.max_row + 1):
            for c in range(m.min_col, m.max_col + 1):
                ws.cell(r, c)
        ws.unmerge_cells(str(m))

    # 4) 삭제 행이 병합 상단이면 값+스타일을 아래 행으로 이동 (라벨 유지)
    #    (병합 해제 후 일반 셀이므로 값/스타일 설정 가능)
    #    스타일을 함께 복사하지 않으면 B열 '가공품' 배경색이 유실됨.
    from openpyxl.utils import column_index_from_string
    from copy import copy as _copy
    for rng in merged:
        m = RANGE_RE.match(rng)
        if not m:
            continue
        r1, r2 = int(m.group(2)), int(m.group(4))
        if r1 == row and r2 > row:
            c1 = column_index_from_string(m.group(1))
            top_cell = ws.cell(row, c1)
            dst_cell = ws.cell(row + 1, c1)
            if top_cell.value is not None:
                dst_cell.value = top_cell.value
            dst_cell._style = _copy(top_cell._style)

    # 5) 행 삭제
    ws.delete_rows(row, 1)

    # 6) 수식 리맵 (행 참조 시프트)
    #    삭제 행 이하를 참조하는 수식만 리맵 (나머지는 참조가 안 변하므로
    #    재작성을 건너뛰어 성능 유지).
    def remap(f: str) -> str:
        return CELL_REF_RE.sub(lambda m: f"{m.group(1)}{new_row(int(m.group(2)))}", f)
    for (r, c, t) in formulas:
        if r == row:
            continue  # 삭제된 행의 수식은 폐기
        if not any(int(m.group(2)) >= row for m in CELL_REF_RE.finditer(t)):
            continue  # 삭제 행 이하 참조 없음 → 변경 불필요
        ws.cell(new_row(r), c).value = remap(t)

    # 7) 병합 재적용 (리맵)
    #    - 삭제 행 "전체"에 있는 병합(예: E21:F21)은 행이 사라지므로 폐기.
    #      폐기하지 않으면 아래로 올라온 병합(예: C22:I22→C21:I21)과 같은 행에서
    #      겹쳐서 Excel이 "내용 복구"를 요구하는 손상 파일이 된다.
    #    - 삭제 행이 병합 범위에 포함되면 범위가 축소될 수 있음.
    #    - 단일 셀로 축소된 범위(예: B23:B24 → B23:B23)는 병합 불가 → 건너뜀.
    for rng in merged:
        m = RANGE_RE.match(rng)
        if not m:
            continue
        c1, r1, c2, r2 = m.group(1), int(m.group(2)), m.group(3), int(m.group(4))
        if r1 == row and r2 == row:
            continue  # 삭제 행 전체에 있는 병합 → 폐기
        nr1, nr2 = new_row(r1), new_row(r2)
        if nr1 == nr2 and c1 == c2:
            continue  # 단일 셀 → 병합 불필요
        ws.merge_cells(f"{c1}{nr1}:{c2}{nr2}")


def _extend_b_merge(ws, b_label_row: int, new_subtotal_row: int):
    """B열(단가유형) 병합을 new_subtotal_row까지 확장 (단가유형 라벨 표시 유지)."""
    target = None
    for m in list(ws.merged_cells.ranges):
        if m.min_col == 2 and m.max_col == 2 and m.min_row == b_label_row:
            target = m
            break
    if target is None:
        return
    if target.max_row >= new_subtotal_row:
        return  # 이미 충분히 넓음
    ws.unmerge_cells(str(target))
    ws.merge_cells(f"B{b_label_row}:B{new_subtotal_row}")


def _fix_summary_cross_refs(ws, total_inserted: int):
    """
    요약 시트의 '2.상세 견적'! 참조를 행 삽입분만큼 보정.
    예: ='2.상세 견적'!J107 → ='2.상세 견적'!J{107+total_inserted}
    """
    if total_inserted == 0:
        return
    import re
    for r in range(1, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            v = ws.cell(r, c).value
            if isinstance(v, str) and v.startswith("=") and "2.상세 견적" in v:
                new_v = re.sub(
                    r"('2\.상세 견적'!)([A-Z]{1,3})(\d+)",
                    lambda m: f"{m.group(1)}{m.group(2)}{int(m.group(3)) + total_inserted}",
                    v,
                )
                if new_v != v:
                    ws.cell(r, c).value = new_v


def _get_formula_rows(ws) -> set:
    """
    수식이 포함된 행 번호를 수집.
    I열(9) 또는 J열(10)에 수식(=로 시작)이 있는 행을 반환.
    """
    formula_rows = set()
    for row_idx in range(1, ws.max_row + 1):
        for col in (9, 10):  # I열(단가), J열(금액)
            cell = ws.cell(row=row_idx, column=col)
            if cell.value and isinstance(cell.value, str) and cell.value.startswith("="):
                formula_rows.add(row_idx)
                break
    return formula_rows


def _safe_set_cell(ws, row: int, col: int, value):
    """병합 셀이 아닌 경우에만 값 설정"""
    from openpyxl.cell.cell import MergedCell
    cell = ws.cell(row=row, column=col)
    if not isinstance(cell, MergedCell):
        cell.value = value


def _merge_df_text(ws, row: int, text: str):
    """
    D~F 셀을 병합하고 합산(Unit) 행 텍스트를 입력.
    
    기존 D~F 병합이 있으면 해제 후 재병합.
    """
    from openpyxl.cell.cell import MergedCell
    from openpyxl.styles import Alignment
    
    # 기존 D~F 범위 병합 해제
    to_unmerge = []
    for m in list(ws.merged_cells.ranges):
        if m.min_row == row and m.max_row == row and m.min_col <= 6 and m.max_col >= 4:
            to_unmerge.append(str(m))
    for m in to_unmerge:
        ws.unmerge_cells(m)
    
    # D~F 병합
    ws.merge_cells(start_row=row, start_column=4, end_row=row, end_column=6)
    
    # D열에 텍스트 입력 (병합 후 상좌 셀)
    cell = ws.cell(row=row, column=4)
    if not isinstance(cell, MergedCell):
        cell.value = text
        cell.alignment = Alignment(horizontal="left", vertical="center")


def _style_sum_row(ws, row: int):
    """
    합산(Unit) 행 스타일 적용:
    - C~K 열: 굵은 글씨
    - C~J 열: 흐린 회색 채움 (D9D9D9)
    """
    from openpyxl.styles import Font, PatternFill
    from openpyxl.cell.cell import MergedCell

    bold = Font(bold=True)
    gray_fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")

    for col in range(3, 12):  # C(3) ~ K(11)
        cell = ws.cell(row=row, column=col)
        if isinstance(cell, MergedCell):
            continue
        # 기존 폰트 유지하면서 bold만 추가
        old_font = cell.font
        cell.font = Font(
            name=old_font.name, size=old_font.size,
            bold=True, italic=old_font.italic,
            color=old_font.color,
        )
        # C~J 열(3~10)에 회색 채움
        if col <= 10:
            cell.fill = gray_fill


def _detect_sections(ws) -> List[Dict[str, Any]]:
    """
    '2.상세 견적' 시트의 섹션 구조를 감지.
    
    템플릿의 A열(구분)과 B열(단가유형) 값을 기준으로 섹션 경계 파악.
    """
    sections = []
    current_gubun = None
    current_dangayuh = None
    section_start = None
    
    for row_idx in range(5, ws.max_row + 1):
        a_val = ws.cell(row=row_idx, column=1).value  # 구분
        b_val = ws.cell(row=row_idx, column=2).value  # 단가유형
        c_val = ws.cell(row=row_idx, column=3).value  # No
        
        # 합계 행이면 현재 섹션 종료
        if c_val and ("합계" in str(c_val) or "Sub합계" in str(c_val) or "TOTAL" in str(c_val)):
            if section_start is not None:
                sections.append({
                    "start_row": section_start,
                    "end_row": row_idx - 1,
                    "gubun": current_gubun or "",
                    "dangayuh": current_dangayuh or "",
                })
                section_start = None
            continue
        
        # 소계 행이면 스킵 (섹션 경계 아님)
        if c_val and "소계" in str(c_val):
            continue
        
        # 새 구분 시작
        if a_val and str(a_val).strip():
            # 이전 섹션 종료
            if section_start is not None:
                sections.append({
                    "start_row": section_start,
                    "end_row": row_idx - 1,
                    "gubun": current_gubun or "",
                    "dangayuh": current_dangayuh or "",
                })
            current_gubun = str(a_val).strip()
            current_dangayuh = str(b_val).strip() if b_val else ""
            section_start = row_idx
        elif b_val and str(b_val).strip() and section_start is not None:
            # 같은 구분 내 단가유형 변경 (구매품 → 가공품)
            prev_dangayuh = current_dangayuh
            current_dangayuh = str(b_val).strip()
            if prev_dangayuh != current_dangayuh:
                sections.append({
                    "start_row": section_start,
                    "end_row": row_idx - 1,
                    "gubun": current_gubun or "",
                    "dangayuh": prev_dangayuh,
                })
                section_start = row_idx
    
    # 마지막 섹션
    if section_start is not None:
        sections.append({
            "start_row": section_start,
            "end_row": ws.max_row,
            "gubun": current_gubun or "",
            "dangayuh": current_dangayuh or "",
        })
    
    return sections


def _filter_items_for_section(
    items: List[Dict[str, Any]],
    gubun: str,
    dangayuh: str,
) -> List[Dict[str, Any]]:
    """
    섹션(구분+단가유형)에 해당하는 항목 필터링.
    
    매핑된 항목의 '구분' 또는 '단가유형' 필드가 없으면
    품명/규격 키워드로 추론.
    """
    if not gubun:
        return []
    
    filtered = []
    for item in items:
        # 명시적 구분/단가유형 필드 확인
        item_gubun = str(item.get("구분", ""))
        item_dangayuh = str(item.get("단가유형", ""))
        
        if item_gubun:
            if item_gubun == gubun and (not item_dangayuh or item_dangayuh == dangayuh):
                filtered.append(item)
        else:
            # 키워드 기반 추론
            name = str(item.get("품명", "")) + str(item.get("규격", ""))
            if _match_section(name, gubun, dangayuh):
                filtered.append(item)
    
    return filtered


def _match_section(name: str, gubun: str, dangayuh: str) -> bool:
    """품명/규격 키워드로 섹션 매칭"""
    name_lower = name.lower()
    
    # 기구부
    if "기구" in gubun:
        if dangayuh == "구매품":
            return any(k in name_lower for k in ["기구", "금형", "프레임", "브라켓", "볼트", "너트"])
        elif dangayuh == "가공품":
            return any(k in name_lower for k in ["가공", "절단", "드릴", "타공", "절곡"])
    
    # 전장/제어
    if "전장" in gubun or "제어" in gubun:
        if dangayuh == "구매품":
            return any(k in name_lower for k in ["전장", "제어", "케이블", "컨넥터", "PCB", "센서", "모터"])
        elif dangayuh == "가공품":
            return any(k in name_lower for k in ["가공", "절단", "드릴"])
    
    # 인건비
    if "인건비" in gubun:
        return any(k in name_lower for k in ["설계", "조립", "셋업", "양산", "엔지니어", "인건"])
    
    # 경비
    if "경비" in gubun:
        return any(k in name_lower for k in ["포장", "운송", "항공", "숙박", "식비", "임대", "경비"])
    
    return False


def _fill_summary_sheet(ws, company_name: str, quote_no: str, items: List[Dict[str, Any]]):
    """
    '1.표준 견적서(갑지)' 시트의 메타 정보 업데이트.
    
    ⚠️ 수식 셀은 절대 변경하지 않는다.
    - C10: =H19 (수식)
    - F19: =H24 (수식)
    - H19: =ROUNDDOWN(D19*F19, -5) (수식)
    - C23: =B19 (수식)
    - G23: ='2.상세 견적'!P107 (수식)
    - H23: =G23*E23 (수식)
    - H24: =SUM(H23) (수식)
    
    변경 가능한 셀:
    - B8: 견적발행일 (텍스트)
    - B9: Subject (텍스트)
    - B19: 견적명 (텍스트)
    """
    from datetime import datetime
    
    # 견적발행일 (Row 8, B열) - 텍스트만 변경
    today = datetime.now()
    ws.cell(row=8, column=2, value=f"견적발행일: {today.year}년 {today.month}월 {today.day}일")
    
    # Subject (Row 9, B열) - 텍스트만 변경
    if quote_no:
        ws.cell(row=9, column=2, value=f"Subject : {quote_no}")
    
    # 견적명 (Row 19, B열) - 텍스트만 변경
    if company_name:
        ws.cell(row=19, column=2, value=company_name)
    
    # ⚠️ C23(=B19), G23(='2.상세 견적'!P107), H23(=G23*E23) 등 수식 셀은 건드리지 않음
    # 총액은 '2.상세 견적'!P107 수식이 자동 계산


# ─── 조정 견적 내보내기 ─────────────────────────────────────────
def export_adjusted_excel(base_bytes: bytes, model) -> bytes:
    """
    표준 양식 Excel(base_bytes)의 '2.상세 견적' 시트를 복사해
    마지막에 '3.조정 견적' 탭을 추가하고, 조정된 금액만 반영한다.

    포멧(수식/병합/스타일)은 그대로 유지하고, 조정 모델(RatioModel)에
    따라 금액만 변경한다:
      - 재료비·가공품: 단가(I열)를 변경 비율(%)만큼 스케일
        (구매품은 개별 변경 불가 → 그대로)
      - 인건비: MD(G열)를 섹션별 (변경/현견적) 비율만큼 스케일
      - 경비: 단가(I열)를 경비 변경 비율(%)만큼 스케일
    소계/합계/TOTAL은 모두 =SUM 수식이므로 자동 재계산된다.

    Args:
        base_bytes: export_standard_excel() 결과 bytes
        model: converters.ratio_adjuster.RatioModel (조정 입력이 적용된 상태)

    Returns:
        조정 탭이 추가된 Excel bytes
    """
    from converters.ratio_adjuster import SECTION_ORDER

    wb = load_workbook(BytesIO(base_bytes))
    if "2.상세 견적" not in wb.sheetnames:
        raise ValueError("표준 양식 Excel에 '2.상세 견적' 시트가 없습니다.")

    src = wb["2.상세 견적"]
    new_title = "3.조정 견적"
    if new_title in wb.sheetnames:
        del wb[new_title]
    new = wb.copy_worksheet(src)
    new.title = new_title
    # 마지막 탭으로 이동
    wb.move_sheet(new_title, offset=len(wb.sheetnames) - 1 - wb.sheetnames.index(new_title))

    # ── 스케일 비율 계산 ──
    def _made_ratio(family: str) -> float:
        return model.made_pct.get(family, 100.0) / 100.0

    # 전장/제어는 시트에 통합되어 있음 → 금액 가중 평균 비율
    _tz_made = model.mat_made.get("전장부", 0.0) + model.mat_made.get("제어부", 0.0)
    if _tz_made > 0:
        _tz_ratio = (
            model.mat_made.get("전장부", 0.0) * model.made_pct.get("전장부", 100.0)
            + model.mat_made.get("제어부", 0.0) * model.made_pct.get("제어부", 100.0)
        ) / _tz_made / 100.0
    else:
        _tz_ratio = 1.0

    _expense_ratio = model.expense_pct / 100.0

    _labor_ratio: Dict[str, float] = {}
    for sec in SECTION_ORDER:
        cur = model.section_sum(sec, "cur")
        adj = model.section_sum(sec, "adj")
        _labor_ratio[sec] = (adj / cur) if cur > 0 else 1.0

    # ── 시트 스캔: 구분/단가유형/인건비 섹션 추적 후 금액 스케일 ──
    cur_gubun = ""
    cur_dangayuh = ""
    cur_labor_sec = None

    for r in range(1, new.max_row + 1):
        a = new.cell(r, 1).value
        b = new.cell(r, 2).value
        c = new.cell(r, 3).value
        cs = str(c).strip() if c is not None else ""

        if a is not None and str(a).strip():
            cur_gubun = str(a).strip()
        if b is not None and str(b).strip():
            cur_dangayuh = str(b).strip()
        if cur_gubun == "인건비":
            for sec in SECTION_ORDER:
                if sec in cs:
                    cur_labor_sec = sec

        # 소계/합계/TOTAL 행은 수식이므로 건너뜀 (자동 재계산)
        if any(k in cs for k in ("소계", "합계", "Sub합계", "TOTAL")):
            continue
        # 섹션 헤더 행 (C열 '(1) 설계' 등): G=1 플레이스홀더 → 스케일링 금지
        if cs.startswith("("):
            continue

        # D열 라벨 (Engineering 등 특정 행 식별용)
        d_val = new.cell(r, 4).value
        d_str = str(d_val).strip() if d_val is not None else ""

        # 0) Engineering (PRI Engineering Fee): 자체 섹션 비율로 MD(G열) 스케일
        if d_str == "Engineering":
            ratio = _labor_ratio.get("(0) PRI Engineering", 1.0)
            if ratio != 1.0:
                g = new.cell(r, 7).value
                if isinstance(g, (int, float)):
                    new.cell(r, 7).value = g * ratio
            continue
        # 1) 재료비·가공품: 단가(I열) 스케일 (구매품은 변경 불가)
        if cur_dangayuh == "가공품" and cur_gubun in ("기구부", "전장/제어"):
            ratio = _made_ratio("기구부") if cur_gubun == "기구부" else _tz_ratio
            if ratio != 1.0:
                i = new.cell(r, 9).value
                if isinstance(i, (int, float)):
                    new.cell(r, 9).value = i * ratio
        # 2) 인건비: MD(G열) 스케일
        elif cur_gubun == "인건비" and cur_labor_sec:
            ratio = _labor_ratio.get(cur_labor_sec, 1.0)
            if ratio != 1.0:
                g = new.cell(r, 7).value
                if isinstance(g, (int, float)):
                    new.cell(r, 7).value = g * ratio
        # 3) 경비: 단가(I열) 스케일
        elif cur_gubun == "경비":
            if _expense_ratio != 1.0:
                i = new.cell(r, 9).value
                if isinstance(i, (int, float)):
                    new.cell(r, 9).value = i * _expense_ratio

    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()


def compute_excel_total(xlsx_bytes: bytes, sheet: str = "2.상세 견적") -> float:
    """
    생성된 표준 양식 Excel에서 'TOTAL합계'를 직접 계산한다.

    Excel의 총액 구조를 그대로 재현:
      - J2111 = 모든 상세행의 G(수량)×I(단가) 합계
        (소계/합계/Sub합계/TOTAL/합산 행, 섹션 헤더 행 제외)
      - J2114 = J2111 × 1.10  (5% 일반관리 + 5% 기업이윤)

    UI의 '총 견적 금액'이 이 함수를 사용해야 Excel과 100% 일치한다.
    (model.grand_total()은 별도 코드 경로로 계산해 Excel과 어긋날 수 있음)

    Args:
        xlsx_bytes: export_standard_excel() 또는 export_adjusted_excel() 결과 bytes
        sheet: 계산할 시트명 (기본 '2.상세 견적')

    Returns:
        TOTAL합계 (원). 시트가 없으면 0.0
    """
    def _num(v) -> float:
        if v is None:
            return 0.0
        if isinstance(v, (int, float)):
            return float(v)
        try:
            return float(str(v).replace(",", "").replace("원", "").strip())
        except (ValueError, TypeError):
            return 0.0

    try:
        wb = load_workbook(BytesIO(xlsx_bytes), data_only=False)
    except Exception:
        return 0.0
    if sheet not in wb.sheetnames:
        return 0.0
    ws = wb[sheet]

    base = 0.0
    for r in range(1, ws.max_row + 1):
        c = ws.cell(r, 3).value
        cs = str(c) if c is not None else ""
        # 소계/합계/Sub합계/TOTAL/합산 행 제외 (수식 행)
        if any(k in cs for k in ("소계", "합계", "Sub합계", "TOTAL", "합산")):
            continue
        # 섹션 헤더 행 (C열 '(1) 설계' 등): G=1 플레이스홀더 → 제외
        if cs.startswith("("):
            continue
        g = _num(ws.cell(r, 7).value)
        i = _num(ws.cell(r, 9).value)
        if g > 0 and i > 0:
            base += g * i
    return base * 1.10
