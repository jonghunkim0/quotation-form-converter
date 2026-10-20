"""
견적 표준 양식 변환기
업체별 견적서(PDF/Excel)를 표준 양식 Excel로 변환하는 Streamlit 웹앱
"""

import streamlit as st
import pandas as pd
import json
from io import BytesIO
from pathlib import Path
from typing import List, Dict, Any

from config import (
    STANDARD_COLUMNS,
    COLUMN_MAPPING_RULES,
    RULES_FILE,
    DEFAULT_TEMPLATE_FILE,
    CONFIDENCE_THRESHOLD,
)
from parsers import parse_pdf, parse_excel
from converters import RuleMapper, export_standard_excel


def _df_for_display(df: pd.DataFrame) -> pd.DataFrame:
    """표시용 DataFrame: 모든 셀을 실제 문자열로 통일.

    숫자/문자/빈값이 섞인 컬럼(Maker, 수량 등)은 pyarrow 직렬화 시
    'Expected bytes, got int' / 'Could not convert' 오류를 유발한다.
    astype(object)만으로는 int가 int로 남아 혼합 타입이 되므로,
    모든 셀을 str()로 변환해 균일한 string 컬럼을 만든다.
    NaN/None은 빈 문자열로 변환.
    """
    def _to_str(v):
        if v is None:
            return ""
        try:
            if pd.isna(v):
                return ""
        except (TypeError, ValueError):
            pass
        return str(v)

    return df.map(_to_str)


def _extract_auto_info(mapped_items: List[Dict[str, Any]]) -> Dict[str, str]:
    """매핑된 항목에서 업체명, 견적일, 통화, 특기사항 자동 추출"""
    info = {
        "업체명": "",
        "견적일": "",
        "견적번호": "",
        "통화": "",
        "특기사항": "",
    }
    
    if not mapped_items:
        return info
    
    # 업체명, 견적일 추출 (첫 항목에서)
    first = mapped_items[0]
    info["업체명"] = str(first.get("업체명", ""))
    info["견적일"] = str(first.get("견적일", ""))
    
    # 통화 추출 (단가/금액에서 통화 기호 탐색)
    currency_symbols = {"₩": "KRW", "$": "USD", "€": "EUR", "¥": "JPY", "£": "GBP"}
    for item in mapped_items[:10]:
        for col in ("단가", "금액"):
            val = str(item.get(col, ""))
            for symbol, code in currency_symbols.items():
                if symbol in val:
                    info["통화"] = code
                    break
        if info["통화"]:
            break
    
    # 특기사항 (비고) 요약
    remarks = [
        str(item.get("비고", "")) 
        for item in mapped_items 
        if item.get("비고")
    ]
    if remarks:
        info["특기사항"] = f"{len(remarks)}개 항목에 비고 있음"
    
    return info


# ─── analyzer 분류 결과 적용 ───────────────────────────────────
# analyzer 카테고리 → 템플릿 섹션 키 매핑
_ANALYZER_TO_SECTION = {
    "기구 구매품": "기구_구매품",
    "기구 가공품": "기구_가공품",
    "제어 구매품": "제어_구매품",
    "전장 구매품": "전장_구매품",
    "전장 가공품": "전장_가공품",
    "인건비-기구 설계": "설계",
    "인건비-전장 설계": "설계",
    "인건비-제어 설계": "설계",
    "인건비-기구 조립": "조립",
    "인건비-전장/제어 조립": "조립",
    "인건비-기구 셋업": "셋업",
    "인건비-전장/제어 셋업": "셋업",
    "인건비-기구 양산대응": "양산대응",
    "인건비-전장/제어 양산대응": "양산대응",
    "인건비-특수제어": "특수제어",
    "경비-포장": "포장",
    "경비-기타": "기타항목",
    "경비-숙박/교통/항공": "경비",
    "사급품": "기구_구매품",
}


def _collect_labor_data(analyzer_qf) -> dict:
    """
    analyzer의 인건비 항목에서 (MD, 일액, 비고)를 카테고리별로 수집.
    
    Returns:
        {analyzer_category: {"md": float, "rate": float, "amount": float, "remark": str}}
        같은 카테고리가 여러 행이면 MD/금액 합산, 일액은 첫 행 값, 비고는 " / "으로 결합.
    """
    if analyzer_qf is None or not analyzer_qf.items:
        return {}
    labor: dict = {}
    for a_item in analyzer_qf.items:
        cat = a_item.final_category or a_item.auto_category
        if not cat.startswith("인건비-"):
            continue
        md = a_item.quantity or 0.0
        rate = a_item.unit_price or 0.0
        amt = a_item.amount or 0.0
        if cat in labor:
            labor[cat]["md"] += md
            labor[cat]["amount"] += amt
            if a_item.remark:
                labor[cat]["remark"] = (labor[cat]["remark"] + " / " + a_item.remark).strip(" /")
        else:
            labor[cat] = {"md": md, "rate": rate, "amount": amt, "remark": a_item.remark or ""}
    return labor


def _apply_analyzer_classification(mapped_items, analyzer_qf):
    """
    원가분석 Tool(analyzer)의 분류 결과를 mapped_items에 적용.
    
    analyzer의 source_row(1-based)와 converter의 _행(0-based)을
    source_row == _행 + 1 로 매칭.
    """
    if analyzer_qf is None or not analyzer_qf.items:
        return
    
    # (source_sheet, source_row) → analyzer item 인덱스
    a_by_key = {}
    for a_item in analyzer_qf.items:
        a_by_key[(a_item.source_sheet, a_item.source_row)] = a_item
    
    # analyzer가 읽은 시트 (첫 시트)
    a_sheets = {k[0] for k in a_by_key}
    
    for m in mapped_items:
        row_1based = m.get("_행", 0) + 1  # 0-based → 1-based
        sheet = m.get("_출처", "")
        # analyzer가 읽지 않은 시트의 항목은 행 번호 매칭 불가 → 건너뜀
        if a_sheets and sheet not in a_sheets:
            continue
        a_item = a_by_key.get((sheet, row_1based))
        if a_item is None:
            continue
        
        # 합산(Unit) 행 감지: converter의 _합산만 사용
        # (analyzer의 is_sum_row는 소계/합계 행도 포함해서 Unit 행으로 오인됨)
        if m.get("_합산"):
            # Unit명: converter가 이미 추출한 값 우선, 없으면 analyzer에서
            if not m.get("_unit_name"):
                if a_item.name:
                    m["_unit_name"] = a_item.name
        
        # 분류 결과 적용
        cat = a_item.final_category or a_item.auto_category
        if cat == "합산금액":
            # 합산 행: 하위 항목 분류를 상속 (classify_items에서 처리)
            continue
        section_key = _ANALYZER_TO_SECTION.get(cat)
        if section_key:
            m["_분류"] = section_key
            m["_분류_신뢰도"] = a_item.confidence
            m["_분류_이유"] = f"analyzer: {cat} → {section_key}"


def _adjust_family_by_gubun(items):
    """
    Raw 견적의 '구분' 컬럼(계열 정보)을 활용해 계열 재조정.
    
    analyzer는 비고의 '- 사급' 표기만 보고 사급품으로 분류하므로
    기구/전장/제어 계열을 구분하지 못함. Raw 견적의 구분 컬럼에
    '전장/제어'로 명시된 항목은 전장/제어 계열로 재조정한다.
    
    대상: analyzer가 사급품/기구 계열로 분류한 항목 중
          Raw 구분이 전장/제어 계열인 경우.
    """
    # 전장/제어 계열 키워드 (품명 기반 판별용)
    CONTROL_KEYWORDS = [
        "plc", "inverter", "인버터", "touch panel", "터치패널", "터치 패널",
        "sensor", "센서", "photo", "포토", "proximity", "프록시",
        "safety", "안전", "hmi", "panel", "패널", "relay", "릴레이",
        "contactor", "콘택터", "timer", "타이머", "counter", "카운터",
    ]
    WIRING_KEYWORDS = [
        "motor", "모터", "cable", "케이블", "harness", "하네스",
        "breaker", "브레이커", "fuse", "퓨즈", "terminal", "터미널",
        "connector", "커넥터", "wire", "와이어", "lamp", "램프",
        "led", "beacon", "비콘", "buzzer", "버저", "horn", "혼",
    ]
    # 강한 제어 키워드: Raw 구분이 전장/제어가 아니어도(예: "기구부")
    # 제어_구매품으로 재조정할 만큼 제어부품이 명확한 경우만 사용.
    # (sensor/panel 등 약한 키워드는 기구부에도 존재 → 오분류 방지)
    STRONG_CONTROL_KEYWORDS = [
        "plc", "inverter", "인버터", "touch panel", "터치패널", "터치 패널",
        "hmi", "safety relay", "안전릴레이", "safety plc",
    ]
    
    for item in items:
        # 현재 분류가 기구 계열(또는 사급품 폴백)인지 확인
        cur = item.get("_분류", "")
        if cur not in ("기구_구매품", "기구_가공품", "미분류"):
            continue  # 이미 전장/제어 계열이면 그대로
        
        gubun = str(item.get("_구분", "")).replace(" ", "")
        is_elec = ("전장" in gubun) or ("제어" in gubun)
        
        # 품명 키워드로 전장 vs 제어 판별
        name = str(item.get("품명", "")).lower()
        spec = str(item.get("규격", "")).lower()
        text = name + " " + spec
        
        is_control = any(kw in text for kw in CONTROL_KEYWORDS)
        is_wiring = any(kw in text for kw in WIRING_KEYWORDS)
        
        if is_elec:
            # Raw 구분이 전장/제어 → 키워드 기반 판별
            if is_control and not is_wiring:
                new_sec = "제어_구매품"
            elif is_wiring and not is_control:
                new_sec = "전장_구매품"
            else:
                # 키워드 판별 불가 → Raw 구분 텍스트 기준
                if "제어" in gubun and "전장" not in gubun:
                    new_sec = "제어_구매품"
                elif "전장" in gubun and "제어" not in gubun:
                    new_sec = "전장_구매품"
                else:
                    new_sec = "제어_구매품"
        elif any(kw in text for kw in STRONG_CONTROL_KEYWORDS):
            # Raw 구분이 전장/제어가 아니어도, 품명에 강한 제어 키워드 → 제어_구매품
            # (Raw 견적의 구분 컬럼이 "기구부"로 되어 있어도 PLC/Inverter 등은 제어부품)
            new_sec = "제어_구매품"
        else:
            continue  # 키워드도, 구분도 판별 불가 → 그대로
        
        item["_분류"] = new_sec
        item["_분류_이유"] = f"계열 재조정: 품명 '{name}' → {new_sec}"


def _reinherit_unit_classification(items):
    """
    analyzer 분류 적용 후, 합산(Unit) 행의 분류를 하위 항목으로 재상속.
    (classify_items에서 이미 1차 상속했지만, analyzer가 세부항목 분류를
    변경했을 수 있으므로 재실행)
    """
    from collections import Counter
    n = len(items)
    for i, item in enumerate(items):
        if not item.get("_합산"):
            continue
        # analyzer가 직접 분류한 Unit(예: "경비-포장" → 포장)은 재상속으로 덮어쓰지 않음
        # (하위 스캔이 다음 섹션 항목까지 닿아 분류가 오염되는 것을 방지)
        if str(item.get("_분류_이유", "")).startswith("analyzer:"):
            continue
        sub_sections = []
        for j in range(i + 1, n):
            if items[j].get("_합산"):
                break
            sec = items[j].get("_분류", "")
            if sec and sec != "미분류":
                sub_sections.append(sec)
        if sub_sections:
            most_common = Counter(sub_sections).most_common(1)[0][0]
            item["_분류"] = most_common
            item["_분류_이유"] = f"Unit 합산 행 → 하위 항목 분류 상속: {most_common}"


# ─── 페이지 설정 ───────────────────────────────────────────────
st.set_page_config(
    page_title="견적 표준 양식 변환기",
    page_icon="📋",
    layout="wide",
)

# ─── 메인 헤더 ───────────────────────────────────────────────
APP_VERSION = "0.2"
APP_UPDATED = "2026. 10. 19."
st.markdown(
    f'<div style="display:flex;align-items:baseline;gap:10px;">'
    f'<h1 style="margin:0;">📋 견적 표준 양식 변환기</h1>'
    f'<span style="font-size:0.75rem;color:#888;">'
    f'ver {APP_VERSION} · 업데이트 {APP_UPDATED}</span>'
    f'</div>',
    unsafe_allow_html=True,
)
st.caption("업체별 견적서(PDF/Excel)를 표준 양식 Excel로 변환합니다.")

# ─── 전역 컴팩트 스타일 (글씨/줄간격 ~30% 축소) ────────────────
st.markdown(
    """
    <style>
    /* 기준 글씨 크기 축소 (rem 단위 전체가 이 값에 비례) */
    html { font-size: 13px !important; }
    /* 본문/캡션 줄간격 축소 */
    div[data-testid="stMarkdownContainer"] p,
    div[data-testid="stCaptionContainer"] { line-height: 1.4 !important; }
    /* 제목 간격 축소 */
    div[data-testid="stHeader"] { margin-top: 0.4rem !important; margin-bottom: 0.2rem !important; }
    div[data-testid="stSubheader"] { margin-top: 0.3rem !important; margin-bottom: 0.15rem !important; }
    /* markdown 컨테이너 기본 하단 여백 축소 */
    div[data-testid="stMarkdownContainer"] { margin-bottom: 0.3rem !important; }
    /* dataframe 글씨/여백 축소 */
    div[data-testid="stDataFrame"] { font-size: 0.78rem !important; }
    div[data-testid="stDataFrame"] .stDataFrameResizable { font-size: 0.78rem !important; }
    /* 버튼 패딩 축소 */
    div[data-testid="stButton"] button { padding: 0.35rem 0.8rem !important; font-size: 0.85rem !important; }
    </style>
    """,
    unsafe_allow_html=True,
)

# ─── 사이드바: 설정 ────────────────────────────────────────────
st.sidebar.header("⚙️ 설정")

# 원본명 표시 옵션
show_original = st.sidebar.checkbox(
    "원본 컬럼명 표시",
    value=False,
    help="표준 양식에 원본 컬럼명을 추가 열로 표시합니다.",
)

# 표준 양식 템플릿 (기본: 내장 템플릿 사용)
from config import DEFAULT_TEMPLATE_FILE
template_file_path = None

if DEFAULT_TEMPLATE_FILE.exists():
    st.sidebar.caption(f"📄 템플릿: `{DEFAULT_TEMPLATE_FILE.name}`")
    
    # 템플릿 교체 옵션 (기본: 사용 안 함)
    with st.sidebar.expander("🔄 템플릿 교체 (고급)", expanded=False):
        st.caption("기본 템플릿을 다른 파일로 교체합니다.")
        template_upload = st.file_uploader(
            "새 템플릿 Excel 업로드",
            type=["xlsx"],
            help="업로드한 파일이 기본 템플릿을 대체합니다.",
        )
        if template_upload is not None:
            template_file_path = "data/custom_template.xlsx"
            Path("data").mkdir(exist_ok=True)
            with open(template_file_path, "wb") as f:
                f.write(template_upload.getvalue())
            st.sidebar.success(f"템플릿 교체: `{template_upload.name}`")
else:
    st.sidebar.warning("기본 템플릿이 없습니다. 파일을 지정하세요.")
    template_upload = st.sidebar.file_uploader(
        "표준 양식 템플릿 Excel 업로드",
        type=["xlsx"],
    )
    if template_upload is not None:
        template_file_path = "data/custom_template.xlsx"
        Path("data").mkdir(exist_ok=True)
        with open(template_file_path, "wb") as f:
            f.write(template_upload.getvalue())

# ─── 사이드바: 매핑 규칙 관리 ──────────────────────────────────
st.sidebar.header("🔗 매핑 규칙")

if RULES_FILE.exists():
    st.sidebar.caption(f"규칙 파일: `{RULES_FILE.name}`")
else:
    st.sidebar.caption("규칙 파일이 없습니다. 아래에서 추가하세요.")

# 매핑 규칙 추가
with st.sidebar.expander("➕ 매핑 규칙 추가", expanded=False):
    rule_source = st.text_input(
        "원본 컬럼명 (업체 견적서에 있는 이름)",
        placeholder="예: Part No",
    )
    rule_target = st.selectbox(
        "표준 컬럼",
        STANDARD_COLUMNS,
    )
    rule_keyword = st.text_input(
        "키워드 (부분 일치 시 적용)",
        placeholder="예: part",
        help="원본 컬럼명에 이 키워드가 포함되면 해당 표준 컬럼으로 매핑",
    )

    if st.button("규칙 저장"):
        if not rule_source and not rule_keyword:
            st.sidebar.error("원본 컬럼명 또는 키워드를 입력하세요.")
        else:
            rules = {}
            if RULES_FILE.exists():
                rules = json.loads(RULES_FILE.read_text(encoding="utf-8"))
            if rule_source:
                rules[rule_source] = rule_target
            if rule_keyword:
                rules[f"~{rule_keyword}"] = rule_target  # ~ 접두사 = 키워드 매칭
            RULES_FILE.parent.mkdir(parents=True, exist_ok=True)
            RULES_FILE.write_text(json.dumps(rules, ensure_ascii=False, indent=2), encoding="utf-8")
            st.sidebar.success("규칙 저장 완료!")
            st.sidebar.rerun()

# 현재 규칙 목록 표시
if RULES_FILE.exists():
    rules = json.loads(RULES_FILE.read_text(encoding="utf-8"))
    if rules:
        st.sidebar.subheader("등록된 규칙")
        for src, tgt in rules.items():
            prefix = "🔍 " if src.startswith("~") else "📌 "
            st.sidebar.text(f"{prefix}{src} → {tgt}")
    else:
        st.sidebar.info("등록된 규칙이 없습니다.")

# ─── 메인: 파일 업로드 ─────────────────────────────────────────
st.header("1️⃣ 견적서 업로드")

uploaded_file = st.file_uploader(
    "견적서 파일 업로드 (PDF, Excel)",
    type=["pdf", "xlsx", "xls"],
    help="업체별 견적서 1개를 업로드하세요.",
)

parsed_items = None  # 파싱된 항목 목록
mapped_items = None  # 매핑된 항목
output_bytes = None  # 생성된 Excel bytes
auto_info = {}  # 자동 추출 정보
labor_data = None   # 인건비 데이터 (analyzer)

# ── 원가분석툴(analyzer) 파이프라인 ──
analyzer_qf = None       # QuotationFile (analyzer 파싱 결과)
analyzer_scenario = None  # Scenario (집계 결과)
analyzer_issues = []      # ValidationIssue 목록

if uploaded_file is not None:
    st.success(f"업로드됨: `{uploaded_file.name}`")

    # ── 파이프라인 캐시: 파일/옵션 변경 시에만 재실행 (조정 입력 시 재실행 방지) ──
    _pipe_key = f"{uploaded_file.name}|{uploaded_file.size}|{show_original}|{template_file_path}"
    _pc = st.session_state.get("_pipe_cache")
    if _pc and _pc.get("key") == _pipe_key:
        (parsed_items, mapped_items, output_bytes, auto_info,
         analyzer_qf, analyzer_scenario, analyzer_issues, labor_data) = _pc["result"]
    else:
        # 진행 상황 표시 (가로 막대)
        progress = st.progress(0, text="시작...")

        try:
            # Step 1: 파싱
            progress.progress(10, text="📄 파싱 중...")
            if uploaded_file.name.lower().endswith(".pdf"):
                parsed_items = parse_pdf(uploaded_file)
            else:
                parsed_items = parse_excel(uploaded_file)

            # Step 2: 매핑
            progress.progress(30, text="🔗 매핑 중...")
            rules = {}
            if RULES_FILE.exists():
                rules = json.loads(RULES_FILE.read_text(encoding="utf-8"))
            mapper = RuleMapper(rules)
            mapped_items = mapper.map_items(parsed_items)
        
            # Step 3: 원가분석툴(analyzer) 파이프라인 (분류 결과 확보)
            progress.progress(50, text="🔍 원가분석 중...")
            try:
                import tempfile, os
                from quotation_analyzer import ExcelParser, AutoClassifier, Aggregator, Validator
                from quotation_analyzer.models.data_models import QuotationFile

                # 업로드 파일을 임시 파일로 저장 (analyzer는 파일 경로 기반)
                tmp_analyzer = os.path.join(tempfile.gettempdir(), f"_analyzer_{uploaded_file.name}")
                with open(tmp_analyzer, "wb") as f:
                    f.write(uploaded_file.getvalue())

                a_parser = ExcelParser()
                sheets = a_parser.list_sheets(tmp_analyzer)
                selected_sheet = sheets[0] if sheets else ""
                analyzer_qf = a_parser.parse_sheet(tmp_analyzer, selected_sheet, uploaded_file.name, "견적가")

                if analyzer_qf.status.value != "오류":
                    a_classifier = AutoClassifier()
                    a_classifier.classify_all(analyzer_qf.items)
                    for it in analyzer_qf.items:
                        it.final_category = it.auto_category

                    a_agg = Aggregator()
                    analyzer_scenario = a_agg.aggregate_scenario(
                        name=Path(uploaded_file.name).stem,
                        items=analyzer_qf.items,
                    )

                    a_validator = Validator()
                    analyzer_issues = a_validator.validate_file(analyzer_qf)

                try:
                    os.unlink(tmp_analyzer)
                except OSError:
                    pass

            except Exception as ae:
                err_msg = str(ae)
                st.error(f"원가분석 실행 중 오류: {err_msg}")
                with st.expander("🔧 해결 방법", expanded=True):
                    if "No module named 'quotation_analyzer'" in err_msg:
                        st.markdown("""
**`quotation_analyzer` 패키지가 설치되어 있지 않거나 서버가 오래된 상태로 실행 중입니다.**

해결 순서:
1. 터미널에서 아래 명령을 실행하세요:
```
cd C:\\AI_Agent\\quotation-form-converter
.venv\\Scripts\\python.exe -m pip install -e C:\\AI_Agent\\quotation-analyzer
```
2. **Streamlit 서버를 재시작**하세요 (Ctrl+C 후 다시 실행):
```
.venv\\Scripts\\python.exe -m streamlit run app.py --server.port 8501
```
3. 이 페이지를 새로고침 후 파일을 다시 업로드하세요.
""")
                    elif "No module named 'plotly'" in err_msg:
                        st.markdown("""
**`plotly` 패키지가 설치되어 있지 않습니다.**

해결:
```
cd C:\\AI_Agent\\quotation-form-converter
.venv\\Scripts\\python.exe -m pip install plotly
```
그 후 Streamlit 서버를 재시작하세요.
""")
                    else:
                        st.markdown(f"""
**오류 내용:** `{err_msg}`

해결 순서:
1. `quotation_analyzer` 패키지 설치 확인:
```
cd C:\\AI_Agent\\quotation-form-converter
.venv\\Scripts\\python.exe -c "import quotation_analyzer; print(quotation_analyzer.__version__)"
```
2. 패키지가 없으면 설치:
```
.venv\\Scripts\\python.exe -m pip install -e C:\\AI_Agent\\quotation-analyzer
```
3. Streamlit 서버 재시작:
```
.venv\\Scripts\\python.exe -m streamlit run app.py --server.port 8501
```
4. 여전히 오류가 나면 터미널의 전체 오류 메시지를 확인하세요.
""")
                analyzer_qf = None
                analyzer_scenario = None
                analyzer_issues = []

            # Step 4: 분류 (템플릿 섹션별) + analyzer 분류 결과 적용
            progress.progress(70, text="📂 분류 중...")
            from classification import classify_items
            mapped_items = classify_items(mapped_items)
            _apply_analyzer_classification(mapped_items, analyzer_qf)
            # Raw 구분 컬럼 기반 계열 재조정 (사급품 → 전장/제어 계열)
            _adjust_family_by_gubun(mapped_items)
            # analyzer 분류 적용 후 Unit 행 분류 재상속
            _reinherit_unit_classification(mapped_items)

            # Step 5: Excel 생성
            progress.progress(85, text="📊 표준 양식 생성 중...")
            auto_info = _extract_auto_info(mapped_items)
            labor_data = _collect_labor_data(analyzer_qf)
            output_bytes = export_standard_excel(
                mapped_items,
                company_name=auto_info.get("업체명", ""),
                quote_no=auto_info.get("견적번호", ""),
                template_file=template_file_path,
                show_original=show_original,
                labor_data=labor_data,
            )

            # 완료
            progress.progress(100, text="✅ 완료!")

        except Exception as e:
            progress.progress(100, text="❌ 오류 발생")
            st.error(f"변환 실패: {e}")
            parsed_items = None
            mapped_items = None
            output_bytes = None
            analyzer_qf = None
            analyzer_scenario = None
            analyzer_issues = []
            labor_data = None

        st.session_state._pipe_cache = {
            "key": _pipe_key,
            "result": (parsed_items, mapped_items, output_bytes, auto_info,
                       analyzer_qf, analyzer_scenario, analyzer_issues, labor_data),
        }

    if parsed_items:
        with st.expander(f"🔍 파싱 결과 (원본) 보기 — {len(parsed_items)}개 항목", expanded=False):
            raw_rows = []
            for item in parsed_items:
                row = item["raw_data"].copy()
                row["_시트/페이지"] = item.get("sheet_name") or f"p.{item.get('page', '?')}"
                raw_rows.append(row)
            df_raw = pd.DataFrame(raw_rows)
            st.dataframe(_df_for_display(df_raw), width="stretch")
            st.caption(f"총 {len(parsed_items)}개 항목 추출")

# ─── 메인: 변환 결과 ───────────────────────────────────────────
st.header("2️⃣ 표준 양식으로 변환")

if parsed_items and mapped_items:

    # ─── 자동 추출 정보 표시 ───────────────────────────────────
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("업체명", auto_info.get("업체명", "-"))
    col2.metric("견적일", auto_info.get("견적일", "-"))
    col3.metric("통화", auto_info.get("통화", "-"))
    col4.metric("특기사항", auto_info.get("특기사항", "-"))

    # ─── 매핑 결과 표시 (접힘) ─────────────────────────────────
    with st.expander(f"🔗 매핑 결과 보기 — {len(mapped_items)}개 항목", expanded=False):
        # 표시용 DataFrame (내부 필드 제외)
        display_cols = list(STANDARD_COLUMNS)
        if show_original:
            display_cols.append("_원본명")
        display_cols.append("_신뢰도")

        df_mapped = pd.DataFrame(mapped_items)[display_cols]
        df_mapped.columns = [
            "원본 컬럼명" if c == "_원본명" else 
            "신뢰도" if c == "_신뢰도" else c 
            for c in display_cols
        ]

        st.dataframe(_df_for_display(df_mapped), width="stretch")

        # 신뢰도 80% 미만 항목 강조
        low_confidence = [
            item for item in mapped_items 
            if item.get("_신뢰도", 1.0) < CONFIDENCE_THRESHOLD
        ]
        if low_confidence:
            st.warning(
                f"⚠️ 신뢰도 80% 미만 항목 {len(low_confidence)}개 - 사용자 확인 필요"
            )
            df_low = pd.DataFrame(low_confidence)[display_cols]
            df_low.columns = [
                "원본 컬럼명" if c == "_원본명" else 
                "신뢰도" if c == "_신뢰도" else c 
                for c in display_cols
            ]
            st.dataframe(_df_for_display(df_low), width="stretch")

    # 검증 결과 요약
    st.subheader("검증 요약")
    error_items = [i for i in mapped_items if "오류" in str(i.get("검증결과", ""))]
    missing_items = [i for i in mapped_items if "확인 필요" in str(i.get("검증결과", ""))]
    dup_items = [i for i in mapped_items if "중복" in str(i.get("검증결과", ""))]
    
    vcol1, vcol2, vcol3 = st.columns(3)
    vcol1.metric("💰 금액 검증 오류", len(error_items))
    vcol2.metric("❓ 확인 필요", len(missing_items))
    vcol3.metric("🔄 중복", len(dup_items))

    # ─── 다운로드 ──────────────────────────────────────────────
    st.header("3️⃣ 표준 양식 다운로드")

    if output_bytes:
        original_name = Path(uploaded_file.name).stem
        st.download_button(
            label="⬇️ 표준 양식 Excel 다운로드",
            data=output_bytes,
            file_name=f"표준견적_{original_name}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )

    # ─── 4. 견적 조정 (비율 조정) ──────────────────────────────
    st.header("4️⃣ 견적 조정")
    st.caption(
        "🟡 **가공품/경비**는 변경 비율(%)로 조정, **인건비**는 대응인원수/대응기간으로 조정. "
        "🩶 단가는 분석 결과 자동 반영. MD = 인원수 × 대응기간, 금액 = MD × 단가."
    )

    # UI: 모든 셀을 동일한 고정 높이로 맞춰 좌우 행을 정렬.
    # st.columns는 컬럼별 독립 흐름이라 number_input(높음)과 텍스트(낮음)의
    # 높이 차이가 누적되어 행이 어긋난다 → 모든 셀을 _ROW_H 높이 flex 박스로 통일.
    _ROW_H = 26
    st.markdown(
        f"""
        <style>
        div[data-testid="stNumberInput"] {{
            height: {_ROW_H}px !important;
            margin: 0 !important;
        }}
        div[data-testid="stNumberInput"] input {{
            height: 20px !important;
            padding: 1px 6px !important;
            font-size: 0.75rem !important;
        }}
        /* markdown 셀 컨테이너의 기본 margin-bottom(-16px)을 0으로 리셋.
           이 음수 마진이 number_input(38px)과 텍스트 셀(22px)의 높이 차이를
           만들어 행이 어긋나게 한다 → adj-cell이 포함된 컨테이너만 리셋. */
        div[data-testid="stMarkdownContainer"]:has(> div.adj-cell) {{
            margin-bottom: 0 !important;
        }}
        </style>
        """,
        unsafe_allow_html=True,
    )

    def _cell(col, html: str, color: str = "#222"):
        """고정 높이(_ROW_H) flex 박스로 셀 렌더 → 컬럼 간 행 정렬."""
        col.markdown(
            f"<div class='adj-cell' style='height:{_ROW_H}px;display:flex;"
            f"align-items:center;font-size:0.75rem;color:{color};'>{html}</div>",
            unsafe_allow_html=True,
        )

    from converters.ratio_adjuster import (
        RatioModel, LaborRow, SECTION_ORDER, build_ratio_model, apply_adjustment,
    )
    from converters import export_adjusted_excel, compute_excel_total

    # 모델 생성 (파일 업로드 시 1회) — 새 파일이면 이전 조정 입력/결과도 리셋
    if "ratio_model" not in st.session_state or st.session_state.get("_ratio_file") != uploaded_file.name:
        st.session_state.ratio_model = build_ratio_model(labor_data, mapped_items)
        st.session_state._ratio_file = uploaded_file.name
        st.session_state.ratio_inputs = {}
        st.session_state._adj_bytes = None
        st.session_state._adj_calculated = False
        # number_input 위젯 키(ratio_*)도 삭제 → 이전 조정값이 위젯에 유지되는 것 방지
        # (ratio_model/ratio_inputs는 제외 — 방금 초기화한 값)
        for _k in [k for k in st.session_state.keys()
                   if k.startswith("ratio_") and k not in ("ratio_model", "ratio_inputs")]:
            del st.session_state[_k]

    model: RatioModel = st.session_state.ratio_model
    if "ratio_inputs" not in st.session_state:
        st.session_state.ratio_inputs = {}

    _fmt_won = lambda v: f"{v:,.0f}" if v is not None else "0"

    # ── 1) 상단 요약: 가공품/경비 비율(%) 입력 ──
    st.subheader("재료비·경비 조정 (변경 비율)")
    _fam_cols = st.columns([2, 1, 1, 1, 1])
    _cell(_fam_cols[0], "구분", "#888")
    _cell(_fam_cols[1], "현견적(원)", "#888")
    _cell(_fam_cols[2], "변경 비율(%)", "#b8860b")
    _cell(_fam_cols[3], "변경(원)", "#888")
    _cell(_fam_cols[4], "증감(원)", "#888")

    for fam in ("기구부", "전장부", "제어부"):
        cur = model.mat_made.get(fam, 0.0)
        key = f"made|{fam}|pct"
        pct = st.session_state.ratio_inputs.get(key, 100.0)
        _cell(_fam_cols[0], f"1. 가공품 · {fam}")
        _cell(_fam_cols[1], _fmt_won(cur))
        new_pct = _fam_cols[2].number_input(
            "변경 비율(%)", min_value=0.0, max_value=500.0, step=5.0, value=float(pct),
            key=f"ratio_{key}", label_visibility="collapsed",
        )
        st.session_state.ratio_inputs[key] = new_pct
        model.made_pct[fam] = new_pct
        adj = model.made_adj(fam)
        _cell(_fam_cols[3], _fmt_won(adj))
        _cell(_fam_cols[4], f"{adj - cur:+,.0f}", "#c0392b" if adj - cur > 0 else "#27ae60")

    # 구매품 (개별 변경 불가 — 표시만)
    for fam in ("기구부", "전장부", "제어부"):
        cur = model.mat_buy.get(fam, 0.0)
        if not cur:
            continue
        _cell(_fam_cols[0], f"2. 구매품 · {fam}")
        _cell(_fam_cols[1], _fmt_won(cur))
        _cell(_fam_cols[2], "개별 변경 불가", "#888")
        _cell(_fam_cols[3], _fmt_won(cur))
        _cell(_fam_cols[4], "0")

    # 경비 (비율 입력)
    _cell(_fam_cols[0], "4. 경비")
    _cell(_fam_cols[1], _fmt_won(model.expense))
    exp_pct = _fam_cols[2].number_input(
        "경비 변경 비율(%)", min_value=0.0, max_value=500.0, step=5.0,
        value=float(st.session_state.ratio_inputs.get("expense|pct", 100.0)),
        key="ratio_expense_pct", label_visibility="collapsed",
    )
    st.session_state.ratio_inputs["expense|pct"] = exp_pct
    model.expense_pct = exp_pct
    exp_adj = model.expense_adj()
    _cell(_fam_cols[3], _fmt_won(exp_adj))
    _cell(_fam_cols[4], f"{exp_adj - model.expense:+,.0f}",
          "#c0392b" if exp_adj - model.expense > 0 else "#27ae60")

    # ── 2) 인건비 조정: 대응인원수/대응기간 입력 ──
    st.subheader("인건비 조정 (대응인원수 · 대응기간)")
    if not model.labor_rows:
        st.info("인건비 항목이 없습니다.")
    else:
        # 헤더 행 (재료비·경비 조정과 동일 양식)
        _lh = st.columns([3, 1, 1, 1, 1, 1])
        _cell(_lh[0], "구분", "#888")
        _cell(_lh[1], "현견적(원)", "#888")
        _cell(_lh[2], "인원수", "#888")
        _cell(_lh[3], "대응일수", "#888")
        _cell(_lh[4], "변경(원)", "#888")
        _cell(_lh[5], "증감(원)", "#888")

        for sec in SECTION_ORDER:
            sec_rows = [r for r in model.labor_rows if r.section == sec]
            if not sec_rows:
                continue
            st.markdown(f"**{sec}**")
            for r in sec_rows:
                key_p = f"{r.section}|{r.job}|people"
                key_d = f"{r.section}|{r.job}|days"
                ic1, ic2, ic3, ic4, ic5, ic6 = st.columns([3, 1, 1, 1, 1, 1])
                _cell(ic1, f"{r.job} "
                    f"<span style='color:#888;font-size:0.68rem;'>단가 {_fmt_won(r.rate)} · "
                    f"현 MD {r.cur_md:,.0f}</span>")
                _cell(ic2, _fmt_won(r.cur_amount))
                new_p = ic3.number_input(
                    "대응인원수", min_value=0.0, step=1.0,
                    value=float(st.session_state.ratio_inputs.get(key_p, r.adj_people) or 0),
                    key=f"ratio_{key_p}", label_visibility="collapsed",
                )
                new_d = ic4.number_input(
                    "대응기간(일)", min_value=0.0, step=1.0,
                    value=float(st.session_state.ratio_inputs.get(key_d, r.adj_days) or 0),
                    key=f"ratio_{key_d}", label_visibility="collapsed",
                )
                adj_md = new_p * new_d
                adj_amt = adj_md * r.rate
                _cell(ic5, f"{adj_amt:,.0f}", "#0066cc")
                _cell(ic6, f"{adj_amt - r.cur_amount:+,.0f}",
                      "#c0392b" if adj_amt - r.cur_amount > 0 else "#27ae60")
                st.session_state.ratio_inputs[key_p] = new_p
                st.session_state.ratio_inputs[key_d] = new_d

    # ── 견적조정 산출 버튼 ──
    if st.button("📐 견적조정 산출", type="primary", use_container_width=True):
        apply_adjustment(model, st.session_state.ratio_inputs)
        try:
            st.session_state._adj_bytes = export_adjusted_excel(output_bytes, model)
            st.session_state._adj_calculated = True
        except Exception as _ae:
            st.session_state._adj_bytes = None
            st.session_state._adj_calculated = False
            st.warning(f"조정견적 생성 실패: {_ae}")

    # ── 산출 결과 (버튼 클릭 후) ──
    if st.session_state.get("_adj_calculated") and st.session_state.get("_adj_bytes"):
        # 변경 결과 요약
        res_rows = []
        # 재료비 (가공품)
        for fam in ("기구부", "전장부", "제어부"):
            cur = model.mat_made.get(fam, 0.0)
            adj = model.made_adj(fam)
            if cur or adj:
                res_rows.append((f"가공품 · {fam}", cur, adj, adj - cur))
        # 재료비 (구매품)
        for fam in ("기구부", "전장부", "제어부"):
            cur = model.mat_buy.get(fam, 0.0)
            if cur:
                res_rows.append((f"구매품 · {fam}", cur, cur, 0))
        # 경비
        if model.expense:
            exp_adj = model.expense_adj()
            res_rows.append(("경비", model.expense, exp_adj, exp_adj - model.expense))
        # 인건비
        for sec in SECTION_ORDER:
            cur = model.section_sum(sec, "cur")
            adj = model.section_sum(sec, "adj")
            if cur or adj:
                res_rows.append((sec, cur, adj, adj - cur))
        res_rows.append(("인건비 합계", model.labor_total("cur"), model.labor_total("adj"),
                         model.labor_total("adj") - model.labor_total("cur")))
        df_res = pd.DataFrame(res_rows, columns=["구분", "현견적(원)", "변경(원)", "증감(원)"])
        st.dataframe(df_res, width="stretch", hide_index=True)

        # 총 견적 금액 (Excel 파일에서 직접 계산 → TOTAL합계와 100% 일치)
        # 조정 견적은 '3.조정 견적' 시트에 있으므로 해당 시트를 읽어야 함
        st.subheader("총 견적 금액")
        cur_total = compute_excel_total(output_bytes)
        adj_total = compute_excel_total(st.session_state._adj_bytes, sheet="3.조정 견적")
        tc1, tc2, tc3 = st.columns(3)
        _delta = adj_total - cur_total
        tc1.markdown(
            f"<div style='font-size:0.7rem;color:#888;'>현견적 총액</div>"
            f"<div style='font-size:0.9rem;font-weight:600;'>{cur_total:,.0f}원</div>",
            unsafe_allow_html=True,
        )
        tc2.markdown(
            f"<div style='font-size:0.7rem;color:#888;'>변경 후 총액</div>"
            f"<div style='font-size:0.9rem;font-weight:600;'>{adj_total:,.0f}원</div>",
            unsafe_allow_html=True,
        )
        tc3.markdown(
            f"<div style='font-size:0.7rem;color:#888;'>증감</div>"
            f"<div style='font-size:0.9rem;font-weight:600;color:{'#c0392b' if _delta > 0 else '#27ae60'};'>{_delta:+,.0f}원</div>",
            unsafe_allow_html=True,
        )

        # 다운로드 버튼
        _adj_name = Path(uploaded_file.name).stem
        st.download_button(
            label="📤 조정견적 내보내기",
            data=st.session_state._adj_bytes,
            file_name=f"조정견적_{_adj_name}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )
    else:
        st.info("📐 '견적조정 산출' 버튼을 클릭하면 조정 결과가 계산됩니다.")

    # ─── 5. 분석 결과 내보내기 (원가분석툴과 동일 문서) ──────────
    st.header("5️⃣ 분석 결과 내보내기")

    if analyzer_qf is None or analyzer_scenario is None:
        st.warning("원가분석 결과가 없습니다. 내보내기를 사용할 수 없습니다.")
        st.caption(
            "🔧 해결: `pip install -e C:\\AI_Agent\\quotation-analyzer` 후 "
            "Streamlit 서버 재시작 → 파일 다시 업로드 (상단의 '🔧 해결 방법' 참고)"
        )
    else:
        from quotation_analyzer import Exporter
        from quotation_analyzer.config import COST_CATEGORIES

        exporter = Exporter(metadata={
            "원본 파일": uploaded_file.name,
            "시트": analyzer_qf.selected_sheet,
            "항목 수": len(analyzer_qf.items),
        })
        all_items = analyzer_qf.items
        scenarios = [analyzer_scenario]
        _quote_name = Path(uploaded_file.name).stem

        def _csv_to_xlsx(csv_content: str) -> bytes:
            """CSV 문자열 → Excel bytes"""
            import io as _io
            df = pd.read_csv(_io.StringIO(csv_content), encoding="utf-8-sig")
            buf = BytesIO()
            df.to_excel(buf, index=False, engine="openpyxl")
            return buf.getvalue()

        # ── 1. 견적 세부분류 (원본 파일 유지 + 끝단에 분류 열 추가) ──
        def _build_detail_classification() -> bytes:
            import openpyxl
            import tempfile as _tf
            import os as _os

            tmp_path = _os.path.join(_tf.gettempdir(), f"_orig_{uploaded_file.name}")
            with open(tmp_path, "wb") as f:
                f.write(uploaded_file.getvalue())

            try:
                wb = openpyxl.load_workbook(tmp_path)
            except Exception:
                df = pd.DataFrame([{
                    "품명": it.name, "수량": it.quantity,
                    "단가": it.unit_price, "금액": it.amount,
                    "자동분류": it.auto_category, "신뢰도(%)": it.confidence,
                    "최종분류": it.final_category, "분류근거": it.reason,
                } for it in all_items])
                buf = BytesIO()
                df.to_excel(buf, index=False, engine="openpyxl")
                return buf.getvalue()

            # (filename, sheet, source_row) → QuotationItem 매핑
            item_map = {}
            for it in all_items:
                item_map[(it.source_file, it.source_sheet, it.source_row)] = it

            new_headers = ["자동분류", "신뢰도(%)", "핵심키워드", "분류근거",
                           "최종분류", "검토상태", "수정여부"]

            for ws in wb.worksheets:
                from quotation_analyzer.parsers.excel_parser import ExcelParser as _EP
                _p = _EP()
                df_tmp = pd.read_excel(tmp_path, sheet_name=ws.title, header=None)
                header_idx = _p.detect_header_row(df_tmp)
                if header_idx < 0:
                    continue

                new_col_start = ws.max_column + 1
                for i, h in enumerate(new_headers):
                    ws.cell(row=header_idx + 1, column=new_col_start + i, value=h)

                for row_idx in range(header_idx + 1, len(df_tmp)):
                    source_row = row_idx + 1  # 1-based
                    key = (uploaded_file.name, ws.title, source_row)
                    if key in item_map:
                        it = item_map[key]
                        ws.cell(row=row_idx + 1, column=new_col_start + 0, value=it.auto_category)
                        ws.cell(row=row_idx + 1, column=new_col_start + 1, value=it.confidence)
                        ws.cell(row=row_idx + 1, column=new_col_start + 2, value=", ".join(it.keywords))
                        ws.cell(row=row_idx + 1, column=new_col_start + 3, value=it.reason)
                        ws.cell(row=row_idx + 1, column=new_col_start + 4, value=it.final_category)
                        ws.cell(row=row_idx + 1, column=new_col_start + 5, value=it.review_status.value)
                        ws.cell(row=row_idx + 1, column=new_col_start + 6, value="Y" if it.is_modified else "N")

            buf = BytesIO()
            wb.save(buf)
            wb.close()
            try:
                _os.unlink(tmp_path)
            except OSError:
                pass
            return buf.getvalue()

        # ── 다운로드 버튼 (원가분석툴 Exporter 직접 호출) ──
        exp_col1, exp_col2, exp_col3 = st.columns(3)
        with exp_col1:
            st.download_button(
                "📄 견적 세부분류",
                data=_build_detail_classification(),
                file_name="견적세부분류.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_detail",
            )
        with exp_col2:
            st.download_button(
                "📊 집계표",
                data=_csv_to_xlsx(exporter.export_summary_csv(scenarios)),
                file_name="집계표.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_summary",
            )
        with exp_col3:
            st.download_button(
                "⚠️ 검증오류",
                data=_csv_to_xlsx(exporter.export_validation_csv(analyzer_issues)),
                file_name="검증오류.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_validation",
            )

        exp_col4, exp_col5, exp_col6 = st.columns(3)
        with exp_col4:
            _full_buf = BytesIO()
            exporter.export_excel(_full_buf, all_items, scenarios, [], analyzer_issues)
            st.download_button(
                "📑 전체 Excel",
                data=_full_buf.getvalue(),
                file_name="원가분석_결과.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_full",
            )
        with exp_col5:
            st.download_button(
                "📊 원가구성 차트",
                data=exporter.export_chart_excel(scenarios, unit="백만원", decimals=1),
                file_name="원가구성비교차트.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_chart",
            )
        with exp_col6:
            if RULES_FILE.exists():
                st.download_button(
                    "📤 규칙 내보내기 (JSON)",
                    data=RULES_FILE.read_text(encoding="utf-8"),
                    file_name="classification_rules.json",
                    mime="application/json",
                    key="dl_rules",
                )
            else:
                st.caption("저장된 규칙이 없습니다.")
else:
    st.info("👆 먼저 견적서 파일을 업로드하세요.")


# ─── 푸터 ─────────────────────────────────────────────────────
st.divider()
st.caption(f"견적 표준 양식 변환기 v{APP_VERSION} | 로컬 처리, 외부 API 미사용")
