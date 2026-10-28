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
    f'<h1 style="margin:0;font-size:2.2rem;">📋 견적 표준 양식 변환기</h1>'
    f'<span style="font-size:0.9rem;color:#888;">'
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
    /* 큰 글씨 20% 축소: 섹션 제목(h2 1.5→1.2rem), 소제목(h3 1.25→1.0rem) */
    div[data-testid="stHeader"] { font-size: 1.2rem !important; margin-top: 0.4rem !important; margin-bottom: 0.2rem !important; }
    div[data-testid="stSubheader"] { font-size: 1.0rem !important; margin-top: 0.3rem !important; margin-bottom: 0.15rem !important; }
    /* 작은 글씨 20% 확대: 캡션 0.7→0.84rem */
    div[data-testid="stCaptionContainer"] { font-size: 0.84rem !important; }
    /* markdown 컨테이너 기본 하단 여백 축소 */
    div[data-testid="stMarkdownContainer"] { margin-bottom: 0.3rem !important; }
    /* dataframe 글씨 20% 확대 (0.78→0.9rem) */
    div[data-testid="stDataFrame"] { font-size: 0.9rem !important; }
    div[data-testid="stDataFrame"] .stDataFrameResizable { font-size: 0.9rem !important; }
    /* metric 값(업체명/특기사항, 검증 요약 수치)을 작은 글씨(0.84rem)와 동일하게 */
    div[data-testid="stMetricValue"] { font-size: 0.84rem !important; }
    /* 인건비 섹션 헤더: 상하 여백 70% 축소 */
    div[data-testid="stMarkdownContainer"]:has(> div.sec-hdr) { margin-top: 4px !important; margin-bottom: 2px !important; }
    div.sec-hdr { font-weight: 600; font-size: 0.82rem; line-height: 1.1; margin: 0 !important; padding: 0 !important; }
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

# ─── 사이드바: 💬 분석 챗봇 (Qwen3.8) ──────────────────────────
st.sidebar.header("💬 분석 챗봇")

# 챗봇 대화 히스토리 (session_state)
if "chat_history" not in st.session_state:
    st.session_state.chat_history = []  # [{"role": "user"|"assistant", "content": str}]
if "chat_log" not in st.session_state:
    st.session_state.chat_log = []  # 표시용 [(role, text)]

# 챗봇 CSS
st.sidebar.markdown(
    """
    <style>
    .chat-log {
        height: 220px; overflow-y: auto; background: #f9fafb;
        border: 1px solid #e5e7eb; border-radius: 6px;
        padding: 8px; margin-bottom: 6px;
    }
    .chat-msg {
        max-width: 92%; padding: 5px 9px; border-radius: 9px;
        font-size: 0.78rem; line-height: 1.45; margin-bottom: 5px;
        word-break: keep-all; white-space: pre-wrap;
    }
    .chat-msg.bot {
        background: #fff; border: 1px solid #e5e7eb;
        border-bottom-left-radius: 3px;
    }
    .chat-msg.user {
        background: #2563eb; color: #fff; margin-left: auto;
        border-bottom-right-radius: 3px;
    }
    .chat-hint { font-size: 0.7rem; color: #6b7280; margin-top: 4px; line-height: 1.4; }
    .chat-hint b { color: #374151; }
    </style>
    """,
    unsafe_allow_html=True,
)

# 대화 로그 표시
_chat_html_parts = []
for role, text in st.session_state.chat_log:
    cls = "user" if role == "user" else "bot"
    _chat_html_parts.append(f'<div class="chat-msg {cls}">{text}</div>')
_chat_html = "".join(_chat_html_parts) or '<div class="chat-msg bot">안녕하세요! 견적 분석을 도와드립니다.<br>예: <b>"가공품/구매품/인건비 비율 알려줘"</b></div>'

st.sidebar.markdown(
    f'<div class="chat-log" id="chatLog">{_chat_html}</div>',
    unsafe_allow_html=True,
)

# 입력框 (전송 후 초기화: _chat_input_next가 있으면 그 값으로 초기화)
if "_chat_input_next" in st.session_state:
    st.session_state.chat_input = st.session_state.pop("_chat_input_next")

_chat_col1, _chat_col2 = st.sidebar.columns([4, 1])
with _chat_col1:
    _chat_input = st.text_input(
        "요청 입력",
        key="chat_input",
        placeholder="예: 가공비 상위 10개 20% 올려줘",
        label_visibility="collapsed",
    )
with _chat_col2:
    _chat_send = st.button("전송", key="chat_send", use_container_width=True)

# 사용 예시
st.sidebar.markdown(
    """
    <div class="chat-hint">
    <b>사용 예시</b><br>
    · 가공품/구매품/인건비/기타 비율 알려줘<br>
    · 가공비 상위 10개 항목 20% 올려줘<br>
    · 미분류 항목 보여줘<br>
    · 인건비 총액 알려줘<br>
    · 경비 110%로 조정해줘
    </div>
    """,
    unsafe_allow_html=True,
)

# 전송 처리
if _chat_send and _chat_input.strip():
    _user_msg = _chat_input.strip()

    # 데이터가 없으면 안내 (mapped_items는 파일 업로드 후 정의됨)
    _pc_result = st.session_state.get("_pipe_cache", {}).get("result")
    if not _pc_result or _pc_result[1] is None:
        _bot_msg = "⚠️ 먼저 견적서 파일을 업로드하세요."
        st.session_state.chat_log.append(("user", _user_msg))
        st.session_state.chat_log.append(("bot", _bot_msg))
        st.session_state.chat_history.append({"role": "user", "content": _user_msg})
        st.session_state.chat_history.append({"role": "assistant", "content": _bot_msg})
        st.sidebar.error(_bot_msg)
    else:
        # LLM 호출 (파이프라인 캐시에서 데이터 참조)
        from chatbot import ask_chatbot, apply_actions

        _pipe_result = st.session_state.get("_pipe_cache", {}).get("result", (None,)*8)
        _mi = _pipe_result[1]   # mapped_items
        _ld = _pipe_result[7]   # labor_data

        with st.spinner("🤖 분석 중..."):
            result = ask_chatbot(
                user_query=_user_msg,
                mapped_items=_mi,
                labor_data=_ld,
                history=st.session_state.chat_history,
            )

        _bot_msg = result.get("message", "")
        st.session_state.chat_log.append(("user", _user_msg))
        st.session_state.chat_log.append(("bot", _bot_msg))
        st.session_state.chat_history.append({"role": "user", "content": _user_msg})
        st.session_state.chat_history.append({"role": "assistant", "content": _bot_msg})

        # 조정 액션 적용
        actions = result.get("actions", [])
        if actions and result.get("type") == "adjust":
            # ratio_model이 있는지 확인
            _rm = st.session_state.get("ratio_model")
            _ri = st.session_state.get("ratio_inputs", {})
            if _rm is not None:
                applied = apply_actions(actions, _rm, _ri, _mi)
                if applied:
                    _bot_msg += "\n\n" + "\n".join(applied)
                    # 마지막 bot 메시지 업데이트
                    st.session_state.chat_log[-1] = ("bot", _bot_msg)
                    st.session_state.chat_history[-1]["content"] = _bot_msg
                    st.session_state._adj_calculated = False  # 재산출 필요

        # 입력框 초기화 (rerun 후 새 값으로 초기화)
        st.session_state["_chat_input_next"] = ""
        st.rerun()

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
    _pipe_key = f"v2|{uploaded_file.name}|{uploaded_file.size}|{show_original}|{template_file_path}"
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

    # 검증 결과 요약 (접기/펼치기)
    with st.expander("검증 요약", expanded=False):
        error_items = [i for i in mapped_items if "오류" in str(i.get("검증결과", ""))]
        missing_items = [i for i in mapped_items if "확인 필요" in str(i.get("검증결과", ""))]
        dup_items = [i for i in mapped_items if "중복" in str(i.get("검증결과", ""))]
        
        vcol1, vcol2, vcol3 = st.columns(3)
        vcol1.metric("💰 금액 검증 오류", len(error_items))
        vcol2.metric("❓ 확인 필요", len(missing_items))
        vcol3.metric("🔄 중복", len(dup_items))

    # ─── 견적 요약 (Excel에서 직접 읽기) ────────────────────────
    from io import BytesIO as _BIO
    from openpyxl import load_workbook as _lwb

    def _read_excel_summary(xlsx_bytes: bytes) -> Dict[str, float]:
        """Excel '2.상세 견적' 시트에서 세부 행의 G(수량)×I(단가)를 합산하여
        소계/합계를 계산한다. J열은 수식이므로 직접 읽을 수 없으므로,
        compute_excel_total()과 동일한 방식으로 G×I를 합산한다.
        
        Returns:
            {UI_라벨: 금액} 딕셔너리
        """
        try:
            wb = _lwb(_BIO(xlsx_bytes), data_only=False)
        except Exception:
            return {}
        if "2.상세 견적" not in wb.sheetnames:
            return {}
        ws = wb["2.상세 견적"]

        def _num(v) -> float:
            if v is None:
                return 0.0
            if isinstance(v, (int, float)):
                return float(v)
            s = str(v)
            for ch in ("₩", "원", ",", " ", "￦", "$", "€", "¥"):
                s = s.replace(ch, "")
            try:
                return float(s)
            except (ValueError, TypeError):
                return 0.0

        # 각 세부 행의 G×I를 (gubun, dangayuh, section)별로 합산
        # key: (gubun, dangayuh, section) → sum(G*I)
        section_sums: Dict[tuple, float] = {}
        cur_gubun = ""
        cur_dangayuh = ""
        cur_section = ""

        for r in range(1, ws.max_row + 1):
            a = ws.cell(r, 1).value
            b = ws.cell(r, 2).value
            c = ws.cell(r, 3).value

            if a is not None and str(a).strip():
                cur_gubun = str(a).strip()
            if b is not None and str(b).strip():
                cur_dangayuh = str(b).strip()

            cs = str(c).strip() if c is not None else ""

            # 섹션 헤더 행 (C열이 "(1) 설계" 등)
            if cs.startswith("("):
                cur_section = cs
                continue

            # 소계/합계/Sub합계/TOTAL/합산 행 제외 (수식 행)
            if any(k in cs for k in ("소계", "합계", "Sub합계", "TOTAL", "합산")):
                continue

            # 세부 행: G(수량) × I(단가) 합산
            g = _num(ws.cell(r, 7).value)
            i = _num(ws.cell(r, 9).value)
            if g > 0 and i > 0:
                # D열(자재명) 확인: PRI Engineering Fee 별도 처리
                # C열이 "0)"이고 D열이 "Engineering"인 행 = PRI Engineering Fee
                d_val = str(ws.cell(r, 4).value or "").strip()
                if cur_gubun == "인건비" and cs == "0)" and d_val == "Engineering":
                    key = (cur_gubun, cur_dangayuh, "__PRI__")
                else:
                    key = (cur_gubun, cur_dangayuh, cur_section)
                section_sums[key] = section_sums.get(key, 0.0) + g * i

        # (gubun, dangayuh, section) → UI 라벨 매핑
        result: Dict[str, float] = {}
        for (gubun, dangayuh, section), amount in section_sums.items():
            if amount == 0:
                continue

            if gubun == "기구부":
                if "구매" in dangayuh:
                    result["기구·구매품 소계"] = result.get("기구·구매품 소계", 0.0) + amount
                elif "가공" in dangayuh:
                    result["기구·가공품 소계"] = result.get("기구·가공품 소계", 0.0) + amount
            elif gubun in ("전장/제어", "전장", "제어"):
                if "구매" in dangayuh:
                    # 섹션 헤더로 제어/전장 구분
                    if "제어" in section:
                        result["제어·구매품 소계"] = result.get("제어·구매품 소계", 0.0) + amount
                    elif "전장" in section:
                        result["전장·구매품 소계"] = result.get("전장·구매품 소계", 0.0) + amount
                    else:
                        # 섹션 구분 불가: 첫 번째 = 제어, 두 번째 = 전장
                        if "제어·구매품 소계" not in result:
                            result["제어·구매품 소계"] = amount
                        else:
                            result["전장·구매품 소계"] = result.get("전장·구매품 소계", 0.0) + amount
                elif "가공" in dangayuh:
                    result["전장·가공품 소계"] = result.get("전장·가공품 소계", 0.0) + amount
            elif gubun == "인건비":
                if section == "__PRI__":
                    result["PRI Engineering"] = result.get("PRI Engineering", 0.0) + amount
                elif "설계" in section:
                    result["설계 인건비"] = result.get("설계 인건비", 0.0) + amount
                elif "셋업" in section:
                    result["셋업 인건비"] = result.get("셋업 인건비", 0.0) + amount
                elif "조립" in section:
                    result["조립 인건비"] = result.get("조립 인건비", 0.0) + amount
                elif "양산" in section or "특수" in section:
                    result["양산대응 인건비"] = result.get("양산대응 인건비", 0.0) + amount
            elif gubun == "경비":
                result["경비 및 미분류"] = result.get("경비 및 미분류", 0.0) + amount

        # 합계 계산
        result["기구부 합계"] = result.get("기구·구매품 소계", 0.0) + result.get("기구·가공품 소계", 0.0)
        result["전장/제어부 합계"] = (
            result.get("제어·구매품 소계", 0.0)
            + result.get("제어·가공품 소계", 0.0)
            + result.get("전장·구매품 소계", 0.0)
            + result.get("전장·가공품 소계", 0.0)
        )
        result["인건비 합계"] = (
            result.get("PRI Engineering", 0.0)
            + result.get("설계 인건비", 0.0)
            + result.get("셋업 인건비", 0.0)
            + result.get("조립 인건비", 0.0)
            + result.get("양산대응 인건비", 0.0)
        )

        # Total 견적 = 모든 세부 행 G×I 합계 × 1.10
        total_base = sum(section_sums.values())
        result["Total 견적"] = total_base * 1.10

        return result

    _excel_summary = _read_excel_summary(output_bytes) if output_bytes else {}

    _SUM_ORDER = [
        ("기구·구매품 소계",   "sub"),
        ("기구·가공품 소계",   "sub"),
        ("기구부 합계",         "fam"),
        ("제어·구매품 소계",   "sub"),
        ("제어·가공품 소계",   "sub"),
        ("전장·구매품 소계",   "sub"),
        ("전장·가공품 소계",   "sub"),
        ("전장/제어부 합계",    "fam"),
        ("PRI Engineering",     "sub"),
        ("설계 인건비",         "sub"),
        ("조립 인건비",         "sub"),
        ("셋업 인건비",         "sub"),
        ("양산대응 인건비",     "sub"),
        ("인건비(특급)",        "sub"),
        ("인건비 합계",         "fam"),
        ("경비 및 미분류",      "sub"),
        ("Total 견적",          "total"),
    ]

    _fmt_won_s = lambda v: f"{v:,.0f}"

    def _sum_row_html(ui_label: str, kind: str) -> str:
        val = _excel_summary.get(ui_label, 0.0)
        if kind == "sub":
            bg, fw = "#fff", "normal"
        elif kind == "fam":
            bg, fw = "#ffe0b2", "bold"
        else:
            bg, fw = "#c8e6c9", "bold"
        return (f"<tr style='background:{bg};'>"
                f"<td style='border:1px solid #ccc;padding:4px 8px;font-weight:{fw};'>{ui_label}</td>"
                f"<td style='border:1px solid #ccc;padding:4px 8px;text-align:right;font-weight:{fw};'>{_fmt_won_s(val)}</td></tr>")

    # 전장/제어부 합계 기준으로 좌우 분할
    _split_idx = next(i for i, (lbl, _) in enumerate(_SUM_ORDER) if lbl == "전장/제어부 합계")
    _left_items = _SUM_ORDER[:_split_idx + 1]
    _right_items = _SUM_ORDER[_split_idx + 1:]

    _th_style = "border:1px solid #ccc;padding:4px 8px;background:#f5f5f5;"
    _sum_html = "<table style='width:100%;border-collapse:collapse;font-size:0.84rem;'>"
    # 헤더 행: 구분(왼쪽) | 구분(오른쪽) | 금액(왼쪽) | 금액(오른쪽)
    _sum_html += f"<tr><th style='{_th_style}text-align:left;width:35%;'>구분</th>"
    _sum_html += f"<th style='{_th_style}text-align:right;width:15%;'>금액(원)</th>"
    _sum_html += f"<th style='{_th_style}text-align:left;width:35%;'>구분</th>"
    _sum_html += f"<th style='{_th_style}text-align:right;width:15%;'>금액(원)</th></tr>"
    # 데이터 행: 왼쪽 8행 + 오른쪽 9행 → 최대 9행
    _max_rows = max(len(_left_items), len(_right_items))
    for i in range(_max_rows):
        _sum_html += "<tr>"
        # 왼쪽
        if i < len(_left_items):
            lbl, kind = _left_items[i]
            val = _excel_summary.get(lbl, 0.0)
            bg, fw = ("#fff", "normal") if kind == "sub" else ("#ffe0b2", "bold") if kind == "fam" else ("#c8e6c9", "bold")
            _sum_html += f"<td style='border:1px solid #ccc;padding:4px 8px;background:{bg};font-weight:{fw};'>{lbl}</td>"
            _sum_html += f"<td style='border:1px solid #ccc;padding:4px 8px;text-align:right;background:{bg};font-weight:{fw};'>{_fmt_won_s(val)}</td>"
        else:
            _sum_html += "<td style='border:1px solid #ccc;padding:4px 8px;'></td>"
            _sum_html += "<td style='border:1px solid #ccc;padding:4px 8px;'></td>"
        # 오른쪽
        if i < len(_right_items):
            lbl, kind = _right_items[i]
            val = _excel_summary.get(lbl, 0.0)
            bg, fw = ("#fff", "normal") if kind == "sub" else ("#ffe0b2", "bold") if kind == "fam" else ("#c8e6c9", "bold")
            _sum_html += f"<td style='border:1px solid #ccc;padding:4px 8px;background:{bg};font-weight:{fw};'>{lbl}</td>"
            _sum_html += f"<td style='border:1px solid #ccc;padding:4px 8px;text-align:right;background:{bg};font-weight:{fw};'>{_fmt_won_s(val)}</td>"
        else:
            _sum_html += "<td style='border:1px solid #ccc;padding:4px 8px;'></td>"
            _sum_html += "<td style='border:1px solid #ccc;padding:4px 8px;'></td>"
        _sum_html += "</tr>"
    _sum_html += "</table>"
    st.markdown(_sum_html, unsafe_allow_html=True)

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
    _ROW_H = 30
    st.markdown(
        f"""
        <style>
        div[data-testid="stNumberInput"] {{
            height: {_ROW_H}px !important;
            margin: 0 !important;
        }}
        div[data-testid="stNumberInput"] .react-aria-TextField {{
            height: {_ROW_H}px !important;
        }}
        div[data-testid="stNumberInput"] .react-aria-TextField > div {{
            height: {_ROW_H}px !important;
        }}
        div[data-testid="stNumberInput"] input {{
            height: {_ROW_H}px !important;
            line-height: {_ROW_H}px !important;
            padding: 0 4px !important;
            font-size: 0.84rem !important;
        }}
        div[data-testid="stMarkdownContainer"]:has(> div.adj-cell) {{
            margin-bottom: 4px !important;
        }}
        </style>
        """,
        unsafe_allow_html=True,
    )

    def _cell(col, html: str, color: str = "#222"):
        """고정 높이(_ROW_H) flex 박스로 셀 렌더 → 컬럼 간 행 정렬."""
        col.markdown(
            f"<div class='adj-cell' style='height:{_ROW_H}px;display:flex;"
            f"align-items:center;font-size:0.84rem;color:{color};'>{html}</div>",
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
    # 인건비 조정과 동일하게 매 행마다 새 st.columns()를 생성하여
    # 컬럼 간 마진 누적에 의한 행 어긋남을 방지.
    _FAM_W = [2, 1, 1.5, 1, 1]
    st.subheader("재료비·경비 조정 (변경 비율)")
    _fh = st.columns(_FAM_W)
    _cell(_fh[0], "구분", "#888")
    _cell(_fh[1], "현견적(원)", "#888")
    _cell(_fh[2], "변경 비율(%)", "#b8860b")
    _cell(_fh[3], "변경(원)", "#888")
    _cell(_fh[4], "증감(원)", "#888")

    for fam in ("기구부", "전장부", "제어부"):
        cur = model.mat_made.get(fam, 0.0)
        key = f"made|{fam}|pct"
        pct = st.session_state.ratio_inputs.get(key, 100.0)
        _fc = st.columns(_FAM_W)
        _cell(_fc[0], f"1. 가공품 · {fam}")
        _cell(_fc[1], _fmt_won(cur))
        new_pct = _fc[2].number_input(
            "변경 비율(%)", min_value=0.0, max_value=500.0, step=5.0, value=float(pct),
            key=f"ratio_{key}", label_visibility="collapsed",
        )
        st.session_state.ratio_inputs[key] = new_pct
        model.made_pct[fam] = new_pct
        adj = model.made_adj(fam)
        _cell(_fc[3], _fmt_won(adj))
        _cell(_fc[4], f"{adj - cur:+,.0f}", "#c0392b" if adj - cur > 0 else "#27ae60")

    # 구매품 (개별 변경 불가 — 표시만)
    for fam in ("기구부", "전장부", "제어부"):
        cur = model.mat_buy.get(fam, 0.0)
        if not cur:
            continue
        _fc = st.columns(_FAM_W)
        _cell(_fc[0], f"2. 구매품 · {fam}")
        _cell(_fc[1], _fmt_won(cur))
        _cell(_fc[2], "개별 변경 불가", "#888")
        _cell(_fc[3], _fmt_won(cur))
        _cell(_fc[4], "0")

    # 경비 (비율 입력)
    _fc = st.columns(_FAM_W)
    _cell(_fc[0], "4. 경비")
    _cell(_fc[1], _fmt_won(model.expense))
    exp_pct = _fc[2].number_input(
        "경비 변경 비율(%)", min_value=0.0, max_value=500.0, step=5.0,
        value=float(st.session_state.ratio_inputs.get("expense|pct", 100.0)),
        key="ratio_expense_pct", label_visibility="collapsed",
    )
    st.session_state.ratio_inputs["expense|pct"] = exp_pct
    model.expense_pct = exp_pct
    exp_adj = model.expense_adj()
    _cell(_fc[3], _fmt_won(exp_adj))
    _cell(_fc[4], f"{exp_adj - model.expense:+,.0f}",
          "#c0392b" if exp_adj - model.expense > 0 else "#27ae60")

    # ── 2) 인건비 조정: 대응인원수/대응기간 입력 ──
    st.subheader("인건비 조정 (대응인원수 · 대응기간)")
    if not model.labor_rows:
        st.info("인건비 항목이 없습니다.")
    else:
        # 헤더 행 (재료비·경비 조정과 동일 양식)
        # 인원수/대응일수 컬럼을 넓혀 number_input 스텝 버튼(−/+)이 표시되도록
        _lh = st.columns([3, 1, 1.5, 1.5, 1, 1])
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
            st.markdown(f"<div class='sec-hdr'>{sec}</div>", unsafe_allow_html=True)
            for r in sec_rows:
                key_p = f"{r.section}|{r.job}|people"
                key_d = f"{r.section}|{r.job}|days"
                ic1, ic2, ic3, ic4, ic5, ic6 = st.columns([3, 1, 1.5, 1.5, 1, 1])
                _cell(ic1, f"{r.job} "
                    f"<span style='color:#888;font-size:0.82rem;'>단가 {_fmt_won(r.rate)}</span>")
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

    # ── 3) 경비 자동산출: Site 선택 → 항공료/숙박비/식비·교통비 자동 입력 ──
    st.subheader("경비 자동산출 (Site 선택)")
    from config import EXPENSE_SITE_RATES, EXPENSE_SECTION_MAP
    _EXP_SECS = ["설계", "조립", "셋업", "양산대응"]
    _EXP_ITEMS = ["항공료", "숙박비", "식비/교통비"]
    _EW = [2.5, 1.5, 1.5, 1.5, 1.5]

    def _calc_expense_amounts(_model) -> dict:
        """Site 선택 + 인건비 행 기준 경비 금액 산출 (UI/산출 공용)
        항공료     = 투입인원수 sum × 항공료단가  (한국 자사/출장 = 0)
        숙박비     = 투입인원수×대응일수 sum(MD) × 숙박비단가  (한국 자사 = 0)
        식비/교통비 = 투입인원수×대응일수 sum(MD) × 식비/교통비단가  (한국 자사 = 0)
        직접입력 섹션은 제외 (수동 입력만)."""
        _res = {}
        for _sec in _EXP_SECS:
            _site = st.session_state.get(f"exp_site|{_sec}", "한국(자사근무)")
            _rates = EXPENSE_SITE_RATES.get(_site)
            if not _rates:
                continue  # 직접입력
            _rows = [r for r in _model.labor_rows
                     if r.section == EXPENSE_SECTION_MAP[_sec]]
            _people = sum(r.adj_people for r in _rows)
            _md = sum(r.adj_md for r in _rows)
            _res[_sec] = {
                "항공료": _people * _rates["항공료"]
                          if _site not in ("한국(자사근무)", "한국(출장대응)") else 0.0,
                "숙박비": _md * _rates["숙박비"]
                          if _site != "한국(자사근무)" else 0.0,
                "식비/교통비": _md * _rates["식비/교통비"]
                               if _site != "한국(자사근무)" else 0.0,
            }
        return _res

    _eh = st.columns(_EW)
    _cell(_eh[0], "구분", "#888")
    for _ci, _s in enumerate(_EXP_SECS):
        _cell(_eh[_ci + 1], _s, "#888")

    # Site 선택 행 (드롭다운)
    _fc = st.columns(_EW)
    _cell(_fc[0], "Site 선택")
    for _ci, _sec in enumerate(_EXP_SECS):
        _fc[_ci + 1].selectbox(
            "Site 선택",
            options=list(EXPENSE_SITE_RATES.keys()) + ["직접입력"],
            index=0,
            key=f"exp_site|{_sec}",
            label_visibility="collapsed",
        )

    # 자동산출: 매 렌더 재계산 (Site/인건비 변경 즉시 반영).
    # 사용자가 수동 수정한 값(마지막 자동산출값과 다름)은 보존.
    _exp_calc = _calc_expense_amounts(model)
    for _sec in _EXP_SECS:
        for _it in _EXP_ITEMS:
            _k = f"exp_{_it}|{_sec}"
            _bk = f"{_k}_base"
            _new = float(_exp_calc.get(_sec, {}).get(_it, 0.0))
            _cur = st.session_state.get(_k)
            if _cur is None or float(_cur) == float(st.session_state.get(_bk, _new)):
                st.session_state[_k] = _new
                st.session_state[_bk] = _new

    # 항공료/숙박비/식비·교통비 행 (자동 입력 + 수동 수정 가능)
    for _it in _EXP_ITEMS:
        _fc = st.columns(_EW)
        _cell(_fc[0], _it)
        for _ci, _sec in enumerate(_EXP_SECS):
            _fc[_ci + 1].number_input(
                _it, min_value=0.0, step=10000.0,
                value=float(st.session_state.get(f"exp_{_it}|{_sec}", 0.0)),
                key=f"exp_{_it}|{_sec}", label_visibility="collapsed",
            )

    # 합계 / 경비 합계 행 (자동 계산)
    _fc = st.columns(_EW)
    _cell(_fc[0], "합계", "#0066cc")
    _exp_grand = 0.0
    for _ci, _sec in enumerate(_EXP_SECS):
        _s = sum(float(st.session_state.get(f"exp_{_it}|{_sec}", 0.0) or 0.0)
                 for _it in _EXP_ITEMS)
        _exp_grand += _s
        _cell(_fc[_ci + 1], f"{_s:,.0f}", "#0066cc")
    _fc = st.columns(_EW)
    _cell(_fc[0], "경비 합계", "#0066cc")
    _cell(_fc[1], f"{_exp_grand:,.0f}원", "#0066cc")

    # ── 견적조정 산출 버튼 ── (on_click 콜백: PRI Engineering 대응일수 자동 반영)
    def _on_calc():
        _m = st.session_state.get("ratio_model")
        _inputs = st.session_state.get("ratio_inputs", {})
        if _m is None:
            return
        apply_adjustment(_m, _inputs)
        # PRI Engineering 대응일수 자동 재산출:
        # 섹션별(설계/조립/셋업/양산대응) 조정된 대응일수 중 최대값을 합산
        _eng = next((r for r in _m.labor_rows
                     if r.section == "(0) PRI Engineering"), None)
        if _eng is not None:
            _nd = 0.0
            for _sec in ("(1) 설계", "(2) 조립", "(3) 셋업", "(4) 양산대응"):
                _sd = [r.adj_days for r in _m.labor_rows
                       if r.section == _sec and r.adj_days > 0]
                if _sd:
                    _nd += max(_sd)
            _eng.adj_days = _nd
            _kd = f"{_eng.section}|{_eng.job}|days"
            _inputs[_kd] = _nd
            st.session_state[f"ratio_{_kd}"] = _nd
        # 경비 자동산출 합계 반영 (Site별 항공/숙박/식비·교통 합)
        from config import EXPENSE_SITE_RATES, EXPENSE_SECTION_MAP
        _exp_grand = 0.0
        # (item, site) → 총 수량 (여러 섹션이 같은 Site를 선택하면 합산)
        _site_data: dict = {}
        for _sec in ("설계", "조립", "셋업", "양산대응"):
            _site = st.session_state.get(f"exp_site|{_sec}", "한국(자사근무)")
            _rates = EXPENSE_SITE_RATES.get(_site)
            if not _rates:
                continue  # 직접입력 → 수동 금액만 합산, 행 매핑 없음
            _lab_sec = EXPENSE_SECTION_MAP[_sec]
            _sec_rows = [r for r in _m.labor_rows if r.section == _lab_sec]
            _people_sum = sum(r.adj_people for r in _sec_rows)
            _md_sum = sum(r.adj_md for r in _sec_rows)
            for _it in ("항공료", "숙박비", "식비/교통비"):
                _amt = float(st.session_state.get(f"exp_{_it}|{_sec}", 0.0) or 0.0)
                _exp_grand += _amt
                _rate = float(_rates[_it])
                _qty = (_amt / _rate) if _rate > 0 else 0.0
                _key = (_it, _site)
                _site_data[_key] = _site_data.get(_key, 0.0) + _qty
        _m.expense_auto_total = _exp_grand
        _m.expense_site_data = _site_data
        try:
            st.session_state._adj_bytes = export_adjusted_excel(
                st.session_state.get("_pipe_cache", {}).get("result", (None,)*8)[2], _m)
            st.session_state._adj_calculated = True
        except Exception as _ae:
            st.session_state._adj_bytes = None
            st.session_state._adj_calculated = False
            st.session_state._calc_err = str(_ae)

    if st.button("📐 견적조정 산출", type="primary", on_click=_on_calc):
        if st.session_state.get("_calc_err"):
            st.warning(f"조정견적 생성 실패: {st.session_state._calc_err}")
            st.session_state._calc_err = None
            st.session_state._adj_calculated = False

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
            f"<div style='font-size:0.84rem;color:#888;'>현견적 총액</div>"
            f"<div style='font-size:0.72rem;font-weight:600;'>{cur_total:,.0f}원</div>",
            unsafe_allow_html=True,
        )
        tc2.markdown(
            f"<div style='font-size:0.84rem;color:#888;'>변경 후 총액</div>"
            f"<div style='font-size:0.72rem;font-weight:600;'>{adj_total:,.0f}원</div>",
            unsafe_allow_html=True,
        )
        tc3.markdown(
            f"<div style='font-size:0.84rem;color:#888;'>증감</div>"
            f"<div style='font-size:0.72rem;font-weight:600;color:{'#c0392b' if _delta > 0 else '#27ae60'};'>{_delta:+,.0f}원</div>",
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
            if not csv_content or not csv_content.strip():
                # 빈 CSV → 헤더만 있는 Excel
                buf = BytesIO()
                pd.DataFrame(columns=["항목"]).to_excel(buf, index=False, engine="openpyxl")
                return buf.getvalue()
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
