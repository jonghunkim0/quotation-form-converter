"""
견적 조정 (비율 조정) 모듈

`비율조정 로직.xlsx`의 구조를 Streamlit UI로 재현한다.

구조 (Excel 기준):
  ┌─ 상단 요약 테이블 ──────────────────────────────────────┐
  │ 1. 가공품  (기구부/전장부/제어부)   금액(자동) │ 변경비율(입력) │
  │ 2. 구매품  (기구부/전장부/제어부)   금액(자동) │ 개별변경불가   │
  │ 3. 인건비                          금액(수식) │ 아래표에서조정 │
  │ 4. 경비                            금액(자동) │ 지역명조정     │
  │ 5. 관리비 및 이윤                  금액(자동) │ 자동반영       │
  │ 총 견적 금액                        (수식)                     │
  └──────────────────────────────────────────────────────────┘
  ┌─ 인건비 상세 테이블 (좌=현견적, 우=변경) ────────────────┐
  │ 직종        │ 인원수 │ 대응기간 │ MD(=인원×기간) │ 단가 │ 금액(=MD×단가) │
  │ (단가=회색 자동, 인원수/대응기간=노란색 입력)              │
  └──────────────────────────────────────────────────────────┘

수식 로직 (Excel과 동일):
  - MD   = 인원수 × 대응기간
  - 금액 = MD × 단가
  - 섹션 합계 = 섹션 내 상세 행 금액 합
  - 인건비 = Engineering + 각 섹션 합계
  - 총 견적 = 가공품 + 구매품 + 인건비 + 경비 + 관리비
"""

from dataclasses import dataclass, field
from typing import List, Dict, Any, Optional, Tuple


# ─── analyzer 카테고리 → (섹션, 직종 라벨) ─────────────────────
_LABOR_SECTION_MAP: Dict[str, Tuple[str, str]] = {
    "인건비-기구 설계":       ("(1) 설계",     "기구설계자"),
    "인건비-전장 설계":       ("(1) 설계",     "전장설계자"),
    "인건비-제어 설계":       ("(1) 설계",     "제어설계자"),
    "인건비-기구 조립":       ("(2) 조립",     "기구조립"),
    "인건비-전장/제어 조립":  ("(2) 조립",     "전장/제어조립"),
    "인건비-기구 셋업":       ("(3) 셋업",     "기구셋업"),
    "인건비-전장/제어 셋업":  ("(3) 셋업",     "전장/제어셋업"),
    "인건비-기구 양산대응":   ("(4) 양산대응", "기구양산대응"),
    "인건비-전장/제어 양산대응": ("(4) 양산대응", "전장/제어양산대응"),
    "인건비-특수제어":        ("(5) 특수제어", "제어엔지니어"),
}

SECTION_ORDER = ["(0) PRI Engineering", "(1) 설계", "(2) 조립", "(3) 셋업", "(4) 양산대응", "(5) 특수제어"]


@dataclass
class LaborRow:
    """인건비 상세 1행 (직종)."""
    section: str          # (1) 설계 ...
    job: str              # 기구설계자 ...
    rate: float = 0.0     # 단가 (회색, 분석 결과)
    # 현견적 (좌 블록)
    cur_people: float = 0.0   # 인원수 (노란색 입력)
    cur_days: float = 0.0     # 대응기간 (노란색 입력)
    # 변경 (우 블록)
    adj_people: float = 0.0   # 인원수 (노란색 입력)
    adj_days: float = 0.0     # 대응기간 (노란색 입력)

    # ── 수식 계산 ──
    @property
    def cur_md(self) -> float:
        return self.cur_people * self.cur_days

    @property
    def cur_amount(self) -> float:
        return self.cur_md * self.rate

    @property
    def adj_md(self) -> float:
        return self.adj_people * self.adj_days

    @property
    def adj_amount(self) -> float:
        return self.adj_md * self.rate


@dataclass
class RatioModel:
    """견적 조정 전체 모델."""
    labor_rows: List[LaborRow] = field(default_factory=list)
    # 상단 요약 (재료비 등, 분석/변환 결과에서 자동 채움)
    mat_made: Dict[str, float] = field(default_factory=dict)   # 가공품 {부: 금액}
    mat_buy: Dict[str, float] = field(default_factory=dict)    # 구매품 {부: 금액}
    expense: float = 0.0        # 경비
    overhead: float = 0.0       # 관리비 및 이윤
    # 변경 비율(%) — 노란색 입력 셀 (100 = 변동 없음)
    made_pct: Dict[str, float] = field(default_factory=dict)   # 가공품 {부: %}
    expense_pct: float = 100.0  # 경비 %

    # ── 섹션별 인건비 합계 ──
    def section_sum(self, section: str, side: str = "cur") -> float:
        key = "cur_amount" if side == "cur" else "adj_amount"
        return sum(getattr(r, key) for r in self.labor_rows if r.section == section)

    def labor_total(self, side: str = "cur") -> float:
        return sum(self.section_sum(s, side) for s in SECTION_ORDER)

    def mat_total(self, kind: str) -> float:
        d = self.mat_made if kind == "made" else self.mat_buy
        return sum(d.values())

    # ── 변경(%) 적용 금액 ──
    def made_adj(self, fam: str) -> float:
        """가공품 변경 금액 = 현견적 × 비율%"""
        return self.mat_made.get(fam, 0.0) * self.made_pct.get(fam, 100.0) / 100.0

    def expense_adj(self) -> float:
        """경비 변경 금액 = 현견적 × 비율%"""
        return self.expense * self.expense_pct / 100.0

    def grand_total(self, side: str = "cur") -> float:
        """총 견적 = (재료비 + 인건비 + 경비) × 1.10 (일반관리 5% + 기업이윤 5%)"""
        if side == "cur":
            base = (self.mat_total("made") + self.mat_total("buy")
                    + self.labor_total("cur") + self.expense)
        else:
            base = (sum(self.made_adj(f) for f in self.mat_made)
                    + self.mat_total("buy")
                    + self.labor_total("adj") + self.expense_adj())
        return base * 1.10


def _parse_amount(val) -> float:
    """금액 문자열/숫자를 float로 변환 (통화 기호/쉼표 제거)."""
    if val is None:
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val)
    for ch in ("₩", "원", ",", " ", "￦", "$", "€", "¥"):
        s = s.replace(ch, "")
    try:
        return float(s)
    except ValueError:
        return 0.0


def _decompose_md(md: float, remark: str) -> Tuple[float, float]:
    """MD를 (인원수, 대응기간)으로 분해. _calc_labor_md_split과 동일한 후보 탐색."""
    import math
    import re
    md = int(round(md or 0))
    if md <= 0:
        return 0.0, 0.0
    pairs = [(int(p), int(d)) for p, d in re.findall(
        r"(\d+)\s*[名名]\s*[xX×*]\s*(\d+)\s*[日日]", remark or "")]
    if pairs:
        total_p = sum(p for p, _ in pairs)
        days = max(1, int(round(md / total_p)))
        return float(total_p), float(days)
    for days in (7, 14, 21, 30, 60):
        for people in range(1, 11):
            if people * days == md:
                return float(people), float(days)
    days = min((7, 14, 21, 30, 60), key=lambda d: abs(md - round(md / d) * d))
    return float(max(1, round(md / days))), float(days)


def build_ratio_model(labor_data: Dict[str, Dict[str, Any]],
                      mapped_items: List[Dict[str, Any]]) -> RatioModel:
    """
    labor_data + mapped_items에서 RatioModel을 구성.

    - 인건비 행: analyzer 카테고리별 (단가=분석, 인원/기간=MD 분해)
    - 재료비: mapped_items의 _분류별 금액 합산
    """
    model = RatioModel()

    # ── 인건비 행 ──
    for cat, info in (labor_data or {}).items():
        if cat not in _LABOR_SECTION_MAP:
            continue
        section, job = _LABOR_SECTION_MAP[cat]
        md = info.get("md", 0) or 0
        rate = info.get("rate", 0) or 0
        remark = info.get("remark", "") or ""
        people, days = _decompose_md(md, remark)
        model.labor_rows.append(LaborRow(
            section=section, job=job, rate=rate,
            cur_people=people, cur_days=days,
            adj_people=people, adj_days=days,  # 변경은 현견적과 동일하게 시작
        ))

    # ── Engineering (PRI Engineering Fee) — 단가 600,000원 고정 ──
    total_labor_md = sum(info.get("md", 0) or 0 for info in (labor_data or {}).values())
    if total_labor_md > 0:
        if total_labor_md < 21:
            eng_md = int(round(total_labor_md / 7.0)) * 7
        else:
            eng_md = int(round(total_labor_md / 30.0)) * 30
    else:
        eng_md = 0
    eng_people, eng_days = _decompose_md(eng_md, "")
    model.labor_rows.insert(0, LaborRow(
        section="(0) PRI Engineering", job="PRI Engineering", rate=600000,
        cur_people=eng_people, cur_days=eng_days,
        adj_people=eng_people, adj_days=eng_days,
    ))

    # ── 재료비 (mapped_items의 _분류별 금액 합산) ──
    made = {"기구부": 0.0, "전장부": 0.0, "제어부": 0.0}
    buy = {"기구부": 0.0, "전장부": 0.0, "제어부": 0.0}
    _FAMILY = {
        "기구_가공품": ("made", "기구부"), "전장_가공품": ("made", "전장부"),
        "제어_가공품": ("made", "제어부"),
        "기구_구매품": ("buy", "기구부"), "전장_구매품": ("buy", "전장부"),
        "제어_구매품": ("buy", "제어부"),
    }
    expense = 0.0
    for it in (mapped_items or []):
        # Unit(합산) 행은 하위 세부항목의 합계이므로 스킵 (이중계산 방지)
        if it.get("_합산"):
            continue
        sec = it.get("_분류", "")
        # Excel과 동일하게 수량 × 단가로 계산 (금액 필드와 다를 수 있음)
        qty = _parse_amount(it.get("수량"))
        price = _parse_amount(it.get("단가"))
        amt = qty * price
        if sec in _FAMILY:
            kind, fam = _FAMILY[sec]
            (made if kind == "made" else buy)[fam] += amt
        elif sec in ("포장", "기타항목", "경비"):
            expense += amt

    model.mat_made = made
    model.mat_buy = buy
    model.expense = expense
    return model


def apply_adjustment(model: RatioModel, inputs: Dict[str, Any]) -> None:
    """
    사용자 입력(변경 블록)을 모델에 적용.

    inputs 키: f"{section}|{job}|people" / f"{section}|{job}|days"
    """
    for row in model.labor_rows:
        key_p = f"{row.section}|{row.job}|people"
        key_d = f"{row.section}|{row.job}|days"
        if key_p in inputs and inputs[key_p] is not None:
            row.adj_people = float(inputs[key_p])
        if key_d in inputs and inputs[key_d] is not None:
            row.adj_days = float(inputs[key_d])
