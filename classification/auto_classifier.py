"""
classification/auto_classifier.py - 자동 분류 로직 (템플릿 기준 수정版)

원가분석률 프로젝트의 AutoClassifier를 기반으로,
고객사 표준 양식 템플릿의 섹션 구조에 맞게 수정.

템플릿 섹션 구조:
  기구부 > 구매품 (Row 5~9)
  기구부 > 가공품 (Row 10~14)
  전장/제어 > (1) 제어부 구매품 (Row 15~20)
  전장/제어 > (1) 제어부 가공품 (Row 21~22)
  전장/제어 > (2) 전장부 구매품 (Row 23~29)
  전장/제어 > (2) 전장부 가공품 (Row 30~33)
  인건비 > (1) 설계 (Row 34~47)
  인건비 > (4) 조립 (Row 48~56)
  인건비 > (5) 셋업 (Row 57~65)
  인건비 > (6) 양산대응 (Row 66~74)
  인건비 > (7) 특수 제어 (Row 75~77)
  경비 > (1) 포장 (Row 78~83)
  경비 > (2) 기타항목 (Row 84~87)
  경비 > (3) 경비 (Row 88~103)
"""

import re
from typing import Optional, Dict, Any, List, Tuple

# ─── 텍스트 정규화 ─────────────────────────────────────────────
def normalize_text(text: str) -> str:
    """텍스트 정규화: 소문자, 공백 제거"""
    if not text:
        return ""
    return re.sub(r"\s+", "", str(text).lower())


# ─── 템플릿 섹션 정의 ─────────────────────────────────────────
# 각 섹션: (section_key, gubun, sub_section, start_row, end_row)
# section_key: 분류 결과를 이 키로 매핑
TEMPLATE_SECTIONS = [
    # (section_key, gubun, sub_section, start_row, end_row)
    # ⚠️ start_row는 '합산항목' 헤더 행(수식 포함)을 제외하고 실제 데이터 행부터 시작.
    #    합산항목 행(예: R5 '(1) 기구부', R10 '가공품')은 절대 건드리지 않는다.
    ("기구_구매품",   "기구부",     "구매품",     6,  8),
    ("기구_가공품",   "기구부",     "가공품",    11, 12),
    ("제어_구매품",   "전장/제어",  "제어부_구매품", 16, 19),
    # 제어부_가공품: R21(서브헤더)~R22(소계) 사이에 데이터 행이 없음 → 빈 범위(쓰기 불가)
    ("제어_가공품",   "전장/제어",  "제어부_가공품", 22, 21),
    ("전장_구매품",   "전장/제어",  "전장부_구매품", 24, 28),
    ("전장_가공품",   "전장/제어",  "전장부_가공품", 31, 31),
    ("설계",         "인건비",     "설계",      35, 47),
    ("조립",         "인건비",     "조립",      49, 56),
    ("셋업",         "인건비",     "셋업",      58, 65),
    ("양산대응",     "인건비",     "양산대응",  67, 74),
    ("특수제어",     "인건비",     "특수제어",  76, 76),
    ("포장",         "경비",       "포장",      79, 83),
    ("기타항목",     "경비",       "기타항목",  85, 87),
    ("경비",         "경비",       "경비",      89, 102),
]

# ─── 분류 키워드 규칙 (템플릿 섹션 기준) ──────────────────────
CLASSIFICATION_RULES = {
    "기구_구매품": [
        "구매품", "구매 part", "구매파트", "상용품", "표준 구매",
        "vendor item", "commercial part", "buy part", "purchased item",
        # 기계 표준 구매 부품
        "모터", "서보", "실린더", "로봇", "lm가이드", "lm 가이드",
        "볼스크류", "볼 스크류", "베어링", "기어", "커플링",
        "벨트", "풀리", "스프링",
        # 기구 관련 구매 부품
        "볼트", "너트", "샤프트", "축", "플랜지", "스페이서",
    ],
    "기구_가공품": [
        "가공품", "프레임", "플레이트", "플라이트", "플레이트재",
        "브라켓", "브래킷", "블라켓", "블랙켓",
        "알루미늄", "프로파일", "강판", "스틸", "구리",
        "베이스", "커버", "하우징", "케이스", "지지대",
        "절단", "드릴", "타공", "절곡", "가공",
        # 약자
        "plt", "blk", "bkt", "plate", "bracket",
    ],
    "제어_구매품": [
        "plc", "hmi", "센서", "인버터", "차단기", "릴레이", "스위치",
        "터미널", "커넥터", "전원공급기", "psu", "컨택터",
        "마그네틱", "피더", "브레이커", "조명", "램프", "버튼",
        "코일", "트랜스", "제어 부품", "제어",
    ],
    "제어_가공품": [
        "제어 가공", "제어부 가공", "컨트롤 박스", "제어 박스",
    ],
    "전장_구매품": [
        "케이블", "전선", "전장품", "라미네이트", "정류기",
        "판넬", "전장", "배선", "결선",
    ],
    "전장_가공품": [
        "전장 가공", "전장부 가공", "판넬 가공",
    ],
    "설계": [
        "기구설계", "기계설계", "구조설계", "cad", "3d 설계", "2d 도면",
        "상세설계", "도면 작성", "기구 설계", "기계 설계", "설계 인건",
        "제어설계", "plc 프로그램", "hmi 프로그램", "소프트웨어 설계",
        "전장설계", "회로설계", "제어 로직", "프로그램 개발",
        "제어 설계", "sw 설계", "프로그래밍",
        "engineering", "설계",
    ],
    "조립": [
        "조립", "기계조립", "기구조립", "장비조립", "제작 인건비",
        "조립 공수", "조립 인건", "설치 인건",
    ],
    "셋업": [
        "셋업", "세팅", "시운전", "디버깅", "제어 튜닝",
        "plc 셋업", "hmi 셋업", "동작 테스트",
        "기구 셋업", "기계 셋업", "기구 조정", "얼라인", "정밀 조정",
        "기구 튜닝", "레벨링", "정렬",
        "현장 배선", "전장 셋업", "판넬 설치", "판넬 결선",
        "케이블 포설", "입출력 확인", "전장 공사",
    ],
    "양산대응": [
        "양산", "양산대응", "양산 대응", "양산인건", "양산 인건",
        "양산체계", "양산대기", "양산사원", "양산지원",
    ],
    "특수제어": [
        "umac", "pmac", "특수 제어", "특수제어",
    ],
    "포장": [
        "포장", "포장비", "패키징", "packaging",
    ],
    "기타항목": [
        "기타", "기타항목", "기타 비용",
    ],
    "경비": [
        "운송비", "출장비", "숙박비", "교통비", "안전관리비",
        "일반관리비", "보험료", "설치 경비", "현장 경비", "배송비",
        "운반비", "관리비", "경비", "이동", "국내이동",
    ],
}

# 우선순위: 명시적 분류 정보 → 최우선
EXPLICIT_CLASSIFICATION = {
    "기구_구매품": ["구매품"],
    "기구_가공품": ["가공품"],
}

# 품명에 플레이트/브라켓 계열이 있으면 구매품 우선 매칭 건너뜀
MECH_PART_GUARD_KWS = [
    "플레이트", "플라이트", "플레이트재", "브라켓", "브래킷", "블라켓", "블랙켓",
    "plt", "blk", "bkt", "plate", "bracket",
]


def _match_kw(kw: str, norm_text: str) -> bool:
    """키워드 매칭. 짧은 영문 약자는 단어 경계로 매칭."""
    nkw = normalize_text(kw)
    if not nkw:
        return False
    if nkw.isascii() and len(nkw) <= 5:
        return re.search(
            rf"(?<![a-z0-9]){re.escape(nkw)}(?![a-z0-9])", norm_text
        ) is not None
    return nkw in norm_text


class AutoClassifier:
    """
    자동 분류기 (템플릿 섹션 기준).
    품명, 규격, 비고를 종합해 템플릿 섹션으로 분류.
    """

    def __init__(self, user_rules: Optional[dict] = None):
        self.user_rules = user_rules or {}

    def classify(self, item: Dict[str, Any]) -> Tuple[str, float, List[str], str]:
        """
        항목을 템플릿 섹션으로 분류.
        
        Args:
            item: {"품명": ..., "규격": ..., "비고": ..., "구분": ...}
            
        Returns:
            (section_key, confidence, keywords, reason)
        """
        # 분석 대상 텍스트
        name = str(item.get("품명", ""))
        spec = str(item.get("규격", ""))
        remark = str(item.get("비고", ""))
        hint = str(item.get("구분", ""))
        
        combined = " ".join(t for t in [name, spec, remark, hint] if t)
        norm_combined = normalize_text(combined)

        if not norm_combined:
            return ("미분류", 0.0, [], "품명, 규격, 비고가 모두 비어 있습니다.")

        # 1) 사용자 규칙 우선 적용
        user_result = self._match_user_rules(norm_combined)
        if user_result:
            return user_result

        # 2) 명시적 분류 (구매품/가공품)
        for category, kws in EXPLICIT_CLASSIFICATION.items():
            matched = [kw for kw in kws if _match_kw(kw, norm_combined)]
            if matched:
                return (
                    category,
                    95.0,
                    matched,
                    f"'{', '.join(matched)}' 명시 → {category}",
                )

        # 3) 인건비/경비 여부 판별
        labor_name = normalize_text(f"{name} {remark}")
        is_labor_item = re.search(
            r"설계|조립|셋업|세팅|인건|시운전|정렬|얼라인|튜닝|레벨링|결선|배선|포설|설치|도면|프로그래밍|개발|테스트|관리비|운송|출장|숙박|배송|포장|보험|양산|umac|pmac",
            labor_name,
        ) is not None

        # 4) 일반 분류 매칭
        best_section = "미분류"
        best_score = 0
        best_keywords = []
        best_reason = ""

        for section_key in CLASSIFICATION_RULES:
            keywords = CLASSIFICATION_RULES[section_key]
            matched = [kw for kw in keywords if _match_kw(kw, norm_combined)]
            if matched:
                score = len(matched) * 10.0
                # 품명에 직접 매칭되면 보너스
                norm_name = normalize_text(name)
                if any(_match_kw(kw, norm_name) for kw in matched):
                    score += 15.0
                # 구분 열에 매칭되면 보너스
                norm_hint = normalize_text(hint)
                if norm_hint and any(_match_kw(kw, norm_hint) for kw in matched):
                    score += 10.0

                if score > best_score:
                    best_score = score
                    best_section = section_key
                    best_keywords = matched
                    best_reason = f"'{', '.join(matched)}' → {section_key}"

        # 5) 신뢰도 계산
        if best_section == "미분류":
            confidence = min(40.0, best_score)
            best_reason = "분류 근거 부족"
        else:
            norm_name = normalize_text(name)
            name_match = any(_match_kw(kw, norm_name) for kw in best_keywords)
            n_kw = len(best_keywords)

            if name_match and n_kw >= 1:
                confidence = min(95.0, 90.0 + n_kw * 2.0)
            elif n_kw >= 2:
                confidence = min(89.0, 75.0 + n_kw * 3.0)
            elif n_kw == 1:
                confidence = 75.0
            else:
                confidence = 55.0

        return (best_section, confidence, best_keywords, best_reason)

    def _match_user_rules(self, norm_text: str) -> Optional[tuple]:
        """사용자 정의 규칙 매칭"""
        for keyword, section in self.user_rules.items():
            if normalize_text(keyword) in norm_text:
                return (section, 95.0, [keyword], f"사용자 규칙 '{keyword}' 적용")
        return None

    def classify_all(self, items: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """모든 항목을 일괄 분류"""
        for item in items:
            section, confidence, keywords, reason = self.classify(item)
            item["_분류"] = section
            item["_분류_신뢰도"] = confidence
            item["_분류_키워드"] = keywords
            item["_분류_이유"] = reason
        
        # 합산(Unit) 행: 하위 항목의 분류를 상속
        self._inherit_unit_classification(items)
        
        return items
    
    def _inherit_unit_classification(self, items: List[Dict[str, Any]]) -> None:
        """
        합산(Unit) 행의 분류를 하위 항목의 분류로 상속.
        
        Unit 행은 품명이 비어 있어 분류가 어려우므로,
        바로 아래에 있는 세부 항목들(다음 Unit 행 전까지)의
        분류 중 가장 많은 것을 상속받는다.
        """
        n = len(items)
        for i, item in enumerate(items):
            if not item.get("_합산"):
                continue
            # 하위 항목 수집 (다음 Unit 행 또는 끝까지)
            sub_sections = []
            for j in range(i + 1, n):
                if items[j].get("_합산"):
                    break
                sec = items[j].get("_분류", "")
                if sec and sec != "미분류":
                    sub_sections.append(sec)
            if sub_sections:
                # 가장 많은 분류 선택
                from collections import Counter
                most_common = Counter(sub_sections).most_common(1)[0][0]
                item["_분류"] = most_common
                item["_분류_이유"] = f"Unit 합산 행 → 하위 항목 분류 상속: {most_common}"


def classify_items(
    items: List[Dict[str, Any]],
    user_rules: Optional[dict] = None,
) -> List[Dict[str, Any]]:
    """
    매핑된 항목들을 템플릿 섹션으로 분류.
    
    Args:
        items: RuleMapper.map_items() 결과
        user_rules: 사용자 정의 분류 규칙 {keyword: section_key}
        
    Returns:
        _분류, _분류_신뢰도, _분류_키워드, _분류_이유 필드 추가된 items
    """
    classifier = AutoClassifier(user_rules)
    return classifier.classify_all(items)
