"""
규칙 기반 컬럼 매핑 및 데이터 변환 로직
"""

import re
from typing import List, Dict, Any, Optional, Tuple
from config import (
    STANDARD_COLUMNS,
    COLUMN_MAPPING_RULES,
    UNIT_NORMALIZATION,
    CURRENCY_SYMBOLS,
    CONFIDENCE_THRESHOLD,
)


class RuleMapper:
    """
    업체별 견적 항목을 표준 컬럼으로 매핑하고 데이터를 정규화.
    """
    
    def __init__(self, custom_rules: Optional[Dict[str, str]] = None):
        """
        Args:
            custom_rules: 사용자 정의 매핑 규칙
                          {"원본컬럼명": "표준컬럼명", "~키워드": "표준컬럼명"}
        """
        self.custom_rules = custom_rules or {}
        self.mapping_log: List[Dict[str, Any]] = []
    
    def map_items(self, items: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """
        파싱된 항목 목록을 표준 형식으로 변환.
        
        Args:
            items: [{"raw_data": {...}, "sheet_name"/"page": ..., "row_index": int}, ...]
            
        Returns:
            표준 컬럼을 가진 항목 목록
        """
        mapped_items = []
        
        for item in items:
            raw_data = item.get("raw_data", {})
            mapped = self._map_single_item(raw_data, item)
            mapped_items.append(mapped)
        
        # 합산(Unit) 행 감지
        self._detect_unit_rows(mapped_items)
        
        # 구분(계열) forward-fill: 병합 셀이라 섹션 첫 행에만 값이 있음
        self._forward_fill_gubun(mapped_items)
        
        # 중복 검출
        self._detect_duplicates(mapped_items)
        
        return mapped_items
    
    def _forward_fill_gubun(self, mapped_items: List[Dict[str, Any]]) -> None:
        """
        구분(계열) 컬럼 forward-fill.
        
        Raw 견적의 '구분' 컬럼은 병합 셀이라 섹션 첫 행에만 값이 있고
        나머지 행은 빈 값. 아래 행으로 전파하여 모든 항목이
        소속 계열(기구부/전장/제어 등)을 알 수 있게 한다.
        """
        last_gubun = ""
        for item in mapped_items:
            if item.get("_구분"):
                last_gubun = item["_구분"]
            elif last_gubun:
                item["_구분"] = last_gubun

    def _detect_unit_rows(self, mapped_items: List[Dict[str, Any]]) -> None:
        """
        합산(Unit) 행 감지 및 Unit명 추출.
        
        원본에서 '(1) Unit명 : BASE_FRAME' 같은 행은 아래 세부항목들의
        합산으로 금액이 산출되는 요약 행이다. No(품목코드)에 'unit명'이
        포함되거나 단위가 '식'인 행으로 판별한다.
        
        각 항목에 다음 필드를 추가:
          _합산: bool  (True이면 합산/Unit 행)
          _unit_name: str (Unit명, 예: "BASE_FRAME")
        """
        for item in mapped_items:
            is_unit = False
            unit_name = ""
            code = str(item.get("품목코드", ""))
            code_nospace = code.lower().replace(" ", "")
            
            if "unit명" in code_nospace:
                is_unit = True
                m = re.search(r"unit명\s*[:：]?\s*(.+)", code, re.IGNORECASE)
                if m:
                    unit_name = m.group(1).strip()
            
            if str(item.get("단위", "")).strip() == "식":
                is_unit = True
            
            item["_합산"] = is_unit
            item["_unit_name"] = unit_name
            
            # 원본 소계/합계 행 감지 (템플릿에 소계 행이 이미 있으므로 제외)
            # 품명 AND 품목코드(No) 둘 다 확인 (소계 텍스트가 No 컬럼에 있는 경우)
            check_text = (str(item.get("품명", "")) + str(item.get("품목코드", ""))).replace(" ", "")
            item["_소계"] = ("소계" in check_text) or ("합계" in check_text)
    
    def _map_single_item(self, raw_data: Dict[str, Any], meta: Dict) -> Dict[str, Any]:
        """단일 항목을 표준 컬럼으로 매핑"""
        result = {col: "" for col in STANDARD_COLUMNS}
        result["검증결과"] = ""
        result["_원본명"] = ""  # 원본 컬럼명 (옵션 표시용)
        result["_신뢰도"] = 1.0
        result["_매핑근거"] = ""
        
        # 원본 컬럼명 → 표준 컬럼명 매핑
        column_mapping = {}  # {표준컬럼: (원본컬럼, 신뢰도)}
        
        for orig_col, value in raw_data.items():
            orig_col_clean = str(orig_col).strip().lower()
            standard_col, confidence, reason = self._match_column(orig_col_clean)
            
            if standard_col and standard_col in STANDARD_COLUMNS:
                # 이미 매핑된 컬럼이면 더 높은 신뢰도 사용
                if standard_col in column_mapping:
                    prev_conf = column_mapping[standard_col][1]
                    if confidence > prev_conf:
                        column_mapping[standard_col] = (orig_col, confidence)
                else:
                    column_mapping[standard_col] = (orig_col, confidence)
        
        # 매핑된 값 채우기
        for std_col, (orig_col, confidence) in column_mapping.items():
            value = raw_data.get(orig_col, "")
            result[std_col] = self._normalize_value(std_col, value)
            
            # 원본명 기록
            if not result["_원본명"]:
                result["_원본명"] = orig_col
            else:
                result["_원본명"] += f", {orig_col}"
            
            # 신뢰도 (최소값)
            result["_신뢰도"] = min(result["_신뢰도"], confidence)
            
            # 매핑 근거
            if not result["_매핑근거"]:
                result["_매핑근거"] = f"{orig_col} → {std_col}"
            else:
                result["_매핑근거"] += f"; {orig_col} → {std_col}"
        
        # 금액 검증: 금액 = 수량 × 단가
        self._validate_amount(result)
        
        # 누락 데이터 플래그
        self._flag_missing(result)
        
        # 메타데이터 추가
        result["_출처"] = meta.get("sheet_name") or f"p.{meta.get('page', '?')}"
        result["_행"] = meta.get("row_index", 0)
        
        # 구분 컬럼 캡처 (Raw 견적의 계열 정보: 기구부/전장/제어 등)
        # 병합 셀이라 첫 행에만 값이 있음 → map_items에서 forward-fill
        for key in ("구분", "구분 ", "분류", "계열"):
            if key in raw_data and raw_data[key] is not None:
                val = str(raw_data[key]).strip()
                if val and val != "nan":
                    result["_구분"] = val
                    break
        
        return result
    
    def _match_column(self, orig_col: str) -> Tuple[Optional[str], float, str]:
        """
        원본 컬럼명을 표준 컬럼으로 매칭.
        
        Returns:
            (표준컬럼명, 신뢰도, 매칭근거)
        """
        orig_col = orig_col.strip().lower()
        # 공백 제거 버전 (예: "품 명" → "품명") - 헤더에 공백이 있는 경우 대응
        orig_col_nospace = orig_col.replace(" ", "")

        # 1. 사용자 정의 규칙 우선
        for rule_key, rule_target in self.custom_rules.items():
            if rule_key.startswith("~"):
                # 키워드 매칭
                keyword = rule_key[1:].lower()
                if keyword in orig_col:
                    return rule_target, 0.9, f"사용자 키워드 규칙: '{keyword}'"
            else:
                if rule_key.lower() == orig_col:
                    return rule_target, 0.95, f"사용자 정의 규칙: '{rule_key}'"
        
        # 2. 기본 매핑 규칙 (완전 일치) - 공백 제거 버전도 시도
        if orig_col in COLUMN_MAPPING_RULES:
            return COLUMN_MAPPING_RULES[orig_col], 0.95, f"완전 일치: '{orig_col}'"
        if orig_col_nospace in COLUMN_MAPPING_RULES:
            return COLUMN_MAPPING_RULES[orig_col_nospace], 0.95, f"완전 일치(공백제거): '{orig_col}'"

        # 3. 부분 일치 (원본 컬럼명이 규칙 키를 포함) - 공백 제거 버전도 시도
        for rule_key, rule_target in COLUMN_MAPPING_RULES.items():
            rk_nospace = rule_key.replace(" ", "")
            if rule_key in orig_col or orig_col in rule_key \
               or rk_nospace in orig_col_nospace or orig_col_nospace in rk_nospace:
                # 길이가 길수록 더 정확한 매칭
                confidence = 0.7 + 0.2 * (len(rule_key) / max(len(orig_col), len(rule_key)))
                return rule_target, min(confidence, 0.9), f"부분 일치: '{orig_col}' ~ '{rule_key}'"

        # 4. 미매칭
        return None, 0.0, "매칭 실패"
    
    def _normalize_value(self, std_col: str, value: Any) -> Any:
        """컬럼별 값 정규화"""
        if value is None or (isinstance(value, float) and value != value):  # NaN check
            return ""
        
        value_str = str(value).strip()
        
        if std_col == "수량":
            return self._parse_number(value_str)
        
        elif std_col == "단가":
            return self._parse_amount(value_str)
        
        elif std_col == "금액":
            return self._parse_amount(value_str)
        
        elif std_col == "단위":
            return self._normalize_unit(value_str)
        
        elif std_col == "견적일":
            return self._normalize_date(value_str)
        
        else:
            return value_str
    
    def _parse_number(self, value: str) -> Any:
        """숫자 파싱 (수량)"""
        if not value:
            return ""
        # 쉼표, 공백 제거
        cleaned = value.replace(",", "").replace(" ", "")
        try:
            num = float(cleaned)
            return int(num) if num == int(num) else num
        except ValueError:
            return value  # 숫자가 아니면 원본 유지
    
    def _parse_amount(self, value: str) -> Any:
        """금액 파싱 (통화 기호, 쉼표 제거)"""
        if not value:
            return ""
        # 통화 기호 제거
        cleaned = value
        for symbol in CURRENCY_SYMBOLS:
            cleaned = cleaned.replace(symbol, "")
        # 쉼표, 공백, "원", "USD" 등 제거
        cleaned = re.sub(r"[,\s원USDJPYEURGBP]", "", cleaned)
        try:
            return float(cleaned)
        except ValueError:
            return value  # 숫자가 아니면 원본 유지
    
    def _normalize_unit(self, value: str) -> str:
        """단위 정규화"""
        if not value:
            return ""
        value_lower = value.strip().lower()
        # 정확히 일치
        if value_lower in UNIT_NORMALIZATION:
            return UNIT_NORMALIZATION[value_lower]
        # 대소문자 무시
        for key, std in UNIT_NORMALIZATION.items():
            if key.lower() == value_lower:
                return std
        return value.strip()
    
    def _normalize_date(self, value: str) -> str:
        """날짜 정규화 (YYYY-MM-DD)"""
        if not value:
            return ""
        # 다양한 날짜 형식 처리
        patterns = [
            (r"(\d{4})[.\-/](\d{1,2})[.\-/](\d{1,2})", r"\1-\2-\3"),
            (r"(\d{4})년\s*(\d{1,2})월\s*(\d{1,2})일", r"\1-\2-\3"),
            (r"(\d{1,2})[.\-/](\d{1,2})[.\-/](\d{4})", r"\3-\1-\2"),
        ]
        for pattern, replacement in patterns:
            match = re.search(pattern, value)
            if match:
                result = re.sub(pattern, replacement, value)
                # 0 패딩
                parts = result.split("-")
                if len(parts) == 3:
                    return f"{parts[0]}-{int(parts[1]):02d}-{int(parts[2]):02d}"
        return value
    
    def _validate_amount(self, result: Dict[str, Any]) -> None:
        """금액 = 수량 × 단가 검증"""
        qty = result.get("수량", "")
        price = result.get("단가", "")
        amount = result.get("금액", "")
        
        if not isinstance(qty, (int, float)) or not isinstance(price, (int, float)):
            return
        if not isinstance(amount, (int, float)):
            return
        
        expected = qty * price
        # 1원 이내 오차 허용
        if abs(expected - amount) > 1:
            result["검증결과"] = "금액 검증 오류"
            result["_검증상세"] = f"수량({qty}) × 단가({price}) = {expected} ≠ 금액({amount})"
        else:
            if not result["검증결과"]:
                result["검증결과"] = "OK"
    
    def _flag_missing(self, result: Dict[str, Any]) -> None:
        """누락 데이터 플래그"""
        missing = []
        for col in STANDARD_COLUMNS:
            if col == "검증결과":
                continue
            if result.get(col, "") == "" or result.get(col) is None:
                missing.append(col)
        
        if missing:
            flag = f"확인 필요: {', '.join(missing)}"
            if result["검증결과"]:
                result["검증결과"] += f" | {flag}"
            else:
                result["검증결과"] = flag
    
    def _detect_duplicates(self, items: List[Dict[str, Any]]) -> None:
        """중복 품목 검출"""
        seen = {}  # {품목코드: [index, ...]}
        
        for i, item in enumerate(items):
            code = str(item.get("품목코드", "")).strip()
            name = str(item.get("품명", "")).strip()
            key = code if code else name
            
            if not key:
                continue
            
            if key in seen:
                # 중복 표시
                for idx in seen[key]:
                    if "중복" not in items[idx]["검증결과"]:
                        items[idx]["검증결과"] = (items[idx]["검증결과"] + " | " if items[idx]["검증결과"] else "") + "중복"
                if "중복" not in item["검증결과"]:
                    item["검증결과"] = (item["검증결과"] + " | " if item["검증결과"] else "") + "중복"
                seen[key].append(i)
            else:
                seen[key] = [i]
    
    def get_mapping_summary(self) -> List[Dict[str, Any]]:
        """매핑 로그 요약 반환"""
        return self.mapping_log
