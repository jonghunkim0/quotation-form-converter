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
        
        # 컬럼 유추: 헤더가 규칙과 안 맞으면 데이터 내용/레이아웃으로 유추
        self._infer_columns_from_data(items)
        # 단가 컬럼 자동 감지 (견적가 등 미지 컬럼 → 인접 총액 수치 관계로 검증)
        self._detect_price_column(items)
        
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

            # "(1) PLC", "(2) 전장" 같은 패턴: Unit명 추출 (기존 "unit명" 패턴 미매칭 시)
            if is_unit and not unit_name:
                m2 = re.match(r"\(\d+\)\s*(.+)", code.strip())
                if m2:
                    unit_name = m2.group(1).strip()
            
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
    
    def _detect_price_column(self, items: List[Dict[str, Any]]) -> None:
        """
        단가 컬럼 자동 감지 (데이터 기반 검증).
        
        헤더에 '단가' 계열 컬럼이 없는 경우, '견적가' 같은 미지 컬럼이
        실제로 단가인지 **인접한 총액(=수량×단가) 컬럼의 수치 관계**로
        검증한다. (예: 헤더 '견적가' 옆에 '총액'이 있고 총액=수량×견적가
        성립하면 견적가=단가로 인식)
        
        검증 조건:
          - 후보: 이름에 견적가/단가/price/cost 포함, 단가로 매칭되지 않는 컬럼
          - 총액: 이름에 총액/금액/amount/total 포함, 금액으로 매칭되는 컬럼
          - 수량: 수량/qty/quantity 컬럼
          - 샘플 200행 중 3행 이상, 80% 이상에서 총액 ≈ 수량×후보 (오차 1원/0.1%)
        """
        if not items:
            return
        headers = list(items[0].get("raw_data", {}).keys())
        
        # 이미 단가로 매칭되는 컬럼이 있으면 스킵
        for h in headers:
            std, _, _ = self._match_column(str(h))
            if std == "단가":
                return
        
        # 총액(금액) 컬럼 찾기
        total_kw = ("총액", "금액", "amount", "total")
        total_col = None
        for h in headers:
            h_str = str(h).strip().lower().replace(" ", "")
            if any(kw in h_str for kw in total_kw):
                std, _, _ = self._match_column(str(h))
                if std == "금액":
                    total_col = h
                    break
        # 수량 컬럼 찾기
        qty_col = None
        for h in headers:
            std, _, _ = self._match_column(str(h))
            if std == "수량":
                qty_col = h
                break
        if qty_col is None:
            return
        
        if total_col is None:
            # 총액 열 이름이 없으면 수치 관계로 단가+금액 열 쌍을 찾음
            unmapped_nums = []
            for h in headers:
                std, _, _ = self._match_column(str(h))
                if std is not None or std == "수량":
                    continue
                n_num = 0
                for item in items[:50]:
                    v = item.get("raw_data", {}).get(h)
                    if v is None:
                        continue
                    s = str(v).strip().replace(",", "")
                    if not s or s.lower() == "nan" or s == "-":
                        continue
                    try:
                        float(s)
                        n_num += 1
                    except ValueError:
                        break
                if n_num >= 3:
                    unmapped_nums.append(h)
            # 쌍 검증: qty × X ≈ Y
            for i, x in enumerate(unmapped_nums):
                for j, y in enumerate(unmapped_nums):
                    if i == j:
                        continue
                    verified = 0
                    checked = 0
                    for item in items[:200]:
                        raw = item.get("raw_data", {})
                        try:
                            qty = float(str(raw.get(qty_col, "")).replace(",", "").replace(" ", "") or 0)
                            xv = float(str(raw.get(x, "")).replace(",", "").replace(" ", "") or 0)
                            yv = float(str(raw.get(y, "")).replace(",", "").replace(" ", "") or 0)
                        except (ValueError, TypeError):
                            continue
                        if qty <= 0 or xv <= 0 or yv <= 0:
                            continue
                        checked += 1
                        if abs(qty * xv - yv) <= max(1.0, yv * 0.001):
                            verified += 1
                    if checked >= 3 and verified >= max(3, int(checked * 0.8)):
                        self.custom_rules[str(x).strip().lower()] = "단가"
                        self.custom_rules[str(y).strip().lower()] = "금액"
                        print(f"[RuleMapper] 수치 유추: '{x}' → 단가, '{y}' → 금액 (qty×{x}≈{y} {verified}/{checked}행)")
                        return
            return
        
        # 후보 컬럼: 견적가/단가/가격/price/cost 포함, 단가로 매칭되지 않음
        price_kw = ("견적가", "단가", "가격", "price", "cost")
        candidates = []
        for h in headers:
            h_str = str(h).strip().lower().replace(" ", "")
            if any(kw in h_str for kw in price_kw):
                std, _, _ = self._match_column(str(h))
                if std != "단가":
                    candidates.append(h)
        if not candidates:
            # 이름 후보 없음 → 미매핑 숫자 컬럼 전체를 후보로 (내용 기반)
            for h in headers:
                std, _, _ = self._match_column(str(h))
                if std is not None:
                    continue
                n_num = 0
                for item in items[:50]:
                    v = item.get("raw_data", {}).get(h)
                    if v is None:
                        continue
                    s = str(v).strip().replace(",", "")
                    if not s or s.lower() == "nan" or s == "-":
                        continue
                    try:
                        float(s)
                        n_num += 1
                    except ValueError:
                        break
                if n_num >= 3:
                    candidates.append(h)
        
        # 총액 컬럼에 가까운 후보 우선 (인접 열 우선)
        try:
            total_idx = headers.index(total_col)
            candidates.sort(key=lambda h: abs(headers.index(h) - total_idx))
        except ValueError:
            pass
        
        # 수치 검증: 총액 ≈ 수량 × 후보값
        for cand in candidates:
            verified = 0
            checked = 0
            for item in items[:200]:
                raw = item.get("raw_data", {})
                try:
                    qty = float(str(raw.get(qty_col, "")).replace(",", "").replace(" ", "") or 0)
                    cand_v = float(str(raw.get(cand, "")).replace(",", "").replace(" ", "") or 0)
                    tot_v = float(str(raw.get(total_col, "")).replace(",", "").replace(" ", "") or 0)
                except (ValueError, TypeError):
                    continue
                if qty <= 0 or cand_v <= 0 or tot_v <= 0:
                    continue
                checked += 1
                if abs(qty * cand_v - tot_v) <= max(1.0, tot_v * 0.001):
                    verified += 1
            if checked >= 3 and verified >= max(3, int(checked * 0.8)):
                self.custom_rules[str(cand).strip().lower()] = "단가"
                print(f"[RuleMapper] 단가 컬럼 자동 감지: '{cand}' → 단가 (총액 '{total_col}' 수치 검증 {verified}/{checked}행)")
                return
    
    # ── 내용 기반 유추: 알려진 단위 어휘 ──
    _KNOWN_UNITS = {
        "ea", "ea.", "pcs", "pc", "set", "set.", "md", "kg", "g", "l", "ml",
        "m", "m2", "m3", "mm", "cm", "box", "roll", "sheet", "pair", "lot",
        "개", "개수", "대", "식", "인", "인원", "일", "주", "월", "년", "회",
        "人", "泊", "回", "日", "台", "セット", "個", "本", "枚", "点", "式",
    }
    _CODE_RE = re.compile(r"^(MECH|CTRL|ELEC|ITEM|PART|PRT)-\d{2,}$|^\d{1,4}[.)]$")
    
    def _infer_columns_from_data(self, items: List[Dict[str, Any]]) -> None:
        """
        데이터 내용 기반 컬럼 유추.
        
        업체마다 헤더 표기가 달라 규칙과 안 맞는 경우, 실제 데이터 값과
        열 위치(레이아웃)를 보고 표준 컬럼을 유추한다:
          1) 단위: 알려진 단위 어휘 (EA, SET, MD, 개, 대, ...)
          2) 품목코드: "N)" 또는 "MECH-0001" 패턴
          3) 수량: 정수·소값 컬럼
          4) 레이아웃: 수량 열의 바로 앞=Maker, 그 앞=규격, 그 앞=품명
             (견적서 표준 레이아웃: 품명 | 규격 | Maker | 수량 | 단위 | 단가 | 금액)
        """
        if not items:
            return
        headers = list(items[0].get("raw_data", {}).keys())
        assigned = {}
        for h in headers:
            std, _, _ = self._match_column(str(h))
            if std:
                assigned[h] = std
        
        def col_values(h):
            vals = []
            for it in items[:300]:
                v = it.get("raw_data", {}).get(h)
                if v is None:
                    continue
                s = str(v).strip()
                if s and s.lower() != "nan" and s != "-":
                    vals.append(s)
            return vals
        
        def register(h, std, reason):
            self.custom_rules[str(h).strip().lower()] = std
            assigned[h] = std
            print(f"[RuleMapper] 내용 유추: '{h}' → {std} ({reason})")
        
        # 0) 데이터 기반 교정: 품목코드로 매칭됐지만 값이 실제 명칭이면 품명으로
        for h in list(assigned.keys()):
            if assigned[h] != "품목코드":
                continue
            vals = col_values(h)
            if not vals:
                continue
            code_hits = sum(1 for v in vals if self._CODE_RE.match(v))
            if code_hits < max(2, int(len(vals) * 0.5)):
                del assigned[h]
                if "품명" not in assigned.values():
                    register(h, "품명", "코드 오매칭 교정")
        
        # 1) 단위: 알려진 단위 어휘
        for h in headers:
            if h in assigned:
                continue
            vals = col_values(h)
            if len(vals) < 3:
                continue
            hits = sum(1 for v in vals if v.lower().replace(" ", "") in self._KNOWN_UNITS)
            if hits >= max(3, int(len(vals) * 0.8)):
                register(h, "단위", f"단위 어휘 {hits}/{len(vals)}")
        
        # 2) 품목코드: "N)" 또는 "CODE-0001" 패턴
        for h in headers:
            if h in assigned:
                continue
            vals = col_values(h)
            if len(vals) < 3:
                continue
            hits = sum(1 for v in vals if self._CODE_RE.match(v))
            if hits >= max(3, int(len(vals) * 0.8)):
                register(h, "품목코드", f"코드 패턴 {hits}/{len(vals)}")
        
        # 3) 수량: 정수·소값 컬럼 (레이아웃상 단가/금액보다 앞)
        if "수량" not in assigned.values():
            for h in headers:
                if h in assigned:
                    continue
                vals = col_values(h)
                if len(vals) < 3:
                    continue
                nums = []
                ok = True
                for v in vals:
                    try:
                        nums.append(float(v.replace(",", "")))
                    except ValueError:
                        ok = False
                        break
                if ok and all(n == int(n) for n in nums) and max(nums) <= 100000:
                    register(h, "수량", f"정수 {len(vals)}행")
                    break
        
        # 4) 레이아웃 유추: 수량 ← Maker ← 규격 ← 품명 (왼쪽에서 오른쪽)
        qty_h = next((h for h, s in assigned.items() if s == "수량"), None)
        if qty_h is not None and "Maker" not in assigned.values():
            qi = headers.index(qty_h)
            if qi > 0:
                h = headers[qi - 1]
                if h not in assigned:
                    vals = col_values(h)
                    if vals and all(len(v) <= 20 and re.search(r"[A-Za-z가-힣]", v) for v in vals):
                        register(h, "Maker", "수량 열 직전")
        maker_h = next((h for h, s in assigned.items() if s == "Maker"), None)
        if maker_h is not None and "규격" not in assigned.values():
            mi = headers.index(maker_h)
            if mi > 0:
                h = headers[mi - 1]
                if h not in assigned:
                    vals = col_values(h)
                    if vals and all(re.search(r"[A-Za-z가-힣0-9]", v) for v in vals):
                        register(h, "규격", "Maker 열 직전")
        spec_h = next((h for h, s in assigned.items() if s == "규격"), None)
        if spec_h is not None and "품명" not in assigned.values():
            si = headers.index(spec_h)
            if si > 0:
                h = headers[si - 1]
                if h not in assigned:
                    vals = col_values(h)
                    if vals and all(re.search(r"[A-Za-z가-힣]", v) for v in vals):
                        register(h, "품명", "규격 열 직전")
    
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
        #    1글자 헤더(A, B, ...)는 부분일치 금지 ("a" in "part no" 등 오매칭 방지)
        if len(orig_col_nospace) >= 2:
            for rule_key, rule_target in COLUMN_MAPPING_RULES.items():
                rk_nospace = rule_key.replace(" ", "")
                if rule_key in orig_col or orig_col in rule_key \
                   or rk_nospace in orig_col_nospace or orig_col_nospace in rk_nospace:
                    # 길이가 길수록 더 정확한 매칭
                    confidence = 0.7 + 0.2 * (len(rule_key) / max(len(orig_col), len(rule_key)))
                    return rule_target, min(confidence, 0.9), f"부분 일치: '{orig_col}' ~ '{rule_key}'"

        # 3.5) 퍼지 매칭 (단어 순서 차이, 오타 등)
        import difflib
        best_ratio = 0.0
        best_key = None
        for rule_key in COLUMN_MAPPING_RULES:
            ratio = difflib.SequenceMatcher(
                None, orig_col_nospace, rule_key.replace(" ", "")).ratio()
            if ratio > best_ratio:
                best_ratio = ratio
                best_key = rule_key
        if best_ratio >= 0.85:
            return COLUMN_MAPPING_RULES[best_key], 0.75, \
                f"퍼지 매칭: '{orig_col}' ~ '{best_key}'"
        
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
