"""
converters/standard_price.py — 구매품 사내 표준단가 검증·반영

기구_구매품 / 제어_구매품 / 전장_구매품(사급품 포함)을 추출해
사내 시스템 표준단가 확인 요청용 Excel을 만들고, 사용자가 다시 업로드한
표준단가 확인 결과를 매칭·반영한다.

템플릿: templates/표준단가_요청템플릿.xlsx (견적Item_Upload_Template_v2.0)
  열: 구분(A,드롭다운) / UNIT(B) / TYPE(C) / No(D) / 품명(E) / 규격(F)
      / Maker(G) / 수량(H) / 단위(I) / 개발단가(J) / Comment(K)
  - 식별번호 열이 없어 Comment(K)에 "MECH-0001 / 사급" 형식으로 기록
  - 표준단가는 개발단가(J) 열에 기록 (결과 파일)
"""
from __future__ import annotations

import re
from io import BytesIO
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from openpyxl import load_workbook
from openpyxl.worksheet.datavalidation import DataValidation

# ─── 상수 ─────────────────────────────────────────────────────
# _분류 → (식별번호 접두, 템플릿 A열 구분값, UNIT 기본값)
_CAT_META = {
    "기구_구매품": ("MECH", "기구부", "Frame Unit"),
    "제어_구매품": ("CTRL", "전장/제어부", "Control Unit"),
    "전장_구매품": ("ELEC", "전장/제어부", "Control Unit"),
}
# 사급 판별 키워드 (원본 비고 기준)
_SUPPLY_KW = ["사급", "고객지급", "발주처지급", "무상사급", "유상사급"]
# 미반영 주석 문구 (스펙 고정)
_MISS_NOTE = "표준단가 미반영-확인 필요"

# 결과 파일 열 이름 → 표준 필드 (의미 동일 열 자동 대응)
_RESULT_COL_ALIASES = {
    "id": ["식별번호", "요청품목식별번호", "요청 품목 식별번호", "item id", "id", "no"],
    "품명": ["품명", "자재명", "item name", "name", "품목명"],
    "Maker": ["maker", "제조사", "공급사", "maker name", "brand"],
    "모델명": ["모델명", "model", "model no", "model no."],
    "규격": ["규격", "spec", "specification", "사양"],
    "단위": ["단위", "uom"],
    # 재업로드 결과 파일의 '견적기준가'(M열)를 최우선으로 반영:
    # 사용자가 사내 시스템에서 확인한 견적기준가를 해당 구매품 단가로 적용.
    # 주의: 실제 헤더는 '견적기준가'(견적+기준+가)이며, '기준견적가'와 단어 순서가
    # 다르다. 둘 다 alias에 포함하되 '견적기준가'를 최우선으로 둬야 '개발단가'(J열,
    # 0원)가 먼저 매핑되는 것을 막을 수 있다.
    "표준단가": ["견적기준가", "기준견적가", "표준단가", "개발단가", "개발 단가", "standard price", "std price", "단가"],
    "통화": ["통화", "currency"],
    "기준일": ["단가 기준일", "기준일", "price date", "date"],
    "적용단위": ["단가 적용 단위", "적용단위", "price unit"],
    "조회결과": ["조회 결과", "조회결과", "result", "status"],
    "비고": ["비고", "comment", "remark", "memo"],
}


# ─── 유틸 ─────────────────────────────────────────────────────
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
    except ValueError:
        return 0.0


def _norm(s) -> str:
    """정규화: 소문자, 공백/구분기호 제거 (대소문자·공백·구분기호 차이 무시)"""
    if s is None:
        return ""
    s = str(s).lower()
    s = re.sub(r"[\s\-_/\\.,()（）\[\]【】]+", "", s)
    return s


def _is_supply(item: Dict[str, Any]) -> bool:
    b = str(item.get("비고", "") or "")
    return any(k in b for k in _SUPPLY_KW)


# ─── 1단계: 대상 구매품 추출 ──────────────────────────────────
def extract_purchase_items(mapped_items: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """기구/제어/전장_구매품(사급품 포함)을 추출해 식별번호를 부여한다.

    - Unit(합산) 행은 제외 (세부 품목만)
    - 원본 행 번호(_행/_출처)를 보존해 재매칭에 사용
    - 식별번호: MECH-####(기구) / CTRL-####(제어) / ELEC-####(전장), 구분별 순번
    """
    items: List[Dict[str, Any]] = []
    counters: Dict[str, int] = {k: 0 for k in _CAT_META}
    for m in mapped_items or []:
        cat = m.get("_분류", "")
        if cat not in _CAT_META:
            continue
        if m.get("_합산"):
            continue  # Unit(합산) 행 제외
        prefix, gubun, unit = _CAT_META[cat]
        counters[cat] += 1
        pid = f"{prefix}-{counters[cat]:04d}"
        supply = "사급" if _is_supply(m) else "비사급"
        items.append({
            "id": pid,
            "구분": gubun,
            "cat": cat,
            "UNIT": unit,
            "품명": str(m.get("품명", "") or "").strip(),
            "Maker": str(m.get("Maker", "") or "").strip(),
            "규격": str(m.get("규격", "") or "").strip(),
            "수량": _num(m.get("수량")),
            "단위": str(m.get("단위", "") or "").strip(),
            "기존단가": _num(m.get("단가")),
            "사급": supply,
            "_row": m.get("_행"),
            "_출처": m.get("_출처", ""),
        })
    return items


def _supply_total(items: List[Dict[str, Any]], price_key: str = "기존단가") -> float:
    """사급품 합계 (수량 × 단가)."""
    return sum(it["수량"] * _num(it.get(price_key)) for it in items if it.get("사급") == "사급")


# ─── 2단계: 요청용 Excel 작성 ─────────────────────────────────
def build_request_excel(items: List[Dict[str, Any]], template_path: str) -> bytes:
    """요청 템플릿에 구매품을 채워 사내 시스템 단가 확인 요청용 Excel을 만든다.

    - 헤더(1행)·서식·드롭다운 유지
    - A열 드롭다운 원본(N열)에 '기구부' 보강 (템플릿에 기구부 누락)
    - 샘플 데이터 행(2~3) 제거 후 데이터 입력
    - 식별번호+사급여부를 Comment(K)에 기록 (재매칭용)
    - 개발단가(J)는 공란 (사용자가 사내 시스템에서 채움)
    """
    wb = load_workbook(template_path)
    ws = wb.active

    # 1) A열 드롭다운 원본(N열) 값 수집 (샘플 행 삭제 전)
    n_values: List[str] = []
    r = 2
    while True:
        v = ws.cell(r, 14).value
        if v is None:
            break
        n_values.append(str(v).strip())
        r += 1

    # 2) 기존 데이터 행(2행~) 지우기 (샘플 포함)
    if ws.max_row >= 2:
        ws.delete_rows(2, ws.max_row - 1)

    # 3) 드롭다운 원본(N열) 재작성: 기구부/전장/제어부/인건비 보강
    _ensure_dropdown(ws, n_values)

    # 4) 데이터 입력 (No는 구분별 순번, 템플릿과 동일하게 "1)" 형식)
    no_counter: Dict[str, int] = {}
    r = 2
    for it in items:
        no_counter[it["구분"]] = no_counter.get(it["구분"], 0) + 1
        ws.cell(r, 1, it["구분"])                       # A 구분
        ws.cell(r, 2, it["UNIT"])                       # B UNIT
        ws.cell(r, 3, "구매품")                          # C TYPE
        gubun = it['구분']
        ws.cell(r, 4, str(no_counter[gubun]) + ')')  # D No
        ws.cell(r, 5, it["품명"])                        # E 품명
        ws.cell(r, 6, it["규격"])                        # F 규격
        ws.cell(r, 7, it["Maker"])                       # G Maker
        ws.cell(r, 8, it["수량"])                         # H 수량
        ws.cell(r, 9, it["단위"])                         # I 단위
        # J 개발단가: 공란 (사내 시스템에서 채움)
        ws.cell(r, 11, f"{it['id']} / {it['사급']}")     # K Comment (식별번호+사급)
        r += 1

    buf = BytesIO()
    wb.save(buf)
    wb.close()
    return buf.getvalue()


def _ensure_dropdown(ws, existing: List[str]):
    """A열 드롭다운 원본(N열)에 필요한 구분값(기구부/전장/제어부/인건비)을 보강한다.

    existing: 샘플 행 삭제 전 N열에 있던 값 목록.
    """
    needed = ["기구부", "전장/제어부", "인건비"]
    values = list(existing)
    for n in needed:
        if n not in values:
            values.append(n)
    # N열(14) 재작성
    for i, v in enumerate(values):
        ws.cell(2 + i, 14, v)
    last = 1 + len(values)
    # DV formula1을 $N$2:$N$<last>로 갱신
    for dv in (ws.data_validations.dataValidation or []):
        if str(dv.sqref).startswith("A"):
            dv.formula1 = f"$N$2:$N${last}"


# ─── 3단계: 표준단가 결과 파일 파싱 ───────────────────────────
def parse_result_file(xlsx_bytes: bytes) -> List[Dict[str, Any]]:
    """표준단가 확인 결과 Excel을 파싱해 행 목록을 반환한다.

    각 행: {id, 품명, Maker, 모델명, 규격, 단위, 표준단가, 통화, 기준일, 적용단위, 조회결과, 비고}
    - 열 이름은 의미 동일이면 대응 (대소문자/공백 무시)
    - 식별번호는 Comment(K)의 첫 토큰 또는 전용 열에서 추출
    """
    wb = load_workbook(BytesIO(xlsx_bytes), data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []

    # 헤더 행 탐색 (첫 10행 중 '품명'/'단가' 포함 행)
    header_idx = 0
    for i, row in enumerate(rows[:10]):
        cells = [str(c).strip() if c is not None else "" for c in row]
        joined = " ".join(cells)
        if ("품명" in joined or "단가" in joined or "식별번호" in joined):
            header_idx = i
            break
    headers = [str(c).strip() if c is not None else "" for c in rows[header_idx]]

    # 헤더 → 표준 필드 매핑 (완전 일치 우선, 없으면 부분 일치)
    col_map: Dict[str, int] = {}
    norm_headers = [(_norm(h), idx) for idx, h in enumerate(headers)]
    for field, aliases in _RESULT_COL_ALIASES.items():
        for a in aliases:
            na = _norm(a)
            for hn, idx in norm_headers:
                if hn == na and field not in col_map:
                    col_map[field] = idx
                    break
            if field in col_map:
                break
    # 부분 일치 보강 (완전 일치로 못 맞춘 필드만)
    for field, aliases in _RESULT_COL_ALIASES.items():
        if field in col_map:
            continue
        for a in aliases:
            na = _norm(a)
            for hn, idx in norm_headers:
                if na and na in hn and field not in col_map:
                    col_map[field] = idx
                    break
            if field in col_map:
                break

    results: List[Dict[str, Any]] = []
    for row in rows[header_idx + 1:]:
        if row is None:
            continue
        def _get(field):
            idx = col_map.get(field)
            if idx is None or idx >= len(row):
                return None
            return row[idx]

        rec = {
            "id": _extract_id(_get("id"), _get("비고"), _get("Comment")),
            "품명": _get("품명"),
            "Maker": _get("Maker"),
            "모델명": _get("모델명"),
            "규격": _get("규격"),
            "단위": _get("단위"),
            "표준단가": _get("표준단가"),
            "통화": _get("통화"),
            "기준일": _get("기준일"),
            "적용단위": _get("적용단위"),
            "조회결과": _get("조회결과"),
            "비고": _get("비고"),
        }
        # 빈 행 제외 (품명/식별번호/표준단가 모두 없으면)
        if rec["id"] is None and rec["품명"] is None and rec["표준단가"] is None:
            continue
        results.append(rec)
    wb.close()
    return results


def _extract_id(id_val, comment_val, comment2_val):
    """식별번호 추출: 전용 열 우선, 없으면 Comment의 첫 토큰(MECH-0001 등)."""
    for v in (id_val, comment_val, comment2_val):
        if v is None:
            continue
        s = str(v).strip()
        m = re.search(r"(MECH|CTRL|ELEC)-\d{4}", s)
        if m:
            return m.group(0)
    return None


# ─── 4단계: 매칭 + 반영 ──────────────────────────────────────
def match_and_apply(items: List[Dict[str, Any]],
                    results: List[Dict[str, Any]]) -> Dict[str, Any]:
    """요청 품목(items)과 표준단가 결과(results)를 매칭해 반영 결과를 만든다.

    Returns:
        {
          "applied": [...],      # 표준단가 반영 완료
          "unchanged": [...],    # 확인 완료-단가 변경 없음
          "missed": [...],       # 표준단가 미반영
          "match_check": [...],  # 매칭 확인 필요
          "extra": [...],        # 원본에 없는 추가 확인 품목
          "summary": {...},
        }
    """
    # 결과 행을 id / 정규화 키로 인덱싱
    by_id: Dict[str, List[Dict[str, Any]]] = {}
    by_key: Dict[str, List[Dict[str, Any]]] = {}
    for rec in results:
        if rec.get("id"):
            by_id.setdefault(str(rec["id"]).strip(), []).append(rec)
        key = _match_key(rec.get("품명"), rec.get("Maker"), rec.get("모델명"), rec.get("규격"))
        if key:
            by_key.setdefault(key, []).append(rec)

    applied, unchanged, missed, match_check = [], [], [], []
    used_result_idx = set()

    for it in items:
        candidates = by_id.get(it["id"], [])
        match_type = "id"
        if not candidates:
            key = _match_key(it["품명"], it["Maker"], None, it["규격"])
            candidates = by_key.get(key, [])
            match_type = "key"
        # 모델명 없는 경우 품명+Maker+규격 보조 매칭
        if not candidates and match_type == "id":
            key2 = _match_key(it["품명"], it["Maker"], None, it["규격"])
            if key2:
                candidates = by_key.get(key2, [])
                match_type = "key"

        if not candidates:
            _push_missed(missed, it, "표준단가 조회 결과 없음")
            continue

        if len(candidates) > 1:
            # 복수 후보 → 자동 반영 금지
            _push_match_check(match_check, it, candidates, "표준단가 결과 복수 후보")
            continue

        rec = candidates[0]
        price = _parse_price(rec.get("표준단가"))
        if price is None:
            reason = _miss_reason(rec)
            _push_missed(missed, it, reason)
            continue

        old = it["기존단가"]
        new_amount = it["수량"] * price
        old_amount = it["수량"] * old
        row = dict(it)
        row["표준단가"] = price
        row["변경후금액"] = new_amount
        row["기준일"] = rec.get("기준일")
        row["통화"] = rec.get("통화")
        row["적용단위"] = rec.get("적용단위")
        row["단가차이"] = price - old
        if old > 0:
            row["증감률"] = (price - old) / old * 100.0
        else:
            row["증감률"] = None  # 비교 불가
        row["매칭"] = match_type
        if abs(price - old) < 1e-9:
            row["결과"] = "표준단가 확인 완료-단가 변경 없음"
            unchanged.append(row)
        else:
            row["결과"] = "표준단가 반영 완료"
            applied.append(row)

    # 원본 요청 목록에 없는 결과 행 → 추가 확인 품목
    known_ids = {it["id"] for it in items}
    known_keys = {_match_key(it["품명"], it["Maker"], None, it["규격"]) for it in items}
    extra = []
    for rec in results:
        if rec.get("id") and rec["id"] not in known_ids:
            extra.append(rec)
        elif not rec.get("id"):
            key = _match_key(rec.get("품명"), rec.get("Maker"), rec.get("모델명"), rec.get("규격"))
            if key and key not in known_keys:
                extra.append(rec)

    summary = _build_summary(items, applied, unchanged, missed, match_check)
    return {
        "applied": applied, "unchanged": unchanged, "missed": missed,
        "match_check": match_check, "extra": extra, "summary": summary,
    }


def _match_key(품명, maker, 모델명, 규격) -> str:
    parts = [_norm(품명), _norm(maker), _norm(모델명), _norm(규격)]
    return "|".join(parts)


def _parse_price(v) -> Optional[float]:
    """표준단가를 float로 파싱. 숫자가 아니거나 음수면 None."""
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v) if v >= 0 else None
    s = str(v)
    for ch in ("₩", "원", ",", " ", "￦", "$", "€", "¥"):
        s = s.replace(ch, "")
    if not s:
        return None
    try:
        f = float(s)
        return f if f >= 0 else None
    except ValueError:
        return None


def _miss_reason(rec: Dict[str, Any]) -> str:
    price = rec.get("표준단가")
    if price is None or (isinstance(price, str) and not price.strip()):
        return "표준단가 공란"
    if _parse_price(price) is None:
        return "표준단가 숫자 아님/음수"
    res = str(rec.get("조회결과") or "")
    if any(k in res for k in ("실패", "없음", "미확인", "fail", "not found")):
        return "조회 실패"
    return "표준단가 유효하지 않음"


def _push_missed(missed, it, reason):
    row = dict(it)
    row["미반영사유"] = reason
    row["주석"] = _MISS_NOTE
    missed.append(row)


def _push_match_check(match_check, it, candidates, reason):
    row = dict(it)
    row["매칭후보"] = candidates
    row["사유"] = reason
    match_check.append(row)


def _build_summary(items, applied, unchanged, missed, match_check) -> Dict[str, Any]:
    by_cat = {"기구_구매품": 0, "제어_구매품": 0, "전장_구매품": 0}
    supply_cnt = 0
    for it in items:
        by_cat[it["cat"]] = by_cat.get(it["cat"], 0) + 1
        if it.get("사급") == "사급":
            supply_cnt += 1
    old_total = sum(it["수량"] * it["기존단가"] for it in items)
    # 변경 후 총액: 반영/변경없음은 표준단가, 미반영/매칭확인은 기존단가
    new_total = 0.0
    std_price = {r["id"]: r["표준단가"] for r in applied}
    std_price.update({r["id"]: r["표준단가"] for r in unchanged})
    for it in items:
        if it["id"] in std_price:
            new_total += it["수량"] * std_price[it["id"]]
        else:
            new_total += it["수량"] * it["기존단가"]
    delta = new_total - old_total
    delta_pct = (delta / old_total * 100.0) if old_total > 0 else None
    return {
        "total": len(items),
        "by_cat": by_cat,
        "supply_cnt": supply_cnt,
        "supply_old_total": _supply_total(items, "기존단가"),
        "applied": len(applied),
        "unchanged": len(unchanged),
        "missed": len(missed),
        "match_check": len(match_check),
        "old_total": old_total,
        "new_total": new_total,
        "delta": delta,
        "delta_pct": delta_pct,
    }


# ─── 5단계: 반영 결과 Excel (요약+반영+미반영+매칭+이력) ──────
def build_result_excel(result: Dict[str, Any], source_name: str) -> bytes:
    """표준단가 반영 결과 Excel(5개 시트)을 만든다."""
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment

    wb = Workbook()
    wb.remove(wb.active)

    hdr_fill = PatternFill("solid", fgColor="1F618D")
    hdr_font = Font(color="FFFFFF", bold=True)
    green_fill = PatternFill("solid", fgColor="D5F5E3")
    red_fill = PatternFill("solid", fgColor="FDEDEC")
    amber_fill = PatternFill("solid", fgColor="FEF5E7")

    def _sheet(title, headers, rows, fills=None):
        ws = wb.create_sheet(title)
        for c, h in enumerate(headers, 1):
            cell = ws.cell(1, c, h)
            cell.fill = hdr_fill
            cell.font = hdr_font
            cell.alignment = Alignment(horizontal="center")
        for ri, row in enumerate(rows, 2):
            for c, v in enumerate(row, 1):
                ws.cell(ri, c, v)
            if fills and fills.get(ri - 2):
                for c in range(1, len(headers) + 1):
                    ws.cell(ri, c).fill = fills[ri - 2]
        # 열 너비
        for c in range(1, len(headers) + 1):
            ws.column_dimensions[ws.cell(1, c).column_letter].width = 16
        return ws

    s = result["summary"]
    # 1. 요약
    _sheet("1.요약", ["항목", "값"], [
        ["원본 파일", source_name],
        ["요청 대상 총 건수", s["total"]],
        ["기구_구매품", s["by_cat"].get("기구_구매품", 0)],
        ["제어_구매품", s["by_cat"].get("제어_구매품", 0)],
        ["전장_구매품", s["by_cat"].get("전장_구매품", 0)],
        ["사급 구매품", s["supply_cnt"]],
        ["사급품 기존 합계", s["supply_old_total"]],
        ["표준단가 반영 완료", s["applied"]],
        ["확인 완료-단가 변경 없음", s["unchanged"]],
        ["표준단가 미반영", s["missed"]],
        ["매칭 확인 필요", s["match_check"]],
        ["변경 전 구매품 총액", s["old_total"]],
        ["변경 후 구매품 총액", s["new_total"]],
        ["총 증감액", s["delta"]],
        ["총 증감률(%)", round(s["delta_pct"], 2) if s["delta_pct"] is not None else "비교 불가"],
    ])

    # 2. 반영 품목
    applied_rows = []
    for r in result["applied"] + result["unchanged"]:
        applied_rows.append([
            r["구분"], r["id"], r["품명"], r["Maker"], r["규격"], r["사급"],
            r["수량"], r["기존단가"], r["표준단가"], r["변경후금액"],
            r["기준일"], r["결과"],
        ])
    _sheet("2.반영품목",
           ["구분", "식별번호", "품명", "Maker", "모델명/규격", "사급", "수량",
            "변경 전 단가", "표준단가", "변경 후 금액", "단가 기준일", "반영 결과"],
           applied_rows)

    # 3. 미반영 품목
    missed_rows = []
    for r in result["missed"]:
        missed_rows.append([
            r["구분"], r["id"], r["품명"], r["Maker"], r["규격"], r["사급"],
            r["수량"], r["기존단가"], r["수량"] * r["기존단가"],
            r["미반영사유"], r["주석"],
        ])
    _sheet("3.미반영품목",
           ["구분", "식별번호", "품명", "Maker", "모델명/규격", "사급", "수량",
            "기존 초기 단가", "기존 단가 적용 금액", "미반영 사유", "주석"],
           missed_rows)

    # 4. 매칭 확인 필요
    mc_rows = []
    for r in result["match_check"]:
        cands = "; ".join(
            f"{c.get('품명')}/{c.get('Maker')}/{c.get('규격')}={c.get('표준단가')}"
            for c in r["매칭후보"]
        )
        mc_rows.append([r["id"], r["품명"], cands, r["사유"], "사용자 확인 필요"])
    _sheet("4.매칭확인필요",
           ["식별번호", "원본 품목", "매칭 후보", "자동 반영 안 한 사유", "사용자 확인 필요 사항"],
           mc_rows)

    # 5. 판단근거
    basis = _basis_text(result, source_name)
    _sheet("5.판단근거", ["판단근거"], [[line] for line in basis])

    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ─── 5-1단계: 기존단가 vs 표준단가 비교 Excel ─────────────────
def build_price_compare_excel(result: Dict[str, Any], source_name: str) -> bytes:
    """'기존파일의단가 vs. 표준견적 반영단가' 비교 Excel을 만든다.

    모든 대상 구매품(반영/변경없음/미반영/매칭확인)을 한 시트에 나열해
    기존 단가·금액과 표준단가(반영단가)·금액, 증감액·증감률을 비교한다.
    미반영/매칭확인 품목은 표준단가가 적용되지 않아 반영단가=기존단가,
    증감 0으로 표시하고 결과 열에 사유를 기록한다.
    """
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment

    wb = Workbook()
    ws = wb.active
    ws.title = "단가비교"

    hdr_fill = PatternFill("solid", fgColor="1F618D")
    hdr_font = Font(color="FFFFFF", bold=True)
    green_fill = PatternFill("solid", fgColor="D5F5E3")   # 반영 완료
    blue_fill = PatternFill("solid", fgColor="EBF5FB")    # 변경 없음
    red_fill = PatternFill("solid", fgColor="FDEDEC")     # 미반영
    amber_fill = PatternFill("solid", fgColor="FEF5E7")   # 매칭확인

    # M열(13열) '견적기준가': 적용할 단가 (사용자 수정 가능).
    # 반영/변경없음 품목은 표준단가 사전 입력, 미반영/매칭확인 품목은 공란으로
    # 두어 사용자가 단가를 채워 재업로드하면 그 단가가 반영된다.
    headers = [
        "식별번호", "구분", "품명", "Maker", "규격", "사급", "수량",
        "기존 단가", "기존 금액",
        "표준(반영) 단가", "표준(반영) 금액",
        "단가 증감", "견적기준가", "금액 증감", "증감률(%)", "결과",
    ]
    for c, h in enumerate(headers, 1):
        cell = ws.cell(1, c, h)
        cell.fill = hdr_fill
        cell.font = hdr_font
        cell.alignment = Alignment(horizontal="center")

    def _row(r, std_price, std_amount, status, fill, apply_price):
        old_price = r["기존단가"]
        old_amount = r["수량"] * old_price
        d_price = std_price - old_price
        d_amount = std_amount - old_amount
        pct = round((d_price / old_price * 100.0), 2) if old_price > 0 else "비교 불가"
        vals = [
            r["id"], r["구분"], r["품명"], r["Maker"], r["규격"], r["사급"],
            r["수량"], old_price, old_amount,
            std_price, std_amount,
            d_price, apply_price, d_amount, pct, status,
        ]
        return vals, fill

    rows = []
    fills = []
    # 반영 완료
    for r in result["applied"]:
        rows.append(_row(r, r["표준단가"], r["변경후금액"], "표준단가 반영 완료", green_fill,
                         r["표준단가"])[0])
        fills.append(green_fill)
    # 변경 없음
    for r in result["unchanged"]:
        rows.append(_row(r, r["표준단가"], r["변경후금액"], "확인 완료-단가 변경 없음", blue_fill,
                         r["표준단가"])[0])
        fills.append(blue_fill)
    # 미반영 (표준단가 미적용 → 기존단가 유지, 견적기준가는 공란으로 사용자 입력 대기)
    for r in result["missed"]:
        rows.append(_row(r, r["기존단가"], r["수량"] * r["기존단가"],
                         f"미반영: {r.get('미반영사유', '')}", red_fill, None)[0])
        fills.append(red_fill)
    # 매칭 확인 필요 (표준단가 미적용 → 기존단가 유지, 견적기준가는 공란으로 사용자 입력 대기)
    for r in result["match_check"]:
        rows.append(_row(r, r["기존단가"], r["수량"] * r["기존단가"],
                         f"매칭 확인 필요: {r.get('사유', '')}", amber_fill, None)[0])
        fills.append(amber_fill)

    for ri, row in enumerate(rows, 2):
        for c, v in enumerate(row, 1):
            ws.cell(ri, c, v)
        if fills:
            for c in range(1, len(headers) + 1):
                ws.cell(ri, c).fill = fills[ri - 2]

    # 합계 행
    if rows:
        hr = len(rows) + 2
        old_amt_sum = sum(r[8] for r in rows)
        new_amt_sum = sum(r[10] for r in rows)
        total_row = ["합계", "", "", "", "", "", "", "", old_amt_sum,
                     "", new_amt_sum, "", "", new_amt_sum - old_amt_sum, "", ""]
        for c, v in enumerate(total_row, 1):
            cell = ws.cell(hr, c, v)
            cell.font = Font(bold=True)
            cell.fill = PatternFill("solid", fgColor="D6EAF8")

    # 열 너비
    widths = [12, 12, 28, 14, 24, 8, 8, 14, 16, 16, 16, 14, 16, 16, 12, 26]
    for c, w in enumerate(widths, 1):
        ws.column_dimensions[ws.cell(1, c).column_letter].width = w
    ws.freeze_panes = "A2"

    # 헤더 행 필터(▼) 적용: 데이터 행만 포함 (합계 행 제외)
    if rows:
        last_col_letter = ws.cell(1, len(headers)).column_letter
        ws.auto_filter.ref = f"A1:{last_col_letter}{len(rows) + 1}"

    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ─── 5-2단계: 단가비교 파일 재업로드 파싱 ─────────────────────
def parse_compare_file(xlsx_bytes: bytes) -> Dict[str, float]:
    """재업로드된 단가비교 파일을 파싱해 {식별번호: 견적기준가}를 반환한다.

    사용자가 단가비교 Excel의 M열 '견적기준가'를 수정해 재업로드하면,
    이 함수가 식별번호별 적용 단가를 추출한다. 견적기준가가 유효한
    숫자인 행만 포함한다 (공란/숫자 아님 행은 제외).
    """
    wb = load_workbook(BytesIO(xlsx_bytes), data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        wb.close()
        return {}

    # 헤더 행 탐색 (첫 10행 중 '식별번호'+'견적기준가' 포함 행)
    header_idx = None
    for i, row in enumerate(rows[:10]):
        cells = [str(c).strip() if c is not None else "" for c in row]
        if any("식별번호" in c for c in cells) and any("견적기준가" in c for c in cells):
            header_idx = i
            break
    if header_idx is None:
        wb.close()
        return {}
    headers = [str(c).strip() if c is not None else "" for c in rows[header_idx]]
    id_idx = next((i for i, h in enumerate(headers) if "식별번호" in h), None)
    price_idx = next((i for i, h in enumerate(headers) if "견적기준가" in h), None)
    if id_idx is None or price_idx is None:
        wb.close()
        return {}

    out: Dict[str, float] = {}
    for row in rows[header_idx + 1:]:
        if row is None:
            continue
        id_val = row[id_idx] if id_idx < len(row) else None
        price_val = row[price_idx] if price_idx < len(row) else None
        if id_val is None:
            continue
        m = re.search(r"(MECH|CTRL|ELEC)-\d{4}", str(id_val))
        if not m:
            continue
        price = _parse_price(price_val)
        if price is None:
            continue
        out[m.group(0)] = price
    wb.close()
    return out


def _basis_text(result: Dict[str, Any], source_name: str) -> List[str]:
    s = result["summary"]
    dpct = f"{s['delta_pct']:.2f}%" if s["delta_pct"] is not None else "비교 불가"
    lines = [
        "[표준단가 반영 결과]",
        f"- 표준단가 확인 요청 대상: 총 {s['total']}건",
        f"- 기구_구매품: {s['by_cat'].get('기구_구매품',0)}건",
        f"- 제어_구매품: {s['by_cat'].get('제어_구매품',0)}건",
        f"- 전장_구매품: {s['by_cat'].get('전장_구매품',0)}건",
        f"- 사급 구매품: {s['supply_cnt']}건",
        f"- 표준단가 반영 완료: {s['applied']}건",
        f"- 표준단가 확인 완료-단가 변경 없음: {s['unchanged']}건",
        f"- 표준단가 미반영: {s['missed']}건",
        f"- 매칭 확인 필요: {s['match_check']}건",
        f"- 변경 전 구매품 총액: {s['old_total']:,.0f}원",
        f"- 변경 후 구매품 총액: {s['new_total']:,.0f}원",
        f"- 총 증감액: {s['delta']:+,.0f}원",
        f"- 총 증감률: {dpct}",
        f"- 원본 파일: {source_name}",
        "",
        "사내 시스템에서 확인된 구매품 표준단가를 기구_구매품, 제어_구매품 및 전장_구매품에 "
        "반영하였다. 사급품도 동일한 검증 대상에 포함하였다. 표준단가가 정상적으로 확인되고 "
        "원본 품목과 정확히 매칭된 항목은 표준단가로 업데이트하였으며, 기존 단가는 변경 이력에 "
        "보존하였다. 표준단가가 확인되지 않았거나 품목 매칭이 불명확한 항목은 기존 초기 단가를 "
        "유지하고 '표준단가 미반영-확인 필요'로 표시하였다.",
    ]
    return lines
