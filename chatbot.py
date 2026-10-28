"""
견적 분석 챗봇 — Qwen3.8 (사내 LLM 서버) 기반

quotation-analyzer의 분석 도우미와 동일한 API를 사용.
- 조회: "가공품/구매품/인건비/기타 비율 알려줘"
- 조정: "가공비 상위 10개 항목 20% 올려줘"
- 분류: "미분류 항목 보여줘"
"""

import json
import re
import ssl
import urllib.request
from typing import List, Dict, Any, Optional

# ─── LLM 설정 (quotation-analyzer와 동일) ──────────────────────
CHAT_API_URL = "https://bai.hpc.lge.com:10602/api/v1/chat/completions"
CHAT_API_KEY = "sk-9370684cc9794f03b2b951b02bab8d98"
CHAT_MODEL = "Qwen/Qwen3.8-27B-FP8"


# ─── 데이터 컨텍스트 빌더 ────────────────────────────────────────

def _parse_amount(val) -> float:
    """금액 문자열/숫자를 float로 변환."""
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


def _build_context(
    mapped_items: List[Dict[str, Any]],
    labor_data: Optional[Dict[str, Dict[str, Any]]],
) -> str:
    """LLM에게 전달할 견적 데이터 컨텍스트 생성."""
    _FAMILY = {
        "기구_가공품": ("가공품", "기구부"), "전장_가공품": ("가공품", "전장부"),
        "제어_가공품": ("가공품", "제어부"),
        "기구_구매품": ("구매품", "기구부"), "전장_구매품": ("구매품", "전장부"),
        "제어_구매품": ("구매품", "제어부"),
    }
    made = {"기구부": 0.0, "전장부": 0.0, "제어부": 0.0}
    buy = {"기구부": 0.0, "전장부": 0.0, "제어부": 0.0}
    expense = 0.0
    labor_total = 0.0
    other = 0.0
    unclassified = []

    items_lines = []
    for i, it in enumerate(mapped_items):
        if it.get("_합산"):
            continue
        qty = _parse_amount(it.get("수량"))
        price = _parse_amount(it.get("단가"))
        amt = qty * price
        sec = it.get("_분류", "미분류")

        items_lines.append(
            f"{i+1}. {str(it.get('품명',''))[:30]} | {amt:,.0f}원 | {sec}"
        )

        if sec in _FAMILY:
            kind, fam = _FAMILY[sec]
            (made if kind == "가공품" else buy)[fam] += amt
        elif sec in ("포장", "기타항목", "경비"):
            expense += amt
        elif sec == "미분류":
            other += amt
            unclassified.append(f"{i+1}. {str(it.get('품명',''))[:30]} | {amt:,.0f}원")

    # 인건비
    labor_lines = []
    if labor_data:
        for cat, info in labor_data.items():
            labor_total += info.get("amount", 0)
            labor_lines.append(
                f"- {cat}: MD={info.get('md',0):.0f}, "
                f"단가={info.get('rate',0):,.0f}원, "
                f"금액={info.get('amount',0):,.0f}원"
            )

    total = sum(made.values()) + sum(buy.values()) + labor_total + expense + other

    ctx = f"## 견적 요약\n"
    ctx += f"- 총견적: {total:,.0f}원\n"
    ctx += f"- 가공품: 기구 {made['기구부']:,.0f} / 전장 {made['전장부']:,.0f} / 제어 {made['제어부']:,.0f}\n"
    ctx += f"- 구매품: 기구 {buy['기구부']:,.0f} / 전장 {buy['전장부']:,.0f} / 제어 {buy['제어부']:,.0f}\n"
    ctx += f"- 인건비: {labor_total:,.0f}원\n"
    ctx += f"- 경비/기타: {expense + other:,.0f}원\n"

    if labor_lines:
        ctx += f"\n## 인건비 세부\n" + "\n".join(labor_lines) + "\n"

    if unclassified:
        ctx += f"\n## 미분류 항목 ({len(unclassified)}개)\n" + "\n".join(unclassified) + "\n"

    ctx += f"\n## 견적 항목 목록 (번호 = 순서)\n"
    ctx += "\n".join(items_lines) + "\n"

    return ctx


# ─── 시스템 프롬프트 ─────────────────────────────────────────────

def _system_prompt(context: str) -> str:
    return (
        '너는 "견적 표준 양식 변환기"의 분석 도우미다. '
        "사용자의 요청에 따라 견적 데이터를 조회하거나 조정 지시를 내린다.\n\n"
        f"## 현재 견적 데이터\n{context}\n\n"
        "## 규칙\n"
        "1. 항상 한국어로 간결하게 답한다.\n"
        '2. "N번째 항목"은 위 목록의 번호 N을 의미한다.\n'
        "3. 조회 요청(비율, 총액, 목록 등)이면 텍스트로만 답한다.\n"
        '4. 조정 요청이면, 답의 마지막 줄에 반드시 아래 형식의 ACTION 줄을 포함하라:\n'
        'ACTION: {"action":"adjust_made_pct","target":"기구부","value":120}\n'
        'ACTION: {"action":"adjust_expense_pct","value":110}\n'
        'ACTION: {"action":"adjust_labor","target":"인건비-기구 설계","value":120}\n'
        'ACTION: {"action":"reclassify","indices":[50],"category":"기구_구매품"}\n'
        "   - adjust_made_pct: target=기구부/전장부/제어부, value=변경 비율(%)\n"
        "   - adjust_expense_pct: value=경비 변경 비율(%)\n"
        "   - adjust_labor: target=인건비 카테고리, value=변경 비율(%)\n"
        "   - reclassify: indices=항목 번호 배열, category=새 분류\n"
        "5. '20% 올려줘' = 120%, '50% 줄여줘' = 50%\n"
        "6. '상위 N개' 요청: 금액 내림차순 정렬 후 N개 추출\n"
        "7. 데이터에 없는 요청이면 그렇게 알려준다.\n"
        "8. 금액은 원(₩) 단위, 천 단위 쉼표 사용.\n"
    )


# ─── LLM 호출 ────────────────────────────────────────────────────

def _send_llm(messages: list) -> str:
    """사내 LLM API 호출 (quotation-analyzer와 동일)."""
    ctx = ssl.create_default_context()
    ctx.check_hostname = False
    ctx.verify_mode = ssl.CERT_NONE

    req = urllib.request.Request(
        CHAT_API_URL,
        data=json.dumps({
            "model": CHAT_MODEL,
            "messages": messages,
            "max_tokens": 2000,
            "temperature": 0,
        }).encode(),
        headers={
            "Content-Type": "application/json",
            "Authorization": f"Bearer {CHAT_API_KEY}",
        },
        method="POST",
    )
    with urllib.request.urlopen(req, timeout=120, context=ctx) as resp:
        data = json.loads(resp.read())
    msg = data.get("choices", [{}])[0].get("message", {})
    # Qwen3.8 reasoning 모델: content가 비면 reasoning_content에 응답이 있을 수 있음
    content = msg.get("content", "") or msg.get("reasoning_content", "")
    return content


def ask_chatbot(
    user_query: str,
    mapped_items: List[Dict[str, Any]],
    labor_data: Optional[Dict[str, Dict[str, Any]]],
    history: Optional[List[Dict[str, str]]] = None,
) -> Dict[str, Any]:
    """
    사용자 요청을 LLM에 전달하고 응답을 반환.

    Returns:
        {"type": "query"|"adjust"|"error", "message": str, "actions": [...]}
    """
    context = _build_context(mapped_items, labor_data)
    messages = [
        {"role": "system", "content": _system_prompt(context)},
    ]
    if history:
        for h in history[-8:]:
            messages.append({"role": h["role"], "content": h["content"]})
    messages.append({"role": "user", "content": user_query})

    try:
        content = _send_llm(messages)
        if not content:
            return {"type": "error", "message": "⚠️ 응답이 비어 있습니다.", "actions": []}

        # ACTION 추출
        actions = []
        for m in re.finditer(r"ACTION:\s*(\{[^}]+\})", content):
            try:
                act = json.loads(m.group(1))
                if act not in actions:
                    actions.append(act)
            except json.JSONDecodeError:
                pass

        # ACTION 줄 제거 후 메시지
        reply = re.sub(r"ACTION:\s*\{[^}]+\}", "", content).strip()

        return {
            "type": "adjust" if actions else "query",
            "message": reply,
            "actions": actions,
        }
    except Exception as e:
        return {
            "type": "error",
            "message": f"❌ LLM 호출 오류: {e}",
            "actions": [],
        }


# ─── 조정 액션 적용 ──────────────────────────────────────────────

def apply_actions(
    actions: List[Dict[str, Any]],
    ratio_model,
    ratio_inputs: Dict[str, Any],
    mapped_items: List[Dict[str, Any]],
) -> List[str]:
    """
    LLM이 반환한 actions를 실제 모델에 적용.

    Returns:
        적용된 조치 설명 목록
    """
    applied = []

    for act in actions:
        action = act.get("action", "")
        target = act.get("target", "")
        value = act.get("value", 0)
        desc = act.get("description", "")

        if action == "adjust_made_pct":
            # 가공품 비율 조정 (target: 기구부/전장부/제어부)
            if target in ("기구부", "전장부", "제어부"):
                key = f"made|{target}|pct"
                ratio_inputs[key] = float(value)
                ratio_model.made_pct[target] = float(value)
                applied.append(f"✅ {target} 가공품 → {value:.0f}% 적용")

        elif action == "adjust_expense_pct":
            ratio_inputs["expense|pct"] = float(value)
            ratio_model.expense_pct = float(value)
            applied.append(f"✅ 경비 → {value:.0f}% 적용")

        elif action == "adjust_labor":
            # 인건비 조정: target = "인건비-기구 설계" 등
            # value = 비율(%) → 현재 금액 × value/100
            # ratio_inputs에 대응하는 key를 찾아 days 조정
            from converters.ratio_adjuster import _LABOR_SECTION_MAP
            if target in _LABOR_SECTION_MAP:
                section, job = _LABOR_SECTION_MAP[target]
                # 해당 섹션의 모든 행에 적용
                for row in ratio_model.labor_rows:
                    if row.section == section and row.cur_amount > 0:
                        # 현재 금액 × value/100 = 목표 금액
                        # 목표 금액 / 단가 = 목표 MD
                        target_md = (row.cur_amount * value / 100.0) / row.rate if row.rate > 0 else 0
                        # 인원수는 유지, days만 조정
                        if row.adj_people > 0:
                            new_days = target_md / row.adj_people
                        else:
                            new_days = target_md
                        key_d = f"{row.section}|{row.job}|days"
                        ratio_inputs[key_d] = round(new_days, 1)
                        row.adj_days = round(new_days, 1)
                applied.append(f"✅ {target} → {value:.0f}% 적용")

        elif action == "reclassify":
            # 항목 재분류 (target: "50", value: "기구_구매품" 등)
            try:
                row_no = int(target)
                new_sec = str(value)
                if 1 <= row_no <= len(mapped_items):
                    mapped_items[row_no - 1]["_분류"] = new_sec
                    applied.append(f"✅ {row_no}번 항목 → {new_sec} 재분류")
            except (ValueError, IndexError):
                applied.append(f"⚠️ {target}번 항목 재분류 실패")

    return applied
