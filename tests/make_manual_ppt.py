# -*- coding: utf-8 -*-
"""견적 표준 양식 변환기 사용 메뉴얼 PPT 생성"""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

# ── 색상 팔레트 ──
C_DARK = RGBColor(0x1F, 0x3A, 0x5F)    # 네이비
C_ACCENT = RGBColor(0x2E, 0x86, 0xAB)  # 블루
C_LIGHT = RGBColor(0xEA, 0xF2, 0xF8)   # 밝은 배경
C_GRAY = RGBColor(0x55, 0x55, 0x55)
C_WHITE = RGBColor(0xFF, 0xFF, 0xFF)
C_GREEN = RGBColor(0x27, 0xAE, 0x60)
C_ORANGE = RGBColor(0xE6, 0x7E, 0x22)

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
BLANK = prs.slide_layouts[6]


def add_title_bar(slide, title, subtitle=""):
    """상단 타이틀 바"""
    bar = slide.shapes.add_shape(1, 0, 0, prs.slide_width, Inches(1.1))
    bar.fill.solid()
    bar.fill.fore_color.rgb = C_DARK
    bar.line.fill.background()
    tf = bar.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.5)
    p = tf.paragraphs[0]
    p.text = title
    p.font.size = Pt(28)
    p.font.bold = True
    p.font.color.rgb = C_WHITE
    if subtitle:
        p2 = tf.add_paragraph()
        p2.text = subtitle
        p2.font.size = Pt(13)
        p2.font.color.rgb = RGBColor(0xBF, 0xD7, 0xEA)


def add_bullets(slide, items, left=0.6, top=1.4, width=12.1, height=5.6, size=18):
    """불릿 텍스트 박스. items: [(text, level, bold, color), ...]"""
    box = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = box.text_frame
    tf.word_wrap = True
    first = True
    for text, level, bold, color in items:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.text = text
        p.level = level
        p.font.size = Pt(size - level * 2)
        p.font.bold = bold
        p.font.color.rgb = color if color else C_GRAY
        p.space_after = Pt(8)
    return box


def add_code_box(slide, text, left, top, width, height):
    """코드/명령어 박스"""
    box = slide.shapes.add_shape(1, Inches(left), Inches(top), Inches(width), Inches(height))
    box.fill.solid()
    box.fill.fore_color.rgb = RGBColor(0x2D, 0x2D, 0x2D)
    box.line.color.rgb = C_ACCENT
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.2)
    tf.margin_top = Inches(0.1)
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(14)
    p.font.name = "Consolas"
    p.font.color.rgb = RGBColor(0x9C, 0xDC, 0xFE)
    return box


# ═══════════════ 1. 표지 ═══════════════
s = prs.slides.add_slide(BLANK)
bg = s.shapes.add_shape(1, 0, 0, prs.slide_width, prs.slide_height)
bg.fill.solid()
bg.fill.fore_color.rgb = C_DARK
bg.line.fill.background()

box = s.shapes.add_textbox(Inches(1), Inches(2.2), Inches(11.3), Inches(3))
tf = box.text_frame
tf.word_wrap = True
p = tf.paragraphs[0]
p.text = "📋 견적 표준 양식 변환기"
p.font.size = Pt(44)
p.font.bold = True
p.font.color.rgb = C_WHITE
p2 = tf.add_paragraph()
p2.text = "사용 메뉴얼"
p2.font.size = Pt(32)
p2.font.color.rgb = RGBColor(0xBF, 0xD7, 0xEA)
p3 = tf.add_paragraph()
p3.text = ""
p4 = tf.add_paragraph()
p4.text = "업체별 견적서(PDF/Excel) → 고객사 표준 양식 Excel 자동 변환"
p4.font.size = Pt(18)
p4.font.color.rgb = RGBColor(0x9C, 0xC9, 0xE8)
p5 = tf.add_paragraph()
p5.text = "ver 0.5 · 2026. 11. 25. · 로컬 처리, 외부 API 미사용"
p5.font.size = Pt(14)
p5.font.color.rgb = RGBColor(0x88, 0xA8, 0xC8)

# ═══════════════ 2. 개요 ═══════════════
s = prs.slides.add_slide(BLANK)
add_title_bar(s, "1. 개요", "이 도구가 하는 일")
add_bullets(s, [
    ("🎯 목적", 0, True, C_DARK),
    ("업체별 형식이 다른 견적서(PDF/Excel)를 고객사 표준 양식 Excel로 자동 변환", 1, False, None),
    ("", 0, False, None),
    ("⚙️ 주요 기능", 0, True, C_DARK),
    ("① 견적서 자동 파싱 — 헤더 자동 감지, 병합 셀 처리, 소계/푸터 행 제외", 1, False, None),
    ("② 원가분석 연동 — quotation-analyzer 패키지 기반 자동 분류 (구매품/가공품/인건비/경비)", 1, False, None),
    ("③ 표준 양식 변환 — 고객사제출_표준양식.xlsx 구조 그대로 데이터 입력", 1, False, None),
    ("④ 견적 조정 — 가공품/경비 비율(%) 조정, 인건비 인원×일수 기반 조정", 1, False, None),
    ("⑤ 분석 챗봇 — \"가공품/구매품/인건비 비율 알려줘\" 등 자연어 질의응답", 1, False, None),
    ("⑥ 자동 저장 — 변환 완료 시 원본 + 표준 양식을 D:\\견적분석 Agent 파일 폴더에 자동 보관", 1, False, None),
    ("", 0, False, None),
    ("🔒 보안", 0, True, C_DARK),
    ("전 과정 로컬 처리 — 외부 API/클라우드 미사용", 1, False, None),
])

# ═══════════════ 3. 실행 방법 ═══════════════
s = prs.slides.add_slide(BLANK)
add_title_bar(s, "2. 실행 방법", "서버 시작 → 브라우저 접속")
add_bullets(s, [
    ("① 서버 시작 (VS Code 또는 PowerShell)", 0, True, C_DARK),
    ("VS Code에서 폴더를 열면 start_server.ps1이 자동 실행됩니다.", 1, False, None),
    ("수동 실행:", 1, False, None),
], size=18)
add_code_box(s, "cd C:\\AI_Agent\\quotation-form-converter\n.venv\\Scripts\\python.exe -m streamlit run app.py --server.port 8501",
             1.2, 3.0, 11.0, 1.1)
add_bullets(s, [
    ("② 브라우저 접속", 0, True, C_DARK),
    ("http://localhost:8501 접속", 1, False, None),
    ("", 0, False, None),
    ("③ 첫 실행 시 자동 처리", 0, True, C_DARK),
    (".venv 생성 → 의존성 자동 설치 (requirements.txt + quotation-analyzer + plotly)", 1, False, None),
    ("기존 8501 포트 서버가 있으면 자동 종료 후 재시작", 1, False, None),
], top=4.4, size=18)

# ═══════════════ 4. Step 1 업로드 ═══════════════
s = prs.slides.add_slide(BLANK)
add_title_bar(s, "3. Step 1 — 견적서 업로드", "1️⃣ 견적서 업로드")
add_bullets(s, [
    ("📤 업로드 방법", 0, True, C_DARK),
    ("화면의 '견적서 파일 업로드 (PDF, Excel)' 영역에 파일 1개 드래그&드롭 또는 클릭 선택", 1, False, None),
    ("지원 형식: PDF, XLSX, XLS (파일당 최대 200MB)", 1, False, None),
    ("", 0, False, None),
    ("⚙️ 사이드바 설정 (선택)", 0, True, C_DARK),
    ("원본 컬럼명 표시 — 출력 Excel M열에 원본 컬럼명/검증 문구 표시", 1, False, None),
    ("🔄 템플릿 교체 (고급) — 기본 표준 양식 템플릿 교체", 1, False, None),
    ("   ⚠️ 표준 양식 시트('1.표준 견적서(갑지)'/'2.상세 견적')가 없는 파일은 자동 거부", 1, False, C_ORANGE),
    ("🔗 매핑 규칙 추가 — 업체별 원본 컬럼명 → 표준 컬럼명 매핑 규칙 등록", 1, False, None),
    ("", 0, False, None),
    ("📌 업로드 후 자동 처리", 0, True, C_DARK),
    ("파싱 → 매핑 → 원가분석(analyzer) → 분류 → 표준 양식 생성 (진행 막대 표시)", 1, False, None),
    ("변환 완료 시 원본 + 표준 양식이 D:\\견적분석 Agent 파일 폴더에 자동 저장", 1, False, C_GREEN),
    ("   (같은 이름 파일이 있으면 _ver2, _ver3... 로 자동 저장)", 1, False, C_GREEN),
])

# ═══════════════ 5. Step 2 변환 결과 ═══════════════
s = prs.slides.add_slide(BLANK)
add_title_bar(s, "4. Step 2 — 변환 결과 확인", "2️⃣ 표준 양식으로 변환")
add_bullets(s, [
    ("📊 화면에 표시되는 내용", 0, True, C_DARK),
    ("🔍 파싱 결과 (원본) — 원본에서 추출된 항목 목록 (접힘)", 1, False, None),
    ("🔗 매핑 결과 — 표준 컬럼 매핑 결과 + 신뢰도 (접힘)", 1, False, None),
    ("📋 섹션별 요약 테이블 — 구매품/가공품/인건비/경비별 금액", 1, False, None),
    ("", 0, False, None),
    ("🏷️ 자동 분류 기준 (quotation-analyzer)", 0, True, C_DARK),
    ("구매품 / 가공품 / 인건비(설계·조립·셋업·양산대응·특수제어) / 경비(포장·기타·숙박교통항공)", 1, False, None),
    ("분류 결과에 '합산금액'이 표시되면 → 해당 행이 합산(소계) 행으로 감지된 것", 1, False, None),
    ("", 0, False, None),
    ("📤 상세 검증 파일", 0, True, C_DARK),
    ("원본 파일 + 자동분류/신뢰도/핵심키워드/분류근거/최종분류 열이 추가된 Excel 다운로드", 1, False, None),
    ("분류 결과 검토 후 '검토상태/수정여부' 열에 직접 수정 가능", 1, False, None),
])

# ═══════════════ 6. Step 3 다운로드 ═══════════════
s = prs.slides.add_slide(BLANK)
add_title_bar(s, "5. Step 3 — 표준 양식 다운로드", "3️⃣ 표준 양식 다운로드")
add_bullets(s, [
    ("⬇️ 다운로드 버튼", 0, True, C_DARK),
    ("'⬇️ 표준 양식 Excel 다운로드' 클릭 → 표준견적_{원본파일명}.xlsx 저장", 1, False, None),
    ("", 0, False, None),
    ("💾 자동 저장 (버튼 클릭 불필요)", 0, True, C_DARK),
    ("변환 완료 시점에 자동으로 아래 폴더에 저장됩니다:", 1, False, None),
], size=18)
add_code_box(s, "D:\\견적분석 Agent 파일\\\n├── {원본 파일명}.xlsx              ← 업로드한 원본 그대로\n└── 표준견적_{원본 파일명}.xlsx    ← 변환된 표준 양식",
             1.2, 4.0, 11.0, 1.2)
add_bullets(s, [
    ("폴더가 없으면 자동 생성 · 같은 이름 파일이 있으면 _ver2, _ver3... 순서로 저장", 0, False, C_GREEN),
], top=5.5, size=16)

# ═══════════════ 7. Step 4 견적 조정 ═══════════════
s = prs.slides.add_slide(BLANK)
add_title_bar(s, "6. Step 4 — 견적 조정", "4️⃣ 견적 조정 (비율 조정)")
add_bullets(s, [
    ("🟡 조정 방식", 0, True, C_DARK),
    ("가공품/경비 — 변경 비율(%)로 조정", 1, False, None),
    ("인건비 — 대응인원수 × 대응기간(MD) 기반 조정 (단가는 분석 결과 자동 반영)", 1, False, None),
    ("", 0, False, None),
    ("📐 사용 절차", 0, True, C_DARK),
    ("① 조정 항목의 비율(%) 또는 인원/기간 입력", 1, False, None),
    ("② '견적조정 산출' 버튼 클릭 → 변경 전/후 총액·증감액 표시", 1, False, None),
    ("③ '📤 조정견적 내보내기' 클릭 → 조정견적_{원본파일명}.xlsx 다운로드", 1, False, None),
    ("", 0, False, None),
    ("📎 구매품 사내 표준단가 반영 (선택)", 0, True, C_DARK),
    ("표준단가 결과 파일 업로드 → '✅ 표준단가 반영' 버튼 클릭", 1, False, None),
    ("미매칭 건 수 배지 표시, 반영 전 상태로 복원 가능", 1, False, None),
])

# ═══════════════ 8. Step 5 분석 내보내기 ═══════════════
s = prs.slides.add_slide(BLANK)
add_title_bar(s, "7. Step 5 — 분석 결과 내보내기", "5️⃣ 분석 결과 내보내기")
add_bullets(s, [
    ("📤 내보내기 항목 (원가분석툴과 동일 문서)", 0, True, C_DARK),
    ("견적 세부분류 — 원본 파일 + 자동분류/신뢰도/핵심키워드/분류근거/최종분류/검토상태/수정여부 열", 1, False, None),
    ("원가분석 내보내기 — 비교안별 항목 + 분류 + 검토 상태 CSV/Excel", 1, False, None),
    ("비용 구성 차트 — Plotly 기반 시각화", 1, False, None),
    ("", 0, False, None),
    ("✏️ 분류 결과 수정", 0, True, C_DARK),
    ("내보낸 Excel에서 '최종분류' 열을 직접 수정 후 다시 업로드하면 수정값 반영", 1, False, None),
    ("", 0, False, None),
    ("💬 분석 챗봇 (사이드바)", 0, True, C_DARK),
    ("예: \"가공품/구매품/인건비/기타 비율 알려줘\"", 1, False, None),
    ("예: \"가공비 상위 10개 항목 20% 올려줘\"", 1, False, None),
    ("예: \"미분류 항목 보여줘\" / \"인건비 총액 알려줘\" / \"경비 110%로 조정해줘\"", 1, False, None),
])

# ═══════════════ 9. 자주 묻는 질문 ═══════════════
s = prs.slides.add_slide(BLANK)
add_title_bar(s, "8. 자주 묻는 질문 (FAQ)")
add_bullets(s, [
    ("Q1. 'No module named quotation_analyzer' 오류가 나면?", 0, True, C_DARK),
    (".venv\\Scripts\\python.exe -m pip install -e C:\\AI_Agent\\quotation-analyzer 후 서버 재시작", 1, False, None),
    ("", 0, False, None),
    ("Q2. M열(자동분류)에 '합산금액'이 표시되면?", 0, True, C_DARK),
    ("해당 행이 소계/합산 행으로 감지된 것입니다. 실제 견적 항목이면 원본의 수량×단가=금액 관계와 '합계' 열 구조를 확인하세요.", 1, False, None),
    ("", 0, False, None),
    ("Q3. 표준 양식 Excel에 빈 행이 섞여 나오면?", 0, True, C_DARK),
    ("원본의 푸터 행(기업이윤/최종금액 등)이 원인일 수 있습니다. v0.4부터 자동 제외됩니다.", 1, False, None),
    ("", 0, False, None),
    ("Q4. 템플릿 교체 후 결과가 이상하면?", 0, True, C_DARK),
    ("data\\custom_template.xlsx를 삭제하면 기본 템플릿으로 복귀합니다. (원본 견적 파일은 템플릿으로 자동 거부됨)", 1, False, None),
    ("", 0, False, None),
    ("Q5. 서버가 안 열리면?", 0, True, C_DARK),
    ("start_server.ps1 재실행 또는 8501 포트 점유 프로세스 종료 후 재시작", 1, False, None),
])

# ═══════════════ 10. 폴더 구조 ═══════════════
s = prs.slides.add_slide(BLANK)
add_title_bar(s, "9. 폴더 구조 & 파일 위치")
add_bullets(s, [
    ("📁 프로젝트 구조", 0, True, C_DARK),
], size=18)
add_code_box(s,
             "C:\\AI_Agent\\\n"
             "├── quotation-form-converter\\     ← 변환기 (이 메뉴얼 대상)\n"
             "│   ├── app.py                    ← Streamlit 메인 앱\n"
             "│   ├── parsers/  converters/  classification/\n"
             "│   ├── templates\\고객사제출_표준양식.xlsx   ← 기본 템플릿\n"
             "│   └── start_server.ps1          ← 서버 자동 시작\n"
             "└── quotation-analyzer\\           ← 원가분석 패키지 (v1.3.15)\n\n"
             "D:\\견적분석 Agent 파일\\            ← 자동 저장 폴더 (자동 생성)",
             0.6, 1.8, 12.1, 4.6)

prs.save(r"D:\견적분석 Agent 파일\견적표준양식변환기_사용메뉴얼.pptx")
print("PPT 저장 완료: D:\\견적분석 Agent 파일\\견적표준양식변환기_사용메뉴얼.pptx")
