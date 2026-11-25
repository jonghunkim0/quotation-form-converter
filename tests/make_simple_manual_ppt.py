# -*- coding: utf-8 -*-
"""견적 표준 양식 변환기 — 간단 사용 메뉴얼 PPT (화면 캡처 포함)

핵심 내용만:
  1. 실행 방법
  2. 견적서 업로드
  3. 표준 양식 다운로드
  4. 견적 조정 (구매품 사내 표준단가 반영 / 재료비·경비 조정 / 인건비 조정)
"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from PIL import Image
import os

# ── 색상 팔레트 ──
C_DARK = RGBColor(0x1F, 0x3A, 0x5F)
C_ACCENT = RGBColor(0x2E, 0x86, 0xAB)
C_GRAY = RGBColor(0x55, 0x55, 0x55)
C_WHITE = RGBColor(0xFF, 0xFF, 0xFF)
C_GREEN = RGBColor(0x27, 0xAE, 0x60)
C_ORANGE = RGBColor(0xE6, 0x7E, 0x22)

SHOTS = os.path.join(os.path.dirname(os.path.abspath(__file__)), "manual_shots")

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
BLANK = prs.slide_layouts[6]


def add_title_bar(slide, title, subtitle=""):
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, Inches(1.0))
    bar.fill.solid()
    bar.fill.fore_color.rgb = C_DARK
    bar.line.fill.background()
    tf = bar.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.5)
    p = tf.paragraphs[0]
    p.text = title
    p.font.size = Pt(26)
    p.font.bold = True
    p.font.color.rgb = C_WHITE
    if subtitle:
        p2 = tf.add_paragraph()
        p2.text = subtitle
        p2.font.size = Pt(12)
        p2.font.color.rgb = RGBColor(0xBF, 0xD7, 0xEA)


def add_bullets(slide, items, left=0.5, top=1.2, width=4.6, height=5.9, size=15):
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
        p.space_after = Pt(6)
    return box


def add_code_box(slide, text, left, top, width, height):
    box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(left), Inches(top), Inches(width), Inches(height))
    box.fill.solid()
    box.fill.fore_color.rgb = RGBColor(0x2D, 0x2D, 0x2D)
    box.line.color.rgb = C_ACCENT
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.2)
    tf.margin_top = Inches(0.1)
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(13)
    p.font.name = "Consolas"
    p.font.color.rgb = RGBColor(0x9C, 0xDC, 0xFE)
    return box


def add_screenshot(slide, filename, left, top, max_width, max_height, caption=None):
    """스크린샷을 테두리 박스 안에 비율 유지하며 삽입"""
    path = os.path.join(SHOTS, filename)
    if not os.path.exists(path):
        print(f"[WARN] 스크린샷 없음: {path}")
        return
    with Image.open(path) as im:
        w, h = im.size
    ratio = w / h
    width = max_width
    height = width / ratio
    if height > max_height:
        height = max_height
        width = height * ratio
    # 테두리
    border = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                    Inches(left), Inches(top), Inches(width), Inches(height))
    border.fill.solid()
    border.fill.fore_color.rgb = C_WHITE
    border.line.color.rgb = C_ACCENT
    border.line.width = Pt(1.5)
    slide.shapes.add_picture(path, Inches(left), Inches(top), Inches(width), Inches(height))
    if caption:
        cap = slide.shapes.add_textbox(Inches(left), Inches(top + height + 0.05),
                                       Inches(width), Inches(0.3))
        cp = cap.text_frame.paragraphs[0]
        cp.text = caption
        cp.font.size = Pt(11)
        cp.font.italic = True
        cp.font.color.rgb = C_GRAY
        cp.alignment = PP_ALIGN.CENTER


# ═══════════════ 1. 표지 ═══════════════
s = prs.slides.add_slide(BLANK)
bg = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
bg.fill.solid()
bg.fill.fore_color.rgb = C_DARK
bg.line.fill.background()

box = s.shapes.add_textbox(Inches(1), Inches(2.4), Inches(11.3), Inches(3))
tf = box.text_frame
tf.word_wrap = True
p = tf.paragraphs[0]
p.text = "📋 견적 표준 양식 변환기"
p.font.size = Pt(44)
p.font.bold = True
p.font.color.rgb = C_WHITE
p2 = tf.add_paragraph()
p2.text = "간단 사용 메뉴얼"
p2.font.size = Pt(32)
p2.font.color.rgb = RGBColor(0xBF, 0xD7, 0xEA)
p3 = tf.add_paragraph()
p3.text = ""
p4 = tf.add_paragraph()
p4.text = "견적서 업로드 → 표준 양식 다운로드 → 견적 조정"
p4.font.size = Pt(18)
p4.font.color.rgb = RGBColor(0x9C, 0xC9, 0xE8)
p5 = tf.add_paragraph()
p5.text = "ver 0.5 · 2026. 11. 25."
p5.font.size = Pt(14)
p5.font.color.rgb = RGBColor(0x88, 0xA8, 0xC8)

# ═══════════════ 2. 실행 방법 ═══════════════
s = prs.slides.add_slide(BLANK)
add_title_bar(s, "실행 방법", "서버 시작 → 브라우저 접속")
add_bullets(s, [
    ("① 서버 시작", 0, True, C_DARK),
    ("VS Code에서 폴더를 열면 자동 실행됩니다.", 1, False, None),
    ("수동 실행:", 1, False, None),
], top=1.3, height=2.0)
add_code_box(s, "cd C:\\AI_Agent\\quotation-form-converter\n.venv\\Scripts\\python.exe -m streamlit run app.py --server.port 8501",
             1.2, 3.1, 11.0, 1.0)
add_bullets(s, [
    ("② 브라우저 접속", 0, True, C_DARK),
    ("http://localhost:8501 접속", 1, False, None),
    ("", 0, False, None),
    ("③ 첫 실행 시", 0, True, C_DARK),
    (".venv 생성 → 의존성 자동 설치 (자동 처리)", 1, False, None),
], top=4.4, height=2.5)

# ═══════════════ 3. Step 1 업로드 ═══════════════
s = prs.slides.add_slide(BLANK)
add_title_bar(s, "Step 1 — 견적서 업로드")
add_bullets(s, [
    ("📤 업로드", 0, True, C_DARK),
    ("'견적서 파일 업로드 (PDF, Excel)' 영역에 파일 1개 드래그&드롭 또는 클릭 선택", 1, False, None),
    ("지원 형식: PDF, XLSX, XLS", 1, False, None),
    ("", 0, False, None),
    ("⚙️ 자동 처리", 0, True, C_DARK),
    ("업로드 즉시 파싱 → 매핑 → 원가분석 → 표준 양식 생성 (진행 막대 표시)", 1, False, None),
    ("", 0, False, None),
    ("💾 자동 저장", 0, True, C_DARK),
    ("변환 완료 시 원본 + 표준 양식이 아래 폴더에 자동 저장", 1, False, None),
    ("D:\\견적분석 Agent 파일\\", 1, True, C_GREEN),
    ("(같은 이름 파일은 _ver2, _ver3... 로 자동 저장)", 1, False, C_GREEN),
], width=4.7)
add_screenshot(s, "02_upload.png", 5.4, 1.3, 7.5, 5.9,
               "업로드 후 자동 변환 진행 화면")

# ═══════════════ 4. Step 2 변환 결과 ═══════════════
s = prs.slides.add_slide(BLANK)
add_title_bar(s, "Step 2 — 변환 결과 확인")
add_bullets(s, [
    ("📊 화면에 표시되는 내용", 0, True, C_DARK),
    ("🔍 파싱 결과 (원본) — 원본에서 추출된 항목 (접힘)", 1, False, None),
    ("🔗 매핑 결과 — 표준 컬럼 매핑 + 신뢰도 (접힘)", 1, False, None),
    ("📋 섹션별 요약 테이블 — 구매품/가공품/인건비/경비별 금액", 1, False, None),
    ("", 0, False, None),
    ("🏷️ 자동 분류 (quotation-analyzer)", 0, True, C_DARK),
    ("구매품 / 가공품 / 인건비(설계·조립·셋업·양산대응) / 경비", 1, False, None),
    ("", 0, False, None),
    ("📤 상세 검증 파일", 0, True, C_DARK),
    ("자동분류/신뢰도/분류근거 열이 추가된 Excel 다운로드", 1, False, None),
    ("'검토상태/수정여부' 열에 직접 수정 가능", 1, False, None),
], width=4.7)
add_screenshot(s, "03_convert.png", 5.4, 1.3, 7.5, 5.9,
               "섹션별 요약 테이블 (Total 견적 확인)")

# ═══════════════ 5. Step 3 다운로드 ═══════════════
s = prs.slides.add_slide(BLANK)
add_title_bar(s, "Step 3 — 표준 양식 다운로드")
add_bullets(s, [
    ("⬇️ 다운로드", 0, True, C_DARK),
    ("'⬇️ 표준 양식 Excel 다운로드' 버튼 클릭", 1, False, None),
    ("→ 표준견적_{원본파일명}.xlsx 저장", 1, False, None),
    ("", 0, False, None),
    ("💾 자동 저장 (버튼 클릭 불필요)", 0, True, C_DARK),
    ("변환 완료 시점에 자동으로 저장:", 1, False, None),
], width=4.7)
add_code_box(s, "D:\\견적분석 Agent 파일\\\n├── {원본 파일명}.xlsx              ← 업로드한 원본\n└── 표준견적_{원본 파일명}.xlsx    ← 변환된 표준 양식",
             0.5, 4.3, 4.7, 1.3)
add_bullets(s, [
    ("폴더가 없으면 자동 생성", 0, False, C_GREEN),
], top=5.8, width=4.7)
add_screenshot(s, "04_download.png", 5.4, 1.3, 7.5, 5.9,
               "다운로드 버튼 위치")

# ═══════════════ 6. Step 4a 표준단가 반영 ═══════════════
s = prs.slides.add_slide(BLANK)
add_title_bar(s, "Step 4 — 견적 조정 ①", "구매품 사내 표준단가 반영")
add_bullets(s, [
    ("📐 사용 절차", 0, True, C_DARK),
    ("① '📐 구매품 사내 표준단가 반영' 버튼 클릭", 1, False, None),
    ("② 대상 구매품 추출 → 사내 시스템 단가 확인 요청용 Excel 다운로드", 1, False, None),
    ("③ 사내 시스템에서 단가 확인 후 결과 파일 재업로드", 1, False, None),
    ("④ '✅ 표준단가 반영' 버튼 클릭 → 표준 양식 Excel에 적용", 1, False, None),
    ("", 0, False, None),
    ("📌 참고", 0, True, C_DARK),
    ("미매칭 건 수 배지 표시", 1, False, None),
    ("반영 전 상태로 복원 가능", 1, False, None),
], width=4.7)
add_screenshot(s, "06_stdprice_dialog.png", 5.4, 1.3, 7.5, 5.9,
               "표준단가 반영 다이얼로그 (4단계 절차)")

# ═══════════════ 7. Step 4b 재료비·경비 조정 ═══════════════
s = prs.slides.add_slide(BLANK)
add_title_bar(s, "Step 4 — 견적 조정 ②", "재료비·경비 조정 (변경 비율)")
add_bullets(s, [
    ("🟡 조정 방식", 0, True, C_DARK),
    ("가공품/경비 — 변경 비율(%)로 조정", 1, False, None),
    ("", 0, False, None),
    ("📐 사용 절차", 0, True, C_DARK),
    ("① '변경 비율(%)' 입력 (예: 110 = 10% 인상)", 1, False, None),
    ("② '📐 견적조정 산출' 버튼 클릭", 1, False, None),
    ("③ 변경(원)/증감(원) 확인", 1, False, None),
    ("④ '📤 조정견적 내보내기' 클릭 → 조정견적_{원본파일명}.xlsx", 1, False, None),
    ("", 0, False, None),
    ("📌 대상 항목", 0, True, C_DARK),
    ("1. 가공품 · 기구부 / 전장부 / 제어부", 1, False, None),
    ("4. 경비", 1, False, None),
], width=4.7)
add_screenshot(s, "07_material_adjust.png", 5.4, 1.3, 7.5, 5.9,
               "재료비·경비 조정 테이블 (노란 칸에 비율 입력)")

# ═══════════════ 8. Step 4c 인건비 조정 ═══════════════
s = prs.slides.add_slide(BLANK)
add_title_bar(s, "Step 4 — 견적 조정 ③", "인건비 조정 (대응인원수 · 대응기간)")
add_bullets(s, [
    ("🟡 조정 방식", 0, True, C_DARK),
    ("인건비 — 대응인원수 × 대응기간(MD) 기반 조정", 1, False, None),
    ("단가는 분석 결과 자동 반영 (수동 수정 불필요)", 1, False, None),
    ("", 0, False, None),
    ("📐 계산 로직", 0, True, C_DARK),
    ("MD = 인원수 × 대응기간", 1, False, None),
    ("금액 = MD × 단가", 1, False, None),
    ("", 0, False, None),
    ("📐 사용 절차", 0, True, C_DARK),
    ("① '인원수'/'대응일수' 입력", 1, False, None),
    ("② '📐 견적조정 산출' → 변경 전/후 총액·증감액 확인", 1, False, None),
    ("③ '📤 조정견적 내보내기' 클릭", 1, False, None),
], width=4.7)
add_screenshot(s, "08_labor_adjust.png", 5.4, 1.3, 7.5, 5.9,
               "인건비 조정 테이블 (직급별 단가 자동 표시)")

prs.save(r"D:\견적분석 Agent 파일\견적표준양식변환기_간단사용메뉴얼.pptx")
print("PPT 저장 완료: D:\\견적분석 Agent 파일\\견적표준양식변환기_간단사용메뉴얼.pptx")
