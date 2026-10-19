# 📋 견적 표준 양식 변환기

업체별 견적서(PDF/Excel)를 고객사 표준 양식 Excel로 변환하는 Streamlit 웹앱.

## 기능

- **PDF/Excel 파싱**: 업체별 견적서에서 항목 자동 추출
- **표준 양식 변환**: 고객사 표준 양식(`고객사제출_표준양식.xlsx`)에 데이터 입력
- **원가분석 연동**: `quotation-analyzer` 패키지 기반 자동 분류
- **견적 조정**:
  - 재료비·경비: 변경 비율(%)로 조정
  - 인건비: 대응인원수/대응기간으로 조정
  - **조정견적 내보내기**: 조정된 금액만 반영한 별도 탭(`3.조정 견적`) 생성
- **분석 결과 내보내기**: 견적 세부분류, 집계표, 검증오류, 전체 Excel, 원가구성 차트

## 실행

```powershell
# 1. venv 생성 + 의존성 설치 (자동)
.\start_server.ps1

# 또는 수동
python -m venv .venv
.venv\Scripts\pip install -r requirements.txt
.venv\Scripts\pip install -e C:\AI_Agent\quotation-analyzer
.venv\Scripts\pip install plotly

# 2. 서버 실행
.venv\Scripts\python -m streamlit run app.py --server.port 8501
```

브라우저에서 `http://localhost:8501` 접속.

## 구조

```
app.py                  # Streamlit 메인 앱
config.py               # 설정 (컬럼 매핑, 템플릿 경로)
start_server.ps1        # 서버 자동 시작 스크립트
classification/         # 자동 분류기
converters/             # 변환 로직
  ratio_adjuster.py     # 견적 조정 모델
  rule_mapper.py        # 컬럼 매핑
  standard_exporter.py  # 표준 양식 Excel 생성 + 조정견적 내보내기
parsers/                # PDF/Excel 파서
templates/              # 표준 양식 템플릿
tests/                  # 테스트/디버그 스크립트
```

## 버전

- **v0.2** (2026. 10. 19.)
  - 견적 조정 UI 정렬 (고정 높이 셀)
  - 조정견적 내보내기 (`3.조정 견적` 탭)
  - 현견적 0일 때 `0` 표기
  - 총 견적 금액 글자 크기 축소
