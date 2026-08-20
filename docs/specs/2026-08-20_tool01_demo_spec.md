# [실행 명세] Tool01 Demo

## 목표

`XLSX/CSV 데이터 + PPTX 템플릿`을 입력받아 각 데이터 행별로 개인화된 PDF를 생성하는 **Tool01 Validation Demo**를 구현한다. 개발·검수는 사용자가 로컬 Windows 환경을 사용하지 않아도 되도록 Claude Code Remote/Linux + GitHub Actions에서 완료할 수 있어야 한다.

이번 작업의 목적은 상용판 전체를 만드는 것이 아니라, B3 Validation Package의 핵심 가치인 아래 흐름을 실제로 증명하는 것이다.

`명단 선택 → 템플릿 선택 → 자동매핑 → Preview → 개인별 PDF 일괄생성`

---

## Baseline 환경

### 사용자 작업 방식 — 확정 제약 (GPT DECISION 2026-08-20)

사용자는 **개발·테스트·검수를 온라인 환경에서만 진행**한다.

따라서 현재 Validation Demo의 Blocking QA Gate는 로컬 Windows 실기 테스트가 아니라 아래 온라인 경로다.

- Claude Code Remote / Linux
- GitHub Repository
- GitHub Actions
- LibreOffice Impress Headless Renderer
- pytest + Cloud Acceptance Script

사용자에게 로컬 Windows PC에서 명령 실행, PowerPoint 설치, W01~W09 수동테스트를 요구하지 않는다.

### 제품 Runtime Baseline

현재 고객용 Runtime 방향은 기존과 같이 Windows 로컬 GUI를 유지한다.

- 입력: XLSX/CSV + PPTX
- 출력: 개인별 PDF
- GUI: Tkinter 기반 단일 화면
- 서버/Web App/Login/DB: 사용하지 않음

### Renderer 정책

두 Renderer Adapter를 유지한다.

1. **LibreOfficeRenderer — Online/Cloud Acceptance Baseline**
   - Linux/Claude Code Remote/GitHub Actions에서 실행 가능
   - LibreOffice Impress headless 사용
   - 현재 Validation Demo의 Blocking QA Renderer
   - 실제 PPTX → PDF 출력물을 생성해 온라인에서 검증한다.

2. **PowerPointRenderer — Windows Compatibility Adapter**
   - Windows + Microsoft PowerPoint Desktop
   - PowerPoint COM 사용
   - PPTX 원본 재현 충실도가 중요할 때 사용할 고객 Runtime 후보
   - 현재 시장 Validation 진입을 막는 Blocking Gate가 아니다.

Renderer 선택은 환경변수 `TOOL01_RENDERER=auto|libreoffice|powerpoint`를 사용한다.

- `auto`: Windows에서는 PowerPoint, 비Windows에서는 LibreOffice
- Cloud Acceptance: `libreoffice` 고정

PowerPoint COM 호환성 실기는 향후 유료 Pilot/패키징 단계에서 필요성이 확인되면 **온라인 Windows VM 또는 외부 Pilot 환경**으로 수행한다. 사용자의 로컬 PC 작업을 전제로 하지 않는다.

### Online Acceptance 범위

작은 Remote Runner에서 Office 프로세스를 과도하게 반복 기동하지 않도록 다음을 Validation Demo 완료기준으로 사용한다.

- Sample 100행 전체: 데이터 로딩/매핑/개인화 PPTX 100/100 검증
- 실제 PDF Renderer: 대표 5행을 LibreOffice로 실제 PDF 5/5 생성
- 실제 생성 PDF에서 한글 내용/한글 파일명 검증
- CSV 입력은 데이터 로딩/매핑/개인화 경로까지 검증하고, PDF Renderer 자체는 대표 Batch에서 별도로 실검증
- 부분 실패/권한/GUI 중복실행은 deterministic pytest로 검증

이는 **시장 Validation용 Demo Gate**다. 100건 PowerPoint COM 실기와 Installer/상용 배포는 이후 단계다.

## 입력

### 1. 데이터 파일

지원 형식:

- `.xlsx`
- `.csv`

Demo 표준 샘플 파일:

`sample_roster.xlsx`

필수 샘플 컬럼:

| 컬럼 | 예시 |
|---|---|
| 수료번호 | CERT-2026-001 |
| 이름 | 홍길동 |
| 과정명 | AI 업무자동화 실무과정 |
| 수료일 | 2026-08-20 |
| 기관명 | ABC 교육원 |

샘플 데이터는 총 100행을 생성한다.

### 2. PowerPoint Template

지원 형식:

- `.pptx`

Demo 표준 샘플 파일:

`certificate_template.pptx`

필수 Placeholder:

- `{{수료번호}}`
- `{{이름}}`
- `{{과정명}}`
- `{{수료일}}`
- `{{기관명}}`

Demo에서는 일반 Text Box 안의 텍스트 Placeholder를 기준으로 구현한다.

### 3. 출력 폴더

사용자가 GUI에서 지정한다.

원본 XLSX/CSV/PPTX가 있는 위치와 같더라도 원본 파일을 변경하거나 삭제해서는 안 된다.

---

## 출력

각 데이터 행마다 1개의 PDF를 생성한다.

Demo 기본 파일명 규칙:

`{수료번호}_{이름}_수료증.pdf`

예:

`CERT-2026-001_홍길동_수료증.pdf`

배치 완료 후 UI에 최소한 다음을 표시한다.

- 전체 대상 건수
- 성공 건수
- 실패 건수
- 완료 상태
- 결과 폴더 열기 동작

---

## 필수 요구사항

### 1. 단일 화면 GUI

한 화면에서 최소 다음 요소를 제공한다.

- 데이터 파일 경로
- 데이터 파일 선택 버튼
- PPTX Template 경로
- Template 선택 버튼
- 출력 폴더 경로
- 출력 폴더 선택 버튼
- 데이터 건수 표시
- 인식 필드 표시
- Mapping/Validation 상태
- `Preview` 버튼
- `PDF 일괄생성` 버튼
- 진행률
- 성공/실패 결과
- 결과 폴더 열기

UI는 기능 검증이 목적이므로 과도한 디자인 작업을 하지 않는다.

### 2. 데이터 로딩

- XLSX/CSV를 정상적으로 읽는다.
- Header Row를 컬럼명으로 사용한다.
- 완전히 빈 행은 처리 대상에서 제외할 수 있다.
- 읽기 실패 시 사용자에게 이해 가능한 오류를 표시한다.
- 데이터 원본은 수정하지 않는다.

### 3. Placeholder 인식

PPTX Template에서 `{{필드명}}` 형식의 Placeholder를 찾는다.

예:

`{{이름}}` → `이름` 컬럼

Placeholder의 필드명이 데이터 파일 Header와 정확히 동일하면 자동매핑한다.

대소문자 개념이 없는 한글 Demo 필드를 기준으로 하며, 공백을 임의로 보정해 다른 필드로 추정하지 않는다.

### 4. Preflight Validation

PDF 생성 전에 최소 다음을 검사한다.

- 데이터 파일 선택 여부
- PPTX 파일 선택 여부
- 출력 폴더 선택 여부
- 파일 존재 여부
- 지원 확장자 여부
- 데이터 Header 존재 여부
- Template Placeholder 존재 여부
- Placeholder에 대응하는 데이터 컬럼 존재 여부
- 처리 가능한 데이터 행 존재 여부

Template에 `{{과정명}}`이 있는데 데이터에 `과정명` 컬럼이 없다면 Batch 생성을 시작하지 않고 누락 필드 이름을 표시한다.

### 5. 자동매핑 결과 표시

사용자가 최소한 다음 사실을 확인할 수 있어야 한다.

- 인식된 Placeholder 목록
- 대응되는 데이터 컬럼
- 누락 필드 여부

Demo에서는 사용자가 Mapping을 수동 편집하는 기능은 만들지 않는다.

### 6. Preview

`Preview`는 첫 번째 유효 데이터 행을 Template에 적용한 실제 결과를 생성한다.

Preview 결과에서 최소 다음을 확인할 수 있어야 한다.

- 이름
- 수료번호
- 과정명
- 수료일
- 기관명
- Template의 기존 디자인

내장 PDF Viewer는 필수사항이 아니다. 구현 복잡도를 줄이기 위해 임시 Preview PDF를 생성하고 Windows 기본 Viewer로 여는 방식도 허용한다.

Preview용 임시파일 때문에 Batch 결과 파일을 덮어쓰거나 오염시키지 않는다.

### 7. PPTX 내용 치환

각 행마다 원본 Template을 직접 수정하지 않고 작업용 사본/메모리 기반 처리 등 안전한 방법을 사용한다.

- Placeholder 부분만 해당 데이터로 치환한다.
- Placeholder가 아닌 문구는 변경하지 않는다.
- Template의 기본 레이아웃/디자인을 불필요하게 변경하지 않는다.
- 한글 문자열이 깨지지 않아야 한다.

### 8. PDF 일괄생성

각 데이터 행을 순차 또는 안전한 방식으로 처리해 개별 PDF를 만든다.

100행이면 정상 입력 기준 100개의 PDF가 생성되어야 한다.

한 건에서 오류가 발생했을 때 가능한 경우 전체 작업을 즉시 중단하지 말고 해당 건을 실패 처리한 후 나머지 건을 계속 처리한다.

단, Template 자체가 잘못되었거나 필수 Mapping이 누락된 시스템성 오류는 Batch 시작 전에 차단한다.

### 9. 파일명 안전처리

Windows 파일명 금지문자를 안전하게 처리한다.

대표 금지문자:

`\\ / : * ? " < > |`

파일명 값의 앞뒤 불필요한 공백도 안전하게 처리한다.

동일한 결과 파일명이 이미 존재할 경우 기존 파일을 절대 덮어쓰지 않는다.

Demo의 기본 정책은 뒤에 순번을 붙여 보존한다.

예:

- `CERT-001_홍길동_수료증.pdf`
- `CERT-001_홍길동_수료증_2.pdf`

### 10. 진행상태

Batch 처리 중 최소 다음을 보여준다.

- 현재 진행 건수 / 전체 건수
- Progress Bar 또는 동등한 진행 표시

프로그램이 장시간 완전히 멈춘 것처럼 보이지 않도록 GUI 응답성을 고려한다.

### 11. 결과 요약

완료 후 최소 다음을 표시한다.

- 총 대상
- 성공
- 실패

실패가 있다면 개발자가 원인을 확인할 수 있도록 최소한의 오류정보를 남긴다.

고객용 UI에 긴 Stack Trace를 그대로 노출하지 않는다.

### 12. 원본 보호

다음 원본은 절대 수정/삭제하지 않는다.

- 데이터 XLSX
- 데이터 CSV
- PPTX Template

작업 중 생성한 임시 파일은 결과물과 구분하고 정상 종료 시 가능한 범위에서 정리한다.

---

## Demo Sample Asset 요구사항

Claude는 구현과 함께 Demo 검증용 샘플을 만든다.

### `sample_roster.xlsx`

- 100행
- 컬럼: 수료번호 / 이름 / 과정명 / 수료일 / 기관명
- 한글 이름 포함
- 중복 없는 수료번호

### `certificate_template.pptx`

- 1개 Slide의 수료증 형태
- Placeholder 5종 포함
- Demo 기능 검증이 목적이므로 지나친 디자인 작업은 필요하지 않음

### 기대 결과

Online Validation에서는 100행 전체 personalization + 대표 5행 실제 PDF를 생성한다.
향후 고객 Runtime Pilot에서는 100건 실제 PDF Batch를 별도 확인한다.

파일명 예:

- `CERT-2026-001_홍길동_수료증.pdf`
- `CERT-2026-002_김민지_수료증.pdf`

---

## 제외 범위

이번 버전에서 다음을 구현하지 않는다.

- 이메일 자동발송
- 문자/SMS/카카오톡 발송
- QR Code 생성
- Web App
- 서버
- Cloud 업로드
- 계정/Login
- DB
- 외부 API
- ERP/LMS 연결
- 결제
- 라이선스 인증
- 자동 업데이트
- 고객별 Template Editor
- Drag & Drop 디자인 편집
- Placeholder 수동 Mapping Editor
- 다중 PPTX Template 동시 Batch
- 여러 종류 문서를 한 번에 혼합 생성
- 다중 Slide 문서용 고급 편집기
- SmartArt/Chart 등 복잡한 개체 내부 Placeholder 대응
- 최종 설치파일/Installer 제작
- 상용 보안/난독화
- 제품 상세페이지 또는 판매페이지 제작

필요성이 보여도 Claude가 임의로 추가하지 않는다.

---

## 오류/예외 처리 기준

최소 다음 케이스를 처리한다.

1. 데이터 파일 미선택
2. PPTX Template 미선택
3. 출력 폴더 미선택
4. 존재하지 않는 파일
5. XLSX/CSV가 아닌 파일
6. PPTX가 아닌 Template
7. 데이터가 0행
8. Header가 없음
9. Placeholder가 없음
10. Placeholder에 대응되는 컬럼 누락
11. 특정 행의 값이 비어 있음
12. 파일명 금지문자 포함
13. 결과 파일명 중복
14. 출력 폴더 쓰기 권한 없음
15. PDF 생성 중 특정 행 처리 실패
16. Excel/PowerPoint/Office 자동화 환경 문제

빈 셀은 해당 Placeholder를 빈 문자열로 처리할 수 있으나, 수료번호/이름 등 파일명 구성에 필요한 값이 비어 결과 파일명을 안전하게 만들 수 없는 경우 해당 행은 실패 처리하고 이유를 남긴다.

---

## 테스트 요구사항

### TEST 01 — Online 표준 Dataset

입력:

- 정상 `sample_roster.xlsx` 100행
- 정상 `certificate_template.pptx`

기대:

- 100행 전체 로딩/매핑 성공
- 개인화 PPTX Dry-run 100/100 성공
- 대표 5행은 LibreOffice 실제 PDF 5/5 생성 / 실패 0
- 파일명 오류 0

100건 PowerPoint COM PDF 실기는 현재 Validation Demo의 Blocking Gate가 아니다.

### TEST 02 — CSV 입력

동일 데이터를 UTF-8-SIG CSV로 입력한다.

기대:

- XLSX와 동일한 Mapping 동작
- 개인화 PPTX 생성 정상
- PDF Renderer 자체의 실제 출력은 TEST01 대표 Batch에서 검증

### TEST 03 — 필드 누락

Template에는 `{{기관명}}`이 있으나 데이터에서 `기관명` 컬럼 제거.

기대:

- Batch 시작 전 차단
- `기관명` 누락을 사용자에게 표시

### TEST 04 — 잘못된 파일 형식

데이터 또는 Template에 지원하지 않는 확장자를 선택한다.

기대:

- Batch 시작 안 함
- 이해 가능한 오류 표시

### TEST 05 — 파일명 금지문자

이름/수료번호 등에 Windows 금지문자가 포함된 데이터를 사용한다.

기대:

- 안전한 파일명 생성
- 프로그램 Crash 없음

### TEST 06 — 동일 파일명

동일 파일명이 발생하도록 데이터 구성.

기대:

- 기존 PDF 덮어쓰기 없음
- 순번 파일명으로 모두 보존

### TEST 07 — 빈 데이터

Header만 있고 행이 없는 파일.

기대:

- Batch 시작 안 함
- 데이터 없음 표시

### TEST 08 — 일부 빈 셀

일부 데이터 필드가 비어 있는 행 포함.

기대:

- 안전하게 처리
- 파일명 필수값 누락 시 해당 행 실패 기록

### TEST 09 — 원본 보호

Batch 전후 원본 파일 비교.

기대:

- XLSX/CSV/PPTX 원본 변경 없음

### TEST 10 — 한글

한글 이름/과정명/기관명 사용.

기대:

- PDF 내용 및 파일명 한글 깨짐 없음

### TEST 11 — Preview

첫 번째 행 Preview 실행.

기대:

- 첫 행 데이터가 적용된 실제 결과 확인 가능
- Batch 결과에 부작용 없음

### TEST 12 — 부분 실패 격리

가능한 테스트 환경에서 특정 행만 실패하도록 입력 구성.

기대:

- 해당 행 실패 기록
- 가능한 경우 나머지 행 계속 생성

### TEST 13 — 출력 폴더 쓰기 권한 없음

입력:

- 쓰기 권한이 없는 출력 경로 사용

기대:

- 오류를 정상 감지할 것
- 프로그램이 강제 종료되지 않을 것
- 사용자가 이해할 수 있는 오류 메시지를 표시할 것
- 입력 원본 XLSX/CSV/PPTX를 변경하지 않을 것
- 부분 생성물이 발생한 경우 결과에 명확히 기록할 것

PowerPoint COM 전용 항목은 현재 Validation Demo Blocking Gate에서 분리한다.

직접 검증하지 않은 PowerPoint COM 호환성은:

`호환성 미검증: 향후 유료 Pilot/온라인 Windows 환경에서 확인`

으로 표시한다.

온라인 Cloud Acceptance에서 수행하지 않은 항목을 수행한 것처럼 PASS 처리하지 않는다.

---

## 성능/사용성 목표

Validation Demo Online Gate:

- 실제 LibreOffice PDF 대표 Batch: 5건 실측값 기록 (현재 Gate에서는 관찰값, non-blocking)
- Sample Dataset 100행 personalization: 100/100 성공
- 핵심 사용자 조작: 최대 6단계 수준
- 사용자 Code/Script 편집: 0회

100건 PDF 2분 이내 목표는 향후 실제 고객 Runtime/유료 Pilot 단계에서 측정한다. 현재 온라인 Validation Gate에서 달성했다고 추정하지 않는다.

## 산출물

최소 다음을 생성한다.

```text
products/tool01/
├─ README.md
├─ src/
│  └─ 실제 Demo 프로그램 소스
├─ tests/
│  └─ 자동화 가능 테스트 및/또는 테스트 코드
└─ docs/
   ├─ demo_run_guide.md
   └─ test_report.md
```

추가로 Demo용 Sample Asset을 `products/tool01/` 하위의 합리적인 위치에 저장한다.

- `sample_roster.xlsx`
- `certificate_template.pptx`
- 생성 Sample PDF

구체적인 내부 폴더 구성은 Claude가 결정할 수 있으나 `src/`, `tests/`, `docs/`의 의미를 훼손하지 않는다.

---

# [완료 조건 / Definition of Done]

현재 Validation Demo 단계에서는 아래 조건을 모두 충족하면 Claude가 `IMPLEMENTATION COMPLETE — ONLINE VALIDATION DEMO`로 보고할 수 있다.

- [ ] Claude Code Remote/Linux에서 전체 Python Regression Test가 통과한다.
- [ ] GitHub Actions `Tool01 Cloud Acceptance` Workflow가 실행 가능한 상태다.
- [ ] LibreOffice headless로 실제 PPTX → PDF 변환이 확인된다.
- [ ] Sample XLSX 100행 전체를 정상 로딩한다.
- [ ] 100행 전체 Placeholder Mapping/개인화 PPTX Dry-run이 100/100 성공한다.
- [ ] 대표 5행 실제 PDF Batch가 5/5 성공한다.
- [ ] XLSX 입력을 지원한다.
- [ ] CSV 입력을 지원한다.
- [ ] PPTX Template을 선택할 수 있다.
- [ ] `{{필드명}}` Placeholder를 인식한다.
- [ ] 동일 컬럼명과 자동매핑한다.
- [ ] 누락 필드를 Batch 전에 검증한다.
- [ ] Preview personalization/임시파일 정리 경로를 검증한다.
- [ ] 데이터 행별 개별 PDF 생성 로직을 보유한다.
- [ ] 기본 파일명 규칙 `{수료번호}_{이름}_수료증.pdf`를 적용한다.
- [ ] Windows 금지문자를 안전하게 처리한다.
- [ ] 기존 결과 PDF를 덮어쓰지 않는다.
- [ ] 진행상태/성공/실패 UI 로직을 검증한다.
- [ ] 원본 XLSX/CSV/PPTX를 수정하지 않는다.
- [ ] 한글 내용 및 한글 파일명을 실제 Cloud PDF에서 검증한다.
- [ ] 부분 실패 격리/권한 오류/Automation 중복실행을 deterministic test로 검증한다.
- [ ] Sample Asset 100행 + PPTX + 실제 Cloud Sample PDF/PNG를 제공한다.
- [ ] `products/tool01/docs/cloud_acceptance_test.md`에 O01~O09 기준을 문서화한다.
- [ ] `docs/results/`에 `[CLAUDE RESULT]` 형식으로 결과보고를 저장한다.
- [ ] `docs/state/PROJECT_STATE.md`를 결과 기준으로 갱신한다.

### 현재 단계에서 Blocking이 아닌 항목

아래는 Validation Demo를 시장에 제시하기 위한 Blocking Gate가 아니다.

- Windows + PowerPoint COM 실제 100건 Batch
- PowerPoint 기존 세션 보존 실기
- Windows Installer/EXE 패키징
- 최종 고객 Runtime 성능 측정

필요성이 확인되면 유료 Pilot/Full MVP 단계에서 온라인 Windows VM 또는 외부 Pilot 환경으로 검증한다.
사용자 로컬 PC 실행을 완료조건으로 추가하지 않는다.

GPT 검수 전에는 `PASS`, `출시 가능`, `최종 완료`라고 판단하지 않는다.

# [Claude 결과보고 형식]

작업 완료 후 `docs/protocol/result-report-format.md`의 `[CLAUDE RESULT]` 형식을 그대로 사용한다.

특히 다음을 빠뜨리지 않는다.

- 구현 완료 내용
- 실제 수행한 테스트
- 테스트하지 못한 항목
- 발견된 문제
- 미해결 사항
- 변경 제안
- GPT 판단 필요사항
- 생성/수정 파일 목록
- 구현 상태 (`IMPLEMENTATION COMPLETE / BLOCKED / PARTIAL`)
