# PROJECT STATE

> 이 파일은 현재 활성 프로젝트의 Baseline만 관리한다. 과거 프로젝트 상세와 이전 결과보고는 `archive/`, `docs/results/`에 보관한다.

## 프로젝트

- 프로젝트: Make Money — Office Automation Tool Business
- 현재 제품: Tool01 — 수료증·상장·확인서·증명서 대량생성
- 사업 판정: GO
- 구현 상태: `PARTIAL — GitHub Cloud Acceptance Re-run Pending`
- 현재 단계: Validation Demo — Online-only QA 전환 완료 / GitHub Actions 재검증 직전
- 날짜: 2026-08-20

## 사용자 작업 제약 — 최우선

사용자는 **온라인 환경에서만 개발·테스트·검수를 진행한다.**

따라서:

- 사용자 로컬 Windows PC 작업을 요구하지 않는다.
- PowerPoint 설치/수동 W01~W09 수행을 사용자 완료조건으로 두지 않는다.
- Claude Code Remote + GitHub + GitHub Actions에서 현재 단계의 Blocking QA를 끝낸다.
- Windows/PowerPoint COM 호환성 실기가 향후 필요하면 온라인 Windows VM 또는 외부 Pilot 환경을 사용한다.

이 제약을 바꾸려면 사용자 승인 필요.

## 사업 Baseline

- 사업 방향: 직장인·소규모 사업자의 반복업무를 줄여주는 Office Automation Tool 판매 + 기존 Tool 기반 맞춤 자동화
- 운영 제약: 하루 약 1시간
- 초기 자본: 100만원 이하
- 제외 분야: 반도체 장비 영업 관련 수익화 아이디어
- 원칙: 저단가 맞춤노동, 반복 수작업, 과도한 고객응대가 필요한 모델은 낮게 평가한다.

## Tool01 Primary ICP

- 소규모 민간 교육기관
- 교육대행사
- 직무교육업체
- 조직 규모: 약 1~20명
- 사용빈도: 월 1회 이상
- 1회 문서 발급량: 약 30~500건

## 핵심 가치 제안

`XLSX/CSV 명단 + PPTX 템플릿`을 선택하면 `{{필드명}}` Placeholder를 자동매핑하고 개인별 PDF 생성 흐름을 자동화한다.

사용자는 VBA, 코드, Apps Script, Mail Merge 설정을 직접 다루지 않는다.

## 가격 Baseline

- Validation 단품 판매가격: 29,000원
- Custom Template Setup: +79,000원
- Custom Template 포함 가설 총액: 108,000원
- Validation 통과 후 목표 단품가격: 39,000원

가격 변경은 사용자/GPT 승인사항이다.

## 승인된 문서

- Brief: `docs/briefs/2026-08-20_tool01_demo_brief.md`
- Spec: `docs/specs/2026-08-20_tool01_demo_spec.md`
- Online QA: `products/tool01/docs/cloud_acceptance_test.md`
- Optional Windows compatibility: `products/tool01/docs/windows_acceptance_test.md`
- 구현 대상: `products/tool01/`

## Renderer Baseline

### Blocking QA Renderer

**LibreOfficeRenderer**

- Claude Code Remote/Linux 실행 가능
- GitHub Actions 실행 가능
- 실제 PPTX → PDF 생성
- 현재 Validation Demo의 Blocking QA 기준

### Compatibility Renderer

**PowerPointRenderer**

- Windows + Microsoft PowerPoint Desktop
- PowerPoint COM
- `Application.Quit()` 호출 금지
- Tool01이 직접 연 Temporary Presentation만 Close
- 시장 Validation 진입의 Blocking Gate가 아님

Renderer 선택:

- `TOOL01_RENDERER=libreoffice` — Cloud Acceptance
- `TOOL01_RENDERER=powerpoint` — Windows compatibility
- `TOOL01_RENDERER=auto` — Windows=PowerPoint / non-Windows=LibreOffice

## Online QA 구현 상태

현재 Repository에는 다음이 포함된다.

1. LibreOffice headless Renderer Adapter
2. PowerPoint COM Compatibility Adapter
3. Renderer 자동선택
4. GitHub Actions Workflow: `.github/workflows/tool01-cloud-acceptance.yml`
5. Cloud Acceptance Script: `products/tool01/scripts/cloud_acceptance.py`
6. Online review artifact 생성: PDF + PNG
7. Sample Asset: `sample_roster.xlsx`, `certificate_template.pptx`
8. Sample Output: `assets/sample_output/cloud_preview.pdf`, `cloud_preview.png`

## 현재 검증 Baseline

GPT 수정환경에서 Online Acceptance 경로를 실제 실행해 다음을 확인했다.

- Python headless Regression: **37 passed / 4 skipped**
- Xvfb GUI-only Regression: **3 passed**
- LibreOffice 실제 PDF 대표 Batch: **5/5 성공**
- Sample 100행 personalization PPTX Dry-run: **100/100 성공**
- 한글 출력 파일명: 확인
- 실제 PDF 내 한글 5개 필드: 확인
- 대표 5건 처리시간: 약 **11.90초** (관찰값, non-blocking)
- Cloud Preview PDF/PNG 생성: 확인

이 결과는 GitHub Actions 업로드 후 동일 Workflow로 재검증해야 한다.

## 현재 Blocking Gate

GitHub에 Repository를 업로드한 뒤:

1. `Tool01 Cloud Acceptance` GitHub Actions 실행
2. Regression Test 통과
3. Cloud Acceptance O01~O09 통과
4. Artifact의 `cloud_preview.png` 또는 PDF를 온라인에서 확인
5. Claude가 결과문서/PROJECT_STATE를 갱신
6. GPT 검수

위 조건이 충족되면 현재 Validation Demo는 `IMPLEMENTATION COMPLETE — ONLINE VALIDATION DEMO` 후보가 된다.

## Windows Compatibility 상태

아래는 **호환성 미검증 / 현재 non-blocking**이다.

- PowerPoint COM 실제 Preview PDF
- PowerPoint COM 100건 Batch
- 기존 사용자 PowerPoint Session 보존 실기
- Windows 실제 성능
- Installer/EXE

사용자가 로컬에서 테스트하지 않는다.
유료 Pilot/Full MVP 단계에서 필요할 경우 온라인 Windows VM 또는 외부 Pilot 환경으로 별도 검증한다.

## 다음 작업

신규 기능 개발보다 먼저 GitHub 온라인 Gate를 완료한다.

1. GitHub 업로드
2. Claude Code Remote에서 문서 정합성 확인
3. Headless Regression 실행
4. GitHub Actions Cloud Acceptance 실행/결과 확인
5. 실패 시 Brief/Spec 범위 안 Blocking Bug만 수정
6. `cloud_acceptance_test.md`, `test_report.md`, `PROJECT_STATE.md` 갱신
7. `docs/results/`에 `[CLAUDE RESULT]` 저장
8. GPT 검수

사용자에게 로컬 명령 실행을 요청하지 않는다.

## 제외 범위

- 이메일/문자 자동발송
- QR
- Login/회원
- DB
- 고객용 Web App/Hosted Service
- ERP/LMS 연동
- 결제
- Template Editor
- 최종 Installer/상용 패키징

> LibreOffice는 고객용 Web/Cloud 서비스가 아니라 **온라인 개발·검수 Renderer Adapter**로 허용한다.

## 사업 Validation 다음 Gate

Online Validation Demo를 타겟 잠재고객 10곳에 제시한다.

다음 중 하나를 달성하면 Full MVP 진행을 `GO` 판단 대상으로 올린다.

- 29,000원 유료 Pilot 2건 이상
- Custom Template Setup 유료 요청 1건 이상

유료 Pilot이 발생하면 그때 실제 고객 Runtime/Windows compatibility 검증의 우선순위를 다시 판단한다.
