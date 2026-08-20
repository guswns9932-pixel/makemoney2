# Claude Code 시작 프롬프트 — Tool01 Online Cloud Acceptance

아래 내용을 Claude Code Remote 새 세션의 첫 메시지로 그대로 입력한다.

---

이 Repository의 `CLAUDE.md`를 최우선 지침으로 따르고 작업을 시작해.

현재 목표는 **신규 기능 개발이 아니라 Tool01 Validation Demo의 Online Cloud Acceptance Gate를 완료하는 것**이다.

사용자는 **온라인 환경에서만 작업**한다. 사용자에게 로컬 Windows PC, PowerPoint 설치, 로컬 명령 실행 또는 W01~W09 수동 테스트를 요구하지 마.

현재 상태는:

`PARTIAL — GitHub Cloud Acceptance Re-run Pending`

이다. GPT 검수 전에는 최종 완료/출시 가능으로 간주하지 않는다.

## 1. 작업 시작 전 필수 확인

먼저 아래 파일을 읽고 서로 일치하는지 확인해.

1. `CLAUDE.md`
2. `docs/state/PROJECT_STATE.md`
3. `docs/briefs/2026-08-20_tool01_demo_brief.md`
4. `docs/specs/2026-08-20_tool01_demo_spec.md`
5. `products/tool01/docs/cloud_acceptance_test.md`
6. `products/tool01/docs/test_report.md`
7. `.github/workflows/tool01-cloud-acceptance.yml`

그다음 `git status`와 현재 branch를 확인해 기존 사용자 변경사항을 보호해.

문서 간 제품 목적, Primary ICP, 가격, Scope, 현재 단계, Renderer 정책, Definition of Done에 충돌이 없으면 재확인 질문 없이 진행해.

충돌이 있으면 구현하지 말고:

`BLOCKED`

및

`[GPT 판단 필요사항]`

으로 보고해.

## 2. 승인된 Online-only Baseline

임의로 변경하지 마.

- 제품: Tool01 — 수료증·상장·확인서·증명서 대량생성
- 입력: XLSX/CSV + PPTX Template
- 출력: 개인별 PDF
- Placeholder: `{{필드명}}`
- Validation 가격: 29,000원
- Custom Template Setup: +79,000원
- 현재 Blocking QA Renderer: `LibreOfficeRenderer`
- Cloud 실행환경: Claude Code Remote/Linux + GitHub Actions
- Cloud Acceptance 시 `TOOL01_RENDERER=libreoffice`
- `PowerPointRenderer`는 Windows Compatibility Adapter로 유지
- PowerPoint COM 실기는 현재 Validation Demo의 Blocking Gate가 아님
- 사용자 로컬 Windows 테스트는 요구하지 않음

다음 기능은 현재 Scope가 아니다.

- 고객용 Web App/Hosted Service
- Login/회원/DB
- 이메일/문자 자동발송
- QR
- 결제
- ERP/LMS 연동
- Template Editor
- Installer/EXE

LibreOffice는 고객용 Cloud 서비스 기능이 아니라 **온라인 개발·검수용 Renderer Adapter**다.

## 3. Remote Regression Test

`products/tool01/`에서 먼저 전체 Regression을 실행해.

```bash
TOOL01_RENDERER=libreoffice python -m pytest tests/ -q -m "not live_renderer"
```

GUI regression은 별도로 다음처럼 실행해.

```bash
xvfb-run -a python -m pytest tests/test_main_gui.py -q
```

`live_renderer` pytest는 Regression에서 제외한다. 실제 LibreOffice 렌더링은 Cloud Acceptance Script가 담당한다.

현재 참고 Baseline:

- Headless: `37 passed / 4 skipped`
- Xvfb GUI-only: `3 passed`

환경에 따라 headless skip 숫자는 달라질 수 있다. 숫자를 억지로 맞추지 말고 unexpected FAIL/ERROR가 있는지를 판단해.

Regression 실패가 실제 코드 문제라면 Brief/Spec 범위 내 Blocking Bug만 수정하고 전체 Regression을 다시 실행해.

## 4. Remote Cloud Acceptance

Regression 통과 후 다음을 실행해.

```bash
TOOL01_RENDERER=libreoffice python scripts/cloud_acceptance.py \
  --output artifacts/cloud_acceptance
```

`products/tool01/docs/cloud_acceptance_test.md`의 O01~O09 기준으로 판정해.

반드시 확인할 것:

- O01: Sample XLSX 100행 + Placeholder 5개 자동매핑
- O02: Preview 개인화/정리 경로
- O03: LibreOffice 실제 PDF 5/5 + personalized PPTX 100/100
- O04: CSV 입력 경로
- O05: 한글 PDF 파일명 + PDF 내부 5개 필드
- O06: 부분 실패 격리
- O07: PermissionError 처리
- O08: Preview/Batch 중복 실행 방지
- O09: 실제 PDF 대표 5건 처리시간 기록 — non-blocking

온라인 Acceptance에서 직접 수행하지 않은 항목을 PASS로 추측하지 마.

## 5. GitHub Actions Gate

Remote 실행이 정상이라면 `.github/workflows/tool01-cloud-acceptance.yml`도 점검해.

가능하면 현재 branch를 GitHub에 push한 뒤 **Tool01 Cloud Acceptance** Workflow를 실행하거나 push로 자동 실행되게 하고 결과를 확인해.

GitHub 권한/인증이 없어 push 또는 Actions 실행/조회가 불가능하면 이를 코드 실패로 처리하지 말고:

`BLOCKED — GitHub permission/authentication`

으로 명확히 보고해. 사용자 로컬 실행으로 우회하지 마.

Workflow 성공 시 Artifact를 확인해.

필수 확인 대상:

- `regression.log`
- `cloud_acceptance.log`
- `cloud_acceptance_report.md`
- `cloud_acceptance_report.json`
- `cloud_preview.pdf`
- `cloud_preview.png`
- 실제 Batch PDF 5개
- 100행 personalized PPTX dry-run

## 6. Online Visual QA

`cloud_preview.png` 또는 `cloud_preview.pdf`를 확인할 수 있으면 온라인에서 직접 검토해.

확인 항목:

- 수료증 레이아웃이 크게 깨지지 않음
- Placeholder 값이 정상 위치에 들어감
- 한글 깨짐 없음
- 글자가 잘려 핵심정보를 읽을 수 없는 문제 없음
- 배경/도형/기본 디자인이 유지됨

Claude가 시각적으로 확인할 수 없는 환경이면 해당 항목만:

`USER CHECK REQUIRED — GitHub Artifact에서 cloud_preview.png 확인`

으로 표시해.

사용자에게 로컬 프로그램 실행을 요구하지 말고 **GitHub 웹 Artifact에서 확인하는 방법만 최대 3단계로 안내**해.

## 7. Bug 발견 시 수정 권한

현재 Brief/Spec을 충족하기 위한 Blocking Bug는 직접 수정 가능하다.

Claude 자율 수정 가능:

- LibreOffice Renderer 안정성
- 파일/경로/임시파일 처리
- 예외처리
- 테스트 코드
- GitHub Actions 안정성
- Thread/GUI regression
- 현재 명세 충족을 위한 내부 함수/모듈 수정

임의 변경 금지:

- 제품 목적
- Primary ICP
- 가격
- 사업모델
- 핵심 기능
- Validation MVP 범위
- 주요 UX Flow
- 제외 기능 신규 추가
- PowerPoint Compatibility Adapter 삭제

범위 변경이 필요하면 구현하지 말고 `[변경 제안]` 또는 `[GPT 판단 필요사항]`으로 보고해.

## 8. 완료 후 문서 갱신

Online Gate 결과에 맞춰 아래를 갱신해.

1. `products/tool01/docs/cloud_acceptance_test.md`
2. `products/tool01/docs/test_report.md`
3. `products/tool01/README.md`
4. `docs/state/PROJECT_STATE.md`
5. `docs/results/`에 새로운 `[CLAUDE RESULT]` 파일

GitHub Actions까지 성공하고 Online Artifact 검수도 완료됐으면 상태를:

`IMPLEMENTATION COMPLETE — ONLINE VALIDATION DEMO`

후보로 보고해.

GitHub Actions 또는 Online Visual QA 등 Blocking Gate가 남아 있으면:

`PARTIAL`

을 유지해.

PowerPoint COM/Windows 항목만 미검증인 경우에는 현재 Gate를 PARTIAL로 유지하는 사유로 사용하지 말고 다음처럼 별도 기록해.

`호환성 미검증: 향후 유료 Pilot/온라인 Windows 환경에서 확인`

## 9. 결과보고 형식

반드시 기존 `[CLAUDE RESULT]` 형식을 따른다.

최소 포함:

- 프로젝트
- 버전/작업단계
- 구현/수정 완료 내용
- Regression 결과
- O01~O09 각각의 결과
- GitHub Actions 결과 및 run 정보
- Online Artifact 목록
- Visual QA 결과 또는 USER CHECK REQUIRED
- 변경사항
- 발견된 문제
- 미해결 사항
- Windows Compatibility 미검증 항목
- 변경 제안
- GPT 판단 필요사항
- 현재 구현 상태

## 10. Git 안전규칙

`docs/protocol/git-workflow.md`를 따른다.

특히:

- 기존 사용자 변경 삭제 금지
- `git reset --hard` 금지
- force push 금지
- 비밀정보 commit 금지
- Scope 밖 파일 임의 수정 금지
- 현재 작업 branch 사용
- main 직접 merge 금지

GitHub Actions 실행을 위해 현재 작업 branch에 commit/push가 필요한 경우, 현재 명세 범위의 코드/문서 변경만 포함해서 수행해도 된다. main merge는 하지 마.

지금부터 위 절차대로 Online Cloud Acceptance Gate를 완료해.
