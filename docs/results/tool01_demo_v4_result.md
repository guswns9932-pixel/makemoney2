[CLAUDE RESULT]

- 프로젝트: Make Money — Office Automation Tool Business / Tool01 — 수료증·상장·확인서·증명서 대량생성
- 버전/작업단계: v4 — Tool01 Validation Demo Online Cloud Acceptance Gate 완료

### 구현 완료 내용

- `.github/workflows/tool01-cloud-acceptance.yml`을 신규 작성. `docs/state/PROJECT_STATE.md`, `products/tool01/docs/cloud_acceptance_test.md`, `products/tool01/docs/test_report.md`는 모두 이 Workflow가 이미 Repository에 존재한다고 기록하고 있었으나 실제 Repository에는 파일이 없었다. Checkout → Python 3.12 → LibreOffice/Xvfb/한글 폰트(fonts-nanum)/poppler-utils/python3-tk 설치 → Logical Regression → Xvfb GUI Regression → Cloud Acceptance O01~O09 → 로그/Artifact 업로드(`if: always()`) 순으로 구성했다.
- `products/tool01/scripts/cloud_acceptance.py`의 O05(한글 내용 검증) 로직을 수정. `pypdf`의 glyph-position 기반 텍스트 추출이 `CERT-2026-001`을 `CERT -2026-001`처럼 하이픈 앞에 공백을 삽입하는 경우가 있었다(`pdftotext -layout`으로 실제 PDF에는 해당 공백이 없음을 대조 확인). 실제 렌더링 콘텐츠는 정상이므로 값 비교 시 공백을 제거하고 비교하도록 변경해 이 추출 아티팩트로 인한 오탐(false negative)을 제거했다.
- 루트 `.gitignore` 신규 추가 (`__pycache__/`, `.pytest_cache/`, `products/tool01/artifacts/`).

### 테스트 결과

Claude Code Remote 환경에서 1차로 직접 실행해 재현했고, 이후 GitHub Actions에서 동일 결과를 재검증했다(Run #1 push / Run #2 pull_request, 둘 다 `completed / success`).

- TEST 01 (Online 표준 Dataset): PASS — 100행 로딩/매핑 성공, 개인화 PPTX Dry-run 100/100, 대표 5행 실제 LibreOffice PDF 5/5, 파일명 오류 0
- TEST 02 (CSV 입력): PASS — XLSX와 동일 Mapping, UTF-8-SIG CSV 10행 Preflight/개인화 성공
- TEST 03~13, 한글, Preview, 부분 실패, 권한 오류, 원본 보호 등: pytest Regression으로 검증 (아래 O01~O09 표 참고)
- Logical Regression (`TOOL01_RENDERER=libreoffice python -m pytest tests/ -q -m "not live_renderer"`): GitHub Actions에서 **38 passed / 3 skipped / 3 deselected** (Remote/Linux 로컬 재현 시 37 passed / 4 skipped — skip 개수는 환경별 차이이며 문서에 명시된 대로 unexpected FAIL/ERROR는 없음)
- Xvfb GUI Regression (`xvfb-run -a python -m pytest tests/test_main_gui.py -q`): **3 passed**
- Cloud Acceptance O01~O09 (`TOOL01_RENDERER=libreoffice TOOL01_LO_BATCH_SIZE=5 python scripts/cloud_acceptance.py`):

| ID | 검증 | 결과 |
|---|---|---|
| O01 | Sample 100행 Preflight/자동매핑 | PASS — 100행 + Placeholder 5개 |
| O02 | Preview 개인화/정리 경로 | PASS |
| O03 | 실제 PDF 5/5 + 100행 개인화 dry-run 100/100 | PASS |
| O04 | CSV 입력 | PASS |
| O05 | 한글 파일명 + PDF 내부 5개 필드 | PASS (pypdf 추출 오탐 수정 후) |
| O06 | 부분 실패 격리 | PASS |
| O07 | 권한 오류 | PASS |
| O08 | Automation 중복 방지 | PASS |
| O09 | 성능 관찰 (non-blocking) | PASS — 대표 5건 2.88초 (GitHub Actions Run #2 관찰값) |

- 실제 LibreOffice PDF 대표 Batch: **5/5**
- Personalized PPTX 100행 dry-run: **100/100**
- 한글 파일명/PDF 내부 텍스트: 확인
- Online Visual QA: `cloud_preview.png`를 GitHub Actions Artifact에서 직접 열어 확인 — 수료증 레이아웃 정상, Placeholder 5개 값이 정상 위치에 삽입됨, 한글 깨짐 없음, 텍스트 잘림 없음, 테두리/기본 디자인 유지. **PASS** (직접 확인, USER CHECK REQUIRED 아님)

### GitHub Actions 결과

- Workflow: `Tool01 Cloud Acceptance` (`.github/workflows/tool01-cloud-acceptance.yml`)
- Run #1 (push, run_number 1): `completed / success` — https://github.com/guswns9932-pixel/makemoney2/actions/runs/32324342906
- Run #2 (pull_request, run_number 2): `completed / success` — https://github.com/guswns9932-pixel/makemoney2/actions/runs/32324360773
- Commit SHA: `6a3c4908eada3125ed03caaae65f82038e0cfa01`
- Branch: `claude/tool01-cloud-acceptance-gate-a2fsco`
- PR: `#1` (Draft) — https://github.com/guswns9932-pixel/makemoney2/pull/1
- Artifact: `tool01-cloud-acceptance` (2.65MB) — `regression.log`, `gui_regression.log`, `cloud_acceptance.log`, `cloud_acceptance_report.md/json`, `cloud_preview.pdf/png`, 실제 Batch PDF 5개, 100행 personalized PPTX dry-run 전량 확인함

### 변경사항

- 기존 대비 변경된 내용:
  - `.github/workflows/tool01-cloud-acceptance.yml` 신규 생성 (기존 문서들이 이미 존재한다고 잘못 기록하고 있던 파일)
  - `products/tool01/scripts/cloud_acceptance.py`의 O05 검증 로직에 공백 정규화 비교 추가 (기능 자체 변경 없음, 검증 정확도 수정)
  - 루트 `.gitignore` 신규 생성
  - `products/tool01/docs/cloud_acceptance_test.md`, `products/tool01/docs/test_report.md`, `products/tool01/README.md`, `docs/state/PROJECT_STATE.md` 결과 갱신
  - 브리프/실행 명세의 제품 목적/ICP/가격/Scope/Renderer 정책은 변경하지 않음

### 발견된 문제

- `.github/workflows/tool01-cloud-acceptance.yml`가 여러 문서(PROJECT_STATE.md, cloud_acceptance_test.md, test_report.md)에 "이미 Repository에 포함되어 있다"고 기록되어 있었으나 실제로는 존재하지 않았음. 이번 작업에서 명세(O01~O09, Definition of Done)에 맞춰 신규 작성해 Blocking Gate를 완료했다.
- `products/tool01/scripts/cloud_acceptance.py`의 O05 검증이 pypdf 텍스트 추출 아티팩트(`CERT -2026-001`처럼 하이픈 앞 공백 삽입)로 인해 오탐 실패했음. 실제 PDF 콘텐츠는 정상임을 `pdftotext -layout`으로 대조 확인 후 비교 로직을 공백-정규화 방식으로 수정.
- Remote 환경 자체에 LibreOffice Impress 필터(`libreoffice-impress`), `python3-tk`, `cffi`(Python 3.11용) 등이 기본 미설치 상태였음. 이는 코드 결함이 아니라 세션 환경 패키지 부재였으며, 로컬 재현을 위해 설치함. GitHub Actions Workflow에는 필요한 패키지 설치 스텝을 포함해 별도 조치가 필요 없도록 구성함.

### 미해결 사항

- 없음 (Blocking Gate 기준)

### Windows Compatibility

`Compatibility Pending — 향후 유료 Pilot 또는 온라인 Windows 환경에서 검증`

다음은 현재 non-blocking이며 검증하지 않았다.

- PowerPoint COM 실제 Preview/Batch PDF 렌더링
- PowerPoint COM 100건 Batch
- 기존 사용자 PowerPoint Session 보존 실기
- Windows 실제 성능
- Installer/EXE

### 변경 제안

- 없음

### GPT 판단 필요사항

- 없음. 현재 Online Validation Demo Gate(GitHub Actions 실행 + O01~O09 + Visual QA + 문서 갱신)는 모두 완료됐다. GPT 최종 검수만 남아 있다.

### 산출물

- 신규: `.github/workflows/tool01-cloud-acceptance.yml`
- 신규: `.gitignore`
- 수정: `products/tool01/scripts/cloud_acceptance.py`
- 수정: `products/tool01/docs/cloud_acceptance_test.md`
- 수정: `products/tool01/docs/test_report.md`
- 수정: `products/tool01/README.md`
- 수정: `docs/state/PROJECT_STATE.md`
- 신규: `docs/results/tool01_demo_v4_result.md` (본 문서)
- GitHub: Branch `claude/tool01-cloud-acceptance-gate-a2fsco`, PR #1 (Draft), Actions Run #1/#2 (success)

### 구현 상태

`IMPLEMENTATION COMPLETE — ONLINE VALIDATION DEMO` (Candidate, GPT 최종 승인 전)
