# Tool01 Demo 테스트 리포트 — Online-only Baseline

날짜: 2026-08-20
현재 구현 상태: `PARTIAL — GitHub Cloud Acceptance Re-run Pending`

이 문서는 **현재 Validation Demo의 Blocking QA**만 기록한다. 사용자는 로컬 Windows PC에서 테스트하지 않는다.

## 1. 현재 QA 정책

현재 Blocking Gate:

- Claude Code Remote / Linux
- GitHub Actions `ubuntu-latest`
- LibreOffice Impress Headless
- Python + pytest
- `products/tool01/scripts/cloud_acceptance.py`

현재 non-blocking Compatibility Gate:

- Windows + Microsoft PowerPoint Desktop
- PowerPoint COM 실제 Preview/Batch
- Windows Installer/EXE

PowerPoint COM 항목은 유료 Pilot/Full MVP에서 필요성이 확인되면 온라인 Windows VM 또는 외부 Pilot 환경으로 검증한다.

## 2. Regression Test

실행 명령:

```bash
cd products/tool01
TOOL01_RENDERER=libreoffice python -m pytest tests/ -q -m "not live_renderer"
```

GPT 온라인 검수환경에서 확인된 Baseline:

- Headless: **37 passed / 4 skipped**
- Xvfb GUI-only: **3 passed**

실 LibreOffice를 직접 기동하는 pytest 3건은 `live_renderer` marker로 Regression Gate에서 제외한다. 실제 Renderer 검증은 바로 뒤의 Cloud Acceptance O03에서 수행한다. 이 분리는 Office 하위 프로세스가 pytest/Xvfb 종료를 지연시키는 CI 불안정성을 피하기 위한 것이다.

Skip 수는 실행환경에 따라 달라질 수 있다. 핵심 판정은 **unexpected FAIL/ERROR가 없는가**이다.

### 자동 검증 범위

- XLSX/CSV 로딩
- Placeholder 추출/자동매핑
- 잘못된 입력/빈 입력
- 파일명 안전처리/중복방지
- 원본 보호
- Preview 임시파일 정리
- Preview/Batch 중복 실행 방지
- Worker Thread 오류 callback
- PermissionError 처리 경로
- 부분 실패 격리
- Renderer 선택/오류 처리
- 기존 기능 Regression

## 3. Cloud Acceptance — O01~O09

실행 명령:

```bash
cd products/tool01
TOOL01_RENDERER=libreoffice python scripts/cloud_acceptance.py \
  --output artifacts/cloud_acceptance
```

현재 기준:

| ID | 검증 | 현재 Baseline |
|---|---|---|
| O01 | Sample 100행 Preflight/자동매핑 | PASS — 100행 + Placeholder 5개 |
| O02 | Preview 개인화/정리 경로 | PASS — 개인화 + regression evidence |
| O03 | 실제 PDF + 100행 개인화 | PASS — LibreOffice PDF 5/5 + PPTX 100/100 |
| O04 | CSV 입력 | PASS — UTF-8-SIG CSV Preflight/개인화 |
| O05 | 한글 | PASS — 실제 PDF 파일명 + PDF 내부 5개 필드 |
| O06 | 부분 실패 격리 | PASS — deterministic pytest |
| O07 | 권한 오류 | PASS — PermissionError mock/validator |
| O08 | Automation 중복 방지 | PASS — GUI regression |
| O09 | 성능 관찰 | PASS(관찰) — 대표 5건 처리시간 기록, non-blocking |

GPT 검수환경의 이전 관찰값:

- 실제 LibreOffice PDF: **5/5**
- 100행 personalized PPTX dry-run: **100/100**
- 대표 5건 실제 PDF 처리시간: 약 **11.90초**
- 한글 내용/파일명: PASS

이 수치는 GitHub Actions에서 다시 실행하여 재확인해야 한다. GitHub Actions 결과를 통과하기 전에는 Online Validation Demo를 최종 완료로 판정하지 않는다.

## 4. GitHub Actions Gate

Workflow:

`.github/workflows/tool01-cloud-acceptance.yml`

Workflow는 다음을 수행한다.

1. Repository Checkout
2. Python 3.12 구성
3. LibreOffice / Xvfb / 한글 폰트 / poppler 설치
4. Python dependency 설치
5. pytest Regression
6. 실제 Cloud Acceptance O01~O09
7. 성공/실패 여부와 관계없이 실행 로그 및 생성 Artifact 업로드

기대 Artifact:

- `regression.log`
- `gui_regression.log`
- `cloud_acceptance.log`
- `cloud_acceptance_report.json`
- `cloud_acceptance_report.md`
- `cloud_preview.pdf`
- `cloud_preview.png`
- 실제 Batch PDF 5개
- 100행 personalized PPTX dry-run

## 5. Sample Output

Repository에 온라인 검수용 Sample Output을 포함한다.

- `assets/sample_output/cloud_preview.pdf`
- `assets/sample_output/cloud_preview.png`

이는 LibreOffice Cloud Acceptance의 예시 출력이다. PowerPoint COM 렌더링 결과를 의미하지 않는다.

## 6. 현재 미검증 — Non-blocking Compatibility

다음은 현재 Validation Demo의 Blocking Gate가 아니다.

- PowerPoint COM 실제 PDF 렌더링 충실도
- PowerPoint COM 100건 Batch
- 기존 사용자 PowerPoint Session 보존 실기
- Windows 실제 성능
- Installer/EXE

보고 표현:

`호환성 미검증: 향후 유료 Pilot/온라인 Windows 환경에서 확인`

사용자에게 로컬 Windows 테스트를 요구하지 않는다.

## 7. 현재 완료 판단

현재 단계에서 `IMPLEMENTATION COMPLETE — ONLINE VALIDATION DEMO` 후보가 되려면 다음이 모두 필요하다.

- [ ] GitHub Actions `Tool01 Cloud Acceptance` 성공
- [ ] Regression unexpected FAIL 없음
- [ ] O01~O09 PASS
- [ ] GitHub Artifact 생성 확인
- [ ] `cloud_preview.png` 또는 PDF 온라인 품질 확인
- [ ] Claude가 결과 문서와 `PROJECT_STATE.md` 갱신
- [ ] GPT 최종 검수

그 전까지 상태는:

`PARTIAL — GitHub Cloud Acceptance Re-run Pending`
