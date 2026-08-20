# Tool01 — 수료증·상장·확인서·증명서 대량생성

현재 Make Money 프로젝트의 활성 제품.

## 목적

Excel/CSV 명단과 PowerPoint Template을 이용해 이름, 수료번호, 과정명, 날짜, 기관명 등이 개인화된 PDF를 대량 생성한다.

사용자는 코드/VBA/Apps Script/Mail Merge를 수정하지 않는다.

## 현재 단계

**Validation Demo — Online-only Cloud Acceptance 준비 완료**

현재 Repository는 사용자가 로컬 Windows PC를 쓰지 않아도 Claude Code Remote/GitHub에서 구현·검수할 수 있도록 구성되어 있다.

현재 상태는 `docs/state/PROJECT_STATE.md`를 기준으로 한다.

## Renderer 구조

### LibreOfficeRenderer — 현재 Blocking QA

- Linux/Claude Code Remote
- GitHub Actions
- 실제 PPTX → PDF
- `TOOL01_RENDERER=libreoffice`

### PowerPointRenderer — Windows Compatibility Adapter

- Windows + PowerPoint Desktop
- PowerPoint COM
- `TOOL01_RENDERER=powerpoint`
- 현재 Validation Demo의 Blocking Gate 아님

### Auto

- Windows → PowerPoint
- non-Windows → LibreOffice

## Online Acceptance

문서:

`docs/cloud_acceptance_test.md`

실행:

```bash
python -m pip install -r requirements.txt
TOOL01_RENDERER=libreoffice python -m pytest tests/ -q
TOOL01_RENDERER=libreoffice python scripts/cloud_acceptance.py --output artifacts/cloud_acceptance
```

현재 GPT 검수환경 Baseline:

- headless Regression: **37 passed / 4 skipped**
- Xvfb GUI-only Regression: **3 passed**
- 실제 LibreOffice PDF 대표 Batch: **5/5 성공**
- 100행 personalized PPTX Dry-run: **100/100 성공**
- 대표 5건: **11.90초** 관찰값
- 실제 한글 PDF/파일명: PASS

GitHub 업로드 후 Actions에서 동일 Gate를 재실행한다.

## GitHub Actions

`.github/workflows/tool01-cloud-acceptance.yml`

성공 시 Artifact:

- `cloud_acceptance_report.md/json`
- `cloud_preview.pdf/png`
- 실제 PDF 5개
- 100행 personalized PPTX dry-run

## Sample Asset

- `assets/sample_roster.xlsx` — 100행
- `assets/certificate_template.pptx` — Placeholder 5종
- `assets/sample_output/cloud_preview.pdf`
- `assets/sample_output/cloud_preview.png`

## Demo 핵심 흐름

`XLSX/CSV → PPTX Template → 자동매핑 → Preview → 개인별 PDF Batch`

## Validation 가격

- 단품: 29,000원
- Custom Template Setup: +79,000원

가격은 GPT/사용자 승인 없이 변경하지 않는다.

## Scope 주의

현재는 Validation Demo다.

다음은 임의로 추가하지 않는다.

- 고객용 Web App/Hosted Service
- Login/DB
- 이메일/문자 발송
- QR
- 결제
- ERP/LMS
- Template Editor
- Installer/상용 패키징

LibreOffice는 고객용 Cloud Service 기능이 아니라 **온라인 개발·검수 Renderer Adapter**다.

## Windows Compatibility

`docs/windows_acceptance_test.md`는 optional/non-blocking 문서다.

사용자가 로컬 Windows에서 수행하지 않는다.
유료 Pilot/Full MVP에서 필요성이 확인되면 온라인 Windows VM/외부 Pilot 환경을 별도 판단한다.
