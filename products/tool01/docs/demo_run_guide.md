# Tool01 Demo 실행 가이드

## 권장 방식 — Online-only

현재 사용자는 로컬 PC가 아니라 GitHub + Claude Code Remote에서 작업한다.

### Claude Code Remote / Linux

```bash
cd products/tool01
python -m pip install -r requirements.txt
TOOL01_RENDERER=libreoffice python -m pytest tests/ -q
TOOL01_RENDERER=libreoffice python scripts/cloud_acceptance.py --output artifacts/cloud_acceptance
```

OS 패키지로 LibreOffice가 필요하다.
GitHub Actions에서는 Workflow가 자동 설치한다.

### GitHub Actions

Workflow:

`.github/workflows/tool01-cloud-acceptance.yml`

Repository에 push하면 Tool01 관련 변경 시 자동 실행되며, 필요하면 GitHub Actions 화면에서 `Run workflow`로 수동 실행할 수도 있다.

성공하면 다음 Artifact가 생성된다.

- Cloud Acceptance Report
- 실제 PDF 5개
- Cloud Preview PDF/PNG
- 100행 personalized PPTX dry-run 결과

## 온라인에서 결과 확인

GitHub Actions 실행 결과의 Artifact를 브라우저에서 확인한다.

특히 `cloud_preview.png`는 별도 프로그램 설치 없이 브라우저에서 디자인/한글/배치를 확인하기 위한 Review Artifact다.

Repository에도 Baseline 샘플이 있다.

- `assets/sample_output/cloud_preview.pdf`
- `assets/sample_output/cloud_preview.png`

## GUI 제품 흐름

고객용 GUI의 기본 흐름은 유지한다.

1. 데이터 파일 선택
2. PPTX Template 선택
3. 출력 폴더 선택
4. 파일 확인 및 자동매핑
5. Preview
6. PDF 일괄생성
7. 결과 폴더 확인

Cloud Acceptance는 GUI를 고객처럼 직접 조작하는 대신 동일 핵심 모듈과 GUI 상태 로직을 자동화 테스트한다.

## Sample Asset

- `assets/sample_roster.xlsx`: 100행
- `assets/certificate_template.pptx`: Placeholder 5종
- `assets/sample_output/cloud_preview.*`: 실제 LibreOffice Cloud Renderer 결과

## Renderer

### Online QA

`TOOL01_RENDERER=libreoffice`

LibreOffice Impress Headless를 사용한다.

### Windows Compatibility Adapter

`TOOL01_RENDERER=powerpoint`

PowerPoint COM을 사용한다. 현재 Validation Demo의 Blocking Gate가 아니며 사용자가 로컬에서 테스트할 필요 없다.

### Auto

`TOOL01_RENDERER=auto`

- Windows → PowerPoint
- non-Windows → LibreOffice

## 알려진 Demo 제한사항

- Placeholder는 `{{필드명}}` 형식이며 하나의 Text Run에 있어야 한다.
- 여러 Run으로 분할된 Placeholder, SmartArt/Chart 내부 Placeholder는 현재 범위 밖이다.
- 고객용 Web App/Hosted Service는 현재 구현하지 않는다. LibreOffice는 온라인 QA Renderer다.
