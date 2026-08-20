# Tool01 Cloud Acceptance Test — O01~O09

이 문서는 사용자가 로컬 Windows PC를 사용하지 않고도 Tool01 Validation Demo를 검증할 수 있도록 만든 **현재 Blocking QA 체크리스트**다.

실행 환경:

- Claude Code Remote / Linux
- GitHub Actions `ubuntu-latest`
- LibreOffice Impress Headless
- Python + pytest

실행 명령:

```bash
cd products/tool01
TOOL01_RENDERER=libreoffice python -m pytest tests/ -q -m "not live_renderer"
xvfb-run -a python -m pytest tests/test_main_gui.py -q
TOOL01_RENDERER=libreoffice python scripts/cloud_acceptance.py --output artifacts/cloud_acceptance
```

GitHub에서는 `.github/workflows/tool01-cloud-acceptance.yml`이 동일 절차를 자동 수행한다.

## 체크리스트

| ID | 검증 | PASS 기준 |
|---|---|---|
| O01 | Sample 100행 Preflight/자동매핑 | 100행 인식 + Placeholder 5개 매핑 |
| O02 | Preview 경로 | 첫 행 personalization 성공 + Preview cleanup/error orchestration regression 통과 |
| O03 | 실제 PDF + 100행 전체 personalization | LibreOffice 실제 PDF 대표 5/5 + personalized PPTX dry-run 100/100 |
| O04 | CSV 입력 | UTF-8-SIG CSV Preflight/Mapping + personalized PPTX 생성 |
| O05 | 한글 | 실제 Cloud PDF 파일명에 한글 + PDF 내부 5개 값 추출 확인 |
| O06 | 부분 실패 격리 | deterministic batch test에서 3행 중 1행 실패/2행 계속 성공 |
| O07 | 권한 오류 | PermissionError mock 경로에서 Crash 없음 + Preflight 실패/오류 메시지 |
| O08 | Automation 중복 방지 | GUI regression에서 Preview/Batch 중복 실행 차단 |
| O09 | 성능 관찰 | 대표 5건 실제 PDF 처리시간 기록. 현재 단계에서는 non-blocking |

## 왜 100개의 실제 PDF를 Cloud Gate에서 강제하지 않는가

현재 단계는 판매용 Installer 완성이 아니라 **시장 Validation용 Demo**다.

작은 Remote Runner에서 LibreOffice Office 프로세스를 100건 반복 변환하는 것은 환경별 변동성이 크다. 따라서 현재 Blocking Gate에서는:

- **100행 전체 데이터/매핑/개인화 로직**은 100/100 검증하고
- **실제 PDF Renderer**는 대표 5행을 실제로 변환한다.

이 방식은 온라인 작업 제약을 지키면서 핵심 기능을 실제 출력물까지 검증한다.

100건 PowerPoint COM 실제 Batch/Windows Installer는 유료 Pilot 또는 Full MVP 단계의 별도 Compatibility Gate다.

## Online Review Artifact

Cloud Acceptance 성공 시 다음을 생성한다.

- `cloud_acceptance_report.json`
- `cloud_acceptance_report.md`
- `cloud_preview.pdf`
- `cloud_preview.png`
- 실제 Batch PDF 5개
- 100행 personalized PPTX dry-run 산출물

GitHub Actions에서는 위 폴더를 Artifact로 업로드한다.
사용자는 브라우저에서 `cloud_preview.png`/PDF를 확인할 수 있으며 로컬 실행은 필요 없다.

## 현재 Baseline 실행 결과

2026-08-20 GPT 검수환경에서:

- O01~O09: PASS
- 실제 LibreOffice PDF: 5/5
- 100행 personalization dry-run: 100/100
- 대표 5건 처리시간: 11.90초
- 한글 PDF 내용/파일명: PASS

GitHub 업로드 후 Actions에서 동일 결과를 다시 확인해야 최종 Online Demo Gate 완료 후보가 된다.
