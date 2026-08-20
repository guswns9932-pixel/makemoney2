# Make Money — Office Automation Tool Business

직장인·소규모 사업자의 반복업무를 줄여주는 **Office Automation Tool 판매 + 기존 Tool 기반 맞춤 자동화** 수익화 프로젝트.

## 협업 구조

사용자(최종 의사결정) → GPT(PM/리서치/사업기획/명세/QA) → Claude(구현/테스트/결과보고) → GPT 검수 → 사용자 최종 결정

Claude Code에서 이 저장소를 열면 `CLAUDE.md`가 핵심 협업 규칙과 현재 Baseline을 로드한다.

## 현재 활성 프로젝트

**Tool01 — 수료증·상장·확인서·증명서 대량생성**

- 현재 단계: Validation Demo 구현
- 현재 판정: GO
- Primary ICP: 소규모 민간 교육기관·교육대행사·직무교육업체
- 핵심 입력: XLSX/CSV + PPTX Template
- 핵심 출력: 개인별 PDF 일괄생성
- Validation 가격: 29,000원
- Custom Template Setup 가설: +79,000원

상세 내용:

- `docs/briefs/2026-08-20_tool01_demo_brief.md`
- `docs/specs/2026-08-20_tool01_demo_spec.md`
- `docs/state/PROJECT_STATE.md`

## 사업 제약

- 하루 약 1시간 운영
- 초기 자본 100만원 이하
- 수익 발생 이후 추가 투자 가능
- 반도체 장비 영업 관련 수익화 아이디어 제외
- 반복적인 저단가 맞춤노동과 과도한 고객응대 모델 지양

## 협업 원칙

- GPT가 제품 방향, 타겟, 가격, MVP 범위와 Definition of Done을 확정한다.
- Claude는 승인된 Brief/Spec 범위 안에서 구현방법과 코드 구조를 결정한다.
- Claude는 Scope를 임의로 넓히지 않는다.
- Claude 구현 후 `[CLAUDE RESULT]`를 작성한다.
- GPT가 구현 결과를 `PASS / REWORK`로 검수한다 (사업 전제에 영향을 주는 문제 발견 시 `HOLD / DROP` 사용 가능).
- 사업/작업 착수 여부는 별도로 `GO / HOLD / DROP`으로 판단한다 (상세: `docs/protocol/result-report-format.md`).

## 폴더 구조

- `CLAUDE.md` — Claude Code 핵심 지침
- `docs/protocol/` — 협업/권한/테스트/Git/REWORK 규칙
- `docs/state/PROJECT_STATE.md` — 현재 활성 프로젝트 Baseline
- `docs/briefs/` — GPT 승인 프로젝트 브리프
- `docs/specs/` — GPT 승인 실행 명세
- `docs/results/` — Claude 결과보고 이력
- `products/tool01/` — 현재 구현 대상
- `archive/` — 과거 프로젝트 기록

## Legacy 주의

`archive/legacy_vba_tools/tool01_file_merge/`는 과거 Excel 다중파일 통합 프로젝트다.

현재 Tool01과 이름만 같으며 별개의 제품이다. 현재 기능이나 사업 범위에 자동으로 포함하지 않는다.
