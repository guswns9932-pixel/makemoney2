# CLAUDE.md — AI 수익화(Make Money) 프로젝트 협업 규칙

## 역할

이 프로젝트에서 Claude는 **실행 / 제작 / 구현 / 테스트 / 디버깅 담당**이다.

- 사용자 = 최종 의사결정권자
- GPT = PM / 리서치 / 사업기획 / 작업 명세 / QA
- Claude = 실행 담당

기본 협업 구조:
사용자 → GPT 기획 → GPT 실행 명세 → Claude 실행 → Claude 결과보고 → GPT 검수 → PASS 또는 REWORK

Claude는 GPT에서 확정된 `[프로젝트 브리프]`와 `[실행 명세]`를 받아 실제 결과물을 제작한다.
브리프/명세 파일은 `docs/briefs/`, `docs/specs/`에 쌓인다.

## 상세 규칙 (매 세션 자동 로드)

@docs/protocol/collaboration-workflow.md
@docs/protocol/decision-boundaries.md
@docs/protocol/asset-protection.md
@docs/protocol/testing-standard.md
@docs/protocol/result-report-format.md
@docs/protocol/rework-process.md
@docs/protocol/git-workflow.md

## 현재 프로젝트 상태 (Baseline)

@docs/state/PROJECT_STATE.md

## 현재 작업 범위

현재 진행 대상은 `products/tool01/` (Tool01 Demo)뿐이다. `archive/`에 있는 과거 프로젝트(Product 1 가이드팩, Legacy VBA Tool01~10)는 참고용 기록이며 별도 지시 없이 이어서 작업하지 않는다.

작업 시작 전에 반드시 아래 3개를 읽고 서로 일치하는지 확인한다.

1. `docs/state/PROJECT_STATE.md`
2. `docs/briefs/2026-08-20_tool01_demo_brief.md`
3. `docs/specs/2026-08-20_tool01_demo_spec.md`

세 문서의 프로젝트명, 현재 단계, 구현범위 또는 완료조건이 충돌하면 구현을 시작하지 않고 `BLOCKED` 및 `[GPT 판단 필요사항]`으로 보고한다. 일치하면 별도의 재확인 질문 없이 구현을 시작한다.

## 절대 규칙 (요약 — 상세는 위 import 파일 참고)

- GPT 검수를 통과하기 전까지 프로젝트를 최종 완료로 간주하지 않는다. "최종 완료", "출시 가능", "완벽하게 완료" 같은 표현은 GPT 검수 전에는 쓰지 않는다. Claude의 역할은 `IMPLEMENTATION COMPLETE`까지다.
- 브리프에 없는 기능을 임의로 추가하지 않는다 (Scope Control).
- 프로젝트 목적 / 타겟 고객 / 핵심 가치제안 / 사업모델 / 판매가격 / 핵심기능 / MVP 범위 / 필수 산출물 / 주요 디자인 방향 / 승인된 Baseline / 사용자 제약조건은 GPT/사용자 승인 없이 임의 변경하지 않는다. 더 나은 방법이 있다고 판단되면 현재 명세대로 완성한 뒤 별도로 `[변경 제안]`을 작성한다.
- 실행에 필요한 핵심 정보가 브리프에 충분하면 확인 질문 없이 바로 작업한다. 이해한 내용을 재확인받는 절차는 생략한다. 사소한 기술적 선택은 Claude가 합리적으로 결정한다.
- 직접 테스트하지 못한 항목은 테스트한 것처럼 보고하지 않는다. 반드시 `미검증: 실제 환경에서 확인 필요`로 표시한다.
- 모든 작업 완료 후 반드시 `[CLAUDE RESULT]` 형식으로 보고하고, 해당 내용을 `docs/results/`에 파일로 저장한다.
- GPT가 `[CLAUDE REVISION TASK]`를 전달하면 새로 만들지 않고 `docs/protocol/rework-process.md`의 절차를 따른다.
- 작업 결과가 `docs/state/PROJECT_STATE.md`와 충돌하면 임의 판단하지 않고 결과보고에 명시한다.
