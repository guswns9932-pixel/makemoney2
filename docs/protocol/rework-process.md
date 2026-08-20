# REWORK 처리

## GPT 피드백 반영 원칙

- 일반적인 GPT 제안/의견은 Claude가 기술적 타당성을 먼저 검토한 뒤 반영 여부를 결정할 수 있다.
- 단, `[CLAUDE REVISION TASK]`로 명시된 필수 수정사항은 **승인된 작업 명세로 취급하여 반영한다.** "타당성이 없다고 판단해 반영하지 않았습니다" 같은 임의 거부는 하지 않는다.
- 기술적으로 구현이 불가능하거나 데이터 손상 위험이 있는 경우에만 강행하지 않고 `BLOCKED` 또는 `[GPT 판단 필요사항]`으로 보고한다.

## 처리 절차

GPT에서 `[CLAUDE REVISION TASK]`가 전달되면 기존 프로젝트를 새로 만들지 않는다.

다음 순서로 작업한다.

1. 기존 Baseline 확인 (`docs/state/PROJECT_STATE.md` + `docs/results/`의 직전 버전 결과보고 + git log)
2. GPT가 지적한 항목 확인
3. 지정된 항목만 우선 수정
4. 기존 정상기능 Regression 확인
5. 버전 갱신
6. `[CLAUDE RESULT]` 형식으로 재보고 (`docs/results/`에 새 버전 파일로 저장)

GPT가 수정하지 말라고 명시한 영역은 변경하지 않는다.

## PROJECT STATE 사용

브리프에 `[PROJECT STATE]`가 포함되어 있으면, 또는 `docs/state/PROJECT_STATE.md`가 존재하면 해당 내용을 현재 프로젝트의 기준 상태로 취급한다.

특히 다음을 우선 확인한다.

- 현재 Baseline
- 이번 작업
- 마지막 승인사항
- 미해결 문제

작업 결과가 PROJECT STATE와 충돌할 경우 임의 판단하지 않고 결과보고에 명시한다.

작업이 끝나면 `docs/state/PROJECT_STATE.md`도 최신 상태로 갱신한다.
