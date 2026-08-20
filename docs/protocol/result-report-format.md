# 결과보고 형식

## 판정 체계 (구현 QA vs 사업 Gate)

이 문서의 `PASS / REWORK`와 사업 판단에 쓰이는 `GO / HOLD / DROP`은 서로 다른 레벨의 판정이다.

- **PASS / REWORK** — 일반적인 구현 결과 QA의 기본 판정이다. Claude가 브리프/실행 명세대로 구현했는지, 테스트를 통과했는지를 GPT가 검수할 때 사용한다.
- **HOLD / DROP** — 구현 중 사업 전제(타겟, 가치제안, 가격 등) 또는 프로젝트 진행 가치에 영향을 주는 문제가 발견되면 GPT는 PASS/REWORK 대신 HOLD 또는 DROP을 사용할 수 있다.
- **GO** — 사업 또는 작업 착수 여부를 승인하는 별도의 판단이며, `docs/state/PROJECT_STATE.md`나 브리프 승인 시점에 쓰인다.

즉 "이번 코드 구현이 통과했는가"는 PASS/REWORK로, "이 사업/기능을 계속 진행할 가치가 있는가"는 GO/HOLD/DROP으로 나눠서 판단한다. Claude는 결과보고에서 항상 PASS/REWORK 체계를 전제로 `IMPLEMENTATION COMPLETE / BLOCKED / PARTIAL`을 보고하며, GO/HOLD/DROP은 GPT/사용자의 사업 판단 영역이다.

## 완료 판단

Claude는 작업이 끝났다고 판단하더라도 프로젝트 자체를 최종 승인하지 않는다.

Claude의 역할은 `IMPLEMENTATION COMPLETE`까지다. 최종 PASS 여부는 GPT가 판단한다.

따라서 결과보고에서 "최종 완료", "출시 가능", "완벽하게 완료" 등의 표현은 GPT 검수 전에는 사용하지 않는다.

## 보고 형식

모든 작업 완료 후 반드시 다음 형식으로 보고한다. 동일한 내용을 `docs/results/{제품명}_{버전}_result.md` 파일로도 저장한다.

```markdown
[CLAUDE RESULT]

- 프로젝트:
- 버전/작업단계:

### 구현 완료 내용

-
-
-

### 테스트 결과

- TEST 01:
- TEST 02:
- TEST 03:

### 변경사항

- 기존 대비 변경된 내용

### 발견된 문제

- 없음 / 내용

### 미해결 사항

- 없음 / 내용

### 변경 제안

- 없음 / 내용
- 현재 명세와 다른 방향을 추천할 경우 이유와 기대효과를 작성

### GPT 판단 필요사항

- 없음 / 내용

### 산출물

- 생성 또는 수정한 파일
- 코드
- 문서
- 기타 결과물

### 구현 상태

IMPLEMENTATION COMPLETE / BLOCKED / PARTIAL
```
