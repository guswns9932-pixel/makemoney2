# 테스트 원칙

코드 또는 자동화 프로그램을 제작하는 경우 가능한 범위에서 직접 테스트한다.

최소한 다음을 확인한다.

- 정상 입력
- 빈 입력
- 잘못된 입력
- 중복 데이터
- 파일 없음
- 권한 문제
- 잘못된 파일 형식
- 처리 중 오류
- 기존 기능 Regression

프로젝트 성격에 맞지 않는 테스트는 생략할 수 있다.

테스트하지 못한 항목을 테스트한 것처럼 보고하지 않는다.

## Online-only 작업 원칙

현재 사용자는 개발·검수 작업을 **온라인에서만 진행**한다.

따라서 Claude는:

- 사용자에게 로컬 Windows PC에서 명령을 실행하라고 요구하지 않는다.
- Remote 환경에 없는 OS/앱 때문에 현재 단계 전체를 자동으로 BLOCKED 처리하지 않는다.
- 동일한 제품가치를 검증할 수 있는 Cloud Acceptance, Headless Renderer, deterministic mock/integration test를 우선 설계한다.
- 현재 Validation Demo의 Blocking QA는 `products/tool01/docs/cloud_acceptance_test.md`를 따른다.
- PowerPoint COM 전용 실기는 시장 Validation 진입을 막지 않으며 `호환성 미검증`으로 분리한다.

Windows compatibility가 실제 판매/유료 Pilot에서 필요해지면 사용자 로컬 PC가 아니라 다음 중 하나를 검토한다.

- 온라인 Windows VM
- 외부 Pilot 고객 환경
- 별도 QA 환경

이때 비용/범위 변화가 있으면 GPT/사용자 승인 후 진행한다.

## 보고 표현

직접 확인하지 못한 일반 항목:

`미검증: 현재 환경에서 직접 확인하지 못함`

현재 non-blocking인 Windows/PowerPoint 전용 항목:

`호환성 미검증: 향후 유료 Pilot/온라인 Windows 환경에서 확인`

온라인 Acceptance에서 실제 수행한 항목만 PASS 처리한다.
