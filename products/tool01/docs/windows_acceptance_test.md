# Tool01 Windows Compatibility Test — Optional / Non-blocking

> 이 문서는 현재 Validation Demo의 Blocking Gate가 아니다.

사용자는 온라인에서만 작업하므로 로컬 Windows PC에서 이 체크리스트를 수행하지 않는다.
현재 Blocking QA는 `cloud_acceptance_test.md`의 O01~O09다.

PowerPoint COM Adapter는 코드에 유지되며, 아래 실기는 **유료 Pilot/Full MVP 단계에서 실제 필요성이 확인될 때만** 수행한다.

가능한 실행환경:

- 온라인 Windows VM + 적법한 Microsoft PowerPoint Desktop 환경
- 외부 Pilot 고객 환경
- 별도 QA 환경

## Compatibility 체크리스트

| ID | 내용 |
|---|---|
| W01 | Windows GUI 정상 실행 |
| W02 | PowerPoint COM Preview 실제 PDF |
| W03 | XLSX 100행 → PowerPoint COM PDF 100개 |
| W04 | CSV → PowerPoint COM PDF |
| W05 | 한글 내용/파일명 |
| W06 | 특정 행 실패 시 나머지 행 계속 처리 |
| W07 | Windows ACL 쓰기 권한 오류 |
| W08 | 기존 PowerPoint 문서/세션 보호 |
| W09 | 100건 실제 Runtime 처리시간 |

## 현재 상태

`호환성 미검증: 향후 유료 Pilot/온라인 Windows 환경에서 확인`

이 상태는 **Online Validation Demo의 시장 Validation 진입을 막지 않는다.**
