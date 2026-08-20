# [ARCHIVE / LEGACY] 엑셀 VBA 자동화 툴 10종 (v2.2.1 기준)

> 이 폴더는 과거 진행분을 보존하기 위한 Legacy 영역이다.
> 현재 진행 중인 `products/tool01/`과 이름이 겹치는 "Tool01"이 있었으나
> 서로 다른 프로젝트이므로 혼동을 막기 위해 여기로 이동했다.
> 최신 사업 방향과 이 Legacy 트랙의 관계(재사용 여부, 폐기 여부 등)는
> 아직 Claude에게 전달되지 않았다 → [GPT 판단 필요사항] 참고.

## 과거 진행 상태 (v2.2.1 시점 스냅샷)

- 공통 모듈(5개, 각 .xlsm에 물리적 내장): modCommonFile, modCommonSafety, modCommonName, modCommonValue, modCommonLog
- tool01_file_merge (다중 파일 통합): v2.2.1 완료
- tool02_file_sort (조건부 파일 분류): v2.2.1 완료
- Tool03 (Excel → PDF 일괄 변환): 개발 시작 단계에서 중단
- Tool04~10: 미착수
- 미해결: `Workbook_Open` MsgBox로 `AutomationSecurity` 차단 동작 확인하는 수동 검증 테스트 미완료 (Windows 환경 필요)
