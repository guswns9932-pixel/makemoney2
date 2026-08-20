import os
import stat
import sys
from unittest import mock

import pytest

from src import validator as v


# TEST 01: 정상 입력 -> Preflight 통과
def test_preflight_ok(sample_xlsx, sample_pptx, tmp_output_dir):
    result = v.run_preflight(sample_xlsx, sample_pptx, tmp_output_dir)
    assert result.ok is True
    assert result.errors == []
    assert len(result.rows) == 100
    assert set(result.placeholders) == {"수료번호", "이름", "과정명", "수료일", "기관명"}
    assert result.mapping["이름"] == "이름"


# 필수 입력 누락
def test_preflight_missing_inputs():
    result = v.run_preflight(None, None, None)
    assert result.ok is False
    assert any("데이터 파일" in e for e in result.errors)
    assert any("PPTX 템플릿" in e for e in result.errors)
    assert any("출력 폴더" in e for e in result.errors)


# TEST 03: Placeholder에 대응하는 컬럼 누락
def test_preflight_missing_mapped_column(sample_pptx, tmp_output_dir, tmp_path):
    csv_path = tmp_path / "data.csv"
    # '기관명' 컬럼을 의도적으로 제거
    csv_path.write_text(
        "수료번호,이름,과정명,수료일\nCERT-001,홍길동,AI과정,2026-08-01\n",
        encoding="utf-8-sig",
    )
    result = v.run_preflight(csv_path, sample_pptx, tmp_output_dir)
    assert result.ok is False
    assert "기관명" in result.missing_fields
    assert any("기관명" in e for e in result.errors)


# TEST 07: 빈 데이터
def test_preflight_empty_data(sample_pptx, tmp_output_dir, tmp_path):
    csv_path = tmp_path / "empty.csv"
    csv_path.write_text("수료번호,이름,과정명,수료일,기관명\n", encoding="utf-8-sig")
    result = v.run_preflight(csv_path, sample_pptx, tmp_output_dir)
    assert result.ok is False


# TEST 04: 잘못된 파일 형식
def test_preflight_wrong_data_extension(sample_pptx, tmp_output_dir, tmp_path):
    bad = tmp_path / "data.txt"
    bad.write_text("not a real data file")
    result = v.run_preflight(bad, sample_pptx, tmp_output_dir)
    assert result.ok is False


# TEST 13: 출력 폴더 쓰기 권한 없음 (GPT DECISION 2026-08-20 반영)
@pytest.mark.skipif(
    os.name == "nt" or os.geteuid() == 0,
    reason="root로 실행되는 샌드박스/Windows에서는 chmod로 쓰기 차단을 재현할 수 없음",
)
def test_preflight_output_dir_not_writable(sample_xlsx, sample_pptx, tmp_path):
    readonly_dir = tmp_path / "readonly_out"
    readonly_dir.mkdir()
    readonly_dir.chmod(stat.S_IREAD | stat.S_IEXEC)  # 쓰기 권한 제거
    try:
        result = v.run_preflight(sample_xlsx, sample_pptx, readonly_dir)
        assert result.ok is False
        assert any("쓰기 권한" in e for e in result.errors)
    finally:
        readonly_dir.chmod(stat.S_IRWXU)  # 정리


# [CLAUDE REVISION TASK 필수 수정 4] 실제 OS ACL과 무관하게, 쓰기 시도에서
# PermissionError가 발생하는 경로를 Mock으로 강제 재현해 항상 실행되는 자동 테스트로 검증한다.
# 실제 Windows 폴더 권한 실기 테스트는 windows_acceptance_test.md의 W07로 별도 유지한다.
def test_check_output_dir_writable_permission_error_mocked(tmp_path):
    target_dir = tmp_path / "mock_out"
    with mock.patch("src.validator.open", side_effect=PermissionError("접근이 거부되었습니다")):
        err = v._check_output_dir_writable(target_dir)
    assert err is not None
    assert "권한" in err
    # 원본/폴더 자체는 실제로 생성되었는지만 확인 (open()만 mock했으므로 mkdir은 정상 동작)
    assert target_dir.exists()


def test_preflight_output_dir_not_writable_mocked(sample_xlsx, sample_pptx, tmp_output_dir):
    """
    Preflight 전체 경로에서 권한 오류가 발생해도:
    - 프로그램이 Crash하지 않고 (예외가 밖으로 새지 않고)
    - ok=False로 정상 실패 처리되며
    - 사용자가 이해할 수 있는 오류 메시지가 담기고
    - 원본 데이터/템플릿 파일은 변경되지 않는지 확인한다.
    """
    before_xlsx = (sample_xlsx.stat().st_mtime, sample_xlsx.stat().st_size)
    before_pptx = (sample_pptx.stat().st_mtime, sample_pptx.stat().st_size)

    with mock.patch(
        "src.validator._check_output_dir_writable",
        return_value="출력 폴더에 쓰기 권한이 없습니다: mocked PermissionError",
    ):
        result = v.run_preflight(sample_xlsx, sample_pptx, tmp_output_dir)

    assert result.ok is False
    assert any("권한" in e for e in result.errors)
    assert (sample_xlsx.stat().st_mtime, sample_xlsx.stat().st_size) == before_xlsx
    assert (sample_pptx.stat().st_mtime, sample_pptx.stat().st_size) == before_pptx
