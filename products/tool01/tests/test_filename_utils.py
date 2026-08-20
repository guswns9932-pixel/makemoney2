import pytest

from src import filename_utils as fu


# TEST 09 관련: 정상 입력 -> 안전한 파일명 생성
def test_build_filename_normal():
    row = {"수료번호": "CERT-2026-001", "이름": "홍길동"}
    name = fu.build_filename("{수료번호}_{이름}_수료증.pdf", row, ["수료번호", "이름"])
    assert name == "CERT-2026-001_홍길동_수료증.pdf"


# TEST 05: 파일명 금지문자 포함
def test_build_filename_forbidden_chars_removed():
    row = {"수료번호": 'CERT/2026:001?*', "이름": '<홍"길동>'}
    name = fu.build_filename("{수료번호}_{이름}_수료증.pdf", row, ["수료번호", "이름"])
    for ch in '\\/:*?"<>|':
        assert ch not in name


def test_sanitize_component_strips_whitespace_and_forbidden():
    assert fu.sanitize_component("  홍길동  ") == "홍길동"
    assert fu.sanitize_component('a/b:c*d?e"f<g>h|i') == "abcdefghi"
    assert fu.sanitize_component(None) == ""


# 오류처리 기준 항목 11: 파일명 필수값이 빈 값 -> 실패 처리 대상
def test_build_filename_missing_required_field_raises():
    row = {"수료번호": "", "이름": "홍길동"}
    with pytest.raises(fu.FilenameBuildError):
        fu.build_filename("{수료번호}_{이름}_수료증.pdf", row, ["수료번호", "이름"])


def test_build_filename_required_field_becomes_empty_after_sanitize():
    # 금지문자만 있던 값이 sanitize 후 빈 문자열이 되는 경우도 실패로 처리되어야 한다
    row = {"수료번호": "///", "이름": "홍길동"}
    with pytest.raises(fu.FilenameBuildError):
        fu.build_filename("{수료번호}_{이름}_수료증.pdf", row, ["수료번호", "이름"])


# TEST 06: 동일 파일명 -> 덮어쓰기 금지, 순번 부여
def test_unique_path_avoids_overwrite(tmp_path):
    (tmp_path / "CERT-001_홍길동_수료증.pdf").write_text("existing")
    p1 = fu.unique_path(tmp_path, "CERT-001_홍길동_수료증.pdf")
    assert p1.name == "CERT-001_홍길동_수료증_2.pdf"
    assert not p1.exists()

    p1.write_text("second")
    p2 = fu.unique_path(tmp_path, "CERT-001_홍길동_수료증.pdf")
    assert p2.name == "CERT-001_홍길동_수료증_3.pdf"


def test_unique_path_no_conflict_returns_original(tmp_path):
    p = fu.unique_path(tmp_path, "new_file.pdf")
    assert p.name == "new_file.pdf"


def test_build_filename_reserved_name_prefixed():
    row = {"수료번호": "CON", "이름": "x"}
    name = fu.build_filename("{수료번호}", row, ["수료번호"])
    assert name != "CON"
    assert name == "_CON"
