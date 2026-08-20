import csv

import pytest

from src import data_loader as dl


# TEST 01 관련: 정상 xlsx 100행 로딩
def test_load_xlsx_normal(sample_xlsx):
    headers, rows = dl.load_data(sample_xlsx)
    assert headers == ["수료번호", "이름", "과정명", "수료일", "기관명"]
    assert len(rows) == 100
    assert rows[0]["수료번호"] == "CERT-2026-001"


# TEST 04: 파일 없음
def test_load_data_missing_file(tmp_path):
    missing = tmp_path / "no_such_file.xlsx"
    with pytest.raises(dl.DataLoadError):
        dl.load_data(missing)


# TEST 04: 잘못된 파일 형식 (지원하지 않는 확장자)
def test_load_data_unsupported_extension(tmp_path):
    bad = tmp_path / "data.txt"
    bad.write_text("hello")
    with pytest.raises(dl.DataLoadError):
        dl.load_data(bad)


# TEST 07: 빈 데이터 (헤더만 있고 행이 없음)
def test_load_csv_header_only_raises(tmp_path):
    p = tmp_path / "empty.csv"
    p.write_text("수료번호,이름\n", encoding="utf-8-sig")
    with pytest.raises(dl.DataLoadError):
        dl.load_data(p)


# 완전히 빈 행은 제외
def test_load_csv_skips_fully_empty_rows(tmp_path):
    p = tmp_path / "data.csv"
    p.write_text(
        "수료번호,이름\nCERT-001,홍길동\n,\nCERT-002,김민지\n",
        encoding="utf-8-sig",
    )
    headers, rows = dl.load_data(p)
    assert len(rows) == 2
    assert rows[1]["수료번호"] == "CERT-002"


# TEST 02 / TEST 10: CSV + 한글 정상 처리
def test_load_csv_korean_utf8_sig(tmp_path):
    p = tmp_path / "data.csv"
    with open(p, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["수료번호", "이름", "기관명"])
        writer.writerow(["CERT-001", "박현준", "ABC 교육원"])
    headers, rows = dl.load_data(p)
    assert headers == ["수료번호", "이름", "기관명"]
    assert rows[0]["이름"] == "박현준"


# TEST 08: 일부 빈 셀
def test_load_csv_partial_empty_cell_kept_as_empty_string(tmp_path):
    p = tmp_path / "data.csv"
    p.write_text(
        "수료번호,이름,기관명\nCERT-001,,ABC 교육원\n",
        encoding="utf-8-sig",
    )
    headers, rows = dl.load_data(p)
    assert rows[0]["이름"] == ""


def test_load_data_not_a_file_but_directory(tmp_path):
    with pytest.raises(dl.DataLoadError):
        dl.load_data(tmp_path)


# TEST 09: 원본 보호 - load_data가 XLSX/CSV 원본을 수정하지 않는지 확인
def test_load_xlsx_does_not_modify_original(sample_xlsx):
    before_mtime = sample_xlsx.stat().st_mtime
    before_size = sample_xlsx.stat().st_size
    dl.load_data(sample_xlsx)
    assert sample_xlsx.stat().st_mtime == before_mtime
    assert sample_xlsx.stat().st_size == before_size


def test_load_csv_does_not_modify_original(tmp_path):
    p = tmp_path / "data.csv"
    p.write_text("수료번호,이름\nCERT-001,홍길동\n", encoding="utf-8-sig")
    before_mtime = p.stat().st_mtime
    before_size = p.stat().st_size
    dl.load_data(p)
    assert p.stat().st_mtime == before_mtime
    assert p.stat().st_size == before_size

