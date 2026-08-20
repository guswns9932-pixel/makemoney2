from pathlib import Path
import platform

import pytest

from src import batch_engine as be
from src import pdf_renderer as pr
from src import validator as v


@pytest.mark.skipif(
    pr.find_libreoffice() is None,
    reason="온라인 실제 Batch 통합테스트에는 LibreOffice가 필요",
)
@pytest.mark.live_renderer
def test_run_batch_real_libreoffice_three_rows(sample_pptx, sample_xlsx, tmp_output_dir, monkeypatch):
    monkeypatch.setenv("TOOL01_RENDERER", "libreoffice")
    preflight = v.run_preflight(sample_xlsx, sample_pptx, tmp_output_dir)
    assert preflight.ok

    result = be.run_batch(
        sample_pptx,
        preflight.rows[:3],
        preflight.mapping,
        tmp_output_dir,
    )
    assert result.total == 3
    assert result.success_count == 3
    assert result.failure_count == 0
    pdfs = list(Path(tmp_output_dir).glob("*.pdf"))
    assert len(pdfs) == 3
    assert all(p.read_bytes().startswith(b"%PDF") for p in pdfs)


def test_default_filename_pattern_matches_spec():
    # 실행 명세 §출력 - 기본 파일명 규칙: {수료번호}_{이름}_수료증.pdf
    assert be.DEFAULT_FILENAME_PATTERN == "{수료번호}_{이름}_수료증.pdf"


# Batch 전체 흐름(파일별 오류 격리, 진행률 콜백, 실제 PDF 생성 성공/실패 집계)은
# PowerPoint COM 렌더러 없이는 끝까지 실행할 수 없다.
# 미검증: 실제 Windows + PowerPoint Desktop 환경에서 100건 Batch 실행 확인 필요


class _PreviewRendererStub:
    def convert(self, pptx_path, pdf_path):
        # Preview 임시 PPTX가 실제로 존재하는 시점에 renderer가 호출되는지 확인하고
        # 최소 PDF-like 파일을 생성한다. PowerPoint COM 자체를 검증하는 테스트는 아니다.
        assert Path(pptx_path).exists()
        Path(pdf_path).write_bytes(b"%PDF-1.4\n% tool01 preview test\n")


def test_preview_cleans_intermediate_pptx_and_previous_preview(
    sample_pptx, sample_xlsx, tmp_output_dir
):
    from pathlib import Path

    preflight = v.run_preflight(sample_xlsx, sample_pptx, tmp_output_dir)
    assert preflight.ok

    work_dir = Path(tmp_output_dir) / "_tool01_preview_tmp"
    renderer = _PreviewRendererStub()

    first = be.generate_preview(
        sample_pptx, preflight.rows[0], preflight.mapping, work_dir, renderer=renderer
    )
    assert first.exists()
    assert list(work_dir.glob("_preview_*.pptx")) == []
    assert list(work_dir.glob("_preview_*.pdf")) == [first]

    second = be.generate_preview(
        sample_pptx, preflight.rows[0], preflight.mapping, work_dir, renderer=renderer
    )
    assert second.exists()
    assert second != first
    assert not first.exists(), "새 Preview 시작 시 이전 Preview PDF가 정리되어야 한다"
    assert list(work_dir.glob("_preview_*.pptx")) == []
    assert list(work_dir.glob("_preview_*.pdf")) == [second]

    be.cleanup_preview_files(work_dir, remove_dir=True)
    assert not work_dir.exists()

class _BatchRendererStub:
    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, tb):
        return None

    def convert_many(self, jobs):
        for _src, dst in jobs:
            Path(dst).write_bytes(b"%PDF-1.4\n% stub\n")
        return [(True, None) for _ in jobs]


def test_batch_partial_failure_isolated_with_stub_renderer(
    sample_pptx, sample_xlsx, tmp_output_dir, monkeypatch
):
    preflight = v.run_preflight(sample_xlsx, sample_pptx, tmp_output_dir)
    assert preflight.ok
    rows = [dict(preflight.rows[0]), dict(preflight.rows[1]), dict(preflight.rows[2])]
    rows[1]["이름"] = ""  # filename required field -> only row 2 should fail

    monkeypatch.setattr(pr, "create_renderer", lambda: _BatchRendererStub())
    result = be.run_batch(sample_pptx, rows, preflight.mapping, tmp_output_dir)

    assert result.total == 3
    assert result.success_count == 2
    assert result.failure_count == 1
    assert result.results[0].success is True
    assert result.results[1].success is False
    assert "필수 필드" in (result.results[1].error or "")
    assert result.results[2].success is True
