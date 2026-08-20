import platform
from pathlib import Path

import pytest

from src import pdf_renderer as pr


@pytest.mark.skipif(
    platform.system() == "Windows",
    reason="PowerPoint COM non-Windows guard는 Windows에서는 의미가 없음",
)
def test_powerpoint_renderer_unavailable_on_non_windows():
    with pytest.raises(pr.RendererUnavailableError):
        with pr.PowerPointRenderer():
            pass


def test_invalid_renderer_name():
    with pytest.raises(pr.RendererUnavailableError):
        pr.create_renderer("does-not-exist")


@pytest.mark.skipif(
    pr.find_libreoffice() is None,
    reason="LibreOffice가 설치된 환경에서만 실제 PDF 변환 통합 테스트 가능",
)
@pytest.mark.live_renderer
def test_libreoffice_renderer_converts_real_pptx(sample_pptx, tmp_path):
    out = tmp_path / "converted.pdf"
    with pr.LibreOfficeRenderer() as renderer:
        result = renderer.convert(sample_pptx, out)
    assert result == out
    assert out.exists()
    assert out.stat().st_size > 1000
    assert out.read_bytes().startswith(b"%PDF")


@pytest.mark.skipif(
    platform.system() == "Windows" or pr.find_libreoffice() is None,
    reason="Linux/온라인 auto renderer 경로 확인용",
)
@pytest.mark.live_renderer
def test_auto_renderer_uses_libreoffice_on_non_windows(sample_pptx, tmp_path, monkeypatch):
    monkeypatch.delenv("TOOL01_RENDERER", raising=False)
    out = tmp_path / "auto.pdf"
    result = pr.convert_single(sample_pptx, out)
    assert result.exists()
    assert result.read_bytes().startswith(b"%PDF")


def test_explicit_env_renderer_selection(monkeypatch):
    monkeypatch.setenv("TOOL01_RENDERER", "libreoffice")
    renderer = pr.create_renderer()
    assert isinstance(renderer, pr.LibreOfficeRenderer)
