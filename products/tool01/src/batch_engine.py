"""
Preview / Batch 생성 오케스트레이션.

실행 명세 §6(Preview), §8(PDF 일괄생성), §9(파일명 안전처리), §12(원본 보호) 통합.

- 원본 XLSX/CSV/PPTX는 read-only로만 접근, 절대 수정하지 않는다.
- 각 행은 임시 폴더에 치환된 PPTX 사본을 만든 뒤 PDF로 변환한다 (원본 Template 불변).
- 한 행 실패가 전체 Batch를 중단시키지 않는다 (파일별 오류 격리).
- 결과 파일명 중복 시 덮어쓰지 않고 순번을 붙인다.
"""

from __future__ import annotations

import shutil
import tempfile
import uuid
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable

from . import filename_utils, pdf_renderer, template_engine, validator

DEFAULT_FILENAME_PATTERN = "{수료번호}_{이름}_수료증.pdf"
DEFAULT_REQUIRED_FIELDS = ["수료번호", "이름"]


@dataclass
class RowResult:
    row_index: int
    success: bool
    output_path: str | None = None
    error: str | None = None


@dataclass
class BatchResult:
    total: int
    success_count: int
    failure_count: int
    results: list[RowResult] = field(default_factory=list)

    @property
    def failures(self) -> list[RowResult]:
        return [r for r in self.results if not r.success]


ProgressCallback = Callable[[int, int], None]  # (현재_건수, 전체_건수)


def cleanup_preview_files(work_dir: str | Path, *, remove_dir: bool = False) -> None:
    """Tool01이 만든 Preview 임시파일만 best-effort로 정리한다.

    사용자가 직접 만든 파일은 건드리지 않도록 ``_preview_*`` 패턴만 삭제한다.
    Viewer가 파일을 잡고 있는 등 삭제할 수 없는 경우에는 실패를 전파하지 않는다.
    """
    work_dir = Path(work_dir)
    if not work_dir.exists():
        return

    for pattern in ("_preview_*.pptx", "_preview_*.pdf"):
        for path in work_dir.glob(pattern):
            try:
                path.unlink()
            except OSError:
                pass

    if remove_dir:
        try:
            work_dir.rmdir()  # Tool01 파일 외 다른 파일이 있으면 안전하게 남긴다.
        except OSError:
            pass



def generate_preview(
    template_path: str | Path,
    row: dict,
    mapping: dict[str, str],
    work_dir: str | Path,
    renderer=None,
) -> Path:
    """
    첫 번째(또는 지정된) 유효 데이터 행을 템플릿에 적용한 실제 결과 PDF를 생성한다.
    Batch 결과 파일과 완전히 분리된 임시 경로에 만든다 (§6 요구사항).

    이전 Preview 잔여물은 새 Preview 시작 전에 정리하고, PowerPoint 변환용 PPTX
    사본은 성공/실패와 관계없이 즉시 삭제한다. 현재 Preview PDF만 Viewer 확인용으로
    남기며 다음 Preview 또는 앱 정상 종료 시 정리한다.
    """
    work_dir = Path(work_dir)
    work_dir.mkdir(parents=True, exist_ok=True)
    cleanup_preview_files(work_dir)

    values = {ph: row.get(col, "") for ph, col in mapping.items()}
    tmp_pptx = work_dir / f"_preview_{uuid.uuid4().hex}.pptx"
    tmp_pdf = work_dir / f"_preview_{uuid.uuid4().hex}.pdf"

    try:
        template_engine.render_pptx(template_path, tmp_pptx, values)

        if renderer is not None:
            renderer.convert(tmp_pptx, tmp_pdf)
        else:
            pdf_renderer.convert_single(tmp_pptx, tmp_pdf)
    except Exception:
        # 실패 중 생성된 불완전 PDF가 다음 Preview에 남지 않게 best-effort 정리.
        try:
            tmp_pdf.unlink(missing_ok=True)
        except OSError:
            pass
        raise
    finally:
        try:
            tmp_pptx.unlink(missing_ok=True)
        except OSError:
            pass

    return tmp_pdf


def run_batch(
    template_path: str | Path,
    rows: list[dict],
    mapping: dict[str, str],
    output_dir: str | Path,
    filename_pattern: str = DEFAULT_FILENAME_PATTERN,
    required_fields: list[str] | None = None,
    progress_cb: ProgressCallback | None = None,
) -> BatchResult:
    """Generate one PDF per row.

    The function first prepares personalized PPTX files and isolates row-level
    filename/template errors. If the active renderer exposes ``convert_many``
    (LibreOffice cloud path), all remaining PPTX files are converted in a single
    process for fast online acceptance. PowerPoint COM keeps the safe sequential
    conversion path.
    """
    required_fields = required_fields or DEFAULT_REQUIRED_FIELDS
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    total = len(rows)
    results_by_index: dict[int, RowResult] = {}
    tmp_root = Path(tempfile.mkdtemp(prefix="tool01_batch_"))

    try:
        jobs: list[tuple[int, Path, Path]] = []

        # 1) Build filenames and personalized PPTX files. Row-level failures are isolated.
        for idx, row in enumerate(rows, start=1):
            try:
                filename = filename_utils.build_filename(
                    filename_pattern, row, required_fields
                )
                final_path = filename_utils.unique_path(output_dir, filename)
                values = {ph: row.get(col, "") for ph, col in mapping.items()}
                tmp_pptx = tmp_root / f"row_{idx}.pptx"
                template_engine.render_pptx(template_path, tmp_pptx, values)
                jobs.append((idx, tmp_pptx, final_path))
            except (
                filename_utils.FilenameBuildError,
                template_engine.TemplateError,
            ) as e:
                results_by_index[idx] = RowResult(row_index=idx, success=False, error=str(e))
                if progress_cb:
                    progress_cb(idx, total)
            except Exception as e:  # noqa: BLE001
                results_by_index[idx] = RowResult(
                    row_index=idx, success=False, error=f"예상하지 못한 오류: {e}"
                )
                if progress_cb:
                    progress_cb(idx, total)

        # 2) Convert all valid jobs using the selected renderer.
        with pdf_renderer.create_renderer() as renderer:
            if hasattr(renderer, "convert_many"):
                outcomes = renderer.convert_many([(src, dst) for _, src, dst in jobs])
                for (idx, _src, dst), (ok, error) in zip(jobs, outcomes):
                    if ok:
                        results_by_index[idx] = RowResult(
                            row_index=idx, success=True, output_path=str(dst)
                        )
                    else:
                        results_by_index[idx] = RowResult(
                            row_index=idx, success=False, error=error or "PDF 변환 실패"
                        )
                    if progress_cb:
                        progress_cb(idx, total)
            else:
                for idx, src, dst in jobs:
                    try:
                        renderer.convert(src, dst)
                        results_by_index[idx] = RowResult(
                            row_index=idx, success=True, output_path=str(dst)
                        )
                    except pdf_renderer.RendererError as e:
                        results_by_index[idx] = RowResult(
                            row_index=idx, success=False, error=str(e)
                        )
                    except Exception as e:  # noqa: BLE001
                        results_by_index[idx] = RowResult(
                            row_index=idx, success=False, error=f"예상하지 못한 오류: {e}"
                        )
                    finally:
                        if progress_cb:
                            progress_cb(idx, total)
    finally:
        shutil.rmtree(tmp_root, ignore_errors=True)

    results = [results_by_index[i] for i in range(1, total + 1)]
    success_count = sum(1 for r in results if r.success)
    return BatchResult(
        total=total,
        success_count=success_count,
        failure_count=total - success_count,
        results=results,
    )
