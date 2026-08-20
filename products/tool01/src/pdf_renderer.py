"""PPTX -> PDF renderer adapters.

Tool01 keeps two renderer paths:

1) ``PowerPointRenderer``
   - Windows + Microsoft PowerPoint Desktop
   - customer/runtime compatibility path
   - highest expected fidelity to a PowerPoint-authored template

2) ``LibreOfficeRenderer``
   - Linux/Claude Code Remote/GitHub Actions compatible
   - online development and Cloud Acceptance baseline
   - uses LibreOffice Impress headless conversion

The active renderer is selected by ``TOOL01_RENDERER``:
``powerpoint`` / ``libreoffice`` / ``auto`` (default).

``auto`` chooses PowerPoint on Windows and LibreOffice on non-Windows systems.
This lets the repository be implemented and validated entirely online while
preserving the Windows PowerPoint adapter for later compatibility testing.
"""

from __future__ import annotations

import os
import platform
import shutil
import subprocess
import tempfile
from pathlib import Path

PP_SAVE_AS_PDF = 32


class RendererError(RuntimeError):
    """PPTX -> PDF 렌더링 중 발생한 오류."""


class RendererUnavailableError(RendererError):
    """요청한 렌더러를 현재 환경에서 사용할 수 없을 때."""


# ---------------------------------------------------------------------------
# PowerPoint COM renderer (Windows compatibility path)


def _require_windows_com():
    if platform.system() != "Windows":
        raise RendererUnavailableError(
            "PowerPoint COM 렌더러는 Windows + Microsoft PowerPoint Desktop 환경에서만 "
            "동작합니다. 온라인/리눅스 환경에서는 LibreOffice 렌더러를 사용하세요."
        )
    try:
        import win32com.client  # noqa: F401
        import pythoncom  # noqa: F401
    except ImportError as e:
        raise RendererUnavailableError(
            "pywin32(win32com/pythoncom)가 설치되어 있지 않습니다. "
            "Windows 환경에서 `pip install pywin32`가 필요합니다."
        ) from e


class PowerPointRenderer:
    """Microsoft PowerPoint COM 기반 Renderer.

    Safety baseline:
    - COM은 Renderer를 사용하는 Worker Thread 안에서 초기화/해제한다.
    - ``PowerPoint.Application.Quit()``은 호출하지 않는다.
    - Tool01이 직접 Open한 Temporary Presentation만 Close한다.
    - 사용자가 이미 열어둔 다른 Presentation을 저장/수정/종료하지 않는다.
    """

    def __init__(self):
        self._app = None
        self._com_initialized = False

    def __enter__(self) -> "PowerPointRenderer":
        _require_windows_com()
        import pythoncom
        import win32com.client

        pythoncom.CoInitialize()
        self._com_initialized = True

        try:
            self._app = self._acquire_application(win32com.client)
        except Exception:
            self._cleanup_com()
            raise
        return self

    @staticmethod
    def _acquire_application(win32com_client):
        try:
            return win32com_client.GetActiveObject("PowerPoint.Application")
        except Exception:
            pass

        try:
            return win32com_client.Dispatch("PowerPoint.Application")
        except Exception as e:  # noqa: BLE001
            raise RendererError(f"PowerPoint 애플리케이션을 시작할 수 없습니다: {e}") from e

    def __exit__(self, exc_type, exc_val, exc_tb) -> None:
        self._app = None
        self._cleanup_com()

    def _cleanup_com(self) -> None:
        if self._com_initialized:
            import pythoncom

            try:
                pythoncom.CoUninitialize()
            finally:
                self._com_initialized = False

    def convert(self, pptx_path: str | Path, pdf_path: str | Path) -> Path:
        if self._app is None:
            raise RendererError("PowerPointRenderer가 초기화되지 않았습니다 (with 블록 안에서 사용).")

        pptx_path = Path(pptx_path).resolve()
        pdf_path = Path(pdf_path).resolve()
        pdf_path.parent.mkdir(parents=True, exist_ok=True)

        presentation = None
        try:
            presentation = self._app.Presentations.Open(str(pptx_path), WithWindow=False)
            presentation.SaveAs(str(pdf_path), PP_SAVE_AS_PDF)
        except Exception as e:  # noqa: BLE001
            raise RendererError(f"PDF 변환 중 오류가 발생했습니다 ({pptx_path.name}): {e}") from e
        finally:
            if presentation is not None:
                try:
                    presentation.Close()
                except Exception:  # noqa: BLE001
                    pass

        if not pdf_path.exists() or pdf_path.stat().st_size == 0:
            raise RendererError(f"PowerPoint가 PDF를 생성하지 못했습니다: {pdf_path.name}")
        return pdf_path


# ---------------------------------------------------------------------------
# LibreOffice renderer (online/cloud acceptance baseline)


def find_libreoffice() -> str | None:
    """Return a usable LibreOffice/soffice executable path if available."""
    configured = os.environ.get("TOOL01_SOFFICE")
    if configured:
        path = shutil.which(configured) or configured
        if Path(path).exists():
            return str(path)
    return shutil.which("soffice") or shutil.which("libreoffice")


class LibreOfficeRenderer:
    """LibreOffice Impress headless PPTX -> PDF renderer.

    A private LibreOffice user profile is created for every conversion so
    sequential/concurrent CI conversions never contend for a shared profile lock.
    LibreOffice documents that its process needs write access to the user profile,
    hence the dedicated temporary profile.
    """

    def __init__(self, executable: str | None = None, timeout_seconds: int = 120):
        self.executable = executable or find_libreoffice()
        self.timeout_seconds = timeout_seconds

    def __enter__(self) -> "LibreOfficeRenderer":
        if not self.executable:
            raise RendererUnavailableError(
                "LibreOffice(soffice)를 찾을 수 없습니다. Claude Code Remote/GitHub Actions에서는 "
                "LibreOffice를 설치한 뒤 TOOL01_RENDERER=libreoffice로 실행하세요."
            )
        return self

    def __exit__(self, exc_type, exc_val, exc_tb) -> None:
        return None

    def convert(self, pptx_path: str | Path, pdf_path: str | Path) -> Path:
        if not self.executable:
            raise RendererUnavailableError("LibreOffice(soffice)를 찾을 수 없습니다.")
        pptx_path = Path(pptx_path).resolve()
        pdf_path = Path(pdf_path).resolve()
        if not pptx_path.exists():
            raise RendererError(f"PPTX 파일이 없습니다: {pptx_path}")
        pdf_path.parent.mkdir(parents=True, exist_ok=True)

        convert_dir = Path(tempfile.mkdtemp(prefix="tool01_lo_convert_"))
        profile_dir = Path(tempfile.mkdtemp(prefix="tool01_lo_profile_"))
        try:
            profile_uri = profile_dir.resolve().as_uri()
            cmd = [
                self.executable,
                "--headless",
                "--nologo",
                "--nodefault",
                "--nofirststartwizard",
                f"-env:UserInstallation={profile_uri}",
                "--convert-to",
                "pdf",
                "--outdir",
                str(convert_dir),
                str(pptx_path),
            ]
            try:
                completed = subprocess.run(
                    cmd,
                    capture_output=True,
                    text=True,
                    timeout=self.timeout_seconds,
                    check=False,
                )
            except subprocess.TimeoutExpired as e:
                raise RendererError(
                    f"LibreOffice PDF 변환이 {self.timeout_seconds}초 안에 끝나지 않았습니다: {pptx_path.name}"
                ) from e

            expected = convert_dir / f"{pptx_path.stem}.pdf"
            if completed.returncode != 0 or not expected.exists() or expected.stat().st_size == 0:
                detail = (completed.stderr or completed.stdout or "출력 파일 없음").strip()
                raise RendererError(
                    f"LibreOffice PDF 변환 실패 ({pptx_path.name}): {detail}"
                )

            shutil.move(str(expected), str(pdf_path))
        finally:
            shutil.rmtree(convert_dir, ignore_errors=True)
            shutil.rmtree(profile_dir, ignore_errors=True)

        return pdf_path

    def convert_many(self, jobs: list[tuple[str | Path, str | Path]]) -> list[tuple[bool, str | None]]:
        """Convert multiple PPTX files in chunked LibreOffice processes.

        LibreOffice startup is expensive, while very large single invocations can
        become slow on small CI runners. The cloud path therefore converts in
        chunks (default 20 files/process), which keeps the 100-row acceptance
        deterministic and substantially faster than 100 separate starts.
        """
        if not self.executable:
            raise RendererUnavailableError("LibreOffice(soffice)를 찾을 수 없습니다.")
        if not jobs:
            return []

        normalized: list[tuple[Path, Path]] = []
        for src, dst in jobs:
            src_p = Path(src).resolve()
            dst_p = Path(dst).resolve()
            dst_p.parent.mkdir(parents=True, exist_ok=True)
            normalized.append((src_p, dst_p))

        batch_size = max(1, int(os.environ.get("TOOL01_LO_BATCH_SIZE", "20")))
        outcomes: list[tuple[bool, str | None]] = [(False, "변환 미실행") for _ in normalized]

        for chunk_start in range(0, len(normalized), batch_size):
            chunk = normalized[chunk_start:chunk_start + batch_size]
            convert_dir = Path(tempfile.mkdtemp(prefix="tool01_lo_batch_"))
            profile_dir = Path(tempfile.mkdtemp(prefix="tool01_lo_profile_"))
            try:
                existing = [(idx, src, dst) for idx, (src, dst) in enumerate(chunk, start=chunk_start) if src.exists()]
                missing = [(idx, src) for idx, (src, _dst) in enumerate(chunk, start=chunk_start) if not src.exists()]
                for idx, src in missing:
                    outcomes[idx] = (False, f"PPTX 파일이 없습니다: {src}")

                completed = None
                if existing:
                    profile_uri = profile_dir.resolve().as_uri()
                    cmd = [
                        self.executable,
                        "--headless",
                        "--nologo",
                        "--nodefault",
                        "--nofirststartwizard",
                        f"-env:UserInstallation={profile_uri}",
                        "--convert-to",
                        "pdf",
                        "--outdir",
                        str(convert_dir),
                        *[str(src) for _idx, src, _dst in existing],
                    ]
                    try:
                        completed = subprocess.run(
                            cmd,
                            capture_output=True,
                            text=True,
                            timeout=self.timeout_seconds,
                            check=False,
                        )
                    except subprocess.TimeoutExpired as e:
                        for idx, _src, _dst in existing:
                            outcomes[idx] = (False, f"LibreOffice Batch 변환 시간초과: {e}")
                        continue

                detail = ""
                if completed is not None:
                    detail = (completed.stderr or completed.stdout or "").strip()

                for idx, src, dst in existing:
                    generated = convert_dir / f"{src.stem}.pdf"
                    if not generated.exists() or generated.stat().st_size == 0:
                        outcomes[idx] = (
                            False,
                            f"LibreOffice PDF 변환 실패 ({src.name}): {detail or '출력 파일 없음'}",
                        )
                        continue
                    try:
                        shutil.move(str(generated), str(dst))
                        outcomes[idx] = (True, None)
                    except Exception as e:  # noqa: BLE001
                        outcomes[idx] = (False, f"PDF 결과 이동 실패 ({dst.name}): {e}")
            finally:
                shutil.rmtree(convert_dir, ignore_errors=True)
                shutil.rmtree(profile_dir, ignore_errors=True)

        return outcomes


# ---------------------------------------------------------------------------
# Renderer selection


def create_renderer(name: str | None = None):
    """Create a renderer context manager.

    Selection order:
    - explicit ``name``
    - ``TOOL01_RENDERER`` environment variable
    - ``auto``

    ``auto`` keeps the existing customer behavior on Windows (PowerPoint COM)
    and chooses LibreOffice in Claude Code Remote/Linux.
    """
    selected = (name or os.environ.get("TOOL01_RENDERER") or "auto").strip().lower()

    if selected == "auto":
        selected = "powerpoint" if platform.system() == "Windows" else "libreoffice"

    if selected in {"powerpoint", "ppt", "com"}:
        return PowerPointRenderer()
    if selected in {"libreoffice", "lo", "soffice"}:
        return LibreOfficeRenderer()

    raise RendererUnavailableError(
        f"지원하지 않는 Renderer: {selected}. 사용 가능: auto, powerpoint, libreoffice"
    )


def convert_single(
    pptx_path: str | Path,
    pdf_path: str | Path,
    *,
    renderer_name: str | None = None,
) -> Path:
    """단발성 변환(Preview) helper."""
    with create_renderer(renderer_name) as renderer:
        return renderer.convert(pptx_path, pdf_path)
