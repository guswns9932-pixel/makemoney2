"""
Tool01 GUI 오류 콜백 경로 테스트.

[CLAUDE REVISION TASK 필수 수정 2 - Final Pre-Acceptance Fix]
Preview/Batch Worker Thread에서 예외 변수(e)를 `lambda: ... e ...` 형태로
after() 콜백에 늦게 참조하면, except 블록 종료 시 파이썬이 e를 자동으로
정리하기 때문에 콜백 실행 시점에 NameError가 발생할 수 있었다.
이 테스트는 실제로 Tkinter 이벤트 루프를 통해 Worker Thread -> after() 콜백
경로를 끝까지 실행시켜, 콜백이 NameError 없이 정상적으로 오류 메시지를
전달하는지 확인한다.

Windows COM이 없는 이 샌드박스에서는 RendererUnavailableError가 자연스럽게
발생하므로, 그 경로를 그대로 이용해 콜백 안전성을 검증한다.

참고: 이 버그는 GIL 스위칭 타이밍에 따라 간헐적으로만 재현되는 Race Condition이다
(worker thread가 `except` 블록을 벗어나며 `e`를 정리하기 전에 main thread가
콜백을 실행하면 우연히 성공한다). 별도의 격리된 최소 재현 스크립트로는
`NameError: cannot access free variable 'e' ...`가 확인되었다 — 이 테스트
스위트에서 매번 재현되지 않더라도 수정(위치 인자로 메시지를 즉시 전달)이
불필요한 것은 아니다.

DISPLAY가 없는 환경(일반 pytest 실행)에서는 스킵된다.
`xvfb-run -a python3 -m pytest tests/test_main_gui.py -v` 로 실행한다.
"""

from __future__ import annotations

import os
import platform
import time

import pytest


def _gui_display_available() -> bool:
    """DISPLAY 문자열 존재가 아니라 실제 Tk 연결 가능 여부를 확인한다."""
    if platform.system() == "Windows" or not os.environ.get("DISPLAY"):
        return False
    try:
        import tkinter as tk

        root = tk.Tk()
        root.withdraw()
        root.update_idletasks()
        root.destroy()
        return True
    except tk.TclError:
        return False


pytestmark = pytest.mark.skipif(
    not _gui_display_available(),
    reason="Tkinter GUI 테스트는 실제 연결 가능한 Linux DISPLAY(예: xvfb-run)에서만 실행",
)


def _run_mainloop_until(app, predicate, timeout=5.0):
    """
    app.mainloop()을 실제로 돌리면서 predicate()가 참이 되거나 timeout에 도달하면
    app.quit()으로 루프를 빠져나온다.

    Tkinter의 after()는 실제로 mainloop이 실행 중이어야 콜백이 정상 등록/실행되므로
    (수동 update() 폴링은 "main thread is not in main loop" 오류를 유발함),
    background thread가 after()를 안전하게 호출할 수 있도록 진짜 mainloop을 돌린다.
    """
    deadline = time.time() + timeout

    def poll():
        if predicate() or time.time() > deadline:
            app.quit()
        else:
            app.after(50, poll)

    app.after(50, poll)
    app.mainloop()
    return predicate()


def test_preview_error_callback_no_late_binding_crash(sample_xlsx, sample_pptx, tmp_output_dir, monkeypatch):
    from src.main import Tool01App

    from src import batch_engine, pdf_renderer

    def _fail_preview(*args, **kwargs):
        raise pdf_renderer.RendererUnavailableError("forced renderer error")

    monkeypatch.setattr(batch_engine, "generate_preview", _fail_preview)
    app = Tool01App()
    try:
        app.data_path = str(sample_xlsx)
        app.template_path = str(sample_pptx)
        app.output_dir = str(tmp_output_dir)
        app._run_preflight_clicked()
        app.update()
        assert app.preflight is not None and app.preflight.ok

        captured: list[str] = []
        app._show_error = lambda message: captured.append(message)  # noqa: SLF001

        app._run_preview_clicked()

        ok = _run_mainloop_until(app, lambda: len(captured) > 0, timeout=5.0)
        assert ok, "Preview 오류 콜백이 호출되지 않았다 (late-binding으로 콜백 내부에서 조용히 실패했을 가능성)"
        assert "forced renderer error" in captured[0]
    finally:
        app.destroy()


def test_batch_error_callback_no_late_binding_crash(sample_xlsx, sample_pptx, tmp_output_dir, monkeypatch):
    from src.main import Tool01App

    from src import batch_engine, pdf_renderer

    def _fail_batch(*args, **kwargs):
        raise pdf_renderer.RendererUnavailableError("forced renderer error")

    monkeypatch.setattr(batch_engine, "run_batch", _fail_batch)
    app = Tool01App()
    try:
        app.data_path = str(sample_xlsx)
        app.template_path = str(sample_pptx)
        app.output_dir = str(tmp_output_dir)
        app._run_preflight_clicked()
        app.update()
        assert app.preflight is not None and app.preflight.ok

        captured: list[str] = []
        app._show_error = lambda message: captured.append(message)  # noqa: SLF001

        app._run_batch_clicked()

        ok = _run_mainloop_until(app, lambda: len(captured) > 0, timeout=5.0)
        assert ok, "Batch 오류 콜백이 호출되지 않았다 (late-binding으로 콜백 내부에서 조용히 실패했을 가능성)"
        assert "forced renderer error" in captured[0]
    finally:
        app.destroy()


def test_automation_busy_disables_overlapping_actions(sample_xlsx, sample_pptx, tmp_output_dir):
    from src.main import Tool01App

    app = Tool01App()
    try:
        app.data_path = str(sample_xlsx)
        app.template_path = str(sample_pptx)
        app.output_dir = str(tmp_output_dir)
        app._run_preflight_clicked()
        app.update()
        assert app.preflight is not None and app.preflight.ok
        assert str(app.preview_button["state"]) == "normal"
        assert str(app.generate_button["state"]) == "normal"

        app._set_automation_busy(True)
        assert app._automation_busy is True
        assert str(app.preview_button["state"]) == "disabled"
        assert str(app.generate_button["state"]) == "disabled"
        assert str(app.check_button["state"]) == "disabled"

        app._set_automation_busy(False)
        assert app._automation_busy is False
        assert str(app.preview_button["state"]) == "normal"
        assert str(app.generate_button["state"]) == "normal"
        assert str(app.check_button["state"]) == "normal"
    finally:
        app.destroy()
