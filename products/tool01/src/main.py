"""
Tool01 Demo — 수료증·상장·확인서·증명서 대량생성 GUI.

실행 명세 §필수요구사항-1 (단일 화면 GUI) 대응.
Windows + Microsoft PowerPoint Desktop 환경에서 실행한다 (Baseline 확정사항).

실행 흐름:
    데이터 파일 선택 → PPTX 템플릿 선택 → 출력 폴더 선택
    → 자동매핑/검증 → Preview → PDF 일괄생성 → 결과 확인 → 결과 폴더 열기

미검증: 실제 Windows + PowerPoint Desktop 환경에서 GUI 동작 확인 필요
(이 코드는 Linux 샌드박스에서 작성되었고 디스플레이 환경이 없어 실행 테스트를 하지 못했다).
"""

from __future__ import annotations

import os
import platform
import subprocess
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

# 패키지/스크립트 양쪽 실행 모두 지원
try:
    from . import batch_engine, pdf_renderer, validator
except ImportError:  # python main.py 로 직접 실행하는 경우
    sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
    from src import batch_engine, pdf_renderer, validator  # type: ignore


APP_TITLE = "Tool01 Demo — 수료증/상장/확인서 대량생성"


class Tool01App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("720x560")
        self.resizable(True, True)

        self.data_path: str | None = None
        self.template_path: str | None = None
        self.output_dir: str | None = None
        self.preflight: validator.PreflightResult | None = None
        self._automation_busy = False

        self._build_ui()
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    # ------------------------------------------------------------------ UI
    def _build_ui(self) -> None:
        pad = {"padx": 10, "pady": 6}

        # 1) 파일 선택 영역
        file_frame = ttk.LabelFrame(self, text="1. 파일 선택")
        file_frame.pack(fill="x", **pad)

        self.data_var = tk.StringVar(value="(선택 안 됨)")
        self.template_var = tk.StringVar(value="(선택 안 됨)")
        self.output_var = tk.StringVar(value="(선택 안 됨)")

        self._add_file_row(file_frame, "데이터 파일 (xlsx/csv)", self.data_var, self._pick_data)
        self._add_file_row(file_frame, "PPTX 템플릿", self.template_var, self._pick_template)
        self._add_file_row(file_frame, "출력 폴더", self.output_var, self._pick_output_dir)

        # 2) 검증/매핑 결과
        check_frame = ttk.LabelFrame(self, text="2. 자동매핑 / 검증")
        check_frame.pack(fill="both", expand=False, **pad)

        self.check_button = ttk.Button(check_frame, text="파일 확인 및 자동매핑", command=self._run_preflight_clicked)
        self.check_button.pack(anchor="w", padx=8, pady=4)

        self.status_text = tk.Text(check_frame, height=8, wrap="word", state="disabled")
        self.status_text.pack(fill="both", expand=True, padx=8, pady=4)

        # 3) Preview / 생성
        action_frame = ttk.LabelFrame(self, text="3. Preview / PDF 일괄생성")
        action_frame.pack(fill="x", **pad)

        btn_row = ttk.Frame(action_frame)
        btn_row.pack(fill="x", padx=8, pady=4)

        self.preview_button = ttk.Button(btn_row, text="Preview (첫 행)", command=self._run_preview_clicked, state="disabled")
        self.preview_button.pack(side="left", padx=(0, 8))

        self.generate_button = ttk.Button(btn_row, text="PDF 일괄생성", command=self._run_batch_clicked, state="disabled")
        self.generate_button.pack(side="left")

        self.progress = ttk.Progressbar(action_frame, mode="determinate")
        self.progress.pack(fill="x", padx=8, pady=(4, 8))

        self.progress_label = ttk.Label(action_frame, text="")
        self.progress_label.pack(anchor="w", padx=8)

        # 4) 결과
        result_frame = ttk.LabelFrame(self, text="4. 결과")
        result_frame.pack(fill="both", expand=True, **pad)

        self.result_text = tk.Text(result_frame, height=8, wrap="word", state="disabled")
        self.result_text.pack(fill="both", expand=True, padx=8, pady=4)

        self.open_folder_button = ttk.Button(
            result_frame, text="결과 폴더 열기", command=self._open_output_folder, state="disabled"
        )
        self.open_folder_button.pack(anchor="e", padx=8, pady=4)

    def _add_file_row(self, parent, label, var, command) -> None:
        row = ttk.Frame(parent)
        row.pack(fill="x", padx=8, pady=3)
        ttk.Label(row, text=label, width=22).pack(side="left")
        ttk.Label(row, textvariable=var, foreground="#333").pack(side="left", fill="x", expand=True)
        ttk.Button(row, text="선택", command=command).pack(side="right")

    # ------------------------------------------------------------- 파일선택
    def _pick_data(self) -> None:
        path = filedialog.askopenfilename(
            title="데이터 파일 선택", filetypes=[("Excel/CSV", "*.xlsx *.csv")]
        )
        if path:
            self.data_path = path
            self.data_var.set(path)
            self._reset_after_input_change()

    def _pick_template(self) -> None:
        path = filedialog.askopenfilename(
            title="PPTX 템플릿 선택", filetypes=[("PowerPoint", "*.pptx")]
        )
        if path:
            self.template_path = path
            self.template_var.set(path)
            self._reset_after_input_change()

    def _pick_output_dir(self) -> None:
        path = filedialog.askdirectory(title="출력 폴더 선택")
        if path:
            self.output_dir = path
            self.output_var.set(path)
            self._reset_after_input_change()

    def _reset_after_input_change(self) -> None:
        self.preflight = None
        self.preview_button["state"] = "disabled"
        self.generate_button["state"] = "disabled"
        self._set_text(self.status_text, "")

    # -------------------------------------------------------------- 검증
    def _run_preflight_clicked(self) -> None:
        result = validator.run_preflight(self.data_path, self.template_path, self.output_dir)
        self.preflight = result

        lines = []
        if result.headers:
            lines.append(f"데이터 컬럼: {', '.join(result.headers)}")
            lines.append(f"데이터 행 수: {len(result.rows)}")
        if result.placeholders:
            lines.append(f"템플릿 Placeholder: {', '.join(result.placeholders)}")
        if result.mapping:
            mapped = ", ".join(f"{{{{{k}}}}} → {v}" for k, v in result.mapping.items())
            lines.append(f"자동매핑: {mapped}")
        if result.errors:
            lines.append("")
            lines.append("[문제 발견]")
            lines.extend(f"- {e}" for e in result.errors)
        else:
            lines.append("")
            lines.append("문제 없음 — Preview / PDF 일괄생성을 진행할 수 있습니다.")

        self._set_text(self.status_text, "\n".join(lines))

        if result.ok:
            self.preview_button["state"] = "normal"
            self.generate_button["state"] = "normal"
        else:
            self.preview_button["state"] = "disabled"
            self.generate_button["state"] = "disabled"

    # -------------------------------------------------------------- Preview
    def _run_preview_clicked(self) -> None:
        if self._automation_busy:
            return
        if not self.preflight or not self.preflight.ok or not self.preflight.rows:
            messagebox.showwarning(APP_TITLE, "먼저 '파일 확인 및 자동매핑'을 통과해야 합니다.")
            return

        # Worker가 시작된 뒤 사용자가 파일 선택을 바꾸더라도 이번 실행은 클릭 시점의
        # 검증된 입력만 사용하도록 snapshot한다.
        template_path = self.template_path
        output_dir = self.output_dir
        first_row = dict(self.preflight.rows[0])
        mapping = dict(self.preflight.mapping)
        self._set_automation_busy(True)

        def task():
            try:
                work_dir = Path(output_dir) / "_tool01_preview_tmp"
                pdf_path = batch_engine.generate_preview(
                    template_path, first_row, mapping, work_dir
                )
                self.after(0, self._open_file, pdf_path)
            except pdf_renderer.RendererUnavailableError as e:
                message = str(e)
                self.after(0, self._show_error, message)
            except Exception as e:  # noqa: BLE001
                message = f"Preview 생성 중 오류: {e}"
                self.after(0, self._show_error, message)
            finally:
                self.after(0, self._set_automation_busy, False)

        threading.Thread(target=task, daemon=True).start()

    # -------------------------------------------------------------- Batch
    def _run_batch_clicked(self) -> None:
        if self._automation_busy:
            return
        if not self.preflight or not self.preflight.ok:
            messagebox.showwarning(APP_TITLE, "먼저 '파일 확인 및 자동매핑'을 통과해야 합니다.")
            return

        template_path = self.template_path
        output_dir = self.output_dir
        rows = [dict(row) for row in self.preflight.rows]
        mapping = dict(self.preflight.mapping)

        self._set_automation_busy(True)
        self.progress["value"] = 0
        self.progress["maximum"] = len(rows)

        def progress_cb(current: int, total: int) -> None:
            self.after(0, self._update_progress, current, total)

        def task():
            try:
                result = batch_engine.run_batch(
                    template_path,
                    rows,
                    mapping,
                    output_dir,
                    progress_cb=progress_cb,
                )
                self.after(0, self._show_batch_result, result)
            except pdf_renderer.RendererUnavailableError as e:
                message = str(e)
                self.after(0, self._show_error, message)
            except Exception as e:  # noqa: BLE001
                message = f"PDF 일괄생성 중 오류: {e}"
                self.after(0, self._show_error, message)
            finally:
                self.after(0, self._set_automation_busy, False)

        threading.Thread(target=task, daemon=True).start()

    def _update_progress(self, current: int, total: int) -> None:
        self.progress["value"] = current
        self.progress_label["text"] = f"{current} / {total} 처리 중..."

    def _set_automation_busy(self, busy: bool) -> None:
        """Preview/Batch PowerPoint 자동화가 동시에 겹치지 않게 UI 상태를 관리한다."""
        self._automation_busy = busy
        if busy:
            self.preview_button["state"] = "disabled"
            self.generate_button["state"] = "disabled"
            self.check_button["state"] = "disabled"
            return

        self.check_button["state"] = "normal"
        can_run = bool(self.preflight and self.preflight.ok)
        self.preview_button["state"] = "normal" if can_run else "disabled"
        self.generate_button["state"] = "normal" if can_run else "disabled"

    def _re_enable_action_buttons(self) -> None:
        # 하위 호환용 내부 wrapper. 신규 코드는 _set_automation_busy(False)를 사용한다.
        self._set_automation_busy(False)

    def _show_batch_result(self, result: batch_engine.BatchResult) -> None:
        lines = [
            f"전체 대상: {result.total}건",
            f"성공: {result.success_count}건",
            f"실패: {result.failure_count}건",
        ]
        if result.failures:
            lines.append("")
            lines.append("[실패 상세]")
            for r in result.failures:
                lines.append(f"- 행 {r.row_index}: {r.error}")

        self._set_text(self.result_text, "\n".join(lines))
        self.progress_label["text"] = "완료"
        self.open_folder_button["state"] = "normal"

    # -------------------------------------------------------------- 유틸
    def _open_output_folder(self) -> None:
        if self.output_dir:
            self._open_file(self.output_dir)

    @staticmethod
    def _open_file(path) -> None:
        path = str(path)
        system = platform.system()
        try:
            if system == "Windows":
                os.startfile(path)  # type: ignore[attr-defined]
            elif system == "Darwin":
                subprocess.run(["open", path], check=False)
            else:
                subprocess.run(["xdg-open", path], check=False)
        except Exception:  # noqa: BLE001
            pass  # 결과 폴더/파일을 여는 데 실패해도 치명적이지 않음

    def _show_error(self, message: str) -> None:
        messagebox.showerror(APP_TITLE, message)

    def _cleanup_preview_files(self) -> None:
        if not self.output_dir:
            return
        work_dir = Path(self.output_dir) / "_tool01_preview_tmp"
        batch_engine.cleanup_preview_files(work_dir, remove_dir=True)

    def _on_close(self) -> None:
        if self._automation_busy:
            messagebox.showwarning(APP_TITLE, "Preview/PDF 생성 작업이 끝난 뒤 종료해 주세요.")
            return
        self._cleanup_preview_files()
        self.destroy()

    @staticmethod
    def _set_text(widget: tk.Text, content: str) -> None:
        widget["state"] = "normal"
        widget.delete("1.0", tk.END)
        widget.insert(tk.END, content)
        widget["state"] = "disabled"


def main() -> None:
    app = Tool01App()
    app.mainloop()


if __name__ == "__main__":
    main()
