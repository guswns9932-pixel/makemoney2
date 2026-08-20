"""
Preflight Validation.

실행 명세 §4 (Preflight Validation) / §오류처리 기준 대응.
PDF 생성(Batch)을 시작하기 전에 시스템성 오류를 미리 차단한다.
행 단위 오류(빈 셀 등)는 여기서 막지 않고 Batch 처리 중 개별 실패로 처리한다.
"""

from __future__ import annotations

import os
from dataclasses import dataclass, field
from pathlib import Path

from . import data_loader, template_engine


@dataclass
class PreflightResult:
    ok: bool
    errors: list[str] = field(default_factory=list)
    headers: list[str] = field(default_factory=list)
    rows: list[dict] = field(default_factory=list)
    placeholders: list[str] = field(default_factory=list)
    mapping: dict[str, str] = field(default_factory=dict)  # placeholder -> 데이터 컬럼명 (동일)
    missing_fields: list[str] = field(default_factory=list)


def _check_output_dir_writable(output_dir: Path) -> str | None:
    """출력 폴더 쓰기 권한을 실제로 시도해서 확인한다. 문제 있으면 오류 메시지, 없으면 None."""
    try:
        output_dir.mkdir(parents=True, exist_ok=True)
    except OSError as e:
        return f"출력 폴더를 생성할 수 없습니다: {e}"

    probe = output_dir / ".tool01_write_test.tmp"
    try:
        with open(probe, "w") as f:
            f.write("write test")
    except OSError as e:
        return f"출력 폴더에 쓰기 권한이 없습니다: {e}"
    finally:
        try:
            if probe.exists():
                probe.unlink()
        except OSError:
            pass  # 삭제 실패는 치명적이지 않음 (권한 문제는 이미 위에서 판정됨)
    return None


def run_preflight(
    data_path: str | Path | None,
    template_path: str | Path | None,
    output_dir: str | Path | None,
) -> PreflightResult:
    errors: list[str] = []

    if not data_path:
        errors.append("데이터 파일이 선택되지 않았습니다.")
    if not template_path:
        errors.append("PPTX 템플릿 파일이 선택되지 않았습니다.")
    if not output_dir:
        errors.append("출력 폴더가 선택되지 않았습니다.")

    if errors:
        return PreflightResult(ok=False, errors=errors)

    headers: list[str] = []
    rows: list[dict] = []
    placeholders: list[str] = []

    try:
        headers, rows = data_loader.load_data(data_path)
    except data_loader.DataLoadError as e:
        errors.append(str(e))

    try:
        placeholders = template_engine.extract_placeholders(template_path)
    except template_engine.TemplateError as e:
        errors.append(str(e))

    if not errors and not placeholders:
        errors.append("템플릿에서 {{필드명}} 형식의 Placeholder를 찾을 수 없습니다.")

    missing_fields: list[str] = []
    mapping: dict[str, str] = {}
    if not errors:
        header_set = set(headers)
        for ph in placeholders:
            if ph in header_set:
                mapping[ph] = ph
            else:
                missing_fields.append(ph)
        if missing_fields:
            errors.append(
                "Placeholder에 대응하는 데이터 컬럼이 없습니다: " + ", ".join(missing_fields)
            )

    output_dir_path = Path(output_dir)
    write_err = _check_output_dir_writable(output_dir_path)
    if write_err:
        errors.append(write_err)

    ok = not errors
    return PreflightResult(
        ok=ok,
        errors=errors,
        headers=headers,
        rows=rows,
        placeholders=placeholders,
        mapping=mapping,
        missing_fields=missing_fields,
    )
