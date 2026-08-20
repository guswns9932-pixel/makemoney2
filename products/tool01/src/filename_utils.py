"""
파일명 안전처리 유틸리티.

- Windows 금지문자 제거/치환
- 앞뒤 공백 제거
- 파일명 필수값(빈 값) 검증
- 동일 파일명 존재 시 순번을 붙여 보존 (덮어쓰기 금지)

실행 명세 §9 (파일명 안전처리) / §12 (원본 보호) 대응.
"""

from __future__ import annotations

import os
import re
from pathlib import Path

# Windows 파일명 금지문자: \ / : * ? " < > |
_FORBIDDEN_CHARS_PATTERN = re.compile(r'[\\/:*?"<>|]')

# Windows 예약 파일명 (확장자 없이 비교)
_RESERVED_NAMES = {
    "CON", "PRN", "AUX", "NUL",
    *(f"COM{i}" for i in range(1, 10)),
    *(f"LPT{i}" for i in range(1, 10)),
}


class FilenameBuildError(ValueError):
    """파일명을 안전하게 구성할 수 없을 때 발생 (예: 필수 필드 값이 비어 있음)."""


def sanitize_component(value: str) -> str:
    """파일명 한 조각(필드 값)에서 금지문자를 제거하고 앞뒤 공백을 정리한다."""
    if value is None:
        value = ""
    value = str(value).strip()
    value = _FORBIDDEN_CHARS_PATTERN.sub("", value)
    # 제어문자 제거
    value = "".join(ch for ch in value if ch.isprintable())
    value = value.strip()
    return value


def build_filename(pattern: str, row: dict, required_fields: list[str]) -> str:
    """
    패턴(예: "{수료번호}_{이름}_수료증.pdf")에 row 데이터를 채워 안전한 파일명을 만든다.

    required_fields에 있는 필드는 sanitize 후 빈 문자열이면 FilenameBuildError를 낸다.
    (해당 행은 실패 처리 대상 — 실행 명세 §11 오류/예외 처리 기준 항목 11)
    """
    values = {}
    for key, raw in row.items():
        values[key] = sanitize_component(raw)

    missing = [f for f in required_fields if not values.get(f)]
    if missing:
        raise FilenameBuildError(
            f"파일명 필수 필드 값이 비어 있음: {', '.join(missing)}"
        )

    try:
        filename = pattern.format(**values)
    except KeyError as e:
        raise FilenameBuildError(f"파일명 패턴에 존재하지 않는 필드 참조: {e}") from e

    filename = filename.strip()
    if not filename or filename in (".", ".."):
        raise FilenameBuildError("생성된 파일명이 비어 있거나 유효하지 않음")

    stem, ext = os.path.splitext(filename)
    if stem.upper() in _RESERVED_NAMES:
        stem = f"_{stem}"
        filename = stem + ext

    return filename


def unique_path(directory: Path, filename: str) -> Path:
    """
    directory/filename이 이미 존재하면 기존 파일을 덮어쓰지 않고
    "_2", "_3" ... 순번을 붙인 새 경로를 반환한다.
    """
    directory = Path(directory)
    stem, ext = os.path.splitext(filename)
    candidate = directory / filename
    counter = 2
    while candidate.exists():
        candidate = directory / f"{stem}_{counter}{ext}"
        counter += 1
    return candidate
