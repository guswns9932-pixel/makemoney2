"""
데이터 파일(XLSX/CSV) 로딩.

실행 명세 §입력-1 (데이터 파일) / §오류처리 항목 4,5,7,8 대응.
- 존재하지 않는 파일
- 지원하지 않는 확장자
- 헤더 없음 / 데이터 0행
- 완전히 빈 행 제외
원본 파일은 읽기 전용으로만 열고 절대 수정하지 않는다.
"""

from __future__ import annotations

import csv
from pathlib import Path

import openpyxl

SUPPORTED_EXTENSIONS = {".xlsx", ".csv"}


class DataLoadError(ValueError):
    """데이터 파일을 정상적으로 읽지 못했을 때 발생."""


def load_data(path: str | Path) -> tuple[list[str], list[dict]]:
    """
    XLSX 또는 CSV 파일을 읽어 (헤더 목록, 행 딕셔너리 목록)을 반환한다.
    완전히 빈 행은 결과에서 제외한다.
    """
    path = Path(path)

    if not path.exists():
        raise DataLoadError(f"데이터 파일을 찾을 수 없습니다: {path}")
    if not path.is_file():
        raise DataLoadError(f"데이터 경로가 파일이 아닙니다: {path}")

    ext = path.suffix.lower()
    if ext not in SUPPORTED_EXTENSIONS:
        raise DataLoadError(
            f"지원하지 않는 데이터 파일 형식입니다 ({ext}). xlsx 또는 csv만 지원합니다."
        )

    if ext == ".xlsx":
        headers, rows = _load_xlsx(path)
    else:
        headers, rows = _load_csv(path)

    if not headers:
        raise DataLoadError("데이터 파일에 헤더(첫 행)가 없습니다.")

    if not rows:
        raise DataLoadError("데이터 파일에 처리 가능한 데이터 행이 없습니다.")

    return headers, rows


def _load_xlsx(path: Path) -> tuple[list[str], list[dict]]:
    try:
        # read_only=True: 원본을 메모리에 스트리밍으로만 읽고 수정하지 않는다.
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    except Exception as e:  # noqa: BLE001 - 사용자에게 이해 가능한 오류로 재포장
        raise DataLoadError(f"XLSX 파일을 여는 중 오류가 발생했습니다: {e}") from e

    try:
        ws = wb.active
        rows_iter = ws.iter_rows(values_only=True)
        try:
            header_row = next(rows_iter)
        except StopIteration:
            return [], []

        headers = [str(h).strip() if h is not None else "" for h in header_row]

        data_rows: list[dict] = []
        for raw_row in rows_iter:
            if raw_row is None:
                continue
            if all(cell is None or str(cell).strip() == "" for cell in raw_row):
                continue  # 완전히 빈 행 제외
            row = {}
            for idx, header in enumerate(headers):
                if not header:
                    continue
                value = raw_row[idx] if idx < len(raw_row) else None
                row[header] = "" if value is None else value
            data_rows.append(row)

        return [h for h in headers if h], data_rows
    finally:
        wb.close()


def _load_csv(path: Path) -> tuple[list[str], list[dict]]:
    try:
        # utf-8-sig: 엑셀에서 저장한 CSV의 BOM을 안전하게 처리 (한글 깨짐 방지)
        with open(path, "r", encoding="utf-8-sig", newline="") as f:
            reader = csv.reader(f)
            try:
                header_row = next(reader)
            except StopIteration:
                return [], []

            headers = [h.strip() for h in header_row]

            data_rows: list[dict] = []
            for raw_row in reader:
                if not raw_row or all(c.strip() == "" for c in raw_row):
                    continue  # 완전히 빈 행 제외
                row = {}
                for idx, header in enumerate(headers):
                    if not header:
                        continue
                    row[header] = raw_row[idx] if idx < len(raw_row) else ""
                data_rows.append(row)

            return [h for h in headers if h], data_rows
    except UnicodeDecodeError as e:
        raise DataLoadError(
            f"CSV 파일 인코딩을 읽을 수 없습니다 (UTF-8 필요): {e}"
        ) from e
    except OSError as e:
        raise DataLoadError(f"CSV 파일을 여는 중 오류가 발생했습니다: {e}") from e
