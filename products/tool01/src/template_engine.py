"""
PPTX 템플릿의 {{필드명}} Placeholder 인식 및 치환.

실행 명세 §3 (Placeholder 인식) / §7 (PPTX 내용 치환) 대응.
- 일반 Text Box 안의 텍스트 Placeholder 기준 (Demo 범위)
- 원본 Template은 절대 수정하지 않음 (읽어서 새 사본에만 반영)
- Placeholder가 아닌 문구는 변경하지 않음
- 대소문자/공백 임의 보정 없음 (정확히 동일한 이름만 매칭)
"""

from __future__ import annotations

import re
import shutil
from pathlib import Path

from pptx import Presentation

_PLACEHOLDER_PATTERN = re.compile(r"\{\{([^{}]+)\}\}")


class TemplateError(ValueError):
    """PPTX 템플릿을 열거나 처리할 수 없을 때 발생."""


def extract_placeholders(pptx_path: str | Path) -> list[str]:
    """
    PPTX의 모든 슬라이드 텍스트에서 {{필드명}} 형식의 Placeholder 이름을
    등장 순서 기준 중복 없이 추출한다.
    """
    pptx_path = Path(pptx_path)
    if not pptx_path.exists():
        raise TemplateError(f"템플릿 파일을 찾을 수 없습니다: {pptx_path}")
    if pptx_path.suffix.lower() != ".pptx":
        raise TemplateError(f"지원하지 않는 템플릿 형식입니다 ({pptx_path.suffix}). pptx만 지원합니다.")

    try:
        prs = Presentation(str(pptx_path))
    except Exception as e:  # noqa: BLE001
        raise TemplateError(f"PPTX 템플릿을 여는 중 오류가 발생했습니다: {e}") from e

    found: list[str] = []
    seen = set()
    for slide in prs.slides:
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    for match in _PLACEHOLDER_PATTERN.finditer(run.text or ""):
                        name = match.group(1).strip()
                        if name and name not in seen:
                            seen.add(name)
                            found.append(name)
    return found


def render_pptx(
    template_path: str | Path,
    output_path: str | Path,
    mapping: dict[str, str],
) -> Path:
    """
    template_path를 읽어(수정하지 않고) mapping({필드명: 값})으로
    Placeholder를 치환한 새 사본을 output_path에 저장한다.

    Run 단위 텍스트 치환만 수행하며 서식(폰트/색상 등)은 그대로 유지된다.
    (Run이 아닌 문구, 즉 Placeholder가 아닌 텍스트는 절대 변경하지 않는다.)
    """
    template_path = Path(template_path)
    output_path = Path(output_path)

    if not template_path.exists():
        raise TemplateError(f"템플릿 파일을 찾을 수 없습니다: {template_path}")

    try:
        prs = Presentation(str(template_path))
    except Exception as e:  # noqa: BLE001
        raise TemplateError(f"PPTX 템플릿을 여는 중 오류가 발생했습니다: {e}") from e

    def _replace(text: str) -> str:
        def _sub(match: re.Match) -> str:
            key = match.group(1).strip()
            return str(mapping.get(key, match.group(0)))

        return _PLACEHOLDER_PATTERN.sub(_sub, text)

    for slide in prs.slides:
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    if run.text and "{{" in run.text:
                        run.text = _replace(run.text)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(output_path))
    return output_path


def copy_template_readonly_check(template_path: str | Path) -> None:
    """
    원본 보호 확인용 헬퍼: 템플릿 파일의 수정시각/크기를 반환해
    호출측(테스트)이 작업 전후 비교로 원본 무변경을 검증할 수 있게 한다.
    """
    template_path = Path(template_path)
    stat = template_path.stat()
    return stat.st_mtime, stat.st_size
