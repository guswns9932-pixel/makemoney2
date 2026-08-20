import sys
from pathlib import Path

# products/tool01 을 sys.path에 추가해 `from src import ...`가 되게 한다.
ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import pytest  # noqa: E402

ASSETS_DIR = ROOT / "assets"


@pytest.fixture
def sample_xlsx() -> Path:
    return ASSETS_DIR / "sample_roster.xlsx"


@pytest.fixture
def sample_pptx() -> Path:
    return ASSETS_DIR / "certificate_template.pptx"


@pytest.fixture
def tmp_output_dir(tmp_path) -> Path:
    d = tmp_path / "output"
    d.mkdir()
    return d
