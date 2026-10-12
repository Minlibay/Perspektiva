import os
import sys
from pathlib import Path

import pytest

BACKEND_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(BACKEND_DIR))

CHECKLISTS_DIR = Path(__file__).resolve().parent / "fixtures" / "checklists"


@pytest.fixture
def checklist_path():
    def _path(name: str) -> Path:
        return CHECKLISTS_DIR / name
    return _path


@pytest.fixture
def samples_dir() -> Path:
    """Размеченные образцы документов («ЧК ИИ/18.09»). В репозитории их нет — там данные
    клиентов; тесты на них запускаются, если задан AUDIT_SAMPLES_DIR."""
    path = os.environ.get("AUDIT_SAMPLES_DIR")
    if not path or not Path(path).is_dir():
        pytest.skip("AUDIT_SAMPLES_DIR не задан — тесты на образцах пропущены")
    return Path(path)
