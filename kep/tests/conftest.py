from pathlib import Path
import pytest

@pytest.fixture
def tempdir(tmpdir):
    yield Path(str(tmpdir))