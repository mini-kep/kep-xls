from kep.files import download_and_unpack
from pathlib import Path
import pytest

@pytest.fixture
def xlpath(tempdir):
    from kep.settings import latest_date
    yield download_and_unpack(*latest_date(), tempdir)
    
def test_download_and_unpack(xlpath):
    import os
    assert os.path.exists(xlpath)
    assert str(xlpath).endswith('.xlsx')