from kep import Source
from files import download_and_unpack, make_xlspath
import xlwings as xw

# FIXME: skip if no Excel found


def test_source(tmpdir):
    download_and_unpack(2019, 7, tmpdir)
    xlspath = make_xlspath(2019, 7, tmpdir)
    gdp_q = Source(xlspath, '1.1. ').anchor('C9').freq('Q').to_series('GDP')
    assert gdp_q['1999'][0] == 901.0
    xw.Book(xlspath).app.quit()
