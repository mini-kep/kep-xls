import os
from pathlib import Path
from io import StringIO
import pandas as pd


OUTPUT_FOLDER = 'output'


def validate(frequency: str):
    if frequency.upper() not in ['A', 'Q', 'M']:
        raise ValueError(
            f"frequency must be 'A', 'Q' or 'M', got '{frequency}'")


def output_csv_path(freq: str, folder: str = OUTPUT_FOLDER):
    validate(freq)
    filename = 'df{}.csv'.format(freq.lower())
    return os.path.join(folder, filename)


def read_csv(source) -> pd.DataFrame():
    """
    Wrapper for pd.read_csv(). Treats first column as time index.
    """
    return pd.read_csv(source,
                       converters={0: pd.to_datetime},
                       index_col=0)


def proxy(path):
    """A workaround for pandas problem with non-ASCII paths on Windows
       See <https://github.com/pandas-dev/pandas/issues/15086>
       Args:
           path (pathlib.Path) - CSV filepath
       Returns:
           io.StringIO with CSV content
    """
    content = Path(path).read_text()
    return StringIO(content)


def read_df(freq):
    path = output_csv_path(freq)
    filelike = proxy(path)
    return read_csv(filelike)


def download_dataframe(frequency):
    validate(frequency)
    freq = frequency.lower()
    url = ('https://raw.githubusercontent.com/mini-kep/kep-xls/master/output/'
           f'df{freq}.csv')
    return pd.read_csv(url, converters={0: pd.to_datetime}, index_col=0)


def download_dataframes():
    dfa, dfq, dfm = [download_dataframe(frequency) for frequency in 'aqm']
    return dfa, dfq, dfm


def download_annual():
    return download_dataframe('a')


def download_quarterly():
    return download_dataframe('q')


def download_monthly():
    return download_dataframe('m')
