LOCATIONS = [
    ('1.1. ', 'B9', 'A', 'GDP_RUB'),
    ('1.1. ', 'C9', 'Q', 'GDP_RUB'),

    ('1.1. ', 'B32', 'A', 'GDP_INDEX'),
    ('1.1. ', 'C32', 'Q', 'GDP_INDEX'),

    ('1.6. ', 'C5', 'Q', 'INVEST_RUB'),
    ('1.6. ', 'B5', 'A', 'INVEST_RUB'),

    ('1.6. ', 'B28', 'A', 'INVEST_INDEX'),
    ('1.6. ', 'C51', 'Q', 'INVEST_INDEX'),

    ('1.5. ', 'G73', 'M', 'COMM_FREIGHT'),

    ('1.12 ', 'G5', 'M', 'RETAIL_SALES'),
    ('1.12 ', 'B5', 'A', 'RETAIL_SALES'),

    ('2.1 ', 'B11', 'A', 'GOV_INC_CONS_ACCUM'),
    ('2.1 ', 'B404', 'A', 'GOV_EXP_CONS_ACCUM'),

    ('3.5 ', 'B7', 'A', 'CPI'),
    ('3.5 ', 'C7', 'Q', 'CPI'),
    ('3.5 ', 'G7', 'M', 'CPI'),

    ('4.2 ', 'B7', 'A', 'WAGE'),
    ('4.2 ', 'C7', 'Q', 'WAGE'),
    ('4.2 ', 'G7', 'M', 'WAGE'),

]


"""Download and unpack Excel file from Rosstat website"""

from pathlib import Path
import subprocess
import platform
import os
import wget

SOURCE_FOLDER = 'source'
OUTPUT_FOLDER = 'output'

URL = 'https://gks.ru/storage/mediabank/ind07(2).rar'
RAR = os.path.join(SOURCE_FOLDER, 'kep_2019_07.rar')
XLS = os.path.join(SOURCE_FOLDER, 'краткосрочные_июль.xlsx')


def make_csvpath(freq: str, folder: str = OUTPUT_FOLDER):
    filename = 'df{}.csv'.format(freq.lower())
    return os.path.join(folder, filename)


def output_xl_path(folder: str = OUTPUT_FOLDER):
    return os.path.join(folder, 'df.xlsx')


def unrar_executable():
    if platform.system() == 'Windows':
        current_folder = Path(__file__).parent
        return str(current_folder / 'bin' / 'UnRAR.exe')
    else:
        return 'unrar'


def unrar(args):
    args = [unrar_executable()] + args
    return subprocess.run(args, capture_output=True)


def unpack(filepath: str, destination_folder: str = '.'):
    # UnRAR wants its folder argument to end with '/'
    folder = "{}{}".format(destination_folder, os.sep)
    return unrar(['e', filepath, folder, '-y', '-o+'])


def list_content(filepath: str):
    """Windows on cp866 only"""
    output = unrar(['lb', filepath])
    return output.stdout.decode('cp866').strip()


def make_url(year, month):
    mapper = {(2019, 7): 'ind07(2).rar'}
    try:
        filename = mapper[(year, month)]
    except KeyError:
        raise ValueError('Supported values are: {}'.format(mapper.keys()))
    return f'https://gks.ru/storage/mediabank/{filename}'


def make_xl_filename(year, month):
    mapper = {(2019, 7): 'краткосрочные_июль.xlsx'}
    try:
        return mapper[(year, month)]
    except KeyError:
        raise ValueError('Supported values are: {}'.format(mapper.keys()))


def make_xlspath(year, month, folder=SOURCE_FOLDER):
    filename = make_xl_filename(year, month)
    return os.path.join(folder, filename)


def make_local_rar_filename(year, month):
    month = str(month).zfill(2)
    return f'kep_{year}_{month}.rar'


assert make_local_rar_filename(2019, 7) == 'kep_2019_07.rar'


def download_and_unpack(year, month, folder=SOURCE_FOLDER):
    url = make_url(year, month)
    rarpath = os.path.join(folder, make_local_rar_filename(year, month))
    xlspath = make_xlspath(year, month, folder)
    if not Path(rarpath).exists():
        wget.download(url, rarpath)
        print('Downloaded', url)
        print('Saved as', rarpath)
    if not Path(xlspath).exists():
        unpack(rarpath, folder)
        print('Unpacked', xlspath)
    return xlspath

from enum import Enum
import re
import os
from dataclasses import dataclass

import pandas as pd
import xlwings as xw


def uncomment_string(s: str):
    """See https://regex101.com/r/LdAJBH/1"""
    pat = r"([\d,.]*)(?:\d+\))*$"
    return re.search(pattern=pat, string=s).group(1).replace(',', '.')


assert uncomment_string('244873)') == '24487'


def uncomment(x, type_):
    if isinstance(x, str):
        x = uncomment_string(x)
    return type_(x)


@dataclass
class Source:
    path: str
    sheet: str

    def xl_sheet(self):
        return xw.Book(self.path).sheets[self.sheet]

    def anchor(self, cell):
        years = find_years(self.xl_sheet(), cell)
        return Location(self.xl_sheet().range(cell), years)


def find_years(sheet, cell: str) -> [int]:
    cell = sheet.range(sheet.range(cell).row - 1, 1)
    result = []
    while True:
        cell = cell.offset(1, 0)
        try:
            year = uncomment(cell.value, int)
        except (ValueError, TypeError):
            break
        result.append(year)
    return result


class Freq(Enum):
    ANNUAL = 'A'
    QTR = 'Q'
    MONTH = 'M'


@dataclass
class Location:
    anchor: xw.Range
    years: [int]

    @property
    def width(self):
        return dict(A=1, Q=4, M=12)[self.freq]

    @property
    def height(self):
        return len(self.years)

    def freq(self, freq):
        self.freq = freq.upper()
        a = self.anchor
        b = a.offset(self.height, self.width - 1)
        self.range = xw.Range(a, b)
        return self

    def values(self):
        return self.range.value

    def flat(self):
        if self.freq == Freq.ANNUAL.value:
            return [x for x in self.values() if x]
        else:
            return [x for xs in self.values() for x in xs if x]

    def observations(self):
        return [uncomment(x, float) for x in self.flat()]

    def make_index(self, n: int):
        start_year = self.years[0]
        return pd.date_range(start=f'1/1/{start_year}',
                             freq=self.freq,
                             periods=n)

    def to_series(self, name):
        data = self.observations()
        ix = self.make_index(len(data))
        ts = pd.Series(data=data, index=ix)
        ts.name = name
        return ts


@dataclass
class Variable:
    sheet: str
    cell: str
    freq: str
    varname: str

    def obj(self, file):
        return Source(file, self.sheet).anchor(self.cell).freq(self.freq)

    def get(self, file):
        return self.obj(file).to_series(self.varname)


FREQUENCIES = 'AQM'


class DataframeDict(dict):
    def __init__(self):
        dict0 = {key: pd.DataFrame() for key in FREQUENCIES}
        super().__init__(dict0)

    def all_variables(self):
        vars = set()
        for freq in FREQUENCIES:
            vars.update(self[freq].columns)
        return sorted(vars)

    def save_csv(self):
        for freq in FREQUENCIES:
            filename = make_csvpath(freq)
            self[freq].to_csv(filename)
            yield filename

    def save_xls(self):
        filepath = output_xl_path()
        save_excel(filepath, self)
        return filepath


def save_excel(filepath: str, df_dict: DataframeDict):
    """Save *df_dict* as .xlsx at *filepath*."""
    sheet_names = dict(A='year', Q='quarter', M='month')
    with pd.ExcelWriter(filepath) as writer:
        for freq in FREQUENCIES:
            df_dict[freq].to_excel(writer, sheet_name=sheet_names[freq])
        # TODO: add variable names
        # self.df_vars().to_excel(writer, sheet_name='variables')


def extract_series(xlspath, locations):
    series = []
    for loc in locations:
        v = Variable(*loc)
        ts = v.get(xlspath)
        print("Finished reading", v.varname, "at frequency", v.freq)
        series.append((v, ts))
    return series


def create_dataframes(xlspath) -> dict:
    from locations import LOCATIONS
    series = extract_series(xlspath, locations=LOCATIONS)
    df = DataframeDict()
    for freq in FREQUENCIES:  # reading via xlwings is slow!
        for v, ts in series:
            if v.freq == freq:
                df[freq] = pd.concat([df[freq], ts], axis=1)
                # Attempt preserving the initial timestamps in index.
                # Without it the timestamps get additional detail.
                try:
                    df[freq].index = ts.index
                except ValueError:
                    pass

    return df
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
