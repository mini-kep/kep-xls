from enum import Enum
import subprocess
from pathlib import Path
import platform
import re
import os
from dataclasses import dataclass

import pandas as pd
import wget
import xlwings as xw


SOURCE_FOLDER = 'source'
OUTPUT_FOLDER = 'output'

# file management


def output_csv_path(freq: str, folder: str = OUTPUT_FOLDER):
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

# operations on strings


def uncomment_string(s: str):
    """See https://regex101.com/r/LdAJBH/1"""
    pat = r"([\d,.]*)(?:\d+\))*$"
    return re.search(pattern=pat, string=s).group(1).replace(',', '.')


assert uncomment_string('244873)') == '24487'


def uncomment(x, type_):
    if isinstance(x, str):
        x = uncomment_string(x)
    return type_(x)

# xlwings-ralated classes


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
        
    def save_xls(self):    
        filepath=output_xl_path()
        save_excel(filepath, self)
        return filepath


def save_excel(filepath: str, df_dict: DataframeDict):
    """Save *dataframes* as .xlsx at *filepath*."""
    sheet_names = dict(A='year', Q='quarter', M='month')
    with pd.ExcelWriter(filepath) as writer:
        for freq in FREQUENCIES:
            df[freq].to_excel(writer, sheet_name=sheet_names[freq])
        # TODO: add variable names
        # self.df_vars().to_excel(writer, sheet_name='variables')


# This is main section


URL = 'https://gks.ru/storage/mediabank/ind07(2).rar'
RAR = os.path.join(SOURCE_FOLDER, 'kep_2019_07.rar')
XLS = os.path.join(SOURCE_FOLDER, 'краткосрочные_июль.xlsx')


if not Path(RAR).exists():
    wget.download(URL, RAR)


if not Path(XLS).exists():
    unrar(XLS, SOURCE_FOLDER)


if False:
    gdp_q = Source(XLS, '1.1. ').anchor('C9').freq('Q').to_series('GDP')
    gdp_a = Source(XLS, '1.1. ').anchor('B9').freq('A').to_series('GDP')
    frei_m = Source(XLS, '1.5. ').anchor(
        'G73').freq('M').to_series('COMM_FREIGHT')
    ret = Source(XLS, '1.12 ').anchor('G5').freq('M').to_series('RETAIL')


params = [
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

    ('2.1 ', 'B404', 'A', 'GOV_EXP_CONS_ACCUM'),
    ('2.1 ', 'BB11', 'A', 'GOV_INC_CONS_ACCUM'),

    ('3.5 ', 'B7', 'A', 'CPI'),
    ('3.5 ', 'C7', 'Q', 'CPI'),
    ('3.5 ', 'G7', 'M', 'CPI'),

    ('4.2 ', 'B7', 'A', 'WAGE'),
    ('4.2 ', 'C7', 'Q', 'WAGE'),
    ('4.2 ', 'G7', 'M', 'WAGE'),

]


df = DataframeDict()

for freq in FREQUENCIES: # reading via xlwings is slow!
    for p in params:
        v = Variable(*p)
        if v.freq == freq:
            ts = v.get(XLS)
            df[freq] = pd.concat([df[freq], ts], axis=1)
            print("Finished processing", v.varname, "at frequency", freq)
            try:
                df[freq].index = ts.index
            except ValueError:
                pass

cols = set()
for freq in 'AQM':
    cols.update(df[freq].columns)
cols = sorted(cols)
print('Total', len(cols), 'variables:', ', '.join(cols))

print('\nSaving CSV files:')
for freq in 'AQM':
    filename = output_csv_path(freq)
    df[freq].to_csv(filename)
    print(os.path.abspath(filename))
print('Done')


print('\nSaving Excel file:')
filename = df.save_xls()
print(os.path.abspath(filename))
print('Done')

# TODO: produce another xls
