from files import output_xl_path, make_csvpath
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
