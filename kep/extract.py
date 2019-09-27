from enum import Enum
from dataclasses import dataclass

import pandas as pd

from excel import load_workbook, to_series
from files import OutputFiles

class Freq(Enum):
    ANNUAL = 'A'
    QTR = 'Q'
    MONTH = 'M'

@dataclass
class Variable:
    sheet: str
    cell: str
    freq: str
    name: str


class ExcelFile:
    def __init__(self, path):
        self.wb = load_workbook(filename = path)
        
    def get_series(self, v: Variable):
        return to_series(self.wb, v)
    
    def all_series(self, vs):
        series = []
        for v in vs:            
            ts = self.get_series(v)
            print("Finished reading", v.name, "at frequency", v.freq)
            series.append((v, ts))
        return series
        
    
    def make_dataframe(self, vs):
        series = self.all_series(vs)
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

    def save_csv(self, folder=None):
        for freq in FREQUENCIES:
            filepath = OutputFiles(folder).default().csvpath(freq)            
            self[freq].to_csv(filepath)
            yield filepath

    def save_xls(self, folder=None):
        filepath = OutputFiles(folder).default().xlpath()
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