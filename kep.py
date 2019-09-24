import subprocess
from pathlib import Path
import re
import itertools
import pandas as pd
from dataclasses import dataclass

import wget
import xlwings as xw


URL = 'https://gks.ru/storage/mediabank/ind07(2).rar'
RAR = 'kep_2019_07.rar'
XLS = 'краткосрочные_июль.xlsx'


def list_content(path: str):
    args = ['unrar', 'lb', path]
    output = subprocess.run(args, capture_output=True)
    return output.stdout.decode('cp866').strip()


if not Path(RAR).exists():
    wget.download(URL, RAR)    
    subprocess.run(['unrar', 'e', RAR])
   

def uncomment_string(s: str):
    """https://regex101.com/r/LdAJBH/1"""
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
        except ValueError:
            break
        result.append(year)         
    return result
    
@dataclass
class Location:
    anchor: xw.Range
    years: [int]
    
    @property
    def width(self):
        return dict(Y=1, Q=4, M=12)[self.freq]
    
    def freq(self, freq):        
        self.freq = freq.upper()
        a = self.anchor
        b = a.offset(len(self.years), self.width-1)
        self.range = xw.Range(a, b)
        return self
    
    def values(self):
        return self.range.value
    
    def flat(self):
        if self.freq == 'Y':
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

gdp_q = Source(XLS, '1.1. ').anchor('C9').freq('Q').to_series('GDP')   
gdp_a = Source(XLS, '1.1. ').anchor('B9').freq('Y').to_series('GDP') 
frei_m = Source(XLS, '1.5. ').anchor('G73').freq('M').to_series('COMFREIGHT') 
ret    = Source(XLS, '1.12 ').anchor('G5').freq('M').to_series('RETAIL') 


#s = find_years(Source(XLS, '1.1. ', 'C9', 'Q').range().xl_sheet(), 'C9')
#
#params = [
#  ('1.1. ', 'C9',  'Q', 'GDP'), 
#  ('1.1. ', 'B9',  'Y', 'GDP'), 
#  ('1.5. ', 'G73', 'M', 'COMFREIGHT')   
#]
#
#df = pd.Dataframe()
#for p in params:
#    ts = Source(XLS, '1.1. ', 'C9', 'Q').range().to_series('GDPQ')  


    