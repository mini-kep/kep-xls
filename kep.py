import subprocess
from pathlib import Path
import re
import itertools

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
    return re.search(pattern=pat, string=s).group(1)   

assert uncomment_string('244873)') == '24487'
    
def uncomment(x):    
    if isinstance(x, str):
        x = float(uncomment_string(x))
    return x

def extract_values(arr, start_year=1999):
    result = []
    years = iter(range(start_year, 2020))
    quarters = itertools.cycle([1, 2, 3, 4])
    for xs in arr:
        year = next(years)
        for x in xs:
            qtr = next(quarters)
            if x:
                result.append((year, qtr, uncomment(x)))
    return result                

    
arr = xw.Book(XLS).sheets['1.1. '].range('C9:F29').value
vs = extract_values(arr, start_year=1999)               
assert vs[-1] == (2019, 1, 24487.0)
assert vs[0] == (1999, 1, 901.0)
