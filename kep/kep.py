from extract import ExcelFile, Variable
from cells import RANGES
from files import download_and_unpack
from reader import download_dataframes

def create_dataframes(xlpath):
    vs = [Variable(*r) for r in RANGES]
    return ExcelFile(xlpath).make_dataframe(vs)