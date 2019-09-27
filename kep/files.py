"""Download and unpack Excel file from Rosstat website"""

from pathlib import Path
import wget
from dataclasses import dataclass

from kep.settings import Namer
from kep.unrar import unpack


def default_folder(name) -> Path:
    home = Path.home() / '.kep' / name
    return home


def default_datasource_folder() -> Path:
    return default_folder('source')


def default_output_folder() -> Path:
    return default_folder('output')

@dataclass
class Files:
    folder: Path 
    default: Path

    def persist(self):
        if not self.folder:
            self.folder = self.default
        if not self.folder.exists():
            self.folder.mkdir(parents=True)
        return self
    

@dataclass
class OutputFiles(Files):
    folder: Path
    default: Path = default_output_folder()

    def csvpath(self, freq):
        filename = 'df{}.csv'.format(freq.lower())
        return self.folder.joinpath(filename)

    def xlpath(self):
        return self.folder.joinpath('df.xlsx')


@dataclass
class Filename:
    year: int
    month: int

    def _base(self, extension):
        month = str(self.month).zfill(2)
        return f'kep_{self.year}_{month}.{extension}'

    def rar(self):
        return self._base('rar')

    def xl(self):
        return self._base('xlsx')


assert Filename(2019, 7).rar() == 'kep_2019_07.rar'
assert Filename(2019, 7).xl() == 'kep_2019_07.xlsx'


@dataclass
class SourceFiles(Files):
    folder: Path
    default: Path = default_datasource_folder()

    def join(self, filename):
        return self.folder.joinpath(filename).__str__()

    def xlpath(self, year, month):
        filename = Filename(year, month).xl()
        return self.join(filename)

    def rarpath(self, year, month):
        filename = Filename(year, month).rar()
        return self.join(filename)

    def xlpath_original(self, year, month):
        filename = Namer(year, month).xl()
        return self.join(filename)


def make_url(year, month):
    return Namer(year, month).url()


def download_and_unpack(year, month, source_folder=None):
    url = make_url(year, month)
    source = SourceFiles(source_folder).persist()
    rarpath = source.rarpath(year, month)
    xlspath0 = source.xlpath_original(year, month)
    xlspath = source.xlpath(year, month)
    if not Path(rarpath).exists():
        wget.download(url, rarpath)
        print('Downloaded', url)
        print('  Saved as', rarpath)
    if not Path(xlspath).exists():
        unpack(rarpath, str(source.folder))
        print('  Unpacked', xlspath0)
        Path(xlspath0).rename(xlspath)
        print('Renamed to', xlspath)
    else:
        print('Already exists:', xlspath)
    return xlspath


if __name__ == '__main__':
    download_and_unpack(2019, 7)
