"""Download and unpack Excel file from Rosstat website"""

from pathlib import Path
import subprocess
import platform
import os
import wget
from dataclasses import dataclass

#URL = 'https://gks.ru/storage/mediabank/ind07(2).rar'
#RAR = os.path.join(SOURCE_FOLDER, 'kep_2019_07.rar')
#XLS = os.path.join(SOURCE_FOLDER, 'краткосрочные_июль.xlsx')


def default_folder(name) -> Path:
    home = Path.home() / '.kep' / name
    home.mkdir(exist_ok=True, parents = True)
    return home


def default_datasource_folder() -> Path:
    return default_folder('source') 


def default_output_folder() -> Path:
    return default_folder('output') 


@dataclass
class Folder:
    folder: Path


@dataclass
class OutputFiles:
    folder: Path = default_output_folder()
    
    def default(self):
       if not self.folder:
          self.folder = default_output_folder()
       return self    
    
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
  

XL_MAPPER = {(2019, 7): 'краткосрочные_июль.xlsx'}
URL_MAPPER = {(2019, 7): 'ind07(2).rar'}

@dataclass
class Namer:
    year: int
    month: int

    def xl(self):
        return XL_MAPPER[(self.year, self.month)] 
    
    def url(self):
        return URL_MAPPER[(self.year, self.month)]
    
    def get(self, mapper):
        key = (self.year, self.month)
        try:
            return mapper[key]
        except KeyError:    
            msg = {'Supported values': mapper.keys(),
                   'got': key}
            raise ValueError(msg)           


@dataclass
class SourceFiles:
    folder: Path = default_datasource_folder()
    
    def default(self):
       if not self.folder:
          self.folder = default_datasource_folder()
       return self  
    
    def _infolder(self, filename):
        return self.folder.joinpath(filename).__str__()
    
    def xlpath(self, year, month):
        filename = Filename(year, month).xl()
        return self._infolder(filename)

    def rarpath(self, year, month):
        filename = Filename(year, month).rar()
        return self._infolder(filename)
    
    def xlpath_original(self, year, month):
        filename = Namer(year, month).xl()
        return self._infolder(filename)
    
def make_url(year, month):
    filename = Namer(year, month).url()
    return f'https://gks.ru/storage/mediabank/{filename}'

def unrar_executable():
    if platform.system() == 'Windows':
        current_folder = Path(__file__).parent
        return str(current_folder.parent / 'bin' / 'UnRAR.exe')
    else:
        return 'unrar'

def unrar(args):
    args = [unrar_executable()] + args
    return subprocess.run(args, capture_output=True)


def unpack(filepath: str, destination_folder: str = '.'):
    # UnRAR wants its folder argument to end with '/'
    folder = "{}{}".format(destination_folder, os.sep)
    return unrar(['e', filepath, folder, '-y', '-o+'])


def list_content(rarpath: str):    
    output = unrar(['lb', rarpath])
    return output.stdout.decode('cp866').strip() # on Windows on cp866 only


def download_and_unpack(year, month, source_folder=None):    
    url = make_url(year, month)
    source = SourceFiles(source_folder) if source_folder else SourceFiles()         
    rarpath = source.rarpath(year, month)
    xlspath0 = source.xlpath_original(year, month)
    xlspath = source.xlpath(year, month)
    if not Path(rarpath).exists():
        wget.download(url, rarpath)
        print('Downloaded', url)
        print('Saved as', rarpath)
    if not Path(xlspath).exists():
        unpack(rarpath, str(source.folder))
        print('Unpacked', xlspath0)
        os.rename(xlspath0, xlspath)
        print('Renamed as', xlspath)
    else:
        print('Already exists:', xlspath)
    return xlspath
    # MAYBE: save Excel under different filename, it has not mention of year now
    
if __name__ == '__main__':    
    download_and_unpack(2019, 7)