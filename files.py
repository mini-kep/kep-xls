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
