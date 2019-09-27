import subprocess
import platform
from pathlib import Path
import os

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


def list_content(rarpath: str):
    output = unrar(['lb', rarpath])
    return output.stdout.decode('cp866').strip()  # on Windows on cp866 only
