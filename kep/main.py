import os
from kep.output import create_dataframes
from kep.settings import get_variable_definitions
from kep.files import download_and_unpack


def run_silent(year, month, output_dir=None, vs=get_variable_definitions()):
    xlspath = download_and_unpack(year, month)
    df_dict = create_dataframes(xlspath, vs)
    df_dict.save_csv(output_dir)
    df_dict.save_xls(output_dir)


def run(year, month, output_dir=None, vs=get_variable_definitions()):
    print('Downloading...')

    xlspath = download_and_unpack(year, month)

    print(f'\nProcessing {year}-{month}...')

    df_dict = create_dataframes(xlspath, vs)
    cols = df_dict.all_variables()

    print('Total', len(cols), 'variables:', ', '.join(cols))
    print('Done')
    print('\nSaving CSV files:')

    for filename in df_dict.save_csv(output_dir):
        print(os.path.abspath(filename))

    print('Done')
    print('\nSaving Excel file:')

    filename = df_dict.save_xls(output_dir)

    print(os.path.abspath(filename))
    print('Done')
