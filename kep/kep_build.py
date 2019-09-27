import os
from pathlib import Path
from kep import download_and_unpack, create_dataframes

xlspath = download_and_unpack(2019, 7)

print('Processing 2019-07...')
df_dict = create_dataframes(xlspath)
cols = df_dict.all_variables()
print('Total', len(cols), 'variables:', ', '.join(cols))

output_dir = Path(__file__).parent.parent / 'output'

print('\nSaving CSV files:')
for filename in df_dict.save_csv(output_dir):
    print(os.path.abspath(filename))
print('Done')

print('\nSaving Excel file:')
filename = df_dict.save_xls(output_dir)
print(os.path.abspath(filename))
print('Done')
