from pathlib import Path
from kep import run, latest_date

output_dir = Path(__file__).parent / 'output'
year, month = latest_date()
run(year, month, output_dir)
