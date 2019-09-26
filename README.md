# kep-xls
Import macroeconomic time series from Excel files published by Rosstat.

- From 2019 Rosstat publishes several macroeconomic time series as Excel files ([source](https://www.gks.ru/compendium/document/50802)) 
- `kep.py` is hardcoded to extract latest data release as of September 2018
- `kep.py` saves time series as local CSV files importable by pandas or R, it also saves an Excel file
- 10 variables currently produced: COMM_FREIGHT, CPI, GDP_INDEX, GDP_RUB, GOV_EXP_CONS_ACCUM, 
GOV_INC_CONS_ACCUM, INVEST_INDEX, INVEST_RUB, RETAIL_SALES, WAGE.

### Usage 

```python 
from read import download_annual, download_quarterly, download_monthly

dfa = download_annual()
dfa.GDP_RUB['2018']
# 2018-12-31    103876.0

dfm = download_monthly()
dfm.CPI.last('3M')

#2019-05-31    100.3
#2019-06-30    100.0
#2019-07-31    100.2
#Name: CPI, dtype: float64
```

### Build dataset locally

```
pip install -r requirements.txt
python kep.py
```
