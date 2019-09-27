# kep-xls

Get Rosstat macroeconomic time series.

- Starting 2019 Rosstat publishes some macroeconomic time series as Excel files ([source](https://www.gks.ru/compendium/document/50802)) 
- `kep_build.py` saves time series as local CSV files importable by pandas or R. It also saves same data in an [Excel file](https://github.com/mini-kep/kep-xls/blob/master/output/df.xlsx?raw=true)
- With `kep.py` you can download data from stable URLs in this repo.

### Usage 

```python 
from kep import download_dataframes

dfa, dfq, dfm = download_dataframes()

# ВВП у нас какой вообще?
gdp = dfa.GDP_RUB['2018']
# 2018-12-31    103876.0

# A инфляция сколько?
cpi = dfm.CPI.last('12M').divide(100).product().round(3) * 100
# 104.5

# А на какой месяц это данные?
dfm.CPI.last('1M').index[0]
# Timestamp('2019-07-31 00:00:00')
```

### Build dataset locally

```
pip install -r requirements.txt
python kep_build.py
```

### Variables

10 variables currently produced: 
`COMM_FREIGHT`, `CPI`, `GDP_INDEX`, `GDP_RUB`, `GOV_EXP_CONS_ACCUM`, 
`GOV_INC_CONS_ACCUM`, `INVEST_INDEX`, `INVEST_RUB`, `RETAIL_SALES`, `WAGE`.
