![](carbon.png)

# kep-xls

Get [Rosstat macroeconomic time series](https://www.gks.ru/compendium/document/50802) 
as clean CSV files, importable by pandas or R.

### Data access

Use  `kep.download_dataframes()` to download the data from stable URLs.

```python 
from kep import download_dataframes

dfa, dfq, dfm = download_dataframes()
```

### Usage example

```python 
# Какой объем ВВП в России?
gdp = dfa.GDP_RUB['2018']
# 2018-12-31    103876.0

# Какая сейчас инфляция?
cpi = dfm.CPI.last('12M').divide(100).product().round(3) * 100
# 104.5

# А за какой месяц это данные?
dfm.CPI.last('1M').index[0].strftime('%Y-%m')
# '2019-07'
```

### Variables

10 variables currently produced: 
`COMM_FREIGHT`, `CPI`, `GDP_INDEX`, `GDP_RUB`, `GOV_EXP_CONS_ACCUM`, 
`GOV_INC_CONS_ACCUM`, `INVEST_INDEX`, `INVEST_RUB`, `RETAIL_SALES`, `WAGE`.

### Still want an Excel file?

We have it [here](https://github.com/mini-kep/kep-xls/blob/master/output/df.xlsx?raw=true).

### Build locally 


`kep_build.py` creates local CSV files. 

```
pip install -r requirements.txt
python kep-build.py
```