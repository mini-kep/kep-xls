from kep import download_dataframes

dfa, dfq, dfm = download_dataframes()

# Какой объем ВВП в России?
gdp = dfa.GDP_RUB['2018']
# 2018-12-31    103876.0

# Какая сейчас инфляция?
cpi = dfm.CPI.last('12M').divide(100).product().round(3) * 100
# 104.5

# А за какой месяц это данные?
dfm.CPI.last('1M').index[0].strftime('%Y-%m')
# '2019-07'