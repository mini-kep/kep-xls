from kep import download_dataframes

dfa, dfq, dfm = download_dataframes()

# Какой в России ВВП?
gdp = dfa.GDP_RUB['2018']
# 2018-12-31    103876.0

# Какая сейчас инфляция?
cpi = dfm.CPI.last('12M').divide(100).product().round(3) * 100
# 104.5

# А на какой месяц это данные?
dfm.CPI.last('1M').index[0]
# Timestamp('2019-07-31 00:00:00')