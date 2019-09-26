from read import download_annual, download_monthly

dfa = download_annual()
dfa.GDP_RUB['2018']
# 2018-12-31    103876.0

dfm = download_monthly()
dfm.CPI.last('3M')

#2019-05-31    100.3
#2019-06-30    100.0
#2019-07-31    100.2
#Name: CPI, dtype: float64


