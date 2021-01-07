import pandas as pd
import datetime
d_date = datetime.datetime.now()
reg_format_date = d_date.strftime("%Y-%m-%d")
print(reg_format_date)
intraday = pd.read_html('https://www.moneycontrol.com/stocks/marketstats/blockdeals/')
print(intraday[0])
intraday[0].to_excel(r'.\ '+reg_format_date+'_Intraday.xlsx')