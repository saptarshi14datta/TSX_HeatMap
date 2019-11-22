from yahoofinancials import YahooFinancials
import pandas as pd
import numpy as np
from datetime import datetime

data = pd.read_csv(r'C:\Users\b0589940\Documents\Python Scripts\TSX_Heat_Map\TSX_Listed_Stocks.csv',encoding = "ISO-8859-1", dtype=str)
data['Date of TSX ListingYYYYMMDD'] = data['Date of TSX ListingYYYYMMDD'].replace(np.nan, '00000000')
data['Date of TSX ListingYYYYMMDD'] = data['Date of TSX ListingYYYYMMDD'] .astype(int)
data['Date of TSX ListingYYYYMMDD'] = pd.to_datetime(data['Date of TSX ListingYYYYMMDD'],format='%Y%m%d', errors='coerce')

data['Root Ticker'] = data['Root Ticker'] .astype(str)

unique_sectors = data['Sector'].unique()
unique_subsectors =  data['SubSector'].unique()

def get_yf_symbol(exchange, stock_ticker):
    if exchange == 'TSX':
        yf_ticker = stock_ticker+".TO"
    if exchange == 'TSXV':    
        yf_ticker = stock_ticker+".V"
    if exchange == 'NYSE' or exchange == 'NYSE':    
        yf_ticker = stock_ticker
    return yf_ticker

stock_ticker_yf = []
mkt_cap = []
cmp = []
low52 = []
high52 = []
divyield = []
SMA200 = []
SMA50 = []
Beta = []
current_vol = []
avg_10d_vol = []
avg_3m_vol = []
eps = []
r&d_ratio = []
payout_ratio = []
pe_ratio = []
price_sales_ratio = []
price_book_ratio = []
shares_outstanding = []

for index,row in data.iterrows():
    exchange = data.iloc[index]['Exchange']
    stock_ticker = data.iloc[index]['Root Ticker']
    stock_ticker_yf = get_yf_symbol(exchange,stock_ticker)
    #stock_ticker_yf.append(get_yf_symbol(exchange,stock_ticker))
    yahoo_financials = YahooFinancials(stock_ticker_yf)
    mkt_cap.append(yahoo_financials.get_market_cap())
    cmp.append(yahoo_financials.get_current_price())
    low52.append(yahoo_financials.get_yearly_low())
    high52.append(yahoo_financials.get_yearly_high())
    divyield.append(yahoo_financials.get_dividend_yield())
    SMA200.append(yahoo_financials.get_50day_moving_avg())
    SMA50.append(yahoo_financials.get_50day_moving_avg())
    Beta.append(yahoo_financials.get_beta())
    current_vol.append(yahoo_financials.get_current_volume())
    avg_10d_vol.append(yahoo_financials.get_ten_day_avg_daily_volume())
    avg_3m_vol.append(yahoo_financials.get_three_month_avg_daily_volume())
    eps.append(yahoo_financials.get_earnings_per_share())
    r&d_ratio.append(yahoo_financials.get_research_and_development()/yahoo_financials.get_total_revenue())
    payout_ratio.append(yahoo_financials.get_payout_ratio())
    pe_ratio.append(yahoo_financials.get_pe_ratio())
    price_sales_ratio.append(yahoo_financials.get_price_to_sales())
    price_book_ratio.append(yahoo_financials.get_book_value()/yahoo_financials.get_num_shares_outstanding(price_type='current'))
    shares_outstanding.append(yahoo_financials.get_num_shares_outstanding(price_type='current'))
    
    print("Data for", stock_ticker_yf, "received.")

#data['Ticker_YF'] = stock_ticker_yf 

'''
yahoo_financials = YahooFinancials(stock_ticker_yf)
cmp.append(yahoo_financials.get_current_price())
low52.append(yahoo_financials.get_yearly_low())
high52.append(yahoo_financials.get_yearly_high())
divyield.append(yahoo_financials.get_dividend_yield())
SMA200.append(yahoo_financials.get_50day_moving_avg())
SMA50.append(yahoo_financials.get_50day_moving_avg())
Beta.append(yahoo_financials.get_beta())
current_vol.append(yahoo_financials.get_current_volume())
avg_10d_vol.append(yahoo_financials.get_ten_day_avg_daily_volume())
avg_3m_vol.append(yahoo_financials.get_three_month_avg_daily_volume())
'''
data['Market Capitalization'] = mkt_cap
data['CMP'] = cmp
data['52W_Low'] = low52
data['52W_High'] = high52
data['Div Yield'] = divyield
data['SMA 200'] = SMA200
data['SMA 50'] = SMA50
data['Beta'] = Beta
data['Current Volume'] = current_vol
data['Avg 10 Day Volume'] = avg_10d_vol
data['Avg 3 Mon. Volume'] = avg_3m_vol
data['EPS'] = eps
data['R&D %'] = r&d_ratio
data['Div. Payout Ratio'] = payout_ratio
data['P/E'] = pe_ratio
data['P/S'] = price_sales_ratio
data['P/B'] = price_book_ratio
data['Shares Outstanding'] = shares_outstanding


export_data = data.to_csv (r'C:\Users\b0589940\Documents\Python Scripts\TSX_Heat_Map\TSX_data.csv', index = None, header=True) #Don't forget to add '.csv' at the end of the path
print("done")