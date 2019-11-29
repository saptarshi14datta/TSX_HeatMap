from yahoofinancials import YahooFinancials
import pandas as pd
import numpy as np
from datetime import datetime


country = input("Which stocks do you want to update:(India,US,Canada)?")

data = pd.read_excel(r'C:\Users\b0589940\Documents\Python Scripts\TSX_Heat_Map\Listed_Stocks.xlsx', sheet_name = country, encoding = "ISO-8859-1", dtype=str)
#data['Date of TSX ListingYYYYMMDD'] = data['Date of TSX ListingYYYYMMDD'].replace(np.nan, '00000000')
#data['Date of TSX ListingYYYYMMDD'] = data['Date of TSX ListingYYYYMMDD'] .astype(int)
#data['Date of TSX ListingYYYYMMDD'] = pd.to_datetime(data['Date of TSX ListingYYYYMMDD'],format='%Y%m%d', errors='coerce')

data['Root Ticker'] = data['Root Ticker'] .astype(str)

unique_sectors = data['Sector'].unique()
unique_subsectors =  data['SubSector'].unique()

def get_yf_symbol(exchange, stock_ticker):
    if exchange == 'TSX':
        yf_ticker = stock_ticker+".TO"
    if exchange == 'TSXV':    
        yf_ticker = stock_ticker+".V"
    if exchange == 'NYSE' or exchange == 'NASDAQ' or exchange == '':    
        yf_ticker = stock_ticker
    if exchange == 'NSE':    
        yf_ticker = stock_ticker+".NS"
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
rd_expense = []
total_revenue = []
payout_ratio = []
pe_ratio = []
price_sales_ratio = []
book_value = []
shares_outstanding = []

def get_mkt_cap():
    try:
        x = yahoo_financials.get_market_cap()
    except:
        x = 0
    return x

def get_cmp():
    try:
        x = yahoo_financials.get_current_price()
    except:
        x = 0
    return x

def get_low52():
    try:
        x = yahoo_financials.get_yearly_low()
    except:
        x = 0
    return x

def get_high52():
    try:
        x = yahoo_financials.get_yearly_high()
    except:
        x = 0
    return x

def get_divyield():
    try:
        x = yahoo_financials.get_dividend_yield()
    except:
        x = 0
    return x

def get_SMA200():
    try:
        x = yahoo_financials.get_50day_moving_avg()
    except:
        x = 0
    return x

def get_SMA50():
    try:
        x = yahoo_financials.get_50day_moving_avg()
    except:
        x = 0
    return x

def get_Beta():
    try:
        x = yahoo_financials.get_beta()
    except:
        x = 0
    return x

def get_current_vol():
    try:
        x = yahoo_financials.get_current_volume()
    except:
        x = 0
    return x

def get_10d_vol():
    try:
        x = yahoo_financials.get_ten_day_avg_daily_volume()
    except:
        x = 0
    return x

def get_3m_vol():
    try:
        x = yahoo_financials.get_three_month_avg_daily_volume()
    except:
        x = 0
    return x

def get_eps():
    try:
        x = yahoo_financials.get_earnings_per_share()
    except:
        x = 0
    return x

def rd_expenses():
    try:
        x = yahoo_financials.get_research_and_development()
    except:
        x = 0
    return x

def get_total_revenue():
    try:
        x = yahoo_financials.get_total_revenue()
    except:
        x = 0
    return x

def get_payout_ratio():
    try:
        x = yahoo_financials.get_payout_ratio()
    except:
        x = 0
    return x

def get_pe_ratio():
    try:
        x = yahoo_financials.get_pe_ratio()
    except:
        x = 0
    return x

def get_ps_ratio():
    try:
        x = yahoo_financials.get_price_to_sales()
    except:
        x = 0
    return x

def get_book_value():
    try:
        x = yahoo_financials.get_book_value()
    except:
        x = 0
    return x

def get_shares_outstanding():
    try:
        x = yahoo_financials.get_num_shares_outstanding(price_type='current')
    except:
        x = 0
    return x


for index,row in data.iterrows():
    exchange = data.iloc[index]['Exchange']
    stock_ticker = data.iloc[index]['Root Ticker']
    stock_ticker_yf = get_yf_symbol(exchange,stock_ticker)
    #stock_ticker_yf.append(get_yf_symbol(exchange,stock_ticker))
    yahoo_financials = YahooFinancials(stock_ticker_yf)
    mkt_cap.append(get_mkt_cap())
    cmp.append(get_cmp())
    low52.append(get_low52())
    high52.append(get_high52())
    divyield.append(get_divyield())
    SMA200.append(get_SMA200())
    SMA50.append(get_SMA50())
    Beta.append(get_Beta())
    current_vol.append(get_current_vol())
    avg_10d_vol.append(get_10d_vol())
    avg_3m_vol.append(get_3m_vol())
    eps.append(get_eps())
    rd_expense.append(rd_expenses())
    total_revenue.append(get_total_revenue())
    payout_ratio.append(get_payout_ratio())
    pe_ratio.append(get_pe_ratio())
    price_sales_ratio.append(get_ps_ratio())
    book_value.append(get_book_value())
    shares_outstanding.append(get_shares_outstanding())
    
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
data['R&D expense'] = rd_expense
data['Revenue'] = total_revenue
data['R&D %'] = data['R&D expense']/data['Revenue']
data['Div. Payout Ratio'] = payout_ratio
data['P/E'] = pe_ratio
data['P/S'] = price_sales_ratio
data['Book Value'] = book_value
data['Shares Outstanding'] = shares_outstanding
data['P/B'] = data['Book Value']/data['Shares Outstanding']



export_data = data.to_excel (r'C:\Users\b0589940\Documents\Python Scripts\TSX_Heat_Map\Stock_data.xlsx', sheet_name=country, index = None, header=True) #Don't forget to add '.csv' at the end of the path
print("done")