import re
import json
import csv
from io import StringIO
from bs4 import BeautifulSoup
import requests
from pandasgui import show
import pandas as pd
from pandas.io.json import json_normalize
from openpyxl import load_workbook 
from csv import reader



url_stats = 'https://finance.yahoo.com/quote/{}/key-statistics?p={}'
url_profile = 'https://finance.yahoo.com/quote/{}/profile?p={}'
url_financials = 'https://finance.yahoo.com/quote/{}/financials?p={}'
url_analysis = 'https://finance.yahoo.com/quote/{}/analysis?p={}'

# Stock ticker
stock  = 'INTU'

response = requests.get(url_financials.format(stock,stock))
soup = BeautifulSoup(response.text, 'html.parser')
pattern = re.compile(r'\s--\sData\s--\s')
script_data = soup.find('script', text=pattern).contents[0]

start = script_data.find('context')-2

json_data = json.loads(script_data[start:-12])

json_data['context']['dispatcher']['stores']['QuoteSummaryStore'].keys()

annual_is = json_data['context']['dispatcher']['stores']['QuoteSummaryStore']['incomeStatementHistory']['incomeStatementHistory']
quarterly_is = json_data['context']['dispatcher']['stores']['QuoteSummaryStore']['incomeStatementHistoryQuarterly']['incomeStatementHistory']

annual_cf = json_data['context']['dispatcher']['stores']['QuoteSummaryStore']['cashflowStatementHistory']['cashflowStatements']
quarterly_cf = json_data['context']['dispatcher']['stores']['QuoteSummaryStore']['cashflowStatementHistoryQuarterly']['cashflowStatements']


annual_bs = json_data['context']['dispatcher']['stores']['QuoteSummaryStore']['balanceSheetHistory']['balanceSheetStatements']
quarterly_bs = json_data['context']['dispatcher']['stores']['QuoteSummaryStore']['balanceSheetHistoryQuarterly']['balanceSheetStatements']


annual_is_stmts=[]

# Income Statements
for s in annual_is:
    statement = {}
    for key, val in s.items():
        try:
            statement[key] = val['raw']
        except TypeError:
            continue
        except KeyError:
            continue
    annual_is_stmts.append(statement)

   
# Cash Flows
annual_cf_stmts = []
quarterly_cf_stmts = []
  
for s in annual_cf:
    statement = {}
    for key, val in s.items():
        try:
            statement[key] = val['raw']
        except TypeError:
            continue
        except KeyError:
            continue
    annual_cf_stmts.append(statement)

for s in quarterly_cf:
    statement = {}
    for key, val in s.items():
        try:
            statement[key] = val['raw']
        except TypeError:
            continue
        except KeyError:
            continue
    quarterly_cf_stmts.append(statement)        
  
# Export Quarter Financials    
user_dict = quarterly_is
df = pd.json_normalize(user_dict)
df_tr1 = df.transpose()


user_dict = quarterly_cf
df = pd.json_normalize(user_dict)
df_tr2 = df.transpose()


user_dict = quarterly_bs
df = pd.json_normalize(user_dict)
df_tr3 = df.transpose()


# Export Annual Financials
user_dict = annual_is_stmts
df = pd.json_normalize(user_dict)
df_tr4 = df.transpose()

   
user_dict = annual_cf_stmts
df = pd.json_normalize(user_dict)
df_tr5 = df.transpose()


user_dict = annual_bs
df = pd.json_normalize(user_dict)
df_tr6 = df.transpose()

         
#---Profile----          
response = requests.get(url_profile.format(stock,stock))
soup = BeautifulSoup(response.text, 'html.parser')
pattern = re.compile(r'\s--\sData\s--\s')
script_data = soup.find('script', text=pattern).contents[0]
script_data[:500]
script_data[-500:]
start = script_data.find('context')-2
json_data = json.loads(script_data[start:-12])         
            

# Export Business Description
bus_des = json_data['context']['dispatcher']['stores']['QuoteSummaryStore']
#df = pd.DataFrame()
#df.at[0, '0'] = bus_des
df = pd.json_normalize(bus_des)
df_tr7 = df.transpose()


# Export SEC Filings
try:
    user_dict = json_data['context']['dispatcher']['stores']['QuoteSummaryStore']['secFilings']['filings'][:3]
    df = pd.json_normalize(user_dict)
    df_tr8 = df.transpose() 
except:
    print('No SEC filings found')

# ---Statistics---           
response = requests.get(url_stats.format(stock,stock))
soup = BeautifulSoup(response.text, 'html.parser')
pattern = re.compile(r'\s--\sData\s--\s')
script_data = soup.find('script', text=pattern).contents[0]
script_data[:500]
script_data[-500:]
start = script_data.find('context')-2
json_data = json.loads(script_data[start:-12])             
   
         
user_dict = json_data['context']['dispatcher']['stores']['QuoteSummaryStore']['defaultKeyStatistics']
df = pd.json_normalize(user_dict)
df_tr9 = df.transpose()
       

user_dict = json_data['context']['dispatcher']['stores']['QuoteSummaryStore']['financialData']
df = pd.json_normalize(user_dict)
df_tr10 = df.transpose()
       

# ---EPS---
response = requests.get(url_analysis.format(stock,stock))
soup = BeautifulSoup(response.text, 'html.parser')
pattern = re.compile(r'\s--\sData\s--\s')
script_data = soup.find('script', text=pattern).contents[0]
script_data[:500]
script_data[-500:]
start = script_data.find('context')-2
json_data = json.loads(script_data[start:-12])             
            

user_dict = json_data['context']['dispatcher']['stores']['QuoteSummaryStore']['earningsHistory']
eps_hist = user_dict['history']
df = pd.json_normalize(eps_hist)
df_tr11 = df.transpose()


user_dict = json_data['context']['dispatcher']['stores']['QuoteSummaryStore']['earningsTrend']
eps_trend = user_dict['trend']
df = pd.json_normalize(eps_trend)
df_tr12 = df.transpose()


# ---Historical Stock Data---

stock_url = 'https://query1.finance.yahoo.com/v7/finance/download/{}?'

params = {
    'range': '5y',
    'interval': '1d',
    'events': 'history'
}     

response = requests.get(stock_url.format(stock), params=params)       

file = StringIO(response.text)       
reader = csv.reader(file)
price_data = list(reader)
df_tr13 = pd.DataFrame(price_data)
#for row in price_data[:5]:
   # print(row)

path() # careate file path here
book = load_workbook(path) # file 
writer = pd.ExcelWriter(path, engine='openpyxl')
writer.book = book
writer.sheets = {ws.title: ws for ws in book.worksheets}

for sheetname in writer.sheets:
    df_tr1.to_excel(writer,sheet_name='Qtr P&L',startrow=1, index = True,header= False)
    df_tr2.to_excel(writer,sheet_name='Qtr CF',startrow=1, index = True,header= False)
    df_tr3.to_excel(writer,sheet_name='Qtr BS',startrow=1, index = True,header= False)
    df_tr4.to_excel(writer,sheet_name='Annual P&L',startrow=1, index = True,header= False)
    df_tr5.to_excel(writer,sheet_name='Annual CF',startrow=1, index = True,header= False)
    df_tr6.to_excel(writer,sheet_name='Annual BS',startrow=1, index = True,header= False)
    df_tr7.to_excel(writer,sheet_name='Bus Des',startrow=1, index = True,header= False)
    try: 
        df_tr8.to_excel(writer,sheet_name='SEC',startrow=1, index = True,header= False)
    except:
        print('No SEC filing')
    df_tr9.to_excel(writer,sheet_name='Key Stats',startrow=1, index = True,header= False)
    df_tr10.to_excel(writer,sheet_name='Fin Data',startrow=1, index = True,header= False)
    df_tr11.to_excel(writer,sheet_name='EPS Hist',startrow=1, index = True,header= False)
    df_tr12.to_excel(writer,sheet_name='EPS Est',startrow=1, index = True,header= False)
    df_tr13.to_excel(writer,sheet_name='Price Data',startrow=1, index = True,header= False)

writer.save()  

            
         
            
         
            
         
            
         
            
            
