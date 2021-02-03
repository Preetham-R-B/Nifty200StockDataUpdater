import datetime

import openpyxl
from nsetools import Nse

file_path = "Nifty200_2_2_2020.xlsx"
file = openpyxl.load_workbook(file_path)

wb = file.active
nse = Nse()
row = 0
for i in wb['A']:
    row += 1
    data = nse.get_quote(i.value)
    if(data != None):
        index = 'B'+str(row)
        wb[index] = data.get('change')
        index = 'C'+str(row)
        wb[index] = data.get('dayHigh')
        index = 'D'+str(row)
        wb[index] = data.get('dayLow')
        index = 'E'+str(row)
        wb[index] = data.get('high52')
        index = 'F'+str(row)
        wb[index] = data.get('low52')
        index = 'G'+str(row)
        wb[index] = data.get('open')
        index = 'H'+str(row)
        wb[index] = data.get('closePrice')
        index = 'I'+str(row)
        wb[index] = data.get('totalBuyQuantity')
        index = 'J'+str(row)
        wb[index] = data.get('totalSellQuantity')
        index = 'K'+str(row)
        wb[index] = data.get('totalTradedValue')
        index = 'L'+str(row)
        wb[index] = data.get('totalTradedVolume')
        index = 'M'+str(row)
        wb[index] = datetime.datetime.now()
file.save(file_path)
