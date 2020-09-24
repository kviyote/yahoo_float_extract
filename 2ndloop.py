import os
import openpyxl
import requests
from bs4 import BeautifulSoup
import time
from random import randint

print(os.getcwd())


def getFloat(site):
    try:
        page = requests.get(site)

        page.raise_for_status()
        
        soup = BeautifulSoup(page.text, 'html.parser')

        float_val = soup.select('#Col1-0-KeyStatistics-Proxy > section > div.Mstart\(a\).Mend\(a\) > div.Fl\(end\).W\(50\%\).smartphone_W\(100\%\) > div > div:nth-child(2) > div > div > table > tbody > tr:nth-child(4) > td.Fw\(500\).Ta\(end\).Pstart\(10px\).Miw\(60px\)')
        
        
        return(float_val[0].text.strip())

    except requests.exceptions.HTTPError:
        print('HTTP Error \n sleep for 120 seconds \n')
        time.sleep(120)

        try:
            page = requests.get(site)

            page.raise_for_status()
            
            soup = BeautifulSoup(page.text, 'html.parser')

            float_val = soup.select('#Col1-0-KeyStatistics-Proxy > section > div.Mstart\(a\).Mend\(a\) > div.Fl\(end\).W\(50\%\).smartphone_W\(100\%\) > div > div:nth-child(2) > div > div > table > tbody > tr:nth-child(4) > td.Fw\(500\).Ta\(end\).Pstart\(10px\).Miw\(60px\)')

            
            return(float_val[0].text.strip())
        except IndexError:
            return('missing after HTTP Error')

    except IndexError:
        return('missing')


save_file_name ='test3.xlsx'
na_list = []
missing_list = []
wb = openpyxl.load_workbook('ticker_list_2.xlsx')
print(wb.sheetnames)

sheet = wb['Sheet2']

col = 'B'
row = str(21)


col_len=len(sheet[col])

for i in range(21,col_len + 1,1):
    url = 'https://finance.yahoo.com/quote/' + str(sheet['A' + str(i)].value) + '/key-statistics?p=' + str(sheet['A' + str(i)].value)
    if sheet[col + str(i)].value == 'N/A':
        try:
            print(str(sheet['A'+str(i)].value))
            na_tick = str(getFloat(url))
            print(na_tick)
            na_list.append(str(i)+ ' ' + na_tick)
            sheet['B' + str(i)] = na_tick
            print()
            time.sleep(randint(1,3))
            wb.save(save_file_name)
            time.sleep(1)
        except KeyboardInterrupt:
            break
            wb.save(save_file_name)
            wb.close()
    
    if sheet[col + str(i)].value == 'missing':
        try:
            print(str(sheet['A'+str(i)].value))
            tick_float = str(getFloat(url))
            print(tick_float)
            missing_list.append(str(i)+ ' ' + tick_float)
            sheet['B' + str(i)] = tick_float
            print()
            time.sleep(randint(1,3))
            wb.save(save_file_name)
            time.sleep(1)
        except KeyboardInterrupt:
            break
            wb.save(save_file_name)
            wb.close()

wb.save(save_file_name)
wb.close()    
    

        
