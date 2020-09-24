import requests
from selenium import webdriver
from bs4 import BeautifulSoup
import time
import random
import openpyxl
import os
from urllib.error import HTTPError

print()

userAgent = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/84.0.4147.135 Safari/537.36 Edg/84.0.522.63'}

float_list = []
na_list = []
list_start = 21
j=0

def getFloat(site):
    try:
        page = requests.get(site, headers = userAgent, timeout=10)

        page.raise_for_status()
        
        soup = BeautifulSoup(page.text, 'html.parser')

        float_val = soup.select('#Col1-0-KeyStatistics-Proxy > section > div.Mstart\(a\).Mend\(a\) > div.Fl\(end\).W\(50\%\).smartphone_W\(100\%\) > div > div:nth-child(2) > div > div > table > tbody > tr:nth-child(4) > td.Fw\(500\).Ta\(end\).Pstart\(10px\).Miw\(60px\)')

        
        return(float_val[0].text.strip())

    except requests.exceptions.HTTPError:
        print('HTTP Error \n sleep for 60 seconds \n')
        time.sleep(60)
        
        page = requests.get(site, headers = userAgent, timeout=10)

        page.raise_for_status()
        
        soup = BeautifulSoup(page.text, 'html.parser')

        float_val = soup.select('#Col1-0-KeyStatistics-Proxy > section > div.Mstart\(a\).Mend\(a\) > div.Fl\(end\).W\(50\%\).smartphone_W\(100\%\) > div > div:nth-child(2) > div > div > table > tbody > tr:nth-child(4) > td.Fw\(500\).Ta\(end\).Pstart\(10px\).Miw\(60px\)')

        
        return(float_val[0].text.strip())
        

    
        


wb = openpyxl.load_workbook('ticker list.xlsx')
print(wb.sheetnames)
print()

s2 = wb['Sheet2']
column = s2['A']


print(len(column))
print()

column_list = [column[x].value for x in range(list_start-1,len(column))]
print(column_list)
print()

for i in column_list:
    url = 'https://finance.yahoo.com/quote/' + str(i) + '/key-statistics?p=' + str(i)
    try:
        float_list.append(getFloat(url))

        
        s2['B'+str(list_start)] = float_list[j]
        print(i)
        print(getFloat(url))
        print()

        
        
    except IndexError:
        float_list.append('missing')
        s2['B'+str(list_start)] = 'missing'
        print(i)
        print('missing')
        print()
        
    j = j+1
    list_start = list_start+1
    wb.save('test2.xlsx')
    
print(float_list)
print(len(float_list))
print(na_list)



