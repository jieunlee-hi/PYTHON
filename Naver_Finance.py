# -*- coding:utf-8 -*-
import re
from datetime import datetime, timedelta
import time
import traceback
import pandas as pd
import sys, os, traceback
from selenium.webdriver.common.keys import Keys
import os
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
import sys, os, traceback, glob
import win32com.client as win32
import sys
import os.path
import pandas as pd
import urllib.request
import urllib.parse
from bs4 import BeautifulSoup
from urllib.request import urlopen

import re
from datetime import datetime,timedelta
import time
import traceback

from selenium import webdriver

options = webdriver.ChromeOptions()
options.add_experimental_option("excludeSwitches", ["enable-logging"])
driver = webdriver.Chrome('C:\\NAVER_FINANCE\chromedriver_win32\chromedriver.exe',options=options)
    
# 암묵적으로 웹 자원 로드를 위해 최대 60초까지 기다려 준다.
driver.implicitly_wait(60)
# NAVER_FINANCE 사이트 주소
driver.get(
        'https://finance.naver.com/sise/sise_group.nhn?type=upjong')
html = driver.page_source
soup = BeautifulSoup(html, 'html.parser')

all_data_frame=[]
for td in soup('td'):  # td 안의
    for a in td('a'):  # a 태그 중에서
        name=a.get_text()  #업종명 데이터 추출
        print(name)
        
        p = re.search('/sise/.*', a['href']).group(0)  # 필요부분추출후 그룹핑
        #업종 url
        url2='https://finance.naver.com'+p
        print(url2)
        driver.get(url2)

        search = driver.find_element_by_xpath('//*[@id="contentarea_left"]/table[1]/tbody/tr[4]/td[3]')
        #소속종목수 데이터 추출
        time.sleep(2)
        print(search.text)

        #소속된 종목수
        count = int(search.text)
        for i in range(1, count+1):
            start = driver.find_element_by_xpath(
                '//*[@id="contentarea"]/div[4]/table/tbody/tr[' + str(i) + ']/td[1]/div/a')
            # 종목명
            print(start.text)
            cod = driver.find_element_by_xpath(
                '//*[@id="contentarea"]/div[4]/table/tbody/tr[' + str(i) + ']/td[1]/div/a').get_attribute("href")


            try:
                webpage = urlopen(cod)
                source = BeautifulSoup(webpage, 'html5lib')
                reviews = source.find_all('td')

                ROE=reviews[272].get_text()
                PER=reviews[277].get_text()
                PBR=reviews[282].get_text()
                Dividend_rate=reviews[299].get_text().strip()
                Dividend_rate=float(re.findall('\d+.\d+', Dividend_rate)[0])
                AVG = reviews[300].get_text().strip()
                AVG = float(re.findall('\d+.\d+', AVG)[0])
                print(ROE,PER,PBR,Dividend_rate,AVG)

                data = {'업종명': [name],
                    '소속종목수': [search.text],
                    '소속종목명': [start.text],
                    'ROE': [ROE],
                    'PER':[PER],
                    'PBR':[PBR],
                    'Dividend_rate':[Dividend_rate],
                    'AVG':[AVG],
                    'url':[cod]
                    }
                df = pd.DataFrame(data, columns=["업종명", "소속종목수", "소속종목명", "ROE","PER","PBR","Dividend_rate","AVG","url"])
                #기타 제외
                regex1 = re.compile(r".*기타.*")
                matchobj1 = regex1.finditer(str(name))
                for r1 in matchobj1:
                    match1 = r1.group(0)
                    if (bool(match1) == True):
                        df = pd.DataFrame(None)
                #print(df)
                # #데이터프레임 이어붙이기
                all_data_frame.append(df)
                df_concat = pd.concat(all_data_frame, axis=0, ignore_index=False)
                print(df_concat)
                # 엑셀파일로 저장하기
                writer = pd.ExcelWriter('naver_finance_22z08.xlsx')
                df_concat.to_excel(writer, sheet_name='Sheet1', index=False, header=True, na_rep=' ',
                               encoding='utf-8')  # 엑셀로 저장
                writer.save()
            except:
                pass
                if int(i) > int(count+1):
                    break





            #
            #      PER = driver.find_element_by_xpath('//*[@id="content"]/div[5]/table/tbody/tr[13]/td[1]')
            # #     # PBR =driver.find_element_by_xpath('//*[@id="content"]/div[5]/table/tbody/tr[14]/td[1]')
            # #     # time.sleep(2)
            # #     # #PER PBR
            # #     # print(PER.text)
            # #     # print(PBR.text)


            

            
           
    
