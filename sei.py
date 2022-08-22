# coding=utf-8
from datetime import datetime
from selenium import webdriver
from urllib.request import urlopen
from bs4 import BeautifulSoup
import sys, os, traceback
import time
from datetime import datetime, timedelta
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
import pandas as pd
import re

def seibro():
    today = datetime.today()
    today = today.strftime('%Y%m%d')
    df=pd.read_excel("C:\\abcp\\종목검색(전체).xlsx",sheet_name='종목검색(전체)',engine='openpyxl')
    print(df)
    # print(df.iloc[0+1, 1])
    # all_data = []
    all_data_frame=[]
    for i in range(0, 9):
        chrome_options = webdriver.ChromeOptions()
        chrome_options.add_argument('headless')

        driver = webdriver.Chrome('C:\\abcp\chromedriver_win32\chromedriver.exe', chrome_options=chrome_options)
        driver.implicitly_wait(30)

        url = 'http://www.seibro.or.kr/websquare/control.jsp?w2xPath=/IPORTAL/user/moneyMarke/BIP_CNTS04008P.xml&ISIN=' + str(df.iloc[i, 1]) + '&ABCP_PLAN_NO=' + str(df.iloc[i, 0])
        print(url)
        driver.get(url)
        time.sleep(5)

        #발행회사번호
        AA1= driver.find_element_by_xpath('//*[@id="AA1"]')
        #최근플랜번호
        AA2 = driver.find_element_by_xpath('//*[@id="AA2"]')
        # 플랜기간시작일
        AA3=driver.find_element_by_xpath('//*[@id="AA3"]')
        AA3 = re.sub('/', '', AA3.text, 0).strip()

        # 기초자산종류
        AA4 = driver.find_element_by_xpath('//*[@id="AA4"]')
        # 기초자산원보유자
        AA5 = driver.find_element_by_xpath('//*[@id="AA5"]')

        # 자산관리자
        AA6 = driver.find_element_by_xpath('//*[@id="AA6"]')
        # PF여부
        AA7 = driver.find_element_by_xpath('//*[@id="AA7"]')
        # 신용보강기관
        AA8 = driver.find_element_by_xpath('//*[@id="AA8"]')
        # 신용보강기간시작일
        AA9 = driver.find_element_by_xpath('//*[@id="AA9"]')
        AA9 = re.sub('/', '', AA9.text, 0).strip()
        # 신용보강종류
        AA10 = driver.find_element_by_xpath('//*[@id="AA10"]')


        # 발행자
        BB1 = driver.find_element_by_xpath('//*[@id="BB1"]')
        # 플랜명
        BB2 = driver.find_element_by_xpath('//*[@id="BB2"]')
        # 플랜기간종료일
        BB3 = driver.find_element_by_xpath('//*[@id="BB3"]')
        BB3 = re.sub('/', '', BB3.text, 0).strip()
        # 기초자산금액
        BB4 = driver.find_element_by_xpath('//*[@id="BB4"]')
        BB4 = re.sub(',', '', BB4.text, 0).strip()

        # 주관사
        BB5 = driver.find_element_by_xpath('//*[@id="BB5"]')
        # 파생상품연계
        BB6 = driver.find_element_by_xpath('//*[@id="BB6"]')
        # 시공사
        BB7 = driver.find_element_by_xpath('//*[@id="BB7"]')

        # 신용보강금액
        BB8 = driver.find_element_by_xpath('//*[@id="BB8"]')
        BB8 = re.sub(',', '', BB8.text, 0).strip()
        # 신용보강기간종료일
        BB9 = driver.find_element_by_xpath('//*[@id="BB9"]')
        BB9 = re.sub('/', '', BB9.text, 0).strip()

        data = {'발행회사번호': [AA1.text],
                '발행자': [BB1.text],
                '최근플랜번호': [AA2.text],
                '플랜명': [BB2.text],
                '플랜기간시작일': [AA3],
                '플랜기간종료일': [BB3],
                '기초자산종류': [AA4.text],
                '기초자산금액': [BB4],
                '기초자산원보유자': [AA5.text],
                '주관사': [BB5.text],
                '자산관리자': [AA6.text],
                '파생상품연계': [BB6.text],
                'PF여부':  [AA7.text],
                '시공사': [BB7.text],
                '신용보강기관':  [AA8.text],
                '신용보강금액': [BB8],
                '신용보강기간시작일': [AA9],
                '신용보강기간종료일': [BB9],
                '신용보강종류': [AA10.text],
                '예탁원발행기관코드':'',
                '입력일자':today
                }

        df1 = pd.DataFrame(data,
                          columns=['발행회사번호', '발행자', '최근플랜번호', '플랜명', '플랜기간시작일', '플랜기간종료일', '기초자산종류', '기초자산금액',
                                   '기초자산원보유자','주관사', '자산관리자', '파생상품연계','PF여부','시공사', '신용보강기관', '신용보강금액','신용보강기간시작일', '신용보강기간종료일', '신용보강종류','예탁원발행기관코드','입력일자'])
        #print(df1)
        all_data_frame.append(df1)
        #print(all_data_frame)
    #
        df_concat = pd.concat(all_data_frame, axis=0, ignore_index=False)
        print(df_concat)
        # 엑셀파일로 저장하기
        writer = pd.ExcelWriter('seibroabcp.xls')
        df_concat.to_excel(writer, sheet_name='Sheet1', index=False, header=False, na_rep=' ',
                       encoding='utf-16')  # 엑셀로 저장
        writer.save()


try :
    seibro()


except:
    #에러가 발생한 경우 StackTrace를 파일로 기록한다.
    outputFile = open('error.txt', 'w')
    traceback.print_exc(file=outputFile)
    outputFile.close()
