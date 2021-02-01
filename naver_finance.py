# -*- coding:utf-8 -*-
import re
from datetime import datetime, timedelta
import pandas as pd
import sys, os, traceback
from selenium.webdriver.common.keys import Keys
import sys, os, traceback, glob
import win32com.client as win32
import urllib.request
import urllib.parse
from bs4 import BeautifulSoup


def FINANCE():

    driver = webdriver.Chrome('C:\\NAVER_FINANCE\chromedriver_win32\chromedriver.exe')
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
            try :
                p = re.search('/sise/.*', a['href']).group(0)  # 필요부분추출후 그룹핑

                url2='https://finance.naver.com'+p
                print(url2)
                driver.get(url2)
                html2 = driver.page_source
                soup2=BeautifulSoup(html2, 'html.parser')
                search = driver.find_element_by_xpath('//*[@id="contentarea_left"]/table[1]/tbody/tr[4]/td[3]')
                #소속종목수 데이터 추출

                time.sleep(2)
                print(search.text)




                count = int(search.text)
                for i in range(3, count + 3):
                    start = driver.find_element_by_xpath(
                        '//*[@id="contentarea"]/div[4]/table/tbody/tr[' + str(i) + ']/td[1]/a')

                    # for tr1 in soup2.find_all('tr['+str(i)+']'):
                    #     for td1 in tr1('td'):
                    #         for a1 in td1('a'):
                    #             p1 = re.search('/item/main.nhn.*', a1['href']).group(0)
                    #             print(p1)
                    cod = driver.find_element_by_xpath(
                        '//*[@id="contentarea"]/div[4]/table/tbody/tr[' + str(i) + ']/td[1]/a').get_attribute("href")
                    p1 = re.search('https://finance.naver.com/item/main.nhn\Wcode=.*', cod).group(0)
                    print(p1)
                    parse = re.sub('https://finance.naver.com/item/main.nhn\Wcode=', '', p1, 0).strip()

                    data = {'업종명': [name],
                    '소속종목수': [search.text],
                    '소속종목명': [start.text],
                            '단축코드':[parse]
                           }
                    # 소속된 종목명데이터 추출



                    df = pd.DataFrame(data,columns=["업종명", "소속종목수", "소속종목명","단축코드"])

                    #print(df)
                    # #데이터프레임 이어붙이기
                    all_data_frame.append(df)
                    df_concat = pd.concat(all_data_frame, axis=0, ignore_index=False)
                    print(df_concat)
                    # 엑셀파일로 저장하기
                    writer = pd.ExcelWriter('naver_finance_0813.xls')
                    df_concat.to_excel(writer, sheet_name='Sheet1', index=False, header=True, na_rep=' ',
                                                       encoding='utf-16')  # 엑셀로 저장
                    writer.save()


            except:
                pass


try:
    FINANCE()


# 에러가 발생한 경우 StackTrace를 파일로 기록한다.
except:
    outputFile = open('error.txt', 'w')
    traceback.print_exc(file=outputFile)
    outputFile.close()
