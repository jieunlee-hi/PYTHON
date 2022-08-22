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
import numpy as np

#데이터 수집
def KOFIA_DATA():
    driver = webdriver.Chrome('C:\KOFIA_EMERGENCY\chromedriver_win32\chromedriver.exe')
    # 암묵적으로 웹 자원 로드를 위해 최대 60초까지 기다려 준다.
    driver.implicitly_wait(60)
    # KOFIA 사이트 주소
    driver.get(
        'http://www.kofiabond.or.kr/websquare/websquare.html?w2xPath=/xml/startest/BISBndSrtPrcDay.xml&divisionId=MBIS01070010000000&divisionNm=%25EC%259D%25BC%25EC%259E%2590%25EB%25B3%2584&tabIdx=1&w2xHome=/xml/&w2xDocumentRoot=')


    time.sleep(10)

    # 가장최근의 파일 다운로드
    signInBtn = driver.find_element_by_xpath('//*[@id="grpExcel"]').send_keys(Keys.ENTER)
    time.sleep(20)
        
    driver.close()



#데이터확장자 변환(xls->xlsx)
def CONVERSION():
    df=pd.read_excel("C:\\Users\\user\\Downloads\\채권시가평가기준수익률.xls",sheet_name='sheet')
    print(df)
    #df.to_excel('C:\\Users\\user\\Downloads\\채권시가평가기준수익률.xlsx',index=True)
    df.to_excel('C:\\Users\\user\\Downloads\\채권시가평가기준수익률.xlsx',sheet_name='sheet', index=False, header=True)

#KOFIA 서식변경
def KOFIA():
    #다운받은 KOFIA 데이터를 불러온다
    df=pd.read_excel("C:\\Users\\user\\Downloads\\채권시가평가기준수익률.xlsx",sheet_name='sheet1')

    #시가평가종류부분
    NAME=df.iloc[0:42,0:3]
    #수익률부분
    NUM=df.iloc[0:42,4:20]
    NUM = np.array(NUM)
    NUM_T=NUM.reshape((672,1))
    #print(NAME)

    #수익률
    RATE_DATA = []
    for i in NUM_T:
        for j in i:
            j=str(j).replace(r"-", "0")
            j=float(j)
            RATE= pd.DataFrame({'수익율' :[j]})

        RATE_DATA.append(RATE)
        EARNING_RATE = pd.concat(RATE_DATA, axis=0, ignore_index=True)
    #print(EARNING_RATE)


    GRP_CODE_DATA = []
    REMAIN_TERM_DATA = []
    remain_num = ['3', '6', '9', '100', '106', '200', '206', '300', '400', '500', '700', '1000','1500',
                  '2000','3000', '5000']
    for i in range(0, 42):
        # 시가평가그룹코드
        code = str(NAME['종류'][i] + str(NAME['종류명'][i]))
        code1 = str(NAME['신용등급'][i])
        namecode = re.sub('\W', '', code)
        namecode1 = re.sub('-', 'z', code1)
        # print(parse)
        name = namecode + namecode1
        parse1 = re.sub('국채국고채권양곡,외평,재정', '1010000', name)
        parse2 = re.sub('국채제2종국민주택채권z', '1020000', parse1)
        parse3 = re.sub('국채제1종국민주택채권기타국채', '1030000', parse2)
        parse4 = re.sub('지방채서울도시철도공채증권z', '2010000', parse3)
        parse5 = re.sub('지방채지역개발공채증권기타지방채', '2020000', parse4)
        parse6 = re.sub('특수채공사채및공단채정부보증채', '3070000', parse5)
        parse7 = re.sub('특수채공사채및공단채AAA', '3030110', parse6)
        parse8 = re.sub('특수채공사채및공단채AA\W', '3040121', parse7)
        parse9 = re.sub('특수채공사채및공단채AA', '3040120', parse8)
        parse10 = re.sub('특수채한국주택금융공사유동화증권MBS', '3060000', parse9)
        parse11 = re.sub('통안증권z', '4000000', parse10)
        parse12 = re.sub('금융채I은행채무보증AAA\W산금채\W', '5010110', parse11)
        parse13 = re.sub('금융채I은행채무보증AAA\W중금채\W', '5020110', parse12)
        parse14 = re.sub('금융채I은행채무보증AAA', '5030110', parse13)
        parse15 = re.sub('금융채I은행채무보증AA', '5040120', parse14)
        parse16 = re.sub('금융채I은행채무보증A\W', '5050131', parse15)
        parse17 = re.sub('금융채II금융기관채무보증AA\W', '6010121', parse16)
        parse18 = re.sub('금융채II금융기관채무보증AA0', '6010122', parse17)
        parse19 = re.sub('금융채II금융기관채무보증AAz', '6010123', parse18)
        parse20 = re.sub('금융채II금융기관채무보증A\W', '6010131', parse19)
        parse21 = re.sub('금융채II금융기관채무보증A0', '6010132', parse20)
        parse22 = re.sub('금융채II금융기관채무보증Az', '6010133', parse21)
        parse23 = re.sub('금융채II금융기관채무보증BBB', '6010210', parse22)
        parse24 = re.sub('회사채I공모사채보증특수은행,우량시중은행', '7020110', parse23)
        parse25 = re.sub('회사채I공모사채보증시중은행', '7020120', parse24)
        parse26 = re.sub('회사채I공모사채보증우량지방은행', '7020130', parse25)
        parse27 = re.sub('회사채I공모사채보증기타금융기관', '7020210', parse26)
        parse28 = re.sub('회사채I공모사채무보증AAA', '7010110', parse27)
        parse29 = re.sub('회사채I공모사채무보증AA\W', '7010121', parse28)
        parse30 = re.sub('회사채I공모사채무보증AA0', '7010122', parse29)
        parse31 = re.sub('회사채I공모사채무보증AAz', '7010123', parse30)
        parse32 = re.sub('회사채I공모사채무보증A\W', '7010131', parse31)
        parse33 = re.sub('회사채I공모사채무보증A0', '7010132', parse32)
        parse34 = re.sub('회사채I공모사채무보증Az', '7010133', parse33)
        parse35 = re.sub('회사채I공모사채무보증BBB\W', '7010211', parse34)
        parse36 = re.sub('회사채I공모사채무보증BBB0', '7010212', parse35)
        parse37 = re.sub('회사채I공모사채무보증BBBz', '7010213', parse36)
        parse38 = re.sub('회사채II사모사채무보증AAA', '8010110', parse37)
        parse39 = re.sub('회사채II사모사채무보증AA', '8010120', parse38)
        parse40 = re.sub('회사채II사모사채무보증A\W', '8010131', parse39)
        parse41 = re.sub('회사채II사모사채무보증A0', '8010132', parse40)
        parse42 = re.sub('회사채II사모사채무보증Az', '8010133', parse41)

        grp_code = pd.DataFrame({'시가평가그룹코드': [parse42, parse42, parse42, parse42, parse42, parse42, parse42, parse42,
                                              parse42, parse42, parse42, parse42, parse42, parse42, parse42,parse42]})

        GRP_CODE_DATA.append(grp_code)
        MRK_PRICE_GRP_CODE = pd.concat(GRP_CODE_DATA, axis=0, ignore_index=True)
    #print(MRK_PRICE_GRP_CODE)

        #잔존기간
        for i in remain_num:
            remain_term_code = pd.DataFrame({'잔존기간': [i]})
            REMAIN_TERM_DATA.append(remain_term_code)
        REMAIN_TERM = pd.concat(REMAIN_TERM_DATA, axis=0, ignore_index=True)
    #print(REMAIN_TERM)

    #데이터프레임
    today = datetime.today()
    today = today.strftime('%Y%m%d')
    frame_data=[]
    for i in range(0,672) :
        data = pd.DataFrame({'일자':[today],
                             '처리구분':'1',
                             '입력일':[today]})
        frame_data.append(data)
        data_f = pd.concat(frame_data, axis=0, ignore_index=True)

    data_f['시가평가그룹코드'] =MRK_PRICE_GRP_CODE
    data_f['잔존기간']=REMAIN_TERM
    data_f['수익율']=EARNING_RATE
    #데이터 프레임
    # 순서 변경하기
    df1=pd.DataFrame(data_f,columns=['일자','시가평가그룹코드','잔존기간','수익율','처리구분','입력일'])
    print(df1)

    #저장
    df1.to_excel('C:\\KOFIA_EMERGENCY/RESULT.xlsx',sheet_name='sheet', index=False, header=False)



try:
    # 실행
    KOFIA_DATA()
    CONVERSION()
    KOFIA()

except:
    # 에러가 발생한 경우 StackTrace를 파일로 기록한다.
    outputFile = open('C:\\KOFIA_EMERGENCY/emergency_error.txt', 'w')
    traceback.print_exc(file=outputFile)
    outputFile.close()
