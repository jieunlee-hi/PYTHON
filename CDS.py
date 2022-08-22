# -*- coding:utf-8 -*-
import re
import traceback

import pandas as pd
from selenium import webdriver
from datetime import datetime, timedelta




def cds():
    all_data_frame = []
    code = ['argentina','australia','austria','belgium','brazil','bulgaria','canada','chile','china',\
    'colombia','czech-republic','denmark','egypt','finland','france','germany','greece',\
    'hong-kong','hungary','indonesia','ireland','israel','italy','japan','kazakhstan']
    for i in code:
        chrome_options = webdriver.ChromeOptions()
        chrome_options.add_argument('headless')
        driver = webdriver.Chrome('C:\\CDS\chromedriver_win32\chromedriver.exe',chrome_options=chrome_options)

        url = 'https://ko.tradingeconomics.com/'+i+'/rating'
        print(url)
        driver.get(url)
        str = ['argentina', 'austria', 'belgium', 'canada', 'colombia', 'finland','mexico','pakistan']
        for j in str:
            if i == j:
                i=j[:-3]
        i = re.sub('australia', 'austla', i, 0).strip()
        i = re.sub('bulgaria', 'bgaria', i, 0).strip()
        i = re.sub('czech-republic', 'czech', i, 0).strip()
        i = re.sub('denmark', 'denk', i, 0).strip()
        i = re.sub('france', 'frtr', i, 0).strip()
        i = re.sub('germany', 'dbr', i, 0).strip()
        i = re.sub('hungary', 'hungaa', i, 0).strip()
        i = re.sub('indonesia', 'indon', i, 0).strip()
        i = re.sub('ireland', 'irelnd', i, 0).strip()
        i = re.sub('kazakhstan', 'kazaks', i, 0).strip()
        i = re.sub('malaysia', 'malays', i, 0).strip()
        i = re.sub('netherlands', 'nethrs', i, 0).strip()
        i = re.sub('new-zealand', 'nz', i, 0).strip()
        i = re.sub('philippines', 'philip', i, 0).strip()
        i = re.sub('portugal', 'portug', i, 0).strip()
        i = re.sub('saudi-arabia', 'saudi', i, 0).strip()
        i = re.sub('south-africa', 'soaf', i, 0).strip()
        i = re.sub('south-korea', 'korea', i, 0).strip()
        i = re.sub('sweden', 'swed', i, 0).strip()
        i = re.sub('switzerland', 'swiss', i, 0).strip()
        i = re.sub('thailand', 'thai', i, 0).strip()
        i = re.sub('united-kingdom', 'ukin', i, 0).strip()
        i = re.sub('united-states', 'usgb', i, 0).strip()
        i = re.sub('venezuela', 'venz', i, 0).strip()
        i = re.sub('vietnam', 'vietnm', i, 0).strip()

        # 국가명
        SYM_CODE = i.upper()
        if SYM_CODE == 'HONG-KONG':
            SYM_CODE='CHINA-HongKong'

        print(SYM_CODE)
        #신용평가기관
        search1 = driver.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_ctl01_GridView1"]/tbody/tr[2]/td[1]')
        print(search1.text)
        CRD_ESTI_INST=search1.text

        #신용평가일
        # %b ,%d %Y
        search2 = driver.find_element_by_xpath(
            '//*[@id="ctl00_ContentPlaceHolder1_ctl01_GridView1"]/tbody/tr[2]/td[4]')
        search2=search2.text
        #날짜 포맷변경
        date = datetime.strptime(search2, '%b %d %Y')
        date1 = datetime.strftime(date, '%Y%m%d')

        #신용평가등급
        search3 = driver.find_element_by_xpath(
            '//*[@id="ctl00_ContentPlaceHolder1_ctl01_GridView1"]/tbody/tr[2]/td[2]')
        #print(search3.text)

        #OUTLOOK
        search4 = driver.find_element_by_xpath(
            '//*[@id="ctl00_ContentPlaceHolder1_ctl01_GridView1"]/tbody/tr[2]/td[3]/span')
        search4=search4.text
        OUTLOOK_TYPE=search4.capitalize()
        #print(search4.text)

        if CRD_ESTI_INST == "S&P":
            CRD_ESTI_INST = "1"
        elif CRD_ESTI_INST=="Moody's":
            CRD_ESTI_INST = "2"
        elif CRD_ESTI_INST== "Fitch":
            CRD_ESTI_INST = "3"


        # 데이터프레임생성
        data = {'국가명':[SYM_CODE],
                '신용평가기관': [CRD_ESTI_INST],
                '신용평가일': [date1],
                '신용평가등급': [search3.text],
                'OUTLOOK':[OUTLOOK_TYPE]
                }
        df = pd.DataFrame(data, columns=["국가명", "신용평가기관", "신용평가일", "신용평가등급", "OUTLOOK"])

        print(df)
        if  CRD_ESTI_INST != 'DBRS':

            # print(df)
            # 데이터프레임 이어붙이기
            all_data_frame.append(df)
            df_concat = pd.concat(all_data_frame, axis=0, ignore_index=False)
            print(df_concat)

            # 엑셀파일로 저장하기
            writer = pd.ExcelWriter('CDS_result1.xls')
            df_concat.to_excel(writer, sheet_name='Sheet1', index=False, header=False, na_rep=' ',
                               encoding='utf-16')  # 엑셀로 저장
            writer.save()
    if len(all_data_frame)==0:
        outputFile = open('data_is_null.txt', 'w')
        traceback.print_exc(file=outputFile)
        outputFile.close()


try:
    # 실행
    cds()
except:
    # 에러가 발생한 경우 StackTrace를 파일로 기록한다.
    outputFile = open('cds_error.txt', 'w')
    traceback.print_exc(file=outputFile)
    outputFile.close()


