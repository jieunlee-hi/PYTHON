import requests
from io import BytesIO
import time
import sys, os, traceback
import xlwt
from selenium.webdriver.common.action_chains import ActionChains
import pandas as pd
from selenium.webdriver.common.keys import Keys
from datetime import datetime, timedelta
from selenium import webdriver
import openpyxl
from pandas import Series, DataFrame


# 한기평 - 기업신용평가
# 다운로드된 파일(등급공시.xls)의 확장자가 .xls 이지만 실제 내용을 보면 HTML 문서 이다.
# POST방식
def kap_comp_master():
    try :
        url = 'http://www.rating.co.kr/disclosure/QDisclosure002Excel.do'
        data = {
            'evalNm': '당일공시',
            'svctyNm': 'ICR',
            'evalDt': '0',
            'svcty': '34',
        }

        r = requests.post(url, data=data)
        dfs = pd.read_html(BytesIO(r.content), header=0, encoding='utf-8')
        df = dfs[2].copy()
        # 첫번째,두번째 행이 병합되어서 세번째 행부터 보여준다
        df = df[1:len(df)]

        # 1건도 없는 경우("공시 자료가 없습니다.") 제외
        #df = df[df['회사명'] != '공시 자료가 없습니다.']
        # 공시가 있는 경우에만 헤더를 정의한다. 공시가 없는 경우는 더미컬럼을 두어 컬럼갯수를 맞춰준다.
        if len(df) == 0:
            df = df.assign(dummy1='')
            df = df.assign(dummy2='')
        if len(df) != 0:
            # 컬럼헤더명 정의
            df.columns = ['회사명', '구분', '직전등급', '직전전망', '현재등급', '현재전망', '평가일', '공시일']
            # 필요한 컬럼만 선택
            df = df[['회사명', '구분', '직전등급', '직전전망', '현재등급', '현재전망', '평가일']]

            # 신용평가기관
            df = df.assign(신용평가기관=3)

            # 날짜 원본형식(YYYY.MM.DD)을 변경(YYYYMMDD)
            df['평가일'] = df['평가일'].str.replace('.', '')
            df['현재등급'] = df['현재등급'].str.replace('↓', '').str.replace('↑', '')
            df['직전등급'] = df['직전등급'].str.replace('↓', '').str.replace('↑', '')
            df['현재전망'] = df['현재전망'].str.replace('안정적', '1').str.replace('긍정적', '2').str.replace('부정적', '3').str.replace(
                '유동적', '4').str.replace('없음', '5')
            df['직전전망'] = df['직전전망'].str.replace('안정적', '1').str.replace('긍정적', '2').str.replace('부정적', '3').str.replace(
                '유동적', '4').str.replace('없음', '5')
            df['구분'] = df['구분'].str.replace('본', '21').str.replace('정기', '22').str.replace('수시', '23')

            # DB 컬럼 정의 순서대로 맞춤
            df = df[['평가일', '신용평가기관', '회사명', '현재등급', '구분', '현재전망', '직전등급', '직전전망']]

            # 평가일 조건
            today = datetime.today().strftime('%Y%m%d')
            df = df[df['평가일'] == today]
        # if df[df['회사명'] == '공시 자료가 없습니다.'] :
        #     pass

        # 색인과 컬럼은 파일에 저장하지 않음. 구분자는 |. 누락값은 한칸공백으로 치환.
        # df.to_csv('kap_comp.csv',index=False,header=False,sep='|',na_rep=' ',encoding='utf-16')
        return df

    except AttributeError: 
        df = DataFrame(columns=('신용평가일', '신용평가기관', '한글회사명', '신용평가등급', '신용평가종류', 'OUTLOOKRATING',
                                       '직전등급', '직전OUTLOOKRATING'))
        return df
    except ValueError:  
        df = DataFrame(columns=('신용평가일', '신용평가기관', '한글회사명', '신용평가등급', '신용평가종류', 'OUTLOOKRATING',
                                       '직전등급', '직전OUTLOOKRATING'))
        return df



# 한신평 - Issuer Rating(기업신용평가)
# 다운로드된 파일(등급공시(오늘일자).xls)의 확장자가 .xls 이지만 실제 내용을 보면 HTML 문서 이다.
# GET방식
def kis_comp_master():
    try :
        today = datetime.today()
        today = today.strftime('%Y%m%d')
        chrome_options = webdriver.ChromeOptions()
        chrome_options.add_argument('headless')

        driver = webdriver.Chrome('C:\credit\chromedriver_win32\chromedriver.exe', chrome_options=chrome_options)
        driver.implicitly_wait(60)
        driver.get('https://www.kisrating.com/ratings/hot_disclosure.do')

        click = driver.find_element_by_xpath('//*[@id="tab4"]')
        actionChains = ActionChains(driver)
        actionChains.double_click(click).perform()
        time.sleep(3)

        # start = driver.find_element_by_xpath('//*[@id="startDt"]')
        # actionChains2 = ActionChains(driver)
        # actionChains2.double_click(start).perform()
        # start.send_keys('20191217')
        # time.sleep(2)
        # driver.find_element_by_xpath('//*[@id="btnSearch"]').send_keys(Keys.ENTER)
        # time.sleep(5)
        cnt = driver.find_element_by_xpath('//*[@id="view4"]/div[1]/h2/span/em')
        cnt = cnt.text

        all_data_frame = []
        for i in range(1, int(cnt)+1):
            # 평가일
            item1 = driver.find_element_by_xpath('//*[@id="issueList"]/tbody/tr[' + str(i) + ']/td[11]')
            item1 = item1.text

            # 날짜 포맷변경
            date = datetime.strptime(item1, '%Y.%m.%d')
            date1 = datetime.strftime(date, '%Y%m%d')

            # 회사명
            item2 = driver.find_element_by_xpath('//*[@id="issueList"]/tbody/tr[' + str(i) + ']/td[2]/a')
            # print(item2.text)

            # 현재등급
            item3 = driver.find_element_by_xpath('//*[@id="issueList"]/tbody/tr[' + str(i) + ']/td[8]')
            # print(item5.text)

            # 평가구분
            item4 = driver.find_element_by_xpath('//*[@id="issueList"]/tbody/tr[' + str(i) + ']/td[4]')
            # print(item6.text)
            item4 = item4.text
            item4 = item4.replace("본", "21")
            item4 = item4.replace("정기", "22")
            item4 = item4.replace("수시", "23")

            # 현재Outlook
            item5 = driver.find_element_by_xpath('//*[@id="issueList"]/tbody/tr[' + str(i) + ']/td[9]')
            # print(item7.text)
            item5 = item5.text
            item5 = item5.replace("안정적", "1")
            item5 = item5.replace("긍정적", "2")
            item5 = item5.replace("부정적", "3")
            item5 = item5.replace("유동적", "4")
            item5 = item5.replace("없음", "5")

            # 직전등급
            item6 = driver.find_element_by_xpath('//*[@id="issueList"]/tbody/tr[' + str(i) + ']/td[5]')
            # print(item6.text)
            # 직전 outlook
            item7 = driver.find_element_by_xpath('//*[@id="issueList"]/tbody/tr[' + str(i) + ']/td[6]')
            item7 = item7.text
            item7 = item7.replace("안정적", "1")
            item7 = item7.replace("긍정적", "2")
            item7 = item7.replace("부정적", "3")
            item7 = item7.replace("유동적", "4")
            item7 = item7.replace("없음", "5")

            # print(item9.text)

            data = {'신용평가일': [date1],
                    '신용평가기관': '1',
                    '한글회사명': [item2.text],
                    '신용평가등급': [item3.text],
                    '신용평가종류': [item4],
                    'OUTLOOKRATING': [item5],
                    '직전등급': [item6.text],
                    '직전OUTLOOKRATING': [item7]
                    }

            df = pd.DataFrame(data,
                              columns=['신용평가일', '신용평가기관', '한글회사명', '신용평가등급', '신용평가종류', 'OUTLOOKRATING',
                                       '직전등급', '직전OUTLOOKRATING'])
            if date1 == today:
                all_data_frame.append(df)
        df_concat = pd.concat(all_data_frame, axis=0, ignore_index=False)
        #print(df_concat)
        return df_concat

    except AttributeError:
        df_concat = DataFrame(columns=('신용평가일', '신용평가기관', '한글회사명', '신용평가등급', '신용평가종류', 'OUTLOOKRATING',
                                       '직전등급', '직전OUTLOOKRATING'))
        return df_concat
    except ValueError:
        df_concat = DataFrame(columns=('신용평가일', '신용평가기관', '한글회사명', '신용평가등급', '신용평가종류', 'OUTLOOKRATING',
                                       '직전등급', '직전OUTLOOKRATING'))
        return df_concat


# 한신정 - 기업신용평가
# 다운로드된 파일(일일등급속보_오늘일자.xls)의 확장자가 .xls 이므로 EXCEL 문서 이다.
# GET방식
def nice_comp_master():
    today = datetime.today().strftime('%Y-%m-%d')
    secuTyp = 'ICR'
    strDate = today
    endDate = today
    url = 'http://www.nicerating.com/disclosure/dayRatingPoiExcel.do?today=' + today + '&cmpCd=&seriesNm=&secuTyp=' + secuTyp + '&strDate=' + strDate + '&endDate=' + endDate

    r = requests.get(url, stream=True)
    dfs = pd.read_excel(BytesIO(r.content), header=0, thousands=',')
    # 컬럼헤더명 정의
    dfs.columns = ['기업명', '평정', '직전등급', '직전전망', '현재등급', '현재전망', '등급결정일', '등급확정일', '유효기간']
    # 필요한 컬럼만 선택
    dfs = dfs[['기업명', '평정', '직전등급', '직전전망', '현재등급', '현재전망', '등급확정일']]
    # df = dfs[0].copy()
    # 신용평가기관
    dfs = dfs.assign(신용평가기관=2)

    # 날짜 원본형식(YYYY.MM.DD)을 변경(YYYYMMDD)
    dfs['등급확정일'] = dfs['등급확정일'].str.replace('.', '')
    dfs['현재등급'] = dfs['현재등급'].str.replace('↓', '').str.replace('↑', '')
    dfs['직전등급'] = dfs['직전등급'].str.replace('↓', '').str.replace('↑', '')
    dfs['현재전망'] = dfs['현재전망'].str.replace('Stable', '1').str.replace('Positive', '2').str.replace('Negative',
                                                                                                  '3').str.replace(
        'Developing', '4').str.replace('None', '5')
    dfs['직전전망'] = dfs['직전전망'].str.replace('Stable', '1').str.replace('Positive', '2').str.replace('Negative',
                                                                                                  '3').str.replace(
        'Developing', '4').str.replace('None', '5')
    dfs['평정'] = dfs['평정'].str.replace('본', '21').str.replace('정기', '22').str.replace('수시', '23')

    # 첫번째,두번째 행이 병합되어서 세번째 행부터 보여준다
    dfs = dfs[2:len(dfs)]
    # DB 컬럼 정의 순서대로 맞춤
    dfs = dfs[['등급확정일', '신용평가기관', '기업명', '현재등급', '평정', '현재전망', '직전등급', '직전전망']]

    # 색인과 컬럼은 파일에 저장하지 않음. 구분자는 |. 누락값은 한칸공백으로 치환.
    # dfs.to_csv('nice_comp.csv',index=False,header=False,sep='|',na_rep=' ',encoding='utf-16')
    #print(dfs)
    return dfs


try:
    # 이전에 에러가 기록된 파일은 삭제한다.
    os.unlink('comp_cred_error.txt')
except:
    pass


try:
    dfm_kap = kap_comp_master()
    dfm_nice = nice_comp_master()
    dfm_kis = kis_comp_master()

    # 3개 평가사의 각 헤더명이 다르므로, 합치기 전 동일하게 맞춰준다.
    dfm_kap.columns = ['신용평가일', '신용평가기관', '한글회사명', '신용평가등급', '신용평가종류', 'OUTLOOKRATING', '직전등급', '직전OUTLOOKRATING']
    dfm_nice.columns = ['신용평가일', '신용평가기관', '한글회사명', '신용평가등급', '신용평가종류', 'OUTLOOKRATING', '직전등급', '직전OUTLOOKRATING']
    dfm_kis.columns = ['신용평가일', '신용평가기관', '한글회사명', '신용평가등급', '신용평가종류', 'OUTLOOKRATING', '직전등급', '직전OUTLOOKRATING']

    dfm_all = pd.concat([dfm_kap, dfm_nice,dfm_kis])
    # 3개 평가사의 결과를 합쳐 파일로 저장한다
    #dfm_all.to_csv('grade_comp.csv', index=False, header=False, sep='|', na_rep=' ', encoding='utf-16')
    writer = pd.ExcelWriter('grade_comp.xls')
    dfm_all.to_excel(writer,sheet_name='Sheet1',index=False,header=False,na_rep=' ',encoding='utf-16')
    writer.save()

except:
    # 에러가 발생한 경우 StackTrace를 파일로 기록한다.
    outputFile = open('comp_cred_error.txt', 'w')
    traceback.print_exc(file=outputFile)
    outputFile.close()
