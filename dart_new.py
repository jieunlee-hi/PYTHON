# coding=utf-8
# open_api 이용
import re
from datetime import datetime, timedelta
import time
import urllib.parse
from urllib.request import urlopen
import pandas as pd
from bs4 import BeautifulSoup as bf
import traceback
import requests
from io import BytesIO
import zipfile

session = requests.Session()
session.verify = False
#session.post(url='https://foo.com', data={'bar':'baz'})


# 전영업일구하기
def preday_search():
    today = datetime.today()
    today = today.strftime('%Y%m%d')
    yesterday = datetime.fromtimestamp(time.time() - 60 * 60 * 24)
    yesterday = datetime.strftime(yesterday, '%Y%m%d')
    preday = datetime.strptime(yesterday, '%Y%m%d').date()
    if datetime.weekday(preday) == 5:  # 토요일
        yesterday = preday - timedelta(days=1)
        yesterday = datetime.strftime(yesterday, '%Y%m%d')
    elif datetime.weekday(preday) == 6:  # 일요일
        yesterday = preday - timedelta(days=2)  # then make it Friday
        yesterday = datetime.strftime(yesterday, '%Y%m%d')
    return yesterday

def dart_elsdls():
    API_KEY = "eb7269185750072f9b91761e607412347467f1f0"

    all_data_frame = []
    append = all_data_frame.append
    for i in range(1, 10):
        # 공시검색 OPENAPI
        url_json = "http://opendart.fss.or.kr/api/list.json"
        url_xml = "http://opendart.fss.or.kr/api/list.xml"
        params = {
            'crtfc_key': API_KEY,
            'bgn_de': preday_search(),
            #'bgn_de':'20220720',# < 날짜따로지정시
            'pblntf_ty': 'C',
            'pblntf_detail_ty': 'C002',
            'pblntf_detail_ty': 'C003',
            'pblntf_detail_ty': 'C005',
            'page_count': '100',
            'page_no': str(i)
        }
        response = requests.get(url_json, params=params)
        data = response.json()
        # 전체페이지수만큼 페이지넘기기
        if str(data['total_page']) < str(i):
            break

        # print(data)
        data_list = data.get('list')
        df_list = pd.DataFrame(data_list)
        # print(df_list)
        # 제목이 투자설명서 / 회사이름에 투자,증권이 들어있는 공시만 추출
        df_list = df_list[df_list['report_nm'].str.contains('투자설명서')]
        df_list = df_list[df_list['report_nm'].str.contains('기재정정') == False]
        df_list = df_list[df_list['corp_name'].str.contains('증권') | df_list['corp_name'].str.contains('투자')]
        df_list = df_list[['corp_name', 'report_nm', 'rcept_no']]
        # 인덱스리셋!
        df_list.reset_index(drop=True, inplace=True)
        #print(df_list)

        # 공시서류원본파일다운로드 OPENAPI
        url = "https://opendart.fss.or.kr/api/document.xml"
        API_KEY = "eb7269185750072f9b91761e607412347467f1f0"

        # 조건을주어 추출한 공시들의 금감원문서번호들을 리스트로 만든다.
        rcept_list = list(df_list['rcept_no'])
        #print(rcept_list)

        # xml파일리스트를 넣을곳
        xmlfile = []

        # 조건준 공시만 파일다운로드
        for i in range(0, len(rcept_list)):
            params = {
                'crtfc_key': API_KEY,
                'rcept_no': rcept_list[i]
            }
            # xml파일리스트저장
            xmlfile.append(rcept_list[i])
            time.sleep(1)

            # 공시원본파일저장
            response = requests.get(url, params=params)
            zf = zipfile.ZipFile(BytesIO(response.content))
            zf.extractall()
            zipinfo = zf.infolist()
            filenames = [x.filename for x in zipinfo]
            filename = filenames[0]
            # print(filename)
            zf.extract(filename)
            zf.close()

            # for enc in ['UTF-8', 'EUC-KR']:
            #     try:
            #         with open(filename, 'r', encoding=enc) as f:
            #             f.read()
            #         break
            #     except:
            #         print(enc, '실패')

            # encoding 깨짐문제해결
            # euc-kr로 읽어서 utf-8로저장
            with open(f'./{filename}', 'r', encoding='euc-kr') as fp:
                lines = fp.readlines()
            with open(f'./{filename}', 'w', encoding='utf-8') as fp:
                fp.writelines(lines)
        # 저장된 xml파일 확인
        print(xmlfile)

        # xml파일 읽기
        cnt = []
        for i in xmlfile:
            with open(f'./{i}.xml', 'r', encoding='utf-8') as f:
                _file = bf(f, features="html.parser")
            # print(_file)
            # 회차가 적힌부분추출
            td_data = _file.find_all('td')
            td = str(td_data[1].get_text) + str(td_data[2].get_text) + str(td_data[3].get_text)
            td = td.replace("\n", "")
            # print("처리전\t" + td)

            regex = re.compile('<td.*\n?</td>')
            matchobj = regex.finditer(td)

            # 데이터 정제

            for r in matchobj:
                match = r.group(0)
                parse = re.sub('<.+?>', '', match, 0).strip()
                parse = re.sub('.*주식회사|.*주\s식\s회\s사|투자위험도|\s', '', parse, 0).strip()
                parse = re.sub(
                    '\W?원금비보장형,고위험\W?|\W?초고위험,원금비보장형\W?|\W?고위험,원금비보장형\W?|\W?원금비보장형,높은위험\W?|\W?원금비보장형,초고위험\W?',
                    '', parse, 0).strip()
                parse = re.sub('\W원금부분지급형/원금비보장형\W|\W원금9\d%부분지급형\W원금비보장\W|\W원금부분지급형\W|\W중위험,원금\d?\d?%부분지급형\W원금비보장\W',
                               '',
                               parse, 0).strip()
                parse = re.sub('초고위험\W|&amp;cr;|\W원금9\d%부분지급형,중위험\W|:\W?|\W?원금부분지급\W?|\W원금비보장형\W|', '', parse)
                parse = re.sub('\W?증권의종목과발행증권수\W?', '', parse, 0).strip()
                parse = re.sub(
                    '금\d.?.?\W\d..\W\d..\W\d..원|\d?\d?\d?\W?\d..?\W?\d..\W\d..원|/|\W?\d등급\W|금\d\d?\W\d\d\d원|\W?\d?amp\Wcr\W?',
                    '', parse, 0).strip()
                parse = re.sub('\d?\d?\d?\W?\d?\d?\d?\W?\d?\d.\W\d..증권|', '', parse)
                parse = re.sub('\W?저위험\W?|\W?보통위험\W?|\W?중위험\W?', '', parse, 0).strip()
                parse = re.sub(
                    '\W?모집가액총액|\W?모집또는매출총액\W?\d?|모집가액또는매출가액의총액\d?\d?\d?\w?\w?|\W?USD\d?\d?\W?\d?\d?\W?\d..\W\d..\W', '',
                    parse, 0).strip()
                parse = re.sub('\WNH', '),NH', parse, 0).strip()
                parse = re.sub('회트루', '회,트루', parse, 0).strip()
                parse = re.sub('\W신한', '),신한', parse, 0).strip()
                parse = re.sub('\W한국', '),한국', parse, 0).strip()
                parse = re.sub('회메리츠', '회,메리츠', parse, 0).strip()
                parse = re.sub('\W미래', '),미래', parse, 0).strip()
                parse = re.sub('\W현대차', '),현대차', parse, 0).strip()
                parse = re.sub('회메리츠', '회,메리츠', parse, 0).strip()
                parse = re.sub('유안타홈런\w', '유안타MY', parse, 0).strip()
                parse = re.sub('DB해피플러스', '해피플러스', parse, 0).strip()
                parse = re.sub('메리츠종합금융증권', '메리츠증권', parse, 0).strip()
                parse = re.sub('대신\WBalance\W', '대신증권Balance', parse, 0).strip()
                parse = re.sub('키움증권HERO', '키움HERO', parse, 0).strip()
                parse = re.sub('유안타파생결합증권', '유안타MY', parse, 0).strip()
                parse = re.sub('호유안타', '호,유안타', parse, 0).strip()
                parse = re.sub('회NH투자증권', '회,NH투자증권', parse, 0).strip()
                parse = re.sub('회키움', '회,키움', parse, 0).strip()
                parse = re.sub('회,키움드림공모', '회키움드림공모', parse, 0).strip()
                parse = re.sub('\W주가연계증권\W키움', '(주가연계증권),키움', parse, 0).strip()
                parse = re.sub('회\W$', '회', parse, 0).strip()
                parse = re.sub('호\W$', '호', parse, 0).strip()
                parse = re.sub('^\W\W?|1.증권신고의효력발생일|>|[미화구백구십만달러]|^\W', '', parse, 0).strip()
                parse = re.sub('한국투자증권트루', '트루', parse, 0).strip()
                parse = re.sub('한화투자증권|한스마트', '한화스마트', parse, 0).strip()
                parse = re.sub('래에셋', '미래에셋', parse, 0).strip()
                parse = re.sub('DB해피플러스|해피플스', '해피플러스', parse, 0).strip()
                parse = re.sub('금일\w?\d?억원|()|금\w?\w?억원\W?\W?|\W일천\W|\WUSD\W', '', parse, 0).strip()
                parse = re.sub('\W주\WBNK', 'BNK', parse, 0).strip()
                parse = re.sub('BNK투자증권BNK투자증권', 'BNK투자증권', parse, 0).strip()
                parse = re.sub('^\W^', '', parse, 0).strip()
                parse = re.sub('주가연계증권\W\W?\W?', '주가연계증권)', parse, 0).strip()
                parse = re.sub('한디럭스', '한화디럭스', parse, 0).strip()
                parse = re.sub('한투자증권', '한화투자증권', parse, 0).strip()
                parse = re.sub('\W한화', '),한화', parse, 0).strip()
                parse = re.sub('\W하이', '),하이', parse, 0).strip()
                parse = re.sub('\W현대차', '),현대차', parse, 0).strip()
                parse = re.sub('\W대신', '),대신', parse, 0).strip()
                parse = re.sub('권대신', '권,대신', parse, 0).strip()
                parse = re.sub('채대신', '채,대신', parse, 0).strip()
                parse = re.sub('\WELS\W', '(ELS)', parse, 0).strip()
                parse = re.sub('권\WNH', '권),NH', parse, 0).strip()
                parse = re.sub('권\W신영', '권),신영', parse, 0).strip()
                parse = re.sub('권\WKB', '권),KB', parse, 0).strip()
                parse = re.sub('\W교보', '),교보', parse, 0).strip()
                parse = re.sub('권교보', '권,교보', parse, 0).strip()
                parse = re.sub('채교보', '채,교보', parse, 0).strip()
                parse = re.sub('권\W하나', '권),하나', parse, 0).strip()
                parse = re.sub('권하나', '권,하나', parse, 0).strip()
                parse = re.sub('채하나', '채,하나', parse, 0).strip()
                parse = re.sub('증권\WIBK', '증권),IBK', parse, 0).strip()
                parse = re.sub('권IBK', '권,IBK', parse, 0).strip()
                parse = re.sub('채IBK', '채,IBK', parse, 0).strip()
                parse = re.sub('권\W키움', '권),키움', parse, 0).strip()
                parse = re.sub('권키움', '권),키움', parse, 0).strip()
                parse = re.sub('주\WBNK', 'BNK', parse, 0).strip()

                #print("처리 후\t" + parse)
            cnt.append(parse)
        # print(cnt)
        cnt = pd.DataFrame({"cnt": cnt})

        # print(parse_concat)

        df = pd.DataFrame({"금감원문서번호": df_list['rcept_no'], "발행기관명": df_list['corp_name'], "종목명": cnt['cnt']})


        # 데이터프레임 이어붙이기
        append(df)
        df_concat = pd.concat(all_data_frame, axis=0, ignore_index=False)
        print(df_concat)
        #els/dls외의 채권 제외
        df_concat= df_concat[(df_concat['종목명'].str.contains('주식워런트')|df_concat['종목명'].str.contains('상장지수증권')|df_concat['종목명'].str.contains('금적립계좌')|df_concat['종목명'].str.contains('무기명식')) == False]
        # 엑셀파일로 저장하기
        writer = pd.ExcelWriter('dartelsdls.xlsx')
        df_concat.to_excel(writer, sheet_name='Sheet1', index=False, header=False, na_rep=' ',
                           encoding='utf-16')  # 엑셀로 저장
        writer.save()

def DELETE():
    import os
    import glob
    # 작업디렉토리내의 다운받은 xml파일 전체 삭제

    py_files = glob.glob('C:\\dart\\*.xml')
    for py_file in py_files:
        try:
            os.remove(py_file)
        except OSError as e:
            print(f"Error:{e.strerror}")

try:
    # 실행
    preday_search()
    dart_elsdls()
    DELETE()
except:
    # 에러가 발생한 경우 StackTrace를 파일로 기록한다.
    outputFile = open('dartelsdls_error.txt', 'w')
    traceback.print_exc(file=outputFile)
    outputFile.close()
