# -*- coding:utf-8 -*-
import re
from datetime import datetime, timedelta
import time
import traceback
import pandas as pd
import sys, os, traceback
import os
import sys, os, traceback, glob
import win32com.client as win32
import sys
import os.path

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


def fund():
    today = datetime.today()
    today = today.strftime('%Y%m%d')

    #print(d)
    df=pd.read_excel("C:\\1008\\1.xlsx",sheet_name='유형별기간설정',engine='openpyxl',thousands = ',')
    # df1=pd.read_excel("C:\\1008\\3.xlsx",sheet_name='유형별기간설정')
    # df2=pd.read_excel("C:\\1008\\2.xlsx",sheet_name='유형별기간설정')
    df3=pd.read_excel("C:\\1008\\7.xlsx",sheet_name='신용공여 잔고 추이',engine='openpyxl',thousands = ',')
    df4=pd.read_excel("C:\\1008\\4.xlsx",sheet_name='유형별기간설정',engine='openpyxl',thousands = ',')
    # df5=pd.read_excel("C:\\1008\\6.xlsx",sheet_name='유형별기간설정')
    # df6=pd.read_excel("C:\\1008\\5.xlsx",sheet_name='유형별기간설정')

    #데이터 일자구하기
    df_3=df.iloc[3, 0]
    #df_3 = df_3.strftime('%Y%m%d')
    df_3 = re.sub('/', '', df_3)
    #print(df_3)
    # 기본키(일자, 일자구분,증자구분,이동평균구분)
    PK = pd.DataFrame({'PK0':[df_3],
                        'PK1': [1],
                       'PK2': [0],
                       'PK3': [0]
                   })
    PK = PK.rename(index={0: 3})

    #0값 생성
    d = pd.DataFrame({'1': [0],
                   '2': [0]
                      })
    d = d.rename(index={0: 3})
    d1 = pd.DataFrame({'1': [0],
                       '2': [0],
                       '3': [0],
                       '4': [0],
                       '5': [0],
                       '6': [0],
                       '7': [0]
                      })
    d1 = d1.rename(index={0: 3})


    #설정원본 전체
    df_1=df.iloc[3:4, 7:8]
    df_2=df.iloc[3:4, 4:5]
    df=df.iloc[3:4, 1:4]
    df['Unnamed1']='0'
    df['Unnamed2']=df_2
    df['Unnamed3']=df_1
    df['Unnamed4']='0'
    df['Unnamed5']='0'
    #print(df)
    #print(df['Unnamed5'])


    #설정원본 해외
    # df1=df1.iloc[3:4, 1:5]
    # df1['Unnamed6']='0'
    # df1['Unnamed7']='0'
    #
    # #설정원본 국내
    # df_2=df2.iloc[3:4, 7:8]
    # df2=df2.iloc[3:4, 1:5]
    # df2['Unnamed8']=df_2
    # df2['Unnamed9']='0'
    # df2['입력일']=today
    d2 = pd.DataFrame({'1': 0,
                       '2': 0,
                       '3': [today]
                       })
    d2 = d2.rename(index={0: 3})

    #예탁증권담보융자 날짜
    df3_d = df3.iloc[3,0]
    df3_d = re.sub('/', '', df3_d)
    #print(df3_d)


    # #예탁증권담보융자 데이터
    if df3_d == df_3 :
        df3 = df3.iloc[3, 8:9]
        df3 = df3.rename(index={'Unnamed: 8': 3})
        print(df3)
    else :
        df3 = df3.iloc[4, 8:9]
        df3 = df3.rename(index={'Unnamed: 8': 3})
        print(df3)
    
#
    # 순자산 전체
    df_4=df4.iloc[3:4, 7:8]
    df4=df4.iloc[3:4, 1:5]
    df4['Unnamed10']=df_4
    #df4['Unnamed11']='0'

    #순자산 해외
    # df5=df5.iloc[3:4, 1:5]
    # df5['Unnamed12']='0'
    #
    # #순자산 국내
    # df_6=df6.iloc[3:4, 7:8]
    # df6=df6.iloc[3:4, 1:5]
    # df6['Unnamed13']=df_6
    # df6['Unnamed14']='0'

    #데이터프레임 이어붙이기
    fund_data=pd.concat([PK,d1,df3,d,df,d,d,d,d,d,d,d,d,d,d,df4,d1,d1,d2,d1,d1], axis=1)
    print(fund_data)

    fund_data.to_excel('C:\\1008/PM_1008.xlsx',sheet_name='sheet', index=False, header=False)



def DELETE():
    # 작업디렉토리
    outpath = "C:\1008\\"
    # 에러기록용 파일명
    error_log_file_name = 'excel_cleansing_error.log'

    if os.path.isfile(error_log_file_name) == True:
        os.unlink(error_log_file_name)

    if os.path.isfile('C:\\1008\\1.xlsx') == True:
        os.unlink('C:\\1008\\1.xlsx')

    if os.path.isfile('C:\\1008\\4.xlsx') == True:
        os.unlink('C:\\1008\\4.xlsx')

    if os.path.isfile('C:\\1008\\7.xlsx') == True:
        os.unlink('C:\\1008\\7.xlsx')

try:
    # 실행
    fund()
    DELETE()
except:
    # 에러가 발생한 경우 StackTrace를 파일로 기록한다.
    outputFile = open('1008_error.txt', 'w')
    traceback.print_exc(file=outputFile)
    outputFile.close()
