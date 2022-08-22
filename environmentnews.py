from konlpy.tag import Okt
from collections import Counter
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from datetime import datetime
data_path ='/content/drive/MyDrive/'
df=pd.read_excel(data_path + '1year_environment_issue.xlsx', header=None, index_col=None)
df.rename(columns={0: "News", 1: "Date"}, inplace=True)
#int64 타입인 date컬럼을 날짜 로 변경
df['DateTime'] = pd.to_datetime(df['Date'].astype(str), format='%Y%m%d')

#연도와 월 추출
df['Year'] = df['DateTime'].dt.year
df['Month'] = df['DateTime'].dt.month


all=[]
appen = all.append

for i in range(1,12):
  news=df['News'][df['Month']==i].to_list()
  #news=df['News'].to_list()
  #print(news)
  # Okt 함수를 이용해 형태소 분석
  okt = Okt()
  all_data_frame=[]
  append = all_data_frame.extend
  line =[]
  for num in news:
    line = okt.pos(num)
    n_adj =[]
    # 명사 또는 형용사인 단어만 n_adj에 넣어주기
    for word, tag in line:
        if tag in ['Noun','Adjective']:
            n_adj.append(word)
    #print(n_adj)

    #제외할 단어 추가
    stop_words = "하자 곳 도 관 환경 등 명 개 낮 위 첫 곳곳 제 올해 종합 감 날 중 회 종 진" #추가할 때 띄어쓰기로 추가해주기
    stop_words = set(stop_words.split(' '))

    # 불용어를 제외한 단어만 남기기
    n_adj = [word for word in n_adj if not word in stop_words]
    #print(n_adj)
    append(n_adj)
    #가장 많이 나온 단어 100개 저장
    counts = Counter(all_data_frame)
    tags = counts.most_common(100)
    print(tags)
  appen(i)
  df1=pd.DataFrame({'month':all,'tags':tags})
  print(df1)
