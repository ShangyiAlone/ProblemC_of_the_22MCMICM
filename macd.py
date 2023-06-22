import pandas as pd
import numpy as np
import datetime
import time
#获取数据
df=pd.read_csv('E:/桌面/美赛代码/macd.csv',encoding='gbk')

#df.columns=['date','code','name','close','high','low','open','preclose',
#'change','change_per','volume','amt']

#df=df[['date','open','high','low','close','volume','amt']]
#print(df.head())
print(df.iloc[0,3])
def get_EMA(df,N):
    for i in range(len(df)):
        if i==0:
            df.iloc[i,12]=df.iloc[i,3]
#            df.ix[i,'ema']=0
        if i>0:
            df.iloc[i,12]=(2*df.iloc[i,3]+(N-1)*df.iloc[i-1,12])/(N+1)
    ema=list(df['ema'])
    return ema
 
def get_MACD(df,short=12,long=26,M=9):
    a=get_EMA(df,short)
    b=get_EMA(df,long)
    df['diff']=pd.Series(a)-pd.Series(b)
    #print(df['diff'])
    for i in range(len(df)):
        if i==0:
            df.iloc[i,14]=df.iloc[i,13]
        if i>0:
            df.iloc[i,14]=((M-1)*df.iloc[i-1,14]+2*df.iloc[i,13])/(M+1)
    df['macd']=2*(df['diff']-df['dea'])
    return df
get_MACD(df,12,26,9)


print(df)


with pd.ExcelWriter('macd_gold.xlsx') as writer:
    df.to_excel(writer, 'df1')
