import numpy as np
import math
import pandas as pd
import numpy as np
import xlwt
import matplotlib.pyplot as plt
#相对强弱指标RSI
#计算公式：
#<1> N日RSI =A/（A+B）×100
#<2> A=N日内收盘涨幅之和的平均 A = (前一日A*(n-1) + 当日涨值)/n
#<3>B=N日内收盘跌幅之和的平均 B = (前一日B*(n-1) + 当日跌值)/n （跌值取正值：如下跌-4.32，跌值=4.32）
#<4> 0<=RSI<=100
# 输入：
#     close_k: 收盘价list
#     periods：周期
#从第periods+1个周期开始预测
def RSI(close_k,periods):
    length = len(close_k)

    ans = [np.nan]*length
    A = 0
    B = 0

    sum1,sum2=0,0
    for j in range(periods):
        up = 0
        down = 0
        if close_k[j]>=close_k[j-1]:
            up = close_k[j]-close_k[j-1]
            down = 0
        else:
            up=0
            down = close_k[j-1]-close_k[j]
        sum1+=up
        sum2+=down

    A=sum1/periods
    B=sum2/periods
        
    for j in range(periods,length):
        up = 0
        down = 0
        if close_k[j]>=close_k[j-1]:
            up = close_k[j]-close_k[j-1]
            down = 0
        else:
            up=0
            down = close_k[j-1]-close_k[j]
        
        #计算N日内增长的均值，A为增加值的和的均值，B为减少值的和的均值
        A = (A*(periods-1)+up)/periods
        B = (B*(periods-1)+down)/periods
        if A + B!=0:
            ans[j] = 100 * A / (A + B)
    return ans


if __name__=='__main__':
    

    df=pd.read_excel("E://桌面//筛选后的数据.xlsx")

    book = xlwt.Workbook(encoding='utf-8',style_compression=0)

    sheet = book.add_sheet('lbma-gold',cell_overwrite_ok=True)

    data=[]
    for i in range(1825):
        data.append(df.values[i][1])    

    ans=RSI(data,6)
    for i in range(1825):
        sheet.write(i,0,ans[i])

    savepath = 'E:/桌面/rsi_mkpur.xlsx'
    book.save(savepath)
    
    #print(RSI(date,12))
    #print(len(RSI(date1,12)))







