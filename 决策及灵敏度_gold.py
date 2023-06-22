import numpy as np
import math
import pandas as pd
import numpy as np
import xlwt
import matplotlib.pyplot as plt #画图的库

mkpru=pd.read_excel("E://桌面//美赛代码//总数据汇总.xlsx",sheet_name="gold")

start=1000 #初始投资额
end=start #结束时的投资额

a = 0.3 #每次投资的最大值占总投资额的百分比

cost_m=0.15 #设置佣金

cost_g=[0.002,0.005,0.007,0.01,0.02,0.03,0.05,0.06,0.08,0.09,0.1,0.2,0.3]

p1,p2,p3=0.2,0.2,0.6 #三项指标的权重，顺序为macd,rsi,预测模型

sum1=0 #投出去的美元
sum2=0 #手中持有的比特币
vast=1000
#vast=100 #每次购买的值

col,row,li=[],[],[]

for n in range(len(cost_g)):
    cnt=0 #计算拒绝的次数
    cnt1=0#计算买入的次数
    cnt2=0#计算卖出的次数

    for i in range(1,1813):      
        #计算macd的得分
        if  mkpru.values[i-1][0]<0 and mkpru.values[i][0]>0:
            point1=5
        elif mkpru.values[i-1][0]>0 and mkpru.values[i][0]<0:
            point1=-5
        else:
            point1=0

        #计算rsi的得分
        if mkpru.values[i][1]>80:
            point2=5
        elif mkpru.values[i][1]<20:
            point2=-5
        else:
            point2=0

        #计算灰色预测的得分
        point3=0
        for j in range(10):
            if mkpru.values[i+j][2]<mkpru.values[i+j+1][2]:
                point3+=1
            else:
                break
        
        for j in range(10):
            if mkpru.values[i+j][2]>mkpru.values[i+j+1][2]:
                point3-=1
            else:
                break

        #预测收益,如果收益率大于佣金的三倍，执行交易,否则不执行
        if mkpru.values[i+1][0]/mkpru.values[i][0]-1 < 3*cost_g[n]:
            cnt+=1
            point1,point2,point3=0,0,0
        
        ratio=(point1*p1+point2*p2+point3*p3)/8

        if ratio>0.1 and end-vast>=0 and mkpru.values[i][3]!=mkpru.values[i-1][3]:
            cnt1+=1
            end-=vast
            sum1+=vast
            sum2+=vast/mkpru.values[i][3]

        earn=sum2*mkpru.values[i][3]*(1-cost_m)-sum1*(1+cost_g[n])
        if ratio<-0.3 and earn>=0:
            cnt2+=1
            earn=sum2*mkpru.values[i][3]-sum1*(1+cost_g[n])
            end+=earn


    #假设最后一天会卖出
    earn=sum2*mkpru.values[1813][3]*(1-cost_m)-sum1*(1+cost_g[n])
    end+=earn
    
    print(end," ",cnt," ",cnt1)
    col.append(cnt)
    row.append(end)
    li.append(cnt1)


    


book = xlwt.Workbook(encoding='utf-8',style_compression=0)

sheet = book.add_sheet('lbma-gold',cell_overwrite_ok=True)  

for i in range(10):
    sheet.write(i,0,row[i])
    sheet.write(i,1,col[i])
    sheet.write(i,2,li[i])

savepath = 'E:/桌面/200.xlsx'
book.save(savepath)












    
