import datetime
import numpy as np
import math
import pandas as pd
import numpy as np
import xlwt
import matplotlib.pyplot as plt


df=pd.read_excel("E://桌面//LBMA-GOLD.xlsm")

book = xlwt.Workbook(encoding='utf-8',style_compression=0)

sheet = book.add_sheet('lbma-gold',cell_overwrite_ok=True)

value,data=[],[]
for i in range(1254):
    data.append(df.values[i][2])
    value.append(df.values[i][1])


#print(str(data[0]).split('/'))
#1253

price=[]
for i in range(125):
    ans=str(data[i]).split('/')
    #print(ans)
    str1='null'
    if len(ans[0])==4:
        if int(ans[0])>2009:
            str1=ans[0][-2:]
        else:
            str1=ans[0][-1:]
            #print(str1)
    else:
        str1=ans[0]
        #print(str1)

    str2=ans[1]
    
    str3=ans[2]

    num1,num2,num3=int("20"+str3),int(str1),int(str2)
    d1 = datetime.datetime(num1,num2,num3)# 第一个日期


    ans=str(data[i+1]).split('/')
    #print(ans)
    str1='null'
    if len(ans[0])==4:
        if int(ans[0])>2009:
            str1=ans[0][-2:]
        else:
            str1=ans[0][-1:]
            #print(str1)
    else:
        str1=ans[0]
        #print(str1)

    str2=ans[1]
    
    str3=ans[2]

    num1,num2,num3=int("20"+str3),int(str1),int(str2)
    d2 = datetime.datetime(num1,num2,num3)


    interval = d2 - d1

    #price.append(interval.days)
    #print(interval.days) 
    for j in range(interval.days):
        price.append(value[i])
for i in range(len(price)):
    sheet.write(i,0,price[i])

savepath = 'E:/桌面/筛选后的数据.xlsx'
book.save(savepath)

        
"""
i=18
ans=str(data[i]).split('/')
    #print(ans)
str1='null'
if len(ans[0])==4:
    if int(ans[0])>2009:
        str1=ans[0][-2:]
        print("%%")
    else:
        str1=ans[0][-1:]
            #print(str1)
else:
    str1=ans[0]
        #print(str1)

str2=ans[1]
    
str3=ans[2]
print(str1,int(str2),str3)

"""
