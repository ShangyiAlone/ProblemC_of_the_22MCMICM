# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""
import numpy as np
import math
import pandas as pd
import numpy as np
import xlwt
import matplotlib.pyplot as plt #画图的库

def predict(history_data):

    n = len(history_data)
    X0 = np.array(history_data)
    #累加生成
    history_data_agg = [sum(history_data[0:i+1]) for i in range(n)]
    X1 = np.array(history_data_agg)

    #计算数据矩阵B和数据向量Y
    B = np.zeros([n-1,2])
    Y = np.zeros([n-1,1])
    for i in range(0,n-1):
        B[i][0] = -0.5*(X1[i] + X1[i+1])
        B[i][1] = 1
        Y[i][0] = X0[i+1]

    #计算GM(1,1)微分方程的参数a和u
    #A = np.zeros([2,1])
    A = np.linalg.inv(B.T.dot(B)).dot(B.T).dot(Y)
    a = A[0][0]
    u = A[1][0]

    #建立灰色预测模型
    XX0 = np.zeros(n)
    XX0[0] = X0[0]
    for i in range(1,n):
        XX0[i] = (X0[0] - u/a)*(1-math.exp(a))*math.exp(-a*(i));


    #模型精度的后验差检验
    e = 0      #求残差平均值
    for i in range(0,n):
        e += (X0[i] - XX0[i])
    e /= n

    #求历史数据平均值
    aver = 0;     
    for i in range(0,n):
        aver += X0[i]
    aver /= n

    #求历史数据方差
    s12 = 0;     
    for i in range(0,n):
        s12 += (X0[i]-aver)**2;
    s12 /= n

    #求残差方差
    s22 = 0;       
    for i in range(0,n):
        s22 += ((X0[i] - XX0[i]) - e)**2;
    s22 /= n

    #求后验差比值
    C = s22 / s12   

    #求小误差概率
    cout = 0
    for i in range(0,n):
        if abs((X0[i] - XX0[i]) - e) < 0.6754*math.sqrt(s12):
            cout = cout+1
        else:
            cout = cout
    P = cout / n

    if (C < 0.5 and P > 0.7):
        #预测精度为一级
        m = 1   #请输入需要预测的年数
        #print('往后m各年负荷为：')
        f = np.zeros(m)
        for i in range(0,m):
            return (X0[0] - u/a)*(1-math.exp(a))*math.exp(-a*(i+n))
        print(f)
    else:
        print('灰色预测法不适用')

if __name__ == '__main__':
    df=pd.read_excel("E://桌面//美赛代码//lbma-gold（筛选）.xlsx")

    df1=pd.read_excel("E://桌面//美赛代码//bchain-mkpru（筛选）.xlsx")

    data=[]
    for i in range(727):
        data.append(df.values[i][1])

    book = xlwt.Workbook(encoding='utf-8',style_compression=0)

    sheet = book.add_sheet('lbma-gold',cell_overwrite_ok=True)    
    #print(data[726])
    cnt=9
    ans=[]
    history_data=[0]*10
    #print(history_data)

    #过去若干个交易日的数据，从第十个数据开始预测
    for i in range(11,727):
        #更新预测模型的数据
        cnt+=1
        if cnt%10==0:
            temp = data[i-10:i]
            #print(temp)
            for j in range(10):
                history_data[j]=temp[j]
            cnt=0
        #预测数据   
        ans.append(predict(history_data))


    #print(ans)
    for i in range(702):
        sheet.write(i,0,ans[i])
        
    savepath = 'E:/桌面/灰色预测_gold.xlsx'
    book.save(savepath)

    plt.figure()
    
    plt.plot(list(range(len(ans))), ans, color='b')

    plt.plot(list(range(727)), data[:727], color='y')

    plt.show()































    
