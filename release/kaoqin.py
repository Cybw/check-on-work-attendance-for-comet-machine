# -*- coding: utf-8 -*-
"""
Created on Mon Mar  4 10:30:40 2019

@author: CyberWang
"""

"""
使用说明：
    1、“coursetab.csv”文件为教务处公布的每日上课时间段；
    2、“coursetime.csv”文件为参与打卡的每位同学的上课时间（所有
    参与打卡的同学必须都包含在表内，学校助管、公共活动等都可按上课处理）
	表中从周一到周五，‘o’为不上课，‘1’为上课；
    3、使用前请将考勤报表上传至脚本同一目录下；
    4、将文件名替换为需要统计的表格名（不带.xls），确认是否需要包括
    上课时间，点击运行即可。

Caution!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

The demo.xls is encoding by 'gbk',but the truely data I used is 'cp1252'. So , 
if u want to run the demo ... please use : 

data = xlrd.open_workbook(filename + '.xls', encoding_override='gbk') (line 43)
	and 
with codecs.open(filename + '.csv', 'w', encoding='gbk') as f: (line 47)

Thanks!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
"""

import xlrd
import pandas as pd 
import csv 
import codecs
import numpy as np

#————————使用前请输入文件名，并确认是否包括上课时间—————————————
#文件名                        
filename = 'demo'
#是否包括上课时间，1为包括       
course_enable = 0
#——————————————————————————————————————————

data = xlrd.open_workbook(filename + '.xls', encoding_override='cp1252')
table = data.sheets()[0]

#转码
with codecs.open(filename + '.csv', 'w', encoding='cp1252') as f:
    write = csv.writer(f)
    for row_num in range(table.nrows):
        row_value = table.row_values(row_num)
        write.writerow(row_value)

#打开
data1 = pd.read_csv(filename + '.csv',encoding='gbk',skiprows=0)
data2 = pd.read_csv('coursetime.csv',encoding='gbk',skiprows=0)
data3 = pd.read_csv('coursetab.csv',encoding='gbk',skiprows=0)
data1.drop(data1.index[[-1]],axis=0,inplace = True)

#替换星期为数字
data1['星期'] = data1['星期'].replace('星期一',int(1))
data1['星期'] = data1['星期'].replace('星期二',int(2))
data1['星期'] = data1['星期'].replace('星期三',int(3))
data1['星期'] = data1['星期'].replace('星期四',int(4))
data1['星期'] = data1['星期'].replace('星期五',int(5))
data1['星期'] = data1['星期'].replace('星期六',int(6))
data1['星期'] = data1['星期'].replace('星期日',int(7))

#为方便计算，将00：00变为00：01，将nan置0
data1 = data1.replace('00:00     ','00:01')
data1 = data1.replace(np.nan,'0:0')
datashape1 = data1.shape
datashape2 = data2.shape
datashape3 = data3.shape

#将小时统一转为分钟，并将dataframe转为矩阵
for i in range(5,datashape1[1]):
    for j in range (0,datashape1[0]):
        t = str(data1.iloc[j,i]).split(':',1)
        data1.iloc[j,i] = int(t[0])*60 + int(t[1])

mdata1 = data1.drop(data1.columns[[0,2,3]],axis=1)
mdata1 = mdata1.values
mdata1 = np.insert(mdata1, mdata1.shape[0], values=np.zeros(datashape1[1]-3) , axis=0)

#标记整周无记录人员
exit = data1['人员编号'].unique()
full = data2.num.size
delete = []
for i in range(0,full):
    if data2.num[i] != exit[i]:
        print (data2.iloc[i,1]+' 上周无记录！')
        delete.extend(i)
data2.drop(data2.index[delete],axis=0,inplace = True)

#将无记录日补为零向量，这样每个人的打卡时间就都是7天，方便循环
for i in range (0,data1['人员编号'].unique().size):
    for j in range (1,8):
        num = i*7 + j - 1
        if (mdata1[num][1] != j) and (num != 0):
            add = np.zeros(datashape1[1]-3)
            add[1] = j
            mdata1 = np.insert(mdata1, num, values=add, axis=0)

mdata1 = np.delete(mdata1,[0,1], axis = 1)
mdata1 = np.delete(mdata1,-1, axis = 0)

#将课时转为分钟，并将dataframe转为矩阵
for i in range(0,datashape3[1]):
    for j in range (0,datashape3[0]):
        t = str(data3.iloc[j,i]).split('∶',1)
        data3.iloc[j,i] = int(t[0])*60 + int(t[1])
mdata3 = data3.values

#将coursetime.csv中的上课时间转为矩阵，并补在打卡时间后面，这样每个人的上课、下课都算打一次卡
satsun = np.zeros(20) #周六日无课

for i in range (0,data1['人员编号'].unique().size):
    for j in range (1,8):
        num = i*7 + j - 1
        if (j!=6) and (j!=7):
            if course_enable == 1:
                x = np.array(list(str(data2.iloc[i][j+1])))
            else:
                x = ['0']*20 
            add = []
            for k in range (0,10) :
                if x[k] == '1':
                    add = np.hstack((add,mdata3[k]))
                else:
                    add = np.hstack((add,[0,0]))
            if num == 0:
                coursetime = add
            else:
                coursetime = np.vstack((coursetime,add))
        else:
            coursetime = np.vstack((coursetime,satsun))
mdata = np.hstack((mdata1,coursetime))

#每一行从小到大排序
mdata.sort(axis=1)

#本着宁放过，不杀错的原则，判断打卡情况，是否有迟到早退
for i in range (0,data1['人员编号'].unique().size):
    for j in range (1,8):
        num = i*7 + j - 1
        if (j!=6) and (j!=7):
            if np.sum((mdata[num] != 0)*(mdata[num] >= 240)*(mdata[num] <= 520)) == 0: #4:00-8:40无记录
                print (data2.iloc[i,1] + ' 星期 '+str(j)+' 上午迟到！')
            if np.sum((mdata[num] != 0)*(mdata[num] >= 680)*(mdata[num] <= 850)) == 0: #11:20-14:10无记录
                print (data2.iloc[i,1] + ' 星期 '+str(j)+' 上午早退&下午迟到！')
            if np.sum((mdata[num] != 0)*(mdata[num] >= 680)*(mdata[num] <= 850)) == 1:
                if np.sum((mdata[num] != 0)*(mdata[num] >= 680)*(mdata[num] <= 780)) == 0: #11:20-13:00无记录
                    print (data2.iloc[i,1] + ' 星期 '+str(j)+' 上午早退！')
                else:
                    print (data2.iloc[i,1] + ' 星期 '+str(j)+' 下午迟到！')
            if np.sum((mdata[num] != 0)*(mdata[num] >= 1040)) == 0: #17:20后无记录
                print (data2.iloc[i,1] + ' 星期 '+str(j)+' 下午早退！')


#本着宁多算，不少算的原则，计算打卡时间

time = pd.DataFrame(np.zeros([data1['人员编号'].unique().size,2]))
    
for i in range (0,data1['人员编号'].unique().size):
    morning = 0
    afternoon = 0
    evening = 0
    last_sunday_evening = mdata[i*7][(mdata[i*7] != 0)*(mdata[i*7] <= 240)] #上周日晚上的记录00:00-4:00
    last_sunday_eveningx = last_sunday_evening.shape[0]
    for j in range (1,8):
        num = i*7 + j - 1
        flag0 = 0
        flag1 = 0
        flag2 = 0
        if j != 7:
            ax = mdata[num][(mdata[num+1] != 0)*(mdata[num+1] <= 240)] #打卡超过零点的记录 00:00-4:00
            a = ax.shape[0]
        bx = mdata[num][(mdata[num] != 0)*(mdata[num] > 240)*(mdata[num] <= 680)] #上午打卡的记录 4:00-11:20
        b = bx.shape[0]
        cx = mdata[num][(mdata[num] != 0)*(mdata[num] > 850)*(mdata[num] <= 1040)] #下午打卡的记录 14:10-17:20
        c = cx.shape[0]
        dx = mdata[num][(mdata[num] != 0)*(mdata[num] > 1040)*(mdata[num] <= 1439)] #晚间打卡的记录 17:20-23:59
        d = dx.shape[0]
        ex = mdata[num][(mdata[num] != 0)*(mdata[num] >= 680)*(mdata[num] <= 850)] #中午打卡的记录 11:20-14:10
        e = ex.shape[0]
        if b != 0:
            #上午有打卡记录的情况
            if e == 0: #上午早退的情况 
                morning = morning +  np.max(bx) - np.min(bx)
            elif e >= 1: #中午有记录那么就上午最早减去中午最早
                morning = morning +  np.min(ex) - np.min(bx)
                flag0 = 1 #标记中午最早记录占用
                if e == 1: 
                    flag1 = 1 #标记中午唯一记录占用
        if c == 0 : 
            #下午无记录,但是中午有记录且未占用，并且晚上也有记录的情况
            if e == 1 and flag1 == 0 and d != 0: 
                afternoon = afternoon + np.min(dx) - np.max(ex)
		flag2 = 1 #标记晚上最早记录占用
            elif e > 1 and flag0 == 1 and d != 0: #如果中午被占用了一个，那么就从第二个（ex[1]）开始
                afternoon = afternoon + np.min(dx) - ex[1]
		flag2 = 1 #标记晚上最早记录占用
            elif e > 1 and flag0 == 0 and d != 0:
                afternoon = afternoon + np.min(dx) - ex[1]
            	flag2 = 1 #标记晚上最早记录占用
        else: 
            #如果下午有记录的话
            if e > 1 and d == 0 and flag1 == 0: #下午早退
                afternoon = afternoon + np.max(cx) - np.max(ex)
            elif e > 1 and d == 0 and flag1 == 1: #下午早退且中午被占用一个
                afternoon = afternoon + np.max(cx) - np.max(ex)
            elif e > 1 and d != 0 and flag1 == 0: #上午没打卡的情况
                afternoon = afternoon + np.max(dx) - np.max(ex)
                flag2 = 1
            elif e > 1 and d != 0 and flag1 == 1: #最正常的情况
                afternoon = afternoon + np.max(dx) - np.max(ex)
                flag2 = 1
            elif e == 1 and d == 0 and flag1 == 0 : #上午没打卡的情况
                afternoon = afternoon + np.max(cx) - np.max(ex)
            elif e == 0 and d == 0: #下午迟到又早退的情况
                afternoon = afternoon + np.max(cx) - np.min(cx)
            elif e == 0 and d != 0: #中午没打卡的情况
                afternoon = afternoon + np.max(dx) - np.min(cx)
                flag2 = 1
            
        if a != 0: 
            #打卡超过零点的情况
            evening2 = np.max(ax)
        if d != 0: 
            #晚上有记录
            if a != 0 and flag2 == 0: #过零点打卡的情况
                evening = evening + 1439 - np.min(dx) + evening2
            if a != 0 and flag2 == 1 and d != 1: #过零点打卡且晚上第一条记录被占用的情况
                evening = evening + 1439 - dx[1] + evening2
            if a == 0 and flag2 == 0: #未过零点打卡的情况
                evening = evening + np.max(dx) - np.min(dx)
            if a == 0 and flag2 == 1 and d != 1: #未过零点打卡的情况且晚上第一条记录被占用的情况
                evening = evening + np.max(dx) - dx[1]
    if last_sunday_eveningx != 0:
        time.iloc[i,1] = round(((morning + afternoon + evening + np.max(last_sunday_evening))/60),3)
    else:
        time.iloc[i,1] = round(((morning + afternoon + evening)/60),3)
    time.iloc[i,0] = data2.iloc[i,1]
    
time.sort_values(by=1,inplace=True,ascending= False)
time = time.reset_index()
time = time.drop(time.columns[[0]],axis=1)
print(time)
