import xlrd

workbook = xlrd.open_workbook("test.xlsx")
 
sheet = workbook.sheet_by_index(0)
y15=[]
y16=[]
y17=[]
y18=[]
y19=[]
y20=[]
# Required variables
sheet.cell_value(1,0)
workbook_datemode = workbook.datemode
for i in range(1,sheet.nrows):
    bb=sheet.row_values(i)
    y, m, d, h, mi, s = xlrd.xldate_as_tuple(bb[0], workbook_datemode)
    if(y==2017):
        y17.append(bb[1])
    elif(y==2018):
        y18.append(bb[1])
    elif(y==2019):
        y19.append(bb[1])
    elif(y==2020):
        y20.append(bb[1])
    elif(y==2016):
        y16.append(bb[1])
    elif(y==2015):
        y15.append(bb[1])
s1=[]
s2=[]
s3=[]
s4=[]
s5=[]
s0=[]
s1.append(y15.count(0.0))
s1.append(y15.count(1.0))
s1.append(y15.count(2.0))
s1.append(y15.count(3.0))
s1.append(y15.count(4.0))
s1.append(y15.count(5.0))
s2.append(y16.count(0.0))
s2.append(y16.count(1.0))
s2.append(y16.count(2.0))
s2.append(y16.count(3.0))
s2.append(y16.count(4.0))
s2.append(y16.count(5.0))
s3.append(y17.count(0.0))
s3.append(y17.count(1.0))
s3.append(y17.count(2.0))
s3.append(y17.count(3.0))
s3.append(y17.count(4.0))
s3.append(y17.count(5.0))
s4.append(y18.count(0.0))
s4.append(y18.count(1.0))
s4.append(y18.count(2.0))
s4.append(y18.count(3.0))
s4.append(y18.count(4.0))
s4.append(y18.count(5.0))
s5.append(y19.count(0.0))
s5.append(y19.count(1.0))
s5.append(y19.count(2.0))
s5.append(y19.count(3.0))
s5.append(y19.count(4.0))
s5.append(y19.count(5.0))
s0.append(y20.count(0.0))
s0.append(y20.count(1.0))
s0.append(y20.count(2.0))
s0.append(y20.count(3.0))
s0.append(y20.count(4.0))
s0.append(y20.count(5.0))
print(s1,s2,s3,s4,s5,s0)
import numpy as np
import matplotlib.pyplot as plt
barWidth=0.15
fig=plt.subplots(figsize=(24,16))
br1=np.arange(6)
br2=[x+barWidth for x in br1]
br3=[x+barWidth for x in br2]
br4=[x+barWidth for x in br3]
br5=[x+barWidth for x in br4]
br6=[x+barWidth for x in br5]
plt.bar(br1,s1,color='r',width=barWidth,
        edgecolor='grey',label='2015')
plt.bar(br2,s2,color='g',width=barWidth,
        edgecolor='grey',label='2016')
plt.bar(br3,s3,color='b',width=barWidth,
        edgecolor='grey',label='2017')
plt.bar(br4,s4,color='y',width=barWidth,
        edgecolor='grey',label='2018')
plt.bar(br5,s1,color='c',width=barWidth,
        edgecolor='grey',label='2019')
plt.bar(br6,s0,color='k',width=barWidth,
        edgecolor='grey',label='2020')
plt.xlabel('Star',fontweight='bold')
plt.ylabel('No. of Stars',fontweight='bold')
plt.xticks([r+barWidth for r in range(len(s1))],['0','1','2','3','4','5'])
plt.legend()
plt.show()
