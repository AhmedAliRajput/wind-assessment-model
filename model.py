import numpy as np 
import matplotlib.pyplot as plt
import math 
import xlrd
import xlsxwriter 
wb=xlrd.open_workbook(r"C:\Users\Kakarotto\Desktop\sadqabad.xls")
sheet=wb.sheet_by_index(0)
workbook   = xlsxwriter.Workbook("wind_out.xls")
wsheet = workbook.add_worksheet()
months=['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']
month=int(input("enter month numnber"))
v2=[]
for row in range(sheet.nrows):
    month1=int(sheet.cell_value(row,0))
    if month1==month:
        v2.append(sheet.cell_value(row,1))
temp=0
f=[]
F=[]
v22=np.linspace(min(v2),max(v2),10)
v=[]
for i in range(0,len(v22)):
    sum=0
    for j in range(0,len(v2)):
        if v2[j]<=v22[i]:
            sum=sum+1
    if (sum-temp)!=0:
        f.append(sum-temp)
        F.append(sum)
        v.append(v22[i])
        temp=sum
f=np.array(f)
#v22=np.array([2,2,3,3,3,3,4,4,4,4,4,4,5,5,5,5,5,5,5,5,6,6,6,6,6,7,7,7,7,8,8])
#v=np.array([2,3,4,5,6,7,8])
#f=np.array([2,4,6,8,5,4,2])

F=np.round(f/np.sum(f),6)
maxf=max(F)
MSE,MABE,MAPE,CHI,R2=[],[],[],[],[]
'======================================================='
'Statistical Error'
'======================================================='
def SE(k,c):
    
    W=((k/c)*np.power((v/c),k-1)*np.exp(-np.power((v/c),k)))
    F1=F/maxf
    MSE=np.average(np.power(W-F1,2))**0.5
    MABE=np.average(np.abs(W-F1))
    MAPE=np.average((W-F1)*100*np.power(F1,-1))
    chi=np.average(np.abs(W-F1)**2*np.power(F1,-1))
    R2=((len(F)*np.sum(W*F1)-np.sum(W)*np.sum(F1))/(len(F1)*np.sum(F1*F1)-np.sum(F1)**2))
    sum=0
    for i in range(0,len(F)):
        sum=sum+np.abs((F[i]-((k/c)*np.power((v[i]/c),k-1)*np.exp(-np.power((v[i]/c),k))))/F[i])
        MAPE=sum/len(F)
    return [MSE,MABE,MAPE,chi,R2,sum]
v3=np.linspace(0,max(v2)+2, 200)
'======================================================='
'define function Weibull distribution'
'======================================================='
def W(k,c):
    W=((k/c)*np.power((v3/c),k-1)*np.exp(-np.power((v3/c),k)))
    return W
#'======================================================='
#'Graphical method'
#'======================================================='
#
#y=np.array([])
#v1=np.array([])
# 
#temp1=0
#for i in range(0,len(F)):
#    temp1=temp1+F[i]
#    if (temp1)<1:
#        temp=math.log(-(np.log(1-temp1)))
#        y=np.append(y,temp)
#        v1=np.append(v1, math.log(v[i]))
#
#k=(len(y)*np.sum(v1*y)-np.sum(y)*np.sum(v1))/(len(y)*np.sum(v1*v1)-np.sum(v1)**2)
#intercept=-(np.sum(y)-k*np.sum(v1))/len(v1)
#c=np.exp(intercept/k)
#y1=W(k,c)
#MSE.append(SE(k,c)[0])
#MABE.append(SE(k,c)[1])
#MAPE.append(SE(k,c)[2])
#CHI.append(SE(k,c)[3])
#wsheet.write(1,1,k)
#wsheet.write(2,1,c)
#wsheet.write(3,1,SE(k,c)[0])
#wsheet.write(4,1,SE(k,c)[1])
#wsheet.write(5,1,SE(k,c)[2])
#wsheet.write(6,1,SE(k,c)[3])
#wsheet.write(7,1,SE(k,c)[4])
#print('Graphical method',k,c)
#%%
'======================================================='
'Empirical method'
'======================================================='
k=((np.sum(F*(v-np.dot(F,v))**2))**0.5/np.sum(F*v))**-1.086
#k=(np.std(v2)/np.average(v2))**-1.086
c=np.dot(v,F)/math.gamma(1+1/k)
y2=W(k,c)
MSE.append(SE(k,c)[0])
MABE.append(SE(k,c)[1])
MAPE.append(SE(k,c)[2])
CHI.append(SE(k,c)[3])
wsheet.write(1,2,k)
wsheet.write(2,2,c)
wsheet.write(3,2,SE(k,c)[0])
wsheet.write(4,2,SE(k,c)[1])
wsheet.write(5,2,SE(k,c)[2])
wsheet.write(6,2,SE(k,c)[3])
wsheet.write(7,2,SE(k,c)[4])

print('Empirical method',k,c)
'======================================================='
'Method of moments'
'======================================================='
x=np.linspace(1,10,1000)
temp=[]
for i in range(0,len(x)):
    k=x[i]
    temp.append(abs(((np.sum(F*(v-np.dot(F,v))**2))**0.5/np.dot(v,F))-(math.gamma(1+2/k)-math.gamma(1+1/k)**2)**0.5/math.gamma(1+1/k)))
n=temp.index(min(temp))
k=x[n]
c=np.dot(v,F)/math.gamma(1+1/k)
y3=W(k,c)
MSE.append(SE(k,c)[0])
MABE.append(SE(k,c)[1])
MAPE.append(SE(k,c)[2])
CHI.append(SE(k,c)[3])
wsheet.write(1,3,k)
wsheet.write(2,3,c)
wsheet.write(3,3,SE(k,c)[0])
wsheet.write(4,3,SE(k,c)[1])
wsheet.write(5,3,SE(k,c)[2])
wsheet.write(6,3,SE(k,c)[3])
wsheet.write(7,3,SE(k,c)[4])

print('Method of moments',k,c)

'======================================================='
'energy pattern factor method'
'======================================================='
k=1+3.69/(np.dot(F,np.power(v,3))/np.dot(v,F)**3)**2
c=np.dot(v,F)/math.gamma(1+1/k)
y4=W(k,c)
MSE.append(SE(k,c)[0])
MABE.append(SE(k,c)[1])
MAPE.append(SE(k,c)[2])
CHI.append(SE(k,c)[3])
wsheet.write(1,4,k)
wsheet.write(2,4,c)
wsheet.write(3,4,SE(k,c)[0])
wsheet.write(4,4,SE(k,c)[1])
wsheet.write(5,4,SE(k,c)[2])
wsheet.write(6,4,SE(k,c)[3])
wsheet.write(7,4,SE(k,c)[4])

print('energy pattern factor method',k,c)
'======================================================='
'Maximum likelihood Method'
'======================================================='
x=np.linspace(2,10,1000)
temp=[]
for i in range(0,len(x)):
    k=x[i]
    temp.append(abs(k-(np.sum(F*np.power(v,k)*np.log(v))/np.sum(F*np.power(v,k))-np.sum(F*np.log(v)))**-1))
n=temp.index(min(temp))
k=x[n]
c=(np.sum(F*np.power(v,k)))**(1/k)
y5=W(k,c)
MSE.append(SE(k,c)[0])
MABE.append(SE(k,c)[1])
MAPE.append(SE(k,c)[2])
CHI.append(SE(k,c)[3])
wsheet.write(1,5,k)
wsheet.write(2,5,c)
wsheet.write(3,5,SE(k,c)[0])
wsheet.write(4,5,SE(k,c)[1])
wsheet.write(5,5,SE(k,c)[2])
wsheet.write(6,5,SE(k,c)[3])
wsheet.write(7,5,SE(k,c)[4])

print('Maximum likelihood Method',k,c)

'======================================================='
' Modified maximum likelihood method'
'======================================================='
x=np.linspace(2,10,10000)
temp=[]
f=f/np.sum(f)
for i in range(0,len(x)):
    k=x[i]
    temp.append(abs(k-(np.sum(np.power(v,k)*np.log(v)*f)/np.sum(np.power(v,k)*f)-np.sum(np.log(v)*f))**-1))
n=temp.index(min(temp))
k=x[n]
c=(np.dot(F,np.power(v,k))**(1/k))
y6=W(k,c)
MSE.append(SE(k,c)[0])
MABE.append(SE(k,c)[1])
MAPE.append(SE(k,c)[2])
CHI.append(SE(k,c)[3])

wsheet.write(1,6,k)
wsheet.write(2,6,c)
wsheet.write(3,6,SE(k,c)[0])
wsheet.write(4,6,SE(k,c)[1])
wsheet.write(5,6,SE(k,c)[2])
wsheet.write(6,6,SE(k,c)[3])
wsheet.write(7,6,SE(k,c)[4])

print('Modified Maximum likelihood Method',k,c)
#'======================================================='
#'Equivalent Energy Method'
#'======================================================='
#x=np.linspace(2,10,10000)
#EEMsum=[]
#for j in range(0,len(x)):
#    k=x[j]
#    temp=[]
#    for i in range(0,len(F)):
#        if (v[i]>1):
#            c=(np.average(np.power(v,1))/math.gamma(1+1/k))
#            temp.append((F[i]-np.exp(-np.power((v[i]-1)/c,k))+np.exp(-np.power(v[i]/c,k)))**2)
#    EEMsum.append(np.sum(temp))
#n=EEMsum.index(min(EEMsum))
#k=x[n]
#c=(np.average(v)/math.gamma(1+1/k))
#y7=W(k,c)
#MSE.append(SE(k,c)[0])
#MABE.append(SE(k,c)[1])
#MAPE.append(SE(k,c)[2])
#CHI.append(SE(k,c)[3])
#wsheet.write(1,7,k)
#wsheet.write(2,7,c)
#wsheet.write(3,7,SE(k,c)[0])
#wsheet.write(4,7,SE(k,c)[1])
#wsheet.write(5,7,SE(k,c)[2])
#wsheet.write(6,7,SE(k,c)[3])
#wsheet.write(7,7,SE(k,c)[4])
#print('Equivalent Energy Method',k,c)
'======================================================='
' Plotting Graphs'
'======================================================='
#plt.plot(v3,y1,label='graphical')
plt.plot(v3,y2,label='EM')
plt.plot(v3,y3,label='MoM')
plt.plot(v3,y4,label='EPFM')
plt.plot(v3,y5,label='MLM')
plt.plot(v3,y6,label='MMLM')
#plt.plot(v3,y7,label='EEM')
maxf=max(f)+0.01
f=f*max(y3)/maxf   
plt.bar(v,f,v[1]-v[0]-0.1)
plt.legend()
plt.xlabel('Wind Speed  (m/s)')
plt.ylabel('Relative Probability')
month1=['January','February','March','April','May']
plt.title(months[month-1])
wsheet.write(0,1,'GM')
wsheet.write(0,2,'EPM')
wsheet.write(0,3,'MoM')
wsheet.write(0,4,'EPFM')
wsheet.write(0,5,'MLM')
wsheet.write(0,6,'MMLM')
wsheet.write(0,7,'EEM')
plt.xlim(0,20)
wsheet.write(1,0,'k')
wsheet.write(2,0,'c')
wsheet.write(3,0,'RMSE')
wsheet.write(4,0,'MABE')
wsheet.write(5,0,'MAPE')
wsheet.write(6,0,'CHI')
wsheet.write(7,0,'R2')
workbook.close()
path="C:/Users/Zaheer\Documents/Research/MSC 2021 thesis group/New folder/Karachi/K"+str(month)+".jpeg"
#plt.savefig(path)
