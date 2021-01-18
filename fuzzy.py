# -*- coding: utf-8 -*-
"""
Created on Mon Nov  9 01:22:10 2020

@author: eyditye / Aditya Ramadhan
"""
import xlsxwriter
import pandas as pd
import matplotlib.pyplot as plt

#read file xls
mhs = pd.read_excel (r'C:\Users\eyditye\Documents\python\Mahasiswa.xls')

#      0 1 2 3 4    5   6    7   8   9   for tracing array index
inc = [0,5,7,8,8.5,12,12.25,13.5,14,20] #this is the boundary for function Penghasilan/Income
       # 0   1   2  3 4 5  6   7   8   9  for tracing array index
spend = [0,4.25,4.5,5,6,7,7.25,8,8.15,12] #this is the boundary for function Pengeluaran/Spending

#fuzzification function
def LowIncome(val):
    if (val<= inc[1]):
        return 1
    elif(val > inc[3]):
        return 0
    elif (val > inc[1] and val <inc[3]):
        return (inc[3]-val)/(inc[3]-inc[1])
    
def MiddleIncome(val):
    if (val <= inc[2] or val>= inc[7]):
        return 0
    elif(val > inc[2] and val<inc[4]):
        return (inc[4]-val)/(inc[4]-inc[2])
    elif (val >= inc[4] and val <=inc[6]):
        return 1
    elif (val > inc[6] and val < inc[7]):
        return (inc[7]-val)/(inc[7]-inc[5])

def UpperIncome(val):
    if (val<= inc[5]):
        return 0
    elif (val > inc[5] and val<inc[8]):
        return (inc[8]-val)/(inc[8]-inc[5])
    elif (val >= inc[8]):
        return 1
    
def SmallSpend(val):
    if (val<=spend[2] ):
        return 1
    elif (val>spend[2] and val< spend[4]):
        return (spend[4]-val)/(spend[4]-spend[2])
    elif (val>=spend[4]):
        return 0

def MedSpend(val):
    if (val<=spend[1] or val>=spend[7]):
        return 0 
    elif (val>spend[1] and val<spend[3]):
        return (spend[3]-val)/(spend[3]-spend[1])
    elif (val>=spend[3] and val <= spend[6]):
        return 1
    elif (val>spend[6] and val< spend[7]):
        return (spend[7]-val)/(spend[7]-spend[6])

def LargeSpend(val):
    if (val<=spend[5]):
        return 0
    elif(val> spend[5] and val<spend[8]):
        return (spend[8]-val)/(spend[8]-spend[5])
    elif (val >=spend[8]):
        return 1
   

#plot of boundary of income function
def IncomeInputBoundaryFunction():
    x1 = [0,5,8,20]
    y1 = [1,1 ,0,0]

    x2 = [0,7,8.5,12.25,13.5,20]
    y2 = [0,0,1,1,0,0]

    x3 = [0,12,14,20]
    y3 = [0,0,1,1]

    plt.plot(x1, y1, label='LowIncome')
    plt.plot(x2, y2, label='MiddleIncome')
    plt.plot(x3, y3, label='UpperIncome')
    plt.legend()
    
#plot of boundary of spending function
def SpendingInputBoundaryFunction():
    x1 = [0,4.5,6,12]
    y1 = [1,1 ,0,0]

    x2 = [0,4.25,5,7.25,8,12]
    y2 = [0,0,1,1,0,0]

    x3 = [0,7,8.15,12]
    y3 = [0,0,1,1]

    plt.plot(x1, y1, label='SmallSpend')
    plt.plot(x2, y2, label='MediumSpend')
    plt.plot(x3, y3, label='LargeSpend')
    plt.legend()
    

#uncomment this section to see plot of income
#IncomeInputBoundaryFunction()    

#uncomment this section to see plot of spending
#SpendingInputBoundaryFunction()    
    
    
#inference rules    
def fuzzyRules(cekinc,cekspend,valinc,valspend,id):
    score =[]
    if(LowIncome(cekinc["Penghasilan"][id])>0 and SmallSpend(cekspend["Pengeluaran"][id])>0):
        score.append(['Considered', (valinc[id][1] and valspend[id][1])])
    if(LowIncome(cekinc["Penghasilan"][id])>0 and MedSpend(cekspend["Pengeluaran"][id])>0):
        score.append(['Accepted', (valinc[id][1] and valspend[id][2])])
    if(LowIncome(cekinc["Penghasilan"][id])>0 and LargeSpend(cekspend["Pengeluaran"][id])>0):
        score.append(['Accepted', (valinc[id][1] and valspend[id][3])])
    if(MiddleIncome(cekinc["Penghasilan"][id])>0 and SmallSpend(cekspend["Pengeluaran"][id])>0):
        score.append(['Rejected', (valinc[id][2] and valspend[id][1])])
    if(MiddleIncome(cekinc["Penghasilan"][id])>0 and MedSpend(cekspend["Pengeluaran"][id])>0):
        score.append(['Considered', (valinc[id][2] and valspend[id][2])])
    if(MiddleIncome(cekinc["Penghasilan"][id])>0 and LargeSpend(cekspend["Pengeluaran"][id])>0):
        score.append(['Accepted', (valinc[id][2] and valspend[id][3])])
    if(UpperIncome(cekinc["Penghasilan"][id])>0 and SmallSpend(cekspend["Pengeluaran"][id])>0):
        score.append(['Rejected', (valinc[id][3] and valspend[id][1])])
    if(UpperIncome(cekinc["Penghasilan"][id])>0 and MedSpend(cekspend["Pengeluaran"][id])>0):
        score.append(['Rejected', (valinc[id][3] and valspend[id][2])])
    if(UpperIncome(cekinc["Penghasilan"][id])>0 and LargeSpend(cekspend["Pengeluaran"][id])>0):
        score.append(['Rejected', (valinc[id][3] and valspend[id][3])])
    return score

#defuzzification with sugeno
def defuzzification(inf):
    return ((inf[0]*100)+(inf[1]*70)+(inf[2]*50))/(inf[0]+inf[1]+inf[2])

#plot of defuzzification boundary function
def defuzboundaryfunction():
    plt.axvline(x=50, color='green', label='rejected')
    plt.axvline(x=70, color='blue', label='considered')
    plt.axvline(x=100, color = 'red', label='accepted')
    plt.legend()

#uncomment this section to see plot of defuzzification
#defuzboundaryfunction()

#main function starts here ! 

valinc = [] #later used for fuzzyrules function
for i in range(len(mhs)):
    value = []
    value.append(mhs["Id"][i])
    value.append(LowIncome(mhs["Penghasilan"][i]))
    value.append(MiddleIncome(mhs["Penghasilan"][i]))
    value.append(UpperIncome(mhs["Penghasilan"][i]))
    valinc.append(value)

valspend = [] #later used for fuzzyrules function
for i in range(len(mhs)):
    value = []
    value.append(mhs["Id"][i])
    value.append(SmallSpend(mhs["Pengeluaran"][i]))
    value.append(MedSpend(mhs["Pengeluaran"][i]))
    value.append(LargeSpend(mhs["Pengeluaran"][i]))
    valspend.append(value) 
    
defuzz=[] #array for containing defuzzification with the id
cekinc=mhs #variable used for fuzzyrules function
cekspend=mhs #variable used for fuzzyrules function

#this is the main main program
for i in range (len(mhs)):
    temp = fuzzyRules(cekinc,cekspend,valinc,valspend,i)
    infern=[]
    ac=[]
    con=[]
    rej=[]
    for j in range (len(temp)):
        if (temp[j][0] == 'Accepted'):
            ac.append(temp[j][1])
        if (temp[j][0] == 'Considered'):
            con.append(temp[j][1])
        if (temp[j][0] == 'Rejected'):
            rej.append(temp[j][1])
    if (len(ac)==0):
        infern.append(0)
    else:
        infern.append(max(ac))
    if (len(con)==0):
        infern.append(0)
    else:
        infern.append(max(con))
    if (len(rej)==0):
        infern.append(0)
    else:
        infern.append(max(rej))
    defuzz.append([defuzzification(infern),i+1])

defuzz.sort(reverse=True)
registsupport=[]
for i in range(20):
    registsupport.append(defuzz[i][1])

print("Id yang mendapat bantuan" ,registsupport)

#writing the output to xls as Bantuan.xls
outWorkBook= xlsxwriter.Workbook("Bantuan.xls")
outSheet = outWorkBook.add_worksheet()
row=0
column=0
for item in registsupport:
    outSheet.write(row,column,item)
    row+=1

outWorkBook.close()



