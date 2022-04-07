#!/usr/bin/env python
# coding: utf-8

# In[1]:


import numpy as np
import pandas as pd
from openpyxl import load_workbook
import itertools
import os



# In[2]:


#importation de données du ficher Exel
df = pd.read_excel (r"DATA.xlsm",sheet_name="DATA3") 
#Liste des MAchines
Machine=df.columns.to_numpy()
Machine=np.delete(Machine,0)
print("\nMachine : ",Machine)
#Liste des Pieces
Piece=df.to_numpy()[:,0]
print("\n Piece :",Piece)
#Relation Machines-Pieces
PM=df.fillna(0).to_numpy()
PM=np.delete(PM,0,1)
PM = PM.astype('int')
print ("\nPM :\n",PM)

def RM(PM,M,P):
    T1=np.zeros(M.shape,dtype=int)
    T2=np.zeros(M.shape,dtype=int)
    T3=np.zeros(M.shape)
    for i in range(len(M)):
        for j in range(len(P)):
            T1[i]+=PM[j][i]
    print("TOTAL des rangs \n",T1)
    for i in range(len(M)):
        for j in range(len(P)):
            if(PM[j][i]!=0):
                 T2[i]+=1
    print("Nombre de rangs \n",T2)
    T3=T1/T2
    print("Rang Moyen :\n",T3)
    PM=np.append(PM,[T1],axis=0)
    PM=np.append(PM,[T2],axis=0)
    PM=np.append(PM,[T3],axis=0)
    print("PM :\n",PM)
    return PM,T3
    
pm,t3=RM(PM,Machine,Piece)

pm=pm[:,np.argsort(t3)]
print(pm)

def tri_selection(t,t1):
    tab=t.copy()
    tab1=t1.copy()
    for i in range(len(tab)):
        # Trouver le min
        min = i
        for j in range(i+1, len(tab)):
            if tab[min] >= tab[j]:
                min = j   
                
        tmp = tab[i]
        tmp1=tab1[i]
        #tmp2=pm1[i]
        
        tab[i] = tab[min]
        tab1[i]=tab1[min]
        #pm1[i]=pm1[min]
        
        tab[min] = tmp
        tab1[min]=tmp1
        #pm1[min]=tmp2
        
    return tab,tab1

t,machine=tri_selection(t3,Machine)
print("Machine",machine)
print(t)
ML=[]
i=0

while(i<=(len(t)-1)):
    k=1
    ml=[]
    for j in range(i+1,len(t)):
        if(t[i]==t[j]):
            k+=1
    for l in range(k):
        ml.append(machine[i])
        i+=1
    ML.append(ml)
    
print(ML)
df2=pd.DataFrame(ML)
print(df2.T)
# #### Dans cette partie on va expoter la matrice triée dans le fichier EXEL

# In[11]:


#Creation de la matrice trier 
df1=pd.DataFrame(pm)
#print(df1)
#print(piece)
df1.columns=machine
df1.index=np.append(Piece,['TR','NR','RM'])
print(df1)


#EXPORT TO EXEL
book = load_workbook('RESULTAT_RM.xlsx')
writer = pd.ExcelWriter('RESULTAT_RM.xlsx', engine='openpyxl')
del book['DATA']
del book['RESULTAT']
writer.book = book
writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
df1.to_excel(writer, 'DATA')
df2.T.to_excel(writer, "RESULTAT")
writer.save()
#Open Exel file
os.system('start excel.exe "%s"' % ("RESULTAT_RM.xlsx", ))

#df1.to_excel(r'C:\Users\bouza\Desktop\cour\semestre 5\Conception & Performance des systèmes de\PROJET\RESULTAT_KUSIAK.xlsx', sheet_name='RESULTAT')


# In[ ]:





# In[ ]:




