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

## fonction 
def ANT(PM,M,P):
    L=[]
    for i in range(len(M)):
        ml=np.array([],dtype=int)
        for j in range(len(P)):
            k=PM[j][i]
            #print("k ",k)
            for l in range(1,k):
                for m in range(len(M)):
                    if(l==PM[j][m]):
                        #print("m ",m)
                        ml=np.append(ml,int(m))
        ml=np.unique(ml)
        #print('\n')
        #print("ml ",ml)
        L.append(ml+1)
    for i in range(len(L)):
        print("Machine",i+1,":",L[i])
    return L

ANT=ANT(PM,Machine,Piece)

def boucle(ANT):
    l=np.array([],dtype=int)
    for i in range(len(ANT)):
        for j in range(len(ANT[i])):
            x=ANT[i][j]
            for k in range(len(ANT[x-1])):
                if(ANT[x-1][k]==i+1):
                    l=np.append(l,i+1)
    
    l=np.unique(l)
    print(l)
    return l


machine=Machine.copy()
#print(machine)
#RAY(ANT,machine)
L=[]
while(len(machine)>=1):
    l=np.array([],dtype=str)
    comp=np.array([],dtype=int)
    k2=0
    k1=len(machine)
    if(k1!=k2):
        for i in range(len(ANT)):
            if(len(ANT[i])==0 and Machine[i] in machine):
                l=np.append(l,Machine[i])
        print("l ;",l)
        if(len(l)!=0):
            L.append(l)
        #l=np.array([],dtype=str)
        for i in range(len(l)):
            for j in range(len(ANT)):
                if(l[i]==Machine[j]):
                    comp=np.append(comp,j+1)
                    print("COMP :",comp)
        for i in (comp):
            machine=np.delete(machine,np.where(machine=='M'+str(i)))
            for j in range(len(ANT)):
                ANT[j]=np.delete(ANT[j],np.where(ANT[j]==i))
    k2=len(machine)
    #Machine on boucle
    
    if(k1==k2):
        S=boucle(ANT)
        B=[S[0],S[1]]
        print("B ",B)
        for i in (B):
            l=np.append(l,Machine[i-1])
        print("l ;",l)
        if(len(l)!=0):
            L.append(l)
        #l=np.array([],dtype=str)
        for i in range(len(l)):
            for j in range(len(ANT)):
                if(l[i]==Machine[j]):
                    comp=np.append(comp,j+1)
                    print("COMP :",comp)
        for i in (comp):
            machine=np.delete(machine,np.where(machine=='M'+str(i)))
            for j in range(len(ANT)):
                ANT[j]=np.delete(ANT[j],np.where(ANT[j]==i))
    
    print("L :",L)
    print("machine ",machine)
    for i in range(len(ANT)):
        print("ANT ",i+1,":",ANT[i])

                    
                       


# #### Dans cette partie on va expoter la matrice triée dans le fichier EXEL

# In[11]:


#Creation de la matrice trier 
df1=pd.DataFrame(L)
print(df1.T)



#EXPORT TO EXEL
book = load_workbook('RESULTAT_ANT.xlsx')
writer = pd.ExcelWriter('RESULTAT_ANT.xlsx', engine='openpyxl')
del book['RESULTAT']
writer.book = book
writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
df1.T.to_excel(writer, "RESULTAT")
writer.save()
#Open Exel file
os.system('start excel.exe "%s"' % ("RESULTAT_ANT.xlsx", ))

#df1.to_excel(r'C:\Users\bouza\Desktop\cour\semestre 5\Conception & Performance des systèmes de\PROJET\RESULTAT_KUSIAK.xlsx', sheet_name='RESULTAT')


# In[ ]:





# In[ ]:




