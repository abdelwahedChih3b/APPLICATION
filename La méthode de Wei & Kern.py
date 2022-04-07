#!/usr/bin/env python
# coding: utf-8

# ### La méthode de Wei & Kern

# In[21]:


import numpy as np
import pandas as pd
from openpyxl import load_workbook
import itertools
import os

# #### Exemple : Atelier de 10 machines pour usiner 7 pieces

# In[22]:


#importation de données du ficher Exel
df = pd.read_excel (r"DATA.xlsm",sheet_name="DATA2") 
#Liste des MAchines
Machine=df.columns.to_numpy()
Machine=np.delete(Machine,0)
print("\nMachine : ",Machine)
#Liste des Pieces
Piece=df.to_numpy()[:,0]
print("\n Piece :",Piece)
#Relation Machines-Pieces
PM=df.to_numpy()
PM=np.delete(PM,0,1)
print ("\nPM :\n",PM)


# In[23]:


""""# Matrice des machines
Machine=np.array(['M1','M2','M3','M4','M5','M6','M7','M8','M9','M10'])

# Matrices de Pieces
Piece=np.array(['P1','P2','P3','P4','P5','P6','P7','P8'])


# Matrice de piece associe au machine,aij = 1 si la machine i usine la pièce j, 0 sinon.

PM=np.array([ [1,1,0,1,0,0,0,1,0,0],
              [0,0,1,0,0,1,0,0,0,1],
              [1,0,0,1,0,0,0,1,0,0],
              [0,0,0,0,1,0,0,0,1,0],
              [0,0,1,0,0,1,1,0,0,0],
              [1,1,0,1,0,0,0,1,0,0],
              [0,0,0,0,1,0,0,0,1,1],
              [1,0,1,0,0,1,1,0,0,0]])"""


# #### cette fonction permet de calculer la similarite entre les machines

# In[24]:


def matrice_similarite(PM,M,P):
    MS=np.zeros((M.shape[0],M.shape[0]),dtype=int)
    for i in range(len(M)):
        for j in range(i+1,len(M)):
            for k in range(len(P)):
                if(PM[k][i]==PM[k][j]):
                    if(PM[k][i]==1):
                        MS[i][j]+=(len(P)-1)
                    else:
                        MS[i][j]+=1
    print(MS)
    return MS


# In[25]:


m=matrice_similarite(PM,Machine,Piece)


# #### Cette fonction permet de generer une liste des listes des ilots des machines similaire.

# In[26]:


def main_WK(PM,M,P):
    MS=matrice_similarite(PM,M,P)
    machine=M.copy()
    print(machine)
    print("\n")
    ILOT=[]
    print(ILOT)
    #Le boucle suivant permet d'affecter la machine a l'ilot qui convient d'apres la methode de WK.
    while(len(machine)>0):
        #On cherche le max de la matrice de similarite
        k=np.max(MS)
        #print("k=",k)
        #print(MS)
        
        #on cherche les deux machines et les indices de ces machines.
        for i in range(len(M)):
            for j in range(i+1,len(M)):
                if(MS[i][j]==k):
                    machine1=M[i]
                    machine2=M[j]
                    i1=i
                    j1=j
        print("\nMachine1:",machine1,"Machine2:",machine2)
        #Si la liste des ilot est vide on affecter les deux premieres machines directement
        if(len(ILOT)==0):
            ILOT.append([machine1,machine2])
            machine=np.delete(machine,np.where(machine==('M'+str(i1+1))))
            machine=np.delete(machine,np.where(machine==('M'+str(j1+1))))
            #print(machine)
            MS[i1][j1]=0
            print(ILOT)
            continue
                  
        #la liste des ilot no vide 
        else:
            #tous ces variables sont des variables booléennes.
            x=''
            w1=''
            w11=''
            w111=''
            w2=''
            w22=''
            w222=''
            w3=''
            w33=''
            w333=''
            w3333=''
            
            #Verifier si la machine 1 existe dans les ilots et la non existance de la machine 2.
            for l in range(len(ILOT)):
                if(machine1 in ILOT[l]):
                    print("Machine1 ",machine1," Founded")
                    w1='a1'
                    z=l
       
            if(w1=='a1'):
                for m in range(len(ILOT)):
                    if(machine2 not in ILOT[m]):
                        w11='a1'
                    else:
                        w111='a2'
            if(w11=='a1' and w111!='a2'):                            
                x='a'
                #print("AAA")
                
                            
            elif(w11=='a1' and w111=='a2'):
                x='d'
                
                
            #Verifier si la machine 2 existe dans les ilots et la non existance de la machine 1.
            for l in range(len(ILOT)):
                if(machine2 in ILOT[l]):
                    print("Machine2 ",machine2," Founded")
                    w2='a1'
                    z=l
                    
                

            if(w2=='a1'):
                for m in range(len(ILOT)):
                    if(machine1 not in ILOT[m]):
                        w22='a1'
                    else:
                        w222='a2'
            if(w22=='a1' and w222!='a2'):                            
                x='b'
                #print("BBB")
                            
            elif(w22=='a1' and w222=='a2'):
                x='d'
            
            #La no existance des deux machines dans les ilots
            for l in range(len(ILOT)):
                if(machine1 not in ILOT[l] ):
                    w3='c1'
                else:
                    w33='c2'
            if(w3=='c1' and w33!='c2'):
                for m in range(len(ILOT)):
                    if(machine2 not in ILOT[m]):
                        w333='c11'
                    else:
                        w3333='c22'
            if(w333=='c11' and w3333!='c22'):
                x='c'
                #print("CCC")
            if(w33=='c2' and w3333=='c22'):
                continue
 
        #print("w1 ",w1,"w11 ",w11,"w111 ",w111,"w2 ",w2,"w22 ",w22,"w222 ",w222,"w3 ",w3,"w33 ",w33,"w333 ",w333,"w3333 ",w3333)
        print("X=",x)
        #print("Z =",z)
        
        #L'ajout des machines dans les ilot qui convient.
        if(x=='a'):
            ILOT[z].append(machine2)
            machine=np.delete(machine,np.where(machine==('M'+str(j1+1))))
            MS[i1][j1]=0
            print(ILOT)
            continue
        if(x=='b'):
            ILOT[z].append(machine1)
            #ILOT[z]=np.append(ILOT[z],machine1)
            machine=np.delete(machine,np.where(machine==('M'+str(i1+1))))
            MS[i1][j1]=0
            print(ILOT)
            continue
        if(x=='c'):
            ILOT.append([machine1,machine2])
            machine=np.delete(machine,np.where(machine==('M'+str(i1+1))))
            machine=np.delete(machine,np.where(machine==('M'+str(j1+1))))
            MS[i1][j1]=0
            print(ILOT)
            continue
        if(x=='d'):
            MS[i1][j1]=0
            #print(ILOT)
            continue
        if(x==''):
            MS[i1][j1]=0
            continue
  
    return ILOT


# In[27]:


ilots=main_WK(PM,Machine,Piece)


# In[28]:


#Creation d'une liste des machines trier selon les ilots
m=[]
for i in range(len(ilots)):
    for j in range(len(ilots[i])):
        m.append(ilots[i][j])
print(m)
#Creation d'une liste qui genere les indices des machines
machine_trier=[]
for i in range(len(m)):
    for j in range(len(Machine)):
        if(m[i]==Machine[j]):
            machine_trier.append(j+1)
print(machine_trier)


# #### Cette fonction permet de trier la matrice PM selon l'ordre des ilots

# In[29]:


def Tri(PM,T1):
    #Tri sur colonne
    PM1=PM.T
    PM2=np.zeros(PM.T.shape,dtype=int)
    for j in range(len(T1)):
        PM2[j]=PM1[T1[j]-1]
    return PM2.T


# #### Dans cette partie on va expoter la matrice triée dans le fichier EXEL

# In[30]:

#Creation de la matrice trier 
df1=pd.DataFrame(Tri(PM,machine_trier))
df1.columns=m
#Creation de la matrice des machines par ilot 
MACHINEINDEX=[]
MACHINECOL=[]
MACHINE=pd.DataFrame(ilots)
for i in range(len(ilots)):
    MACHINEINDEX.append('ILOT'+str(i+1))
MACHINE.index=MACHINEINDEX
MACHINE.columns=itertools.repeat('MACHINE',MACHINE.columns.stop)

#EXPORT TO EXEL
book = load_workbook("RESULTAT_WEI_KERN.xlsx")
writer = pd.ExcelWriter("RESULTAT_WEI_KERN.xlsx", engine='openpyxl')
del book['DATA']
del book['RESULTAT']
del book['ILOTS']
writer.book = book
writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
df.to_excel(writer, "DATA",index=False)
df1.to_excel(writer, "RESULTAT", columns=m)
MACHINE.to_excel(writer, "ILOTS")
writer.save()
#Open Exel file
os.system('start excel.exe "%s"' % ("RESULTAT_WEI_KERN.xlsx", ))
#df1.to_excel(r'C:\Users\bouza\Desktop\cour\semestre 5\Conception & Performance des systèmes de\PROJET\RESULTAT_WEI_KERN.xlsx', sheet_name='RESULTAT')


# In[ ]:




