#!/usr/bin/env python
# coding: utf-8

# ###  La méthode de King

# In[1]:


from calendar import c
from ctypes.wintypes import LONG
from pickletools import long1
from subprocess import CalledProcessError
import numpy as np
import pandas as pd
from openpyxl import load_workbook
import itertools
import os


from pyparsing import Or


# #### Exemple : Atelier de 7 machines pour usiner 8 pieces

# #### Dans cette partie on va importer les données d'un fichier EXEL

# In[2]:


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


# In[3]:


""""# Matrice des machines
Machine=np.array(['M1','M2','M3','M4','M5','M6','M7'])

# Matrices de Pieces
Piece=np.array(['P1','P2','P3','P4','P5','P6','P7'])


# Matrice de piece associe au machine,aij = 1 si la machine i usine la pièce j, 0 sinon.

PM=np.array([ [1,0,0,0,0,0,1],
              [0,1,0,0,1,1,0],
              [1,0,1,1,0,0,0],
              [0,0,1,1,1,0,0],
              [0,0,0,1,0,0,1],
              [0,1,1,0,1,0,0],
              [0,1,0,0,1,1,0]])"""


# #### Cette fonction permet de calculer les poids sur les lignes

# In[4]:


def POIDLIG(pm,machine,piece):
    V=np.zeros(machine.shape)
    S=np.zeros(piece.shape)
    for i in range(len(machine)):
        V[i]=2**(len(machine)-i-1)
    for i in range(len(piece)):
        for j in range(len(machine)):
            S[i]+=pm[i][j]*V[j]
    
    print("V=",V)
    print("S=",S)
    return S,V


# #### Cette fonction permet de trier la matrice PM selon les poids des pieces dans l'ordre decroissant

# In[5]:


def TRIP(PM,P,s):
    PM1=PM.T.copy()
    Piece=P.copy()
    s1=s.copy()
    PM1=np.append(PM1,[s1],axis=0)
    PM1=PM1[:,np.argsort(-1*PM1[len(Machine)])]
    PM1=np.delete(PM1,len(Machine),axis=0)
    return PM1.T


# #### Cette fonction permet de calculer les poids sur les colonnes

# In[6]:


def POIDCOL(pm,machine,piece):
    W=np.zeros(piece.shape)
    T=np.zeros(machine.shape)
    for i in range(len(piece)):
        W[i]=2**(len(piece)-i-1)
    for i in range(len(machine)):
        for j in range(len(piece)):
            T[i]+=pm[j][i]*W[j]
    
    print("W=",W)
    print("T=",T)
    return W,T


# #### Cette fonction permet de trier la matrice PM selon les poids des machines dans l'ordre decroissant

# In[7]:


def TRIM(PM,M,t):
    PM1=PM.copy()
    Machine=M.copy()
    t1=t.copy()
    PM1=np.append(PM1,[t1],axis=0)
    PM1=PM1[:,np.argsort(-1*PM1[len(Piece)])]
    PM1=np.delete(PM1,len(Piece),axis=0)
    return PM1


# #### Cette fonction permet de trier les liste des machines et des pieces selon les poids

# In[8]:


def tri_selection(t,t1):
    tab=t.copy()
    tab1=t1.copy()
    for i in range(len(tab)):
        # Trouver le min
        min = i
        for j in range(i+1, len(tab)):
            if tab[min] <= tab[j]:
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


# #### Dans cette fonction main on va repeter les etapes précedantes jusqu'a la no mise a jour de la matrice PM

# In[9]:


def main_king(pm,Machine,Piece):
    i = 0
    PM = np.longlong
    PM=pm.copy()
    print("\nMachine : \n",Machine)
    print("\nPiece   : \n",Piece)
    print("\nPM      : \n",PM)
    p1=Piece
    m1=Machine
    Bool=True
    try : 
        while(Bool):
            i+= 1 
            k1=PM
            print("\nTRI SUR LIGNE\n")
            #TRI SUR LIGNES
            s,v=POIDLIG(PM,Machine,Piece)
            s1,p1=tri_selection(s,p1)
            print("\nS Trier :",s1,"\nPiece Trier :",p1)
            PM=TRIP(PM,Piece,s)
            print("PM :\n",PM)
            print("\n TRI SUR COLONNE")
            #TRI SUR COLONNES
            w,t=POIDCOL(PM,Machine,Piece)
            t1,m1=tri_selection(t,m1)
            print("\nT Trier :",t1,"\nMachine Trier :",m1)
            PM=TRIM(PM,Machine,t)
            print("\nPM : \n",PM)
            k2=PM
            
            if((k1==k2).all()) :
                print("NO UPDATE!")
                Bool=False
            else:
                
                print("UPDATED!")
                continue

    except KeyboardInterrupt:
        pass
    return PM,p1,m1


# In[10]:



PM1,P1,M1=main_king(PM,Machine,Piece)
print("PM1 :\n",PM1,"\nP1 :\n",P1,"\nM1 :\n",M1)


# #### Dans cette partie on va expoter la matrice triée dans le fichier EXEL

# In[11]:


df1=pd.DataFrame(PM1)
#print(df2)
machine=[]
piece=[]
for i in range(len(M1)):
    machine.append(M1[i])
#print(machine)
for j in range(len(P1)):
    piece.append(P1[j])
#print(piece)
df1.columns=machine
df1.index=piece

#EXPORT TO EXEL
book = load_workbook("RESULTAT_KING.xlsx")
writer = pd.ExcelWriter("RESULTAT_KING.xlsx", engine='openpyxl')
del book['DATA']
del book['RESULTAT']
writer.book = book
writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
df.to_excel(writer, "DATA",index=False)
df1.to_excel(writer, "RESULTAT", columns=machine)
writer.save()
#Open Exel file
os.system('start excel.exe "%s"' % ("RESULTAT_KING.xlsx", ))
#df1.to_excel(r'C:\Users\bouza\Desktop\cour\semestre 5\Conception & Performance des systèmes de\PROJET\RESULTAT_KING.xlsx', sheet_name='RESULTAT')


# In[ ]:





# In[ ]:
