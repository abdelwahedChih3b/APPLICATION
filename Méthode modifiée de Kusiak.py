#!/usr/bin/env python
# coding: utf-8

# In[75]:


import numpy as np
import pandas as pd
from openpyxl import load_workbook
import itertools
import os

# #### Exemple : Atelier de 7 machines pour usiner 7 pieces

# #### Dans cette partie on va importer les données d'un fichier EXEL
# 

# In[76]:


#importation de données du ficher Exel
df = pd.read_excel (r"DATA.xlsm",sheet_name="DATA2") 
#Liste des MAchines
Machine1=df.columns.to_numpy()
Machine1=np.delete(Machine1,0)
print("\nMachine : ",Machine1)
#Liste des Pieces
Piece1=df.to_numpy()[:,0]
print("\n Piece :",Piece1)
#Relation Machines-Pieces
PM1=df.to_numpy()
PM1=np.delete(PM1,0,1)
print ("\nPM :\n",PM1)


# In[77]:


""""# Matrice des machines
Machine1=np.array(['M1','M2','M3','M4','M5','M6','M7'])

# Matrices de Pieces
Piece1=np.array(['P1','P2','P3','P4','P5','P6','P7'])


# Matrice de piece associe au machine,aij = 1 si la machine i usine la pièce j, 0 sinon.

PM1=np.array([[0,1,0,0,1,0,0],
              [0,0,0,1,0,1,0],
              [0,0,1,1,1,0,0],
              [1,0,0,0,0,0,1],
              [0,1,0,0,1,0,0],
              [0,0,0,1,0,1,0],
              [0,1,1,0,0,0,0]])"""


# #### Cette fonction permet de parcourir la premier ligne (Machine 1) et extraire les "1" rencontrée.
# 

# In[78]:


def Int1(PM,K):
    M=np.array([],dtype=int)
    for j in range(len(Machine1)):
        if(PM[K][j]==1):
            M=np.append(M,int(j))
    print("\nMachine :\n")
    print(M+1)
    return M


# #### Cette fonction permet de parcourir les Colonnes au "1" recontrées par la fonction precédente.

# In[79]:


def PIECE1(PM,M):
    P=np.array([],dtype=int)
    for j in M:
        for i in range(len(Piece1)):
            #print(PM1[i][int(j)])
            if(PM[i][int(j)]==1):
                P=np.append(P,int(i))
    P=np.unique(P)
    print("\nPiece :\n")
    print(P+1)
    return P


# #### Cette fonction permet de parcourir la liste des pieces recontrées par la fonction precédente avec la satisfaction de contraine de 50% .

# In[80]:


def PIECE50(PM,M,P):
    P1=np.array([],dtype=int)
    for i in P:
        k=0
        for j in M:
            if(PM[int(i)][int(j)]==1):
                #print("PM1[int(",i,")][int(,",j,",)]=",PM1[int(i)][int(j)])
                k+=1
            #print("peice ",i,"k=",k)
        #print("SUm",np.sum(PM1[int(i)]))
        if(k/np.sum(PM1[int(i)])>= 0.5):
            #print("OK")
            P1=np.append(P1,i)
            continue
    P1=np.unique(P1)
    print("\nPiece qui satisfait la contrainte de 50% :\n")
    print(P1+1)
    return P1


# #### Cette fonction permet de parcourir les lignes au "1" recontrées par la fonction precédente.

# In[81]:


def MACHINE1(PM,M,P):
    for i in P:
        for j in range(len(Machine1)):
            if(PM[int(i)][j]==1):
                M=np.append(M,int(j))
    M=np.unique(M)
    print("\nMachine :\n")
    print(M+1)
    return M


# #### Cette fonction permet de parcourir la liste des machines recontrées par la fonction precédente avec la satisfaction de contraine de 50% .

# In[82]:


def MACHINE50(PM,M,P):
    M1=np.array([],dtype=int)
    for i in M:
        SUM=0
        k=0
        for j in P:
            if(PM[int(j)][int(i)]==1):
                k+=1
            #print("Machine ",i+1,"k=",k)
        for l in range (len(Piece1)):
            #print(PM1[l][int(i)])
            SUM+=PM1[l][int(i)]
            #print("SUm=",SUM)
        #print("k/SUM=",k/SUM)
        if(k/SUM>= 0.5):
            #print("\nOK\n")
            M1=np.append(M1,i)
            continue
    M1=np.unique(M1)
    print("\nMachine qui satisfait la contrainte de 50% :\n")
    print(M1+1)
    return M1


# #### Cette fonction prend come input la matrice PM est result comme output 2 tableau (ILOT(Machine,Piece) + 2 tableau  machine,piece sans les machines,les pieces allouées.

# In[83]:


def ILOT1(PM,piece,machine):
    # Deux variable pour comparait les tailles des matrices p et m dans le boucle while.
    k1=k2=0
    Bool=True
    #initialisation de matrice machine et piece
    p=[]
    m=[]
    print("machine : ",machine)
    print("piece : ",piece)
    i = Piece1.tolist().index(piece[0]) # i will return index of 2
    print (i)
    m=Int1(PM,i)
    p=PIECE1(PM,m)
    p=PIECE50(PM,m,p)
    m=MACHINE1(PM,m,p)
    m=MACHINE50(PM,m,p)
    #ilot=np.array([])
    #Boucler dans la matrice jusqu'a le parcour de tout les "1".
    #print("BOUCLE WHILE\n")
    while(Bool):
        k1=len(p)+len(m)
        print("\nLongeur de m+p :",k1)
        p=PIECE1(PM,m)
        p=PIECE50(PM,m,p)
        m=MACHINE1(PM,m,p)
        m=MACHINE50(PM,m,p)
        k2=len(p)+len(m)
        print("\nLongeur de m+p apres l'update :",k2)
        if(k2!=k1):
            print("Updated!\n")
            continue
        else:
            print("No Update!\n")
            Bool=False
            
    #print("m" ,m)
    #ilot=np.append(ilot,m)
    #print("ILOT :",ilot)
    # Suppression des machine et des pieces
    
    print("\n",piece,"\n")
    for j in range (len(p)):
        print(int(p[j])+1)
        piece=np.delete(piece,np.where(piece==('P'+str(int(p[j]+1)))))
        #piece=np.delete(piece,int(p[j]-j))
        print(piece)
        
    print("\n",machine,"\n")
    for j in range (len(m)):
        print(int(m[j])+1)
        machine=np.delete(machine,np.where(machine==('M'+str(int(m[j]+1)))))
        #machine=np.delete(machine,int(m[j]-j))
        print(machine)
    for i in p:
        for j in range(len(Machine1)):
            PM[i][j]=0
    for i in m:
        for j in range(len(Piece1)):
            PM[j][i]=0

    return m,p,piece,machine


# In[84]:


def main_Kusiak_Modifier():
    M=P=np.array([],dtype=int)
    m=[]
    p=[]
    pm1=PM1.copy()
    m1,p1,piece1,machine1=ILOT1(pm1,Piece1,Machine1)
    print("\n\nILOT 1 :\n")
    print("\nM :",m1+1,"\nP :",p1+1)
    M=np.append(M,m1+1)
    m.append((m1+1).tolist())
    P=np.append(P,p1+1)
    p.append((p1+1).tolist())
    k=2
    while(len(machine1)>=1):
        print("\n\n NEW ILOT :\n\n")
        m1,p1,piece1,machine1=ILOT1(pm1,piece1,machine1)
        print("\nILOT ",k,":\n")
        print("\nM :",m1+1,"\nP :",p1+1)
        M=np.append(M,m1+1)
        m.append((m1+1).tolist())
        P=np.append(P,p1+1)
        p.append((p1+1).tolist())
        k+=1
    print("\nLes Machine par Ilot :\n",m)
    print("\nLes pieces par Ilot :\n",p)
    return M,P,m,p


# In[85]:


M,P,m,p=main_Kusiak_Modifier()


# #### Cette fonction permet de trier la matrice PM selon l'ordre des ilots

# In[86]:


def Tri(PM,T1,T2):
    #Tri sur ligne
    PM1=np.zeros(PM.shape,dtype=int)
    for i in range(len(T1)):
        PM1[i]=PM[T1[i]-1]
    #Tri sur colonne
    PM2=PM1.T
    PM3=np.zeros(PM.T.shape,dtype=int)
    for j in range(len(T2)):
        PM3[j]=PM2[T2[j]-1]
    return PM3.T


# #### Dans cette partie on va expoter la matrice triée dans le fichier EXEL

# In[87]:


#Creation de la matrice trier 
df1=pd.DataFrame(Tri(PM1,P,M))
#print(df1)
machine=[]
piece=[]
for i in range(len(M)):
    machine.append('M'+str(M[i]))
#print(machine)
for j in range(len(P)):
    piece.append(('P')+str(P[j]))
#print(piece)
df1.columns=machine
df1.index=piece
#Creation de la matrice des machines par ilot 
MACHINEINDEX=[]
MACHINECOL=[]
MACHINE=pd.DataFrame(m)
for i in range(len(m)):
    MACHINEINDEX.append('ILOT'+str(i+1))
MACHINE.index=MACHINEINDEX
MACHINE.columns=itertools.repeat('MACHINE',MACHINE.columns.stop)
#Creation de la matrice des pieces par ilot 
PIECEINDEX=[]
PIECECOL=[]
PIECE=pd.DataFrame(p)
for i in range(len(p)):
    PIECEINDEX.append('ILOT'+str(i+1))
PIECE.index=PIECEINDEX
PIECE.columns=itertools.repeat('PIECE',PIECE.columns.stop)

#EXPORT TO EXEL
book = load_workbook("RESULTAT_KUSIAK_MODIFIER.xlsx")
writer = pd.ExcelWriter("RESULTAT_KUSIAK_MODIFIER.xlsx", engine='openpyxl')
del book['DATA']
del book['RESULTAT']
del book['ILOTS']
writer.book = book
writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
df.to_excel(writer, "DATA",index=False)
df1.to_excel(writer, "RESULTAT", columns=machine)
MACHINE.to_excel(writer, "ILOTS")
PIECE.to_excel(writer, "ILOTS",startrow=15)
writer.save()
#Open Exel file
os.system('start excel.exe "%s"' % ("RESULTAT_KUSIAK_MODIFIER.xlsx", ))

#df1.to_excel(r'C:\Users\bouza\Desktop\cour\semestre 5\Conception & Performance des systèmes de\PROJET\RESULTAT_KUSIAK.xlsx', sheet_name='RESULTAT')


# In[ ]:




