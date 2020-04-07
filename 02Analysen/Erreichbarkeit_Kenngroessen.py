#!/usr/bin/python
# -*- coding: cp1252 -*-
#-------------------------------------------------------------------------------
# Name:        Erreichbarkeiten aus Kenngrößen
# Purpose: Dieses Tool dient der Erreichbarkeitsberechnung auf Ebene von Bezirken
# mit Hilfe von Kenngrößenmatrizen.
#
# Author:      mape
#
# Created:     27/08/2015
# Update:      24/05/2016
# Copyright:   (c) mape 2015
# Licence:     <your licence>
#-------------------------------------------------------------------------------


#---Vorbereitung---#
import win32com.client.dynamic
import numpy as np
import time
import h5py
start_time = time.clock() ##Ziel: am Ende die Berechnungsdauer ausgeben

from pathlib import Path
path = Path.home() / 'python32' / 'python_dir.txt'
f = open(path, mode='r')
for i in f: path = i
path = Path.joinpath(Path(r'C:'+path),'VISUM_Tools','Erreichbarkeit_Kenngr.txt')
f = path.read_text()
f = f.split('\n')

##############################
#--Parameter--#
##############################
Netz = 'V:'+f[0]
HDF = 'V:'+f[1]
Tabelle_E = f[2]


##############################
#--HDF5--#
##############################
file5 = h5py.File(HDF,'r+') ##HDF5-File
group5 = file5["MRH"] ##Gruppe

try:
    del group5[Tabelle_E]
except:
    pass

#--Attribute HDF5--#
Spalten = [("ID","int32"),("EW30","int32"),("EW60","int32"),("AP30","int32"),("AP60","int32"),("Anzahl","int32"),("Fernbahn","<f8"),("APUH0","int32"),("APBH20","int32")]
Spalten = np.dtype(Spalten) ##Wandle Spalten-Tuple in dtype um


##############################
#--VISUM--#
##############################
print "--Binde VISUM ein und lade Version--"
VISUM = win32com.client.dynamic.Dispatch("Visum.Visum.15")
VISUM.loadversion(Netz) ##Netz,Szenrario,VISUM-Instanz
VISUM.Filters.InitAll()

print "--Folgende Matrizen gibt es:--"
for Matrix in VISUM.Net.SkimMatrices:
    Nr = Matrix.AttValue("No")
    Name = Matrix.AttValue("Name")
    print Name+" mit der Nummer "+str(Nr)


#--Erstelle Liste mit Strukturvariablen--#
Liste = VISUM.Lists.CreateZoneList
Liste.AddColumn("No")
Liste.AddColumn("ValStructuralProp(EW)")
Liste.AddColumn("ValStructuralProp(A)")
Liste.AddColumn("Min:OrigConnectors\Node\Min:StopAreas\TypeNo") ##Spalte mit dem niedrigsten Typ der angebundenen Haltestellen

Liste_np = np.array(Liste.SaveToArray())

##############################
#--Berechnung--#
##############################
"""
1.
"""
matRZ = np.array(VISUM.Net.SkimMatrices.ItemByKey(1).GetValues())
matUH = np.array(VISUM.Net.SkimMatrices.ItemByKey(3).GetValues())
matBH = np.array(VISUM.Net.SkimMatrices.ItemByKey(4).GetValues())

#--Indikatoren--#
T_E = [] ##Ergebnistabelle
for i,RZ in enumerate(matRZ):
    Werte = [0,0,0,0,0,0,999,0,0]
    Werte[0] = int(float(Liste_np[i][0]))
    for n,e in enumerate(RZ):
        #Reisezeitbudget
        if e < 30:
            Werte[1] = Werte[1]+int(float(Liste_np[n][1]))
            Werte[3] = Werte[3]+int(float(Liste_np[n][2]))
        if e < 60:
            Werte[2] = Werte[2]+int(float(Liste_np[n][1]))
            Werte[4] = Werte[4]+int(float(Liste_np[n][2]))
        #erreichbare Ziele
        if e < 500: ##Anzahl der Verkehrszellen, die überhaupt erreicht werden.
            Werte[5] = Werte[5]+1
        #Zeit Fernbahn
        try: ##klappt nur, wenn es eine ÖV-Anbindung gibt
            if int(float(Liste_np[n][3]))==1: ##Um die Reisezeit zum nächsten Fernbahnhalt zu ermitteln (Typ==1)
                if Werte[6] > e:
                    Werte[6] = e
        except:
            pass
        #Umsteigehäufigkeit
        if matUH[i][n]==0:
            Werte[7] = Werte[7]+int(float(Liste_np[n][2]))
        #Bedienhäufigkeit
        if matBH[i][n]>30: ##wegen Werten in unerreichten Zellen
            Werte[8] = Werte[8]+int(float(Liste_np[n][2]))


    T_E.append(Werte)


##############################
#--Ergebnistabelle--#
##############################
Daten = list(map(tuple, T_E))
data = np.array(Daten,Spalten)
group5.create_dataset(Tabelle_E, data = data, dtype = Spalten, maxshape = (None,))
file5.flush()
file5.close()


##############################
#--Ergebnistabelle--#
##############################
print "----fertig----"
