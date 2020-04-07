#!/usr/bin/python
# -*- coding: cp1252 -*-
#-------------------------------------------------------------------------------
# Name: Erreichbarkeiten aus Kenngrößen der Haltestellenbereiche
# Purpose: Dieses Tool dient der Umwandlung von HstBer-Kenngrößen in Isochronen
# Es ist außerdem mögliche, die Umwandlung auf bestimme Gebiete einzugrenzen
#
# Author:      mape
#
# Created:     19/01/2016
# Copyright:   (c) mape 2016
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
path = Path.joinpath(Path(r'C:'+path),'VISUM_Tools','Kenngroessen_Iso.txt')
f = path.read_text()
f = f.split('\n')

#--Parameter--#
Netz = 'C:'+f[0]
HDF = 'C:'+f[1]
group5 = f[2]
Text = f[3]
I_Name = f[4]
Zeitschranke = 120 ##Wähle nur die Verbindungen, die Zeitschranke erfüllen
Startwartezeit = 1
Stundenintervall = 180 ##Anzahl der Minuten im Zeitintervall. Zur Berechung der Startwartezeit.


#--HDF5--#
file5 = h5py.File(HDF,'r+') ##HDF5-File
group5 = file5[group5] ##Gruppe


#--VISUM Version und Szenario--#
print "--Binde VISUM ein und lade Version--"
VISUM = win32com.client.dynamic.Dispatch("Visum.Visum.17")
VISUM.loadversion(Netz) ##Netz,Szenrario,VISUM-Instanz
VISUM.Filters.InitAll()
##VISUM.Procedures.Execute()


#--Erstelle die IsoChronen-Tabelle--#
Liste_HB = []
Spalten = [('StartHstBer', 'i4'),('ZielHstBer', 'i4'),('Kosten', '<f8'),('UH', 'i2'),('BH', 'i2')]
Spalten = np.dtype(Spalten) ##Wandle Spalten-Tuple in dtype um
data = np.array(Liste_HB,Spalten)
if I_Name in group5.keys():
        del group5[I_Name]
group5.create_dataset(I_Name, data = data, dtype = Spalten, maxshape = (None,))
file5.flush()

#--HstBer-Array--#
HstBerNr = np.array(VISUM.Net.StopAreas.GetMultiAttValues("No")).astype("int")[:,1]
Pointer = len(HstBerNr)
Array = range(1,Pointer+1) ##um einen Array von 1 bis Pointer zu erhalten

#--Matrizen Liste--#
for Matrix in VISUM.Net.Matrices:
    if Matrix.AttValue("ObjectTypeRef") == 'OBJECTTYPEREF_STOPAREA':
        if Matrix.AttValue("Code") == "JRT": ##Reisezeit
            JRT = int(Matrix.AttValue("No"))

        if Matrix.AttValue("Code") == "NTR": ##Umsteigehäufigkeit
            NTR = int(Matrix.AttValue("No"))

        if Matrix.AttValue("Code") == "SFQ": ##Bedienhäufigkeit
            SFQ = int(Matrix.AttValue("No"))


#--Befülle Tabelle--#
IsoChronen = group5[I_Name]

for i in Array:
    Liste_HB = []
    #--Kennwerte--#
    matJRT = np.array(VISUM.Net.Matrices.ItemByKey(JRT).GetRow(i))
    matNTR = np.array(VISUM.Net.Matrices.ItemByKey(NTR).GetRow(i))
    matSFQ = np.array(VISUM.Net.Matrices.ItemByKey(SFQ).GetRow(i))

    #--Einzelverbindungen--#
    for e,Nr in enumerate(HstBerNr):
        Start = int(HstBerNr[i-1]) ##Da hier der Index bei 0 beginnt und bei 'Array' bei 1
        Ziel = int(Nr)
        Zeit = matJRT[e]
        if Startwartezeit == 1:
            try: ##Fehler wenn BH = 0 bzw. Zeit=0. Dann Binnenverkehr
                SWZ = 0.53*((Stundenintervall/float(matSFQ[e]))**0.75)
                if SWZ > 10:SWZ=10
                Zeit+= SWZ
            except:
                pass
        Zeit = int(Zeit)
        UH = int(matNTR[e])
        BH = int(matSFQ[e])


        Liste_HB.append((Start,Ziel,Zeit,UH,BH))

    #--Bereinigung--#
    if len(Liste_HB) == 0:
        continue
    Liste_HB = np.array(Liste_HB)
    Liste_HB = Liste_HB[Liste_HB[:,2]<Zeitschranke] ##wähle nur die Zeilen, die Zeitschranke erfüllen.

    #--Fülle Daten in HDF5-Tabelle--#
    oldsize = len(IsoChronen)
    sizer = oldsize + len(Liste_HB) ##Neue Länge der Liste (bisherige Länge + Läne VISUM-IsoChronen
    IsoChronen.resize((sizer,))
    Liste_HB = list(map(tuple, Liste_HB))
    IsoChronen[oldsize:sizer] = Liste_HB
    file5.flush()

#--Parameter Text--#
IsoChronen.attrs.create("Parameter",str(Text))
file5.flush()

file5.close()