#!/usr/bin/python
# -*- coding: cp1252 -*-
#-------------------------------------------------------------------------------
# Name: Matrices_to_Isochrones
# Purpose: Transformation and export of PTV VIsum SkimMatrices to HDF5-Dataframe
#
#
# Author:      mape
#
# Created:     19/01/2016 (new Version November 2021)
# Copyright:   (c) mape 2016
# Licence:     <your licence>
#-------------------------------------------------------------------------------
import win32com.client.dynamic
import numpy as np
import h5py
from pathlib import Path
f = open(Path.home() / 'python32' / 'python_dir.txt', mode='r')
for i in f: path = i
path = Path.joinpath(Path(path),'VISUM_Tools','Matrices_to_Isochrones.txt')
f = path.read_text().split('\n')

#--Parameters--#
Network = f[0]
HDF5 = f[1]
group5 = f[2]
Text = f[3]
I_Name = f[4]
Zeitschranke = 120 
Origin_wait_time = 1
Stundenintervall = 180 ##for pre 


def calc(Network):
    print("> "+Network)
    VISUM = win32com.client.dynamic.Dispatch("Visum.Visum.20")
    VISUM.loadversion(Network)
    # VISUM.Filters.InitAll()
    try: VISUM.Procedures.Execute()
    except: print("> error calculating skim matrix")
    print("\n")
    return VISUM


print("> starting \n")

#--VISUM--#
VISUM = calc(Network)
for Matrix in VISUM.Net.Matrices:
    if Matrix.AttValue("ObjectTypeRef") == 'OBJECTTYPEREF_STOPAREA':
        if Matrix.AttValue("Code") == "JRT": JRT = int(Matrix.AttValue("No")) ##Traveltime    
        if Matrix.AttValue("Code") == "NTR": NTR = int(Matrix.AttValue("No")) ##Transfers
        if Matrix.AttValue("Code") == "SFQ": SFQ = int(Matrix.AttValue("No")) ##Service frequency
            
StopAreaNo = np.array(VISUM.Net.StopAreas.GetMultiAttValues("No")).astype("int")[:,1]
Array = range(1,len(StopAreaNo)+1) 


#--results--#
result_array = []
Columns = np.dtype([('FromArea', 'i4'),('ToArea', 'i4'),('Time', '<f8'),('UH', 'i2'),('BH', 'i2')])
data = np.array(result_array,Columns)





#--HDF5--#
file5 = h5py.File(HDF5,'r+')
group5 = file5[group5]
if I_Name in group5.keys():
        del group5[I_Name]
group5.create_dataset(I_Name, data = data, dtype = Columns, maxshape = (None,))
file5.flush()

IsoChronen = group5[I_Name]



for i in Array:
    result_array = []
    #--Kennwerte--#
    matJRT = np.array(VISUM.Net.Matrices.ItemByKey(JRT).GetRow(i))
    matNTR = np.array(VISUM.Net.Matrices.ItemByKey(NTR).GetRow(i))
    matSFQ = np.array(VISUM.Net.Matrices.ItemByKey(SFQ).GetRow(i))

    #--Einzelverbindungen--#
    for e,Nr in enumerate(StopAreaNo):
        Start = int(StopAreaNo[i-1]) ##Da hier der Index bei 0 beginnt und bei 'Array' bei 1
        Ziel = int(Nr)
        Zeit = matJRT[e]
        if Origin_wait_time == 1:
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