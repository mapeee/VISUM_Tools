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
Datafile= f[1]
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
    try:
        VISUM.Procedures.Execute()
        print("> skim matrices calculated")
    except: print("> error calculating skim matrix")
    return VISUM

def HDF5(Data):
    result_array = []
    Columns = np.dtype([('FromStop', 'i4'),('ToStop', 'i4'),('Time', '<f8'),('UH', 'i2'),('BH', 'i2')])
    data = np.array(result_array,Columns)
    #--HDF5--#
    file5 = h5py.File(Data,'r+')
    gr = file5[group5]
    if I_Name in gr.keys(): del gr[I_Name]
    gr.create_dataset(I_Name, data = data, dtype = Columns, maxshape = (None,))
    file5.flush()
    IsoChrones = gr[I_Name]
    return file5, IsoChrones

def matrices(VISUM):
    for Matrix in VISUM.Net.Matrices:
        if Matrix.AttValue("ObjectTypeRef") == 'OBJECTTYPEREF_STOPAREA':
            if Matrix.AttValue("Code") == "JRT": JRT = int(Matrix.AttValue("No")) ##Traveltime    
            if Matrix.AttValue("Code") == "NTR": NTR = int(Matrix.AttValue("No")) ##Transfers
            if Matrix.AttValue("Code") == "SFQ": SFQ = int(Matrix.AttValue("No")) ##Service frequency
    return JRT, NTR, SFQ
  
###############  
#--Preparing--#
###############
print("> starting")
file5, IsoChrones = HDF5(Datafile)
VISUM = calc(Network)
JRT, NTR, SFQ = matrices(VISUM)
  
################          
#--Isochrones--#
################
StopAreaNo = np.array(VISUM.Net.StopAreas.GetMultiAttValues("No",False)).astype("int")[:,1]
StopAreaNoActive = np.array(VISUM.Net.StopAreas.GetMultiAttValues("No",True)).astype("int")[:,1]

for i in range(1,len(StopAreaNo)+1):
    result_array = []
    From = int(StopAreaNo[i-1])
    if From not in StopAreaNoActive: continue
    
    #--Values--#
    matJRT = np.array(VISUM.Net.Matrices.ItemByKey(JRT).GetRow(i))
    matNTR = np.array(VISUM.Net.Matrices.ItemByKey(NTR).GetRow(i))
    matSFQ = np.array(VISUM.Net.Matrices.ItemByKey(SFQ).GetRow(i))

    for e,Nr in enumerate(StopAreaNo):
        To = int(Nr)
        TTime = matJRT[e]
        if Origin_wait_time == 1:
            try: ##error if BH = 0 or Time = 0 --> Internal Trips
                SWZ = 0.53*((Stundenintervall/float(matSFQ[e]))**0.75)
                if SWZ > 10:SWZ = 10
                TTime+= SWZ
            except: pass
        TTime = int(round(TTime))
        UH = int(matNTR[e])
        BH = int(matSFQ[e])

        result_array.append((From,To,TTime,UH,BH))

    if len(result_array) == 0: continue
    result_array = np.array(result_array)
    result_array = result_array[result_array[:,2]<Zeitschranke]

    #--Data into HDF5 table--#
    oldsize = len(IsoChrones)
    sizer = oldsize + len(result_array)
    IsoChrones.resize((sizer,))
    result_array = list(map(tuple, result_array))
    IsoChrones[oldsize:sizer] = result_array
    file5.flush()

#--end--#
VISUM = False
IsoChrones.attrs.create("Parameter",str(Text))
file5.flush()
file5.close()
print("> finished")