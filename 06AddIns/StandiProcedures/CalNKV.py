# -*- coding: utf-8 -*-
"""
Created on Fri Feb  2 13:33:29 2024
"""
import numpy as np
import pandas as pd
from VisumPy.AddIn import AddIn
from VisumPy import helpers
_ = AddIn.gettext


def calc(Visum,Table):
    HfE = 300
    HfS = 250
    
    emissionfactor = dict((x, y) for x, y in Visum.Net.TableDefinitions.ItemByKey("Standi-Energie").TableEntries.GetMultipleAttributes(["ENERGIEART","EMISSIONSFAKTOR"]))
    emissioncosts = dict((x, y) for x, y in Visum.Net.TableDefinitions.ItemByKey("Standi-Energie").TableEntries.GetMultipleAttributes(["ENERGIEART","EMISSIONSKOSTENSATZ"]))
    scenario = dict((x, y) for x, y in Visum.Net.TableDefinitions.ItemByKey("Standi-Szenario").TableEntries.GetMultipleAttributes(["CODE","WERT"]))
    bcpc = dict((x, y) for x, y in Visum.Net.TableDefinitions.ItemByKey("Ohnefall-Mitfall-Vergleich").TableEntries.GetMultipleAttributes(["CODE","DIFF_ABS"]))
    parameter = dict((x, y) for x, y in Visum.Net.TableDefinitions.ItemByKey("Standi-Parameter").TableEntries.GetMultipleAttributes(["CODE","WERT"]))
    HF_PT = parameter["HF_OEV"]
    HF_PrT = parameter["HF_MIV"]
    
    matrices = ['Matrix([CODE] = "verlOEV")','Matrix([CODE] = "OEV_E_WidDiff")','Matrix([CODE] = "OEV_BefW_M")']
    
    if not _checkMatrExisting(Visum, matrices): return False
    if not _checkMatrValues(Visum, matrices): return False
    if not _checkBDAs(Visum): return False
    if not _checkSzenarioTable(Visum): return False
    _checkNKVTable(Visum, Table)

    #FgNu: Fahrgastnutzen
    FgNu = (Visum.Net.Matrices.ItemsByRef('Matrix([CODE] = "OEV_E_WidDiff")').GetAll[0].AttValue("Sum") * HfE +
             Visum.Net.Matrices.ItemsByRef('Matrix([CODE] = "OEV_S_WidDiff")').GetAll[0].AttValue("Sum") * HfS) * 0.001
    _TableAddValue(Table, "FgNu", FgNu)  
    
    #FahrG: Saldo ÖPNV-Fahrgeld
    NoVerl = Visum.Net.Matrices.ItemsByRef('Matrix([CODE] = "verlOEV")').GetAll[0].AttValue("NO")
    NoInd = Visum.Net.Matrices.ItemsByRef('Matrix([CODE] = "indOEV")').GetAll[0].AttValue("NO")
    NoBefW = Visum.Net.Matrices.ItemsByRef('Matrix([CODE] = "OEV_BefW_M")').GetAll[0].AttValue("NO")
    FahrG = np.sum((helpers.GetODMatrixRaw(Visum, NoVerl) + helpers.GetODMatrixRaw(Visum, NoInd)) * helpers.GetODMatrixRaw(Visum, NoBefW))
    FahrG = FahrG * HfE * 0.001
    _TableAddValue(Table, "FahrG", FahrG)  
      
    #BetrK: Saldo der ÖPNV-Betriebskosten
    FzgK = np.array(Visum.Net.VehicleUnits.GetMultiAttValues("FAHRZEUGK_M"))[:,1].sum() - np.array(Visum.Net.VehicleUnits.GetMultiAttValues("FAHRZEUGK_O"))[:,1].sum()
    EnergieK = np.array(Visum.Net.Lines.GetMultiAttValues("ENERGIEK_M"))[:,1].sum() - np.array(Visum.Net.Lines.GetMultiAttValues("ENERGIEK_O"))[:,1].sum()
    PersoK = np.array(Visum.Net.Lines.GetMultiAttValues("PERSONALK_M"))[:,1].sum() - np.array(Visum.Net.Lines.GetMultiAttValues("PERSONALK_O"))[:,1].sum()
    
    _TableAddValue(Table, "BetrK", FzgK + EnergieK + PersoK) 
    _TableAddValue(Table, "BetrK", FzgK, 1) 
    _TableAddValue(Table, "BetrK", EnergieK, 2) 
    _TableAddValue(Table, "BetrK", PersoK, 3)
        
    #UnterM - Unterhaltungskosten für die ortsfeste Infrastruktur im Mitfall

    BK = scenario["BK"] * 1000
    BKa = scenario["BKa"] / 100
    UnterM = (BK / scenario["I2016"] * 100 ) * BKa * scenario["UK"] * 0.001
    _TableAddValue(Table, "UnterM", UnterM)  
            
    #UnterO - Unterhaltungskosten für die ortsfeste Infrastruktur im Ohnefall
    _TableAddValue(Table, "UnterO", scenario["KapIO"] * 0.2)  
    
    #UnfFK - Saldo der Unfallfolgekosten
    #MIV
    UnfFk_MIV = bcpc["Pkw_Fzgkm"] * HF_PrT * 0.001 * 8.5 * 0.01
    
    #OEV
    TSys = (np.array(Visum.Net.Lines.GetMultiAttValues("TSYSCODE"))[:,1])
    Length_O = (np.array(Visum.Net.Lines.GetMultiAttValues("LL_O"))[:,1])
    Length_M = (np.array(Visum.Net.Lines.GetMultiAttValues("LL_M"))[:,1])
    Lines = pd.DataFrame(np.stack((TSys, Length_O, Length_M), axis=1), columns = ["TSYSCODE", "LL_O", "LL_M"])
    Lines["LL_O"] = Lines["LL_O"].astype(float)
    Lines["LL_M"] = Lines["LL_M"].astype(float)
    Lines["Diff"] = (Lines["LL_M"] - Lines["LL_O"]) * HF_PT * 0.001 * 0.01

    value_SPNV = Lines.groupby(["TSYSCODE"])["Diff"].sum()["RV":"S"].sum() * 36.4
    value_U = Lines.groupby(["TSYSCODE"])["Diff"].sum()["U":"U"].sum() * 19.8
    value_Bus = Lines.groupby(["TSYSCODE"])["Diff"].sum()["Bus":"Bus"].sum() * 21.3
    value_W = Lines.groupby(["TSYSCODE"])["Diff"].sum()["W":"W"].sum() * 21.3
    UnfFk_OEV = value_SPNV + value_U + value_Bus + value_W
    
    _TableAddValue(Table, "UnfFK", UnfFk_MIV + UnfFk_OEV) 
    _TableAddValue(Table, "UnfFK", UnfFk_MIV, 1)
    _TableAddValue(Table, "UnfFK", UnfFk_OEV, 2)

    #CO2E - Saldo der CO2-Emissionen
    #MIV
    MIV_B = bcpc["Pkw_Fzgkm"] * HF_PrT * 0.001 * 127 * 0.001 * 0.001
    MIV_H = bcpc["Pkw_Fzgkm"] * HF_PrT * 0.001 * 41 * 0.001 * 0.001
                
    #OEV
    #Betrieb
    VehCombNo = np.array(Visum.Net.VehicleCombinations.GetMultiAttValues("NO"))[:,1]
    Energy = np.array(Visum.Net.VehicleCombinations.GetMultiAttValues("MIN:VEHUNITS\ENERGIE"))[:,1]
    Ev_O = np.array(Visum.Net.VehicleCombinations.GetMultiAttValues("ENERGIEV_O"))[:,1]
    Ev_M = np.array(Visum.Net.VehicleCombinations.GetMultiAttValues("ENERGIEV_M"))[:,1]
    VehComb = pd.DataFrame(np.stack((VehCombNo, Energy, Ev_O, Ev_M), axis=1), columns = ["VehCombNo","Energy","Ev_O","Ev_M"])
    VehComb["Ev_O"] = VehComb["Ev_O"].astype(float)
    VehComb["Ev_M"] = VehComb["Ev_M"].astype(float)
    VehComb["CO2Diff"] = VehComb.apply(_CO2Diff_PuTService, _emissionfactor = emissionfactor, axis=1)
    OEV_B = VehComb["CO2Diff"].sum() * 0.001
    
    #Herstellung
    VehUnitNo = np.array(Visum.Net.VehicleUnits.GetMultiAttValues("NO"))[:,1]
    TSys = np.array(Visum.Net.VehicleUnits.GetMultiAttValues("TSYSSET"))[:,1]
    Weight = np.array(Visum.Net.VehicleUnits.GetMultiAttValues("LEERMASSE"))[:,1]
    THG = np.array(Visum.Net.VehicleUnits.GetMultiAttValues("THG"))[:,1]
    FZG_O = np.array(Visum.Net.VehicleUnits.GetMultiAttValues("FZG_O"))[:,1]
    FZG_M = np.array(Visum.Net.VehicleUnits.GetMultiAttValues("FZG_M"))[:,1]
    VehUnit = pd.DataFrame(np.stack((VehUnitNo, TSys, Weight, THG, FZG_O, FZG_M), axis=1), columns = ["VehUnitNo","TSys","Weight","THG","FZG_O","FZG_M"])
    VehUnit["THG"] = VehUnit["THG"].astype(float)
    VehUnit["Weight"] = VehUnit["Weight"].astype(float)
    VehUnit["FZG_O"] = VehUnit["FZG_O"].astype(float)
    VehUnit["FZG_M"] = VehUnit["FZG_M"].astype(float)
    VehUnit["OEV_H"] = VehUnit.apply(_THGVehUnit, axis=1)
    
    OEV_H = VehUnit["OEV_H"].sum() * 0.001 * 0.001
    
    _TableAddValue(Table, "CO2E", MIV_B + MIV_H + OEV_B + OEV_H) 
    _TableAddValue(Table, "CO2E", MIV_B, 1)
    _TableAddValue(Table, "CO2E", OEV_B, 2)
    _TableAddValue(Table, "CO2E", MIV_H, 3)
    _TableAddValue(Table, "CO2E", OEV_H, 4)

    #SchadK - Saldo der Schadstoffemissionskosten
    #MIV
    SchadK_MIV = bcpc["Pkw_Fzgkm"] * HF_PrT * 0.001 * 0.4 * 0.01
    #OEV 
    VehComb["SchadKDiff"] = VehComb.apply(_THGDiff_PuTService, _emissioncosts = emissioncosts, axis=1) 
    SchadK_OEV = VehComb["SchadKDiff"].sum()  
    
    _TableAddValue(Table, "SchadK", SchadK_MIV + SchadK_OEV) 
    _TableAddValue(Table, "SchadK", SchadK_MIV, 1)
    _TableAddValue(Table, "SchadK", SchadK_OEV, 2)

    #LAERM - Saldo der Geräuschbelastung
    _TableAddValue(Table, "LAERM", scenario["LAERM"])

    #NGaufI - Nutzen gesellschaftlich auferlegter Investitionen
    _TableAddValue(Table, "NGaufI", scenario["NGaufI"])
    
    #NaNN - Nutzen anderer Netznutzer
    _TableAddValue(Table, "NaNN", scenario["NaNN"])
            
    #FVSysF - Funktionsfähigkeit der Verkehrssysteme / Flächenverbrauch
    _TableAddValue(Table, "FVSysF", scenario["FVSysF"])
    
    #ENERGIE - Primärenergieverbrauch
    _TableAddValue(Table, "ENERGIE", scenario["ENERGIE"])
    
    #DVRaum - Daseinsvorsorge / raumordnerische Aspekte
    _TableAddValue(Table, "DVRaum", scenario["DVRaum"])
    
    #RSchiene - Resilienz von Schienennetzen
    _TableAddValue(Table, "RSchiene", scenario["RSchiene"])

    #MBewEN - Summe der monetär bewerteter Einzelnutzen
    MBewEN = 0
    for i in Table.TableEntries:
        if i.AttValue("CODE") == "MBewEN": break
        if i.AttValue("CODE") != "":
            MBewEN = MBewEN + i.AttValue("BEWERTUNG")
    _TableAddValue(Table, "MBewEN", MBewEN)

    #KapI - Kapitaldienst für die ortsfeste Infrastruktur
    #Mitfall
    BZ = scenario["BZ"]
    AF = scenario["AF"]
    I2016 = scenario["I2016"]
    AFBZ = (pow(1.017, BZ) - 1) / (0.017 * BZ)
    KapI_M = (BK / I2016 * 100) * AF * AFBZ
    
    #Ohnefall
    KapI = KapI_M + (scenario["KapIO"] * -1)
    _TableAddValue(Table, "KapI", KapI)
    _TableAddValue(Table, "KapI", KapI_M, 10)
    _TableAddValue(Table, "KapI", scenario["KapIO"] * -1, 20)

    #NKD - Nutzen-Kosten-Differenz
    NKD = MBewEN - KapI
    _TableAddValue(Table, "NKD", NKD)
    
    #NKV - Nutzen-Kosten-Verhältnis
    NKV = MBewEN / KapI
    _TableAddValue(Table, "NKV", NKV)

    #Standi-Szenario
    for i in Visum.Net.TableDefinitions.ItemByKey("Standi-Szenario").TableEntries:
        if i.AttValue("CODE") == "MBewEN": i.SetAttValue("WERT", MBewEN)
        if i.AttValue("CODE") == "KapI": i.SetAttValue("WERT", KapI)
        if i.AttValue("CODE") == "NKD": i.SetAttValue("WERT", NKD)
        if i.AttValue("CODE") == "NKV": i.SetAttValue("WERT", NKV)
    
    return True

def _addTableEntries(Visum):
    addIn = AddIn(Visum)
    Visum.Net.TableDefinitions.ItemByKey("Standi-NKV").TableEntries.RemoveAll()
    Visum.IO.LoadAccessDatabase(addIn.DirectoryPath + "Data\\Standi_NKVEntries.accdb", True)

def _checkBDAs(Visum):
    for i in ["FAHRZEUGK_M"]:
        if not Visum.Net.VehicleUnits.AttrExists(i):
            Visum.Log(12288,_("VehicleUnits UDA '%s' is missing!") %(i))
            return False
    for i in ["ENERGIEV_M"]:
        if not Visum.Net.VehicleCombinations.AttrExists(i):
            Visum.Log(12288,_("VehicleCombinations UDA '%s' is missing!") %(i))
            return False
    for i in ["ENERGIEK_M","PERSONALK_M"]:
        if not Visum.Net.Lines.AttrExists(i):
            Visum.Log(12288,_("Lines UDA '%s' is missing!") %(i))
            return False
    return True

def _checkMatrExisting(Visum, matrices):
    for i in matrices:
        if len(Visum.Net.Matrices.ItemsByRef(i).GetAll) == 0:
            Visum.Log(12288,_("'%s' is missing!") %(i))
            return False
    return True
    
def _checkNKVTable(Visum, Table):
    if Table.TableEntries.Count != 33:
        _addTableEntries(Visum)
        return
    
    for i in ["FgNu","FahrG","BetrK","UnterM","UnterO","UnfFK","CO2E","SchadK","LAERM","NGaufI","NaNN",
              "FVSysF","ENERGIE","DVRaum","RSchiene","MBewEN","KapI","KapIM","KapIO","NKD","NKV"]:
        if i not in np.array(Table.TableEntries.GetMultiAttValues("CODE"))[:,1]:
            _addTableEntries(Visum)
            return
    
    for i in ["BEWERTUNG","BEWERTUNGSANSATZ","BEWERTUNGSANSATZ_E","CODE","MESSGROESSE","MESSGROESSE_D","QUELLE","TEILINDIKATOR"]:
        if not Table.TableEntries.AttrExists(i):
            _addTableEntries(Visum)
            return

def _checkSzenarioTable(Visum):
    Table = Visum.Net.TableDefinitions.ItemByKey("Standi-Szenario")
    for i in ["M","BK","BZ","BKa","UK","AF","I2016","NGaufI","NaNN","LAERM","FVSysF","ENERGIE","DVRaum","RSchiene",
              "KapIO","MBewEN","KapI","NKD","NKV"]:
        if i not in np.array(Table.TableEntries.GetMultiAttValues("CODE"))[:,1]:
            Visum.Log(12288,_("CODE '%s' in Table 'Standi-Szenario' is missing!") %(i))
            return False
    for i in ["CODE","Name","WERT"]:
        if not Table.TableEntries.AttrExists(i):
            Visum.Log(12288,_("UDA '%s' in Table 'Standi-Szenario' is missing!") %(i))
            return False

    i = Table.TableEntries.Iterator
    while i.Valid:
        if i.Item.AttValue("CODE") in ["BK","BKa","I2016","BZ","UK","AF"] and i.Item.AttValue("WERT") == 0:
            Visum.Log(12288,_("Value for CODE '%s' in Table 'Standi-Szenario' must not be 0!") %(i.Item.AttValue("CODE")))
            return False
        i.Next()
    
    return True

def _checkMatrValues(Visum, matrices):  
    for i in matrices: 
        if Visum.Net.Matrices.ItemsByRef(i).GetAll[0].AttValue("Sum") == 0:
            Visum.Log(12288,_("Matrix '%s' is empty!") %(Visum.Net.Matrices.ItemsByRef(i).GetAll[0].AttValue("CODE")))
            return False
    return True

def _CO2Diff_PuTService(data, _emissionfactor):
    if data["Energy"] == None: return 0
    O = (data["Ev_O"] * (_emissionfactor[data["Energy"]])) * 0.001
    M = (data["Ev_M"] * (_emissionfactor[data["Energy"]])) * 0.001
    CO2 = M - O
    return CO2

def _TableAddValue(Table, CODE, value, row=None):
    for i in Table.TableEntries:
        if i.AttValue("CODE") == CODE:
            if row == None: i.SetAttValue("MESSGROESSE",value)
            else:
                i_row = Table.TableEntries.ItemByKey(i.AttValue("NO") + row)
                i_row.SetAttValue("MESSGROESSE",value)
                
def _THGDiff_PuTService(data, _emissioncosts):
    if data["Energy"] == None: return 0
    O = (data["Ev_O"] * (_emissioncosts[data["Energy"]])) * 0.01
    M = (data["Ev_M"] * (_emissioncosts[data["Energy"]])) * 0.01
    SchadK = M - O
    return SchadK

def _THGVehUnit(data):
    if data["TSys"] == "Bus": OEV_H = (data["THG"] * data["FZG_M"]) - (data["THG"] * data["FZG_O"])
    else: OEV_H = (data["THG"] * data["FZG_M"] * data["Weight"]) - (data["THG"] * data["FZG_O"] * data["Weight"])
    return OEV_H