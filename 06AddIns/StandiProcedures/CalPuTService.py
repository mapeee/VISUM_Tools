# -*- coding: utf-8 -*-
"""
Created on Fri Feb  27 13:33:29 2024
"""
import math
import numpy as np
import pandas as pd
from VisumPy.AddIn import AddIn
from VisumPy.helpers import SetMulti
_ = AddIn.gettext


def PuTCosts(Visum,bc):
    Visum.Log(20480,_("PuT costs: Starting!"))
    
    energycosts = dict((x, y) for x, y in Visum.Net.TableDefinitions.ItemByKey("Standi-Energie").TableEntries.GetMultipleAttributes(["ENERGIEART","ENERGIEPREIS"]))
    stopenergy_a = dict((x, y) for x, y in Visum.Net.TableDefinitions.ItemByKey("Standi-Energie").TableEntries.GetMultipleAttributes(["ENERGIEART","HALTEVERBRAUCH_A"]))
    stopenergy_b = dict((x, y) for x, y in Visum.Net.TableDefinitions.ItemByKey("Standi-Energie").TableEntries.GetMultipleAttributes(["ENERGIEART","HALTEVERBRAUCH_B"]))
    parameter = dict((x, y) for x, y in Visum.Net.TableDefinitions.ItemByKey("Standi-Parameter").TableEntries.GetMultipleAttributes(["CODE","WERT"]))
    HF_PT = parameter["HF_OEV"]
    WZ = parameter["WZ"] / 100 + 1
    
    if not _CheckTSys(Visum):
        return False
    if not _CheckVehComb(Visum):
        return False
        
    if bc:_ev, _fk, _nVU, _ll, _pk, _ek, _ukll = "ENERGIEV_O", "FAHRZEUGK_O", "FZG_O", "LL_O", "PERSONALK_O", "ENERGIEK_O", "UKLL_O"
    else: _ev, _fk, _nVU, _ll, _pk, _ek, _ukll = "ENERGIEV_M", "FAHRZEUGK_M", "FZG_M", "LL_M", "PERSONALK_M", "ENERGIEK_M", "UKLL_M"
    
    SetMulti(Visum.Net.VehicleCombinations,_ev,[0]*Visum.Net.VehicleCombinations.Count,False)
    SetMulti(Visum.Net.VehicleUnits,_fk,[0]*Visum.Net.VehicleUnits.Count,False)
    SetMulti(Visum.Net.VehicleUnits,_ll,[0]*Visum.Net.VehicleUnits.Count,False)
    SetMulti(Visum.Net.Lines,_ll,[0]*Visum.Net.Lines.Count,False)
    SetMulti(Visum.Net.Lines,_pk,[0]*Visum.Net.Lines.Count,False)
    SetMulti(Visum.Net.Lines,_ek,[0]*Visum.Net.Lines.Count,False)
    SetMulti(Visum.Net.Lines,_ukll,[0]*Visum.Net.Lines.Count,False)
    
    #VehicleCosts
    ll_l = []
    c_l = []
    for VehUnit in Visum.Net.VehicleUnits.GetAll:
        if VehUnit.AttValue(r"DISTINCT:VEHCOMBS\NO") == "":
            ll_l.append(0)
            c_l.append(0)
            continue
        
        #Kapitaldienst
        if VehUnit.AttValue("TSYSSET") == "Bus": c1 = VehUnit.AttValue("ANSCHAFFUNGSK") * VehUnit.AttValue(_nVU) * 0.0928
        else: c1 = VehUnit.AttValue("ANSCHAFFUNGSK") * VehUnit.AttValue(_nVU) * 0.0428
        
        #UkT
        if VehUnit.AttValue("TSYSSET") == "Bus": c2 = VehUnit.AttValue(_nVU) * VehUnit.AttValue("UKT") * 0.001
        else: c2 = VehUnit.AttValue(_nVU) * VehUnit.AttValue("UKT") * VehUnit.AttValue("LEERMASSE") * 0.001
        
        #UkLL - VehicleUnit
        VehCombos = VehUnit.AttValue(r"DISTINCT:VEHCOMBS\NO")
        VehCombos = list(map(int, VehCombos.split(",")))
        ll = 0
        for i in VehCombos:
            Combo = Visum.Net.VehicleCombinations.ItemByKey(i)
            if Combo.AttValue(r"SUM:VEHJOURNEYSECTIONS\LENGTH") == None: continue
            nCombo = list(map(int, Combo.AttValue(r"CONCATENATE:VEHUNITS\NO").split(","))).count(int(VehUnit.AttValue("NO")))
            ll = ll + ((Combo.AttValue(r"SUM:VEHJOURNEYSECTIONS\LENGTH") * nCombo) * HF_PT)
            
        ll_l.append(round(ll / HF_PT / 1000, 1))
        if VehUnit.AttValue("TSYSSET") == "Bus": c3 = ll * VehUnit.AttValue("UKLL") * 0.001
        else: c3 = VehUnit.AttValue("UKLL") * (ll * VehUnit.AttValue("LEERMASSE") * 0.001) * 0.001
        c_l.append(round(c1+c2+c3, 1))
  
    SetMulti(Visum.Net.VehicleUnits,_ll,ll_l,False)
    SetMulti(Visum.Net.VehicleUnits,_fk,c_l,False)
    
    #Energy costs - distance
    Weight = np.array(Visum.Net.VehicleJourneySections.GetMultiAttValues(r"VEHCOMB\SUM:VEHUNITS\LEERMASSE"))[:,1]
    Length = np.array(Visum.Net.VehicleJourneySections.GetMultiAttValues("LENGTH"))[:,1]
    EVS = np.array(Visum.Net.VehicleJourneySections.GetMultiAttValues(r"VEHCOMB\MAX:VEHUNITS\EVS"))[:,1]
    Energy = np.array(Visum.Net.VehicleJourneySections.GetMultiAttValues(r"VEHCOMB\MIN:VEHUNITS\ENERGIE"))[:,1]
    TSys = np.array(Visum.Net.VehicleJourneySections.GetMultiAttValues("TSYSCODE"))[:,1]
    Line = np.array(Visum.Net.VehicleJourneySections.GetMultiAttValues("VEHJOURNEY\LINEROUTE\LINENAME"))[:,1]
    UKLL = np.array(Visum.Net.VehicleJourneySections.GetMultiAttValues(r"VEHCOMB\SUM:VEHUNITS\UKLL"))[:,1]
    UKTKM = np.array(Visum.Net.VehicleJourneySections.GetMultiAttValues(r"VEHCOMB\SUM:VEHUNITS\UKTKM"))[:,1]
    STOPS = np.array(Visum.Net.VehicleJourneySections.GetMultiAttValues(r"VEHJOURNEY\STOPSSERVED"))[:,1]
    DURATION = np.array(Visum.Net.VehicleJourneySections.GetMultiAttValues(r"DURATION"))[:,1]
    SUM_DEP = np.array(Visum.Net.VehicleJourneySections.GetMultiAttValues(r"VEHJOURNEY\SUM:VEHJOURNEYITEMS\EXTDEPARTURE"))[:,1]
    SUM_ARR = np.array(Visum.Net.VehicleJourneySections.GetMultiAttValues(r"VEHJOURNEY\SUM:VEHJOURNEYITEMS\EXTARRIVAL"))[:,1]
    VehCombNo = np.array(Visum.Net.VehicleJourneySections.GetMultiAttValues("VEHCOMBNO"))[:,1]
    
    VJ = pd.DataFrame(np.stack((Weight, Length, EVS, Energy, TSys, Line, UKLL, UKTKM, STOPS, DURATION, SUM_DEP, SUM_ARR, VehCombNo), axis=1))
    VJ.columns = ["Weight","Length","EVS","Energy","TSys","Line","UKLL","UKTKM","STOPS","DURATION","SUM_DEP","SUM_ARR","VehCombNo"]
    VJ[["_ek","_ev"]] = VJ.apply(_EnergyCostVehSect, _energycosts = energycosts, _stopenergy_a = stopenergy_a, _stopenergy_b = stopenergy_b, _HF_PT = HF_PT,
                                 axis=1, result_type="expand")
    
    #UkLL - Lines
    VJ["_ukll"] = VJ.apply(_LengthOpCostLine, _HF_PT = HF_PT, axis=1)
    
    ek_l = _SetMultiList(VJ,"Line","_ek",Visum.Net.Lines.GetMultipleAttributes(["NAME"]))    
    evComb = _SetMultiList(VJ,"VehCombNo","_ev",Visum.Net.VehicleCombinations.GetMultipleAttributes(["NO"]))
    ukll_l = _SetMultiList(VJ,"Line","_ukll",Visum.Net.Lines.GetMultipleAttributes(["NAME"]))

    SetMulti(Visum.Net.Lines,_ek,ek_l,False)   
    SetMulti(Visum.Net.VehicleCombinations,_ev,evComb,False)
    SetMulti(Visum.Net.Lines,_ukll,ukll_l,False)
    SetMulti(Visum.Net.Lines,_ll,np.array(Visum.Net.Lines.GetMultiAttValues("SERVICEKM(AP)"))[:,1],False)
    
    #Personnel costs
    pk = []
    for line in Visum.Net.Lines.GetAll:
        h = line.AttValue(r"SERVICETIME(AP)") / 60 / 60
        h = h * WZ * HF_PT
        if line.AttValue("TSYSCODE") == "Bus": pk.append(h * 39 * 0.001)
        elif line.AttValue("TSYSCODE") == "W": pk.append(h * 60 * 0.001)
        else: pk.append(h * 46 * 0.001)
    SetMulti(Visum.Net.Lines,_pk,pk,False)

    return True
        
def VehicleUnits(Visum,bc,):
    Visum.Log(20480,_("PuT vehicle units: Starting!"))
    
    parameter = dict((x, y) for x, y in Visum.Net.TableDefinitions.ItemByKey("Standi-Parameter").TableEntries.GetMultipleAttributes(["CODE","WERT"]))
    WZ = parameter["WZ"] / 100 + 1
    BWR = parameter["BWR"] / 100 + 1
    
    if bc: _nVU, _nLine = "FZG_O", "FAHRTEN_O"
    else: _nVU, _nLine = "FZG_M", "FAHRTEN_M"
    
    SetMulti(Visum.Net.VehicleUnits,_nVU,[0]*Visum.Net.VehicleUnits.Count,False)
    SetMulti(Visum.Net.Lines,_nLine,[0]*Visum.Net.Lines.Count,False)

    mins = range(360,540)
    
    Dep = (np.array(Visum.Net.VehicleJourneySections.GetMultiAttValues("DEP"))[:,1]/60).astype(int)
    Arr = (np.array(Visum.Net.VehicleJourneySections.GetMultiAttValues("ARR"))[:,1]/60).astype(int)
    Lines = np.array(Visum.Net.VehicleJourneySections.GetMultiAttValues(r"VEHJOURNEY\LINENAME"))[:,1]
    VehNO = np.array(Visum.Net.VehicleJourneySections.GetMultiAttValues("VEHCOMBNO"))[:,1]
    
    VJ = pd.DataFrame(np.stack((Dep, Arr, Lines, VehNO), axis=1))
    VJ.columns = ["Dep", "Arr", "Lines", "VehNO"]
    VJ["Dep"] = VJ["Dep"].astype(int)
    VJ["Arr"] = VJ["Arr"].astype(int)
    VJ["Lines"] = VJ["Lines"].astype(str)
    VJ["VehNO"] = VJ["VehNO"].astype(str)
    VJ_VU = VJ.loc[VJ["VehNO"] != 'None']
    VJ_VU["VehNO"] = VJ_VU["VehNO"].astype(float)
    VJ_VU["VehNO"] = VJ_VU["VehNO"].astype(int)
    
    #VehicleUnits
    FZG = []
    for VehUnit in Visum.Net.VehicleUnits.GetAll:
        nVHUnits = 0
        VehComb = VehUnit.AttValue(r"DISTINCT:VEHCOMBS\NO")
        if VehComb == "":
            FZG.append(0)
            continue
    
        VehComb = list(map(int, VehComb.split(",")))
        VJ_Comb = VJ_VU[VJ_VU["VehNO"].isin(VehComb)]
        
        for VJ_Comb_Lines in VJ_Comb["Lines"].unique():
            Vehs_l = 0
            VJ_Comb_l = VJ_Comb[VJ_Comb["Lines"]==VJ_Comb_Lines]
            for minute in mins:
                Vehs = 0
                n = VJ_Comb_l[(VJ_Comb_l["Dep"]<=minute) & (VJ_Comb_l["Arr"]>minute)]
                item_counts = n["VehNO"].value_counts()
                for i in VehComb:
                    if not i in item_counts: continue
                    NVehComb = item_counts[i]
                    nVUVehCom = list(map(int, Visum.Net.VehicleCombinations.ItemByKey(i).AttValue(r"CONCATENATE:VEHUNITS\NO").split(","))).count(VehUnit.AttValue("NO"))
                    Vehs = (NVehComb * nVUVehCom) + Vehs
                Vehs_l = max(Vehs_l, Vehs)
                
            nVHUnits = nVHUnits + round(Vehs_l * WZ, 2)

        FZG.append(round(nVHUnits * BWR, 2))
    SetMulti(Visum.Net.VehicleUnits,_nVU,FZG,False)
  
    #overlapping VJ Lines
    FzgE = []
    for Line in Visum.Net.Lines.GetAll:
        nLines = 0
        
        VJ_Lines = VJ[VJ["Lines"]==Line.AttValue("NAME")]
        for minute in mins:
            n = VJ_Lines[(VJ_Lines["Dep"]<=minute) & (VJ_Lines["Arr"]>minute)]
            nLines = max(nLines, len(n))
        
        FzgE.append(nLines)  
    SetMulti(Visum.Net.Lines,_nLine,FzgE,False)
    
def _CheckTSys(Visum):
    for TSys in Visum.Net.TSystems.GetAll:
        if TSys.AttValue("TYPE") != "PUT": continue
        if TSys.AttValue("CODE") not in ["S","Bus","U","RV","FV","W"]:
            Visum.Log(12288,_("TSys %s is not predefined!") %(TSys.AttValue("CODE")))
            return False
    return True

def _CheckVehComb(Visum):
    for Comb in Visum.Net.VehicleCombinations.GetAll:
        if Comb.AttValue(r"MIN:VEHUNITS\ENERGIE") not in ["Diesel", "Strom konv.", "Strom regen.", "eFuel", "H2"]:
            Visum.Log(12288,_("VehicleCombination %s is not predefined!") %(Comb.AttValue("Name")))
            return False
    return True

def _EnergyCostVehSect(data, _energycosts, _stopenergy_a, _stopenergy_b, _HF_PT):
    
    if data["Weight"] == None: return 0
    if data["TSys"] == "Bus":
        ek = data["Length"] * data["EVS"] * _HF_PT * 0.001 * (_energycosts[data["Energy"]])
        ev = data["Length"] * data["EVS"] * _HF_PT * 0.001
    else:
        ek = data["Weight"] * data["Length"] * data["EVS"] * _HF_PT * 0.001 * 0.001 * (_energycosts[data["Energy"]])
        ev = data["Weight"] * data["Length"] * data["EVS"] * _HF_PT * 0.001 * 0.001
    #Stops
    if data["TSys"] in ["RV","S"]:
        Tf = data["DURATION"] / 60
        Th = (data["SUM_DEP"] / 60) - (data["SUM_ARR"] / 60)
        a = _stopenergy_a[data["Energy"]]
        b = _stopenergy_b[data["Energy"]]
        speed = 3.6 / (a * (data["STOPS"] - 1)) * ((55.6 * (Tf-Th))-math.sqrt(max(pow(55.6 * (Tf-Th), 2) - (2 * a * data["Length"] * 1000 * (data["STOPS"] - 1)),0)))
        
        e = b * pow(speed, 2) * data["Weight"] * 0.000001
        
        ek_stops = e * (data["STOPS"] - 1) * _HF_PT * 0.001 * (_energycosts[data["Energy"]])
        ev_stops = e * (data["STOPS"] - 1) * _HF_PT * 0.001

        ek = ek + ek_stops
        ev = ev + ev_stops

    return ek, ev

def _LengthOpCostLine(data, _HF_PT):
    if data["UKLL"] == None: return 0
    if data["TSys"] == "Bus":
        v = data["Length"] * data["UKLL"] * _HF_PT * 0.001
    else:
        v = data["Length"] * data["UKTKM"] * _HF_PT * 0.001
    return v

def _SetMultiList(data,groupby,sumval,v):
    l = []
    l_gb = data.groupby([groupby])[sumval].sum()
    for i in v:
        if i[0] in l_gb.index: l.append(l_gb[i[0]])
        else: l.append(0)
    return l