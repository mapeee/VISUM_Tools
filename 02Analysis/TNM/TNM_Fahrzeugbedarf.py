# -*- coding: utf-8 -*-
"""
Created on Mon Sep  1 09:47:03 2025

@author: peter
"""

import numpy as np
import pandas as pd
import sys
import time
from pathlib import Path
path = Path.home() / 'python32' / 'python_dir.txt'
path = open(path, mode='r').readlines()
path = path[0] +"\\"+'VISUM_Tools'+"\\"+'Line_import.txt'
f = open(path, mode='r').read().splitlines()
sys.path.insert(0,f[3])
from VisumPy.helpers import SetMulti

try:
    Visum
except:
    import win32com.client.dynamic
    Visum = win32com.client.dynamic.Dispatch("Visum.Visum.25")
    Visum.IO.loadversion(r"C:\Users\peter\hvv.de\B-GR-Bereich B - Planungstool hvv\VisumTest\VisumTest.ver")


Teilnetze = ["PI 1-3"]
Saisons = ["S", "F"]
Tage = ["MO", "DI", "MI", "DO", "FR", "SA", "SO"]

def Add2Visum(_TN, _AddList, _FZG = False):
    if _TN == False:
        for col_name, col_data in _AddList.items():
            if col_name == "LINENAME": continue
            SetMulti(Visum.Net.Lines, col_name, col_data.tolist(), True)
    else:
        FZGtable = Visum.Net.TableDefinitions.ItemByKey(f"{_TN} Fahrzeugbedarf")
        FZGtable.TableEntries.RemoveAll()
        FZGtable.AddMultiTableEntries(range(1, len(_AddList) + 1))
        SetMulti(FZGtable.TableEntries, "KENNZAHL", _AddList["KENNZAHL"].tolist(), True)
        SetMulti(FZGtable.TableEntries, "TYP", _AddList["TYP"].tolist(), True)
        for f in _FZG["CODE"].tolist():
            SetMulti(FZGtable.TableEntries, f, _AddList[f].tolist(), True)

def calculateFZG_Lines(_FZG, _VJTable, _Saisons, _Tage, _Gebiete):
    for s in _Saisons:
        for t in Tage:
            AddListVisum = pd.DataFrame()
            for line in np.sort(_VJTable["LINE"].unique()):
                AddListVisum = pd.concat([AddListVisum, _build_timeline(_VJTable, _FZG, s, t, line, _Gebiete)], ignore_index=True)
            Add2Visum(False, AddListVisum)

def calculateFZG_TN(_FZG, _VJTable, _Saisons, _Tage, _TN):
    columns = ["KENNZAHL", "TYP"] + _FZG["CODE"].tolist()
    AddListVisum = pd.DataFrame(columns = columns)
    for s in _Saisons:
        for t in _Tage:
            timeline = _build_timeline(_VJTable, _FZG, s, t)
            AddListVisum = _valuesToList(timeline, _FZG, AddListVisum, s, t)
    Add2Visum(_TN, AddListVisum, _FZG)
         
def _build_timeline(_VJTable, _FZG, _s, _t, _line = None, _Gebiete = None):
    _timeline = pd.DataFrame({"minute": np.arange(24 * 60)})
    _timeline["diff"] = 0
    rang = _FZG["RANG"].min()
    LineList = pd.DataFrame(columns=["LINENAME"])
    
    if _line:
        LineList.loc[len(LineList)] = [_line]
        lineV = Visum.Net.Lines.ItemByKey(_line)
        filterVJ = _VJTable[((_VJTable["LINE"] == _line) if _line else True) & (_VJTable[_s] == 1) & (_VJTable[_t] == 1)]
        _timeline["FZGM1.1"] = _fill_timeline(filterVJ).max()
    for index, f in _FZG.iterrows():
        filterVJ = _VJTable[((_VJTable["LINE"] == _line) if _line else True) & (_VJTable["VEH"] == f["CODE"]) & (_VJTable[_s] == 1) & (_VJTable[_t] == 1)]
        _timeline[f["CODE"]] = _fill_timeline(filterVJ)
        if _line:
            _timeline[f"{f['CODE']}M1.2"] = _timeline[f["CODE"]].max()
        _timeline[f"{f['CODE']}diff"] = _timeline[f["CODE"]].max() - _timeline[f["CODE"]]
        if f["RANG"] == rang:
            _timeline["diff"] = _timeline["diff"] + _timeline[f"{f['CODE']}diff"]
        else:
            _timeline[f"{f['CODE']}fromdiff"] = 0
            _timeline[f"{f['CODE']}new"] = 0
            _timeline.loc[_timeline[f["CODE"]] >= _timeline["diff"], f"{f['CODE']}fromdiff"] = _timeline["diff"]
            _timeline.loc[_timeline[f["CODE"]] >= _timeline["diff"], f"{f['CODE']}new"] = _timeline[f["CODE"]] - _timeline["diff"]
            _timeline["diff"] = _timeline["diff"] - _timeline[f"{f['CODE']}fromdiff"] + (_timeline[f"{f['CODE']}new"].max() - _timeline[f"{f['CODE']}new"])            
        rang = f["RANG"]
        if _line:
            col = f"{f['CODE']}new" if f"{f['CODE']}new" in _timeline.columns else f"{f['CODE']}"
            LineList[f"{_s}{_t}_{f['CODE']}M2"] = _timeline[col].max()
    if _line:
        m1_cols = _timeline.filter(like = "M1.2")
        _timeline["FZGM1_sum"] = m1_cols.sum(axis=1)
        for index, f in _FZG.iterrows():
            _timeline[f"{f['CODE']}M1"] = _timeline["FZGM1.1"] / _timeline["FZGM1_sum"] * _timeline[f"{f['CODE']}M1.2"]
            _timeline[f"{f['CODE']}M1"] = _timeline[f"{f['CODE']}M1"].fillna(0)
            LineList[f"{_s}{_t}_{f['CODE']}M1"] = _timeline[f"{f['CODE']}M1"].max()
        for g in _Gebiete:
            for index, f in _FZG.iterrows():
                Gesamt_STD = lineV.AttValue(f"{_s}{_t}_{f['CODE']}STD")
                g_STD = lineV.AttValue(f"{g}_{_s}{_t}_{f['CODE']}STD")
                if Gesamt_STD == 0:
                    LineList[f"{g}_{_s}{_t}_{f['CODE']}M1"] = 0
                    LineList[f"{g}_{_s}{_t}_{f['CODE']}M2"] = 0
                    continue
                FZGM1 = _timeline[f"{f['CODE']}M1"] .max()
                FZGM2 = LineList[f"{_s}{_t}_{f['CODE']}M2"][0]
                FZGM1 = FZGM1 * (g_STD / Gesamt_STD)
                FZGM2 = FZGM2 * (g_STD / Gesamt_STD)
                LineList[f"{g}_{_s}{_t}_{f['CODE']}M1"] = FZGM1
                LineList[f"{g}_{_s}{_t}_{f['CODE']}M2"] = FZGM2
        return LineList
    return _timeline
    
def _fill_timeline(_VJ, n_minutes = (24 * 60) + (6 * 60)):
    _timeline = np.zeros(n_minutes + 1, dtype=int)
    for dep, arr in zip(_VJ["DEP"], _VJ["ARR"]):
        _timeline[int(dep)] += 1
        _timeline[int(arr)] -= 1
        # _timeline[int(arr) + 1] -= 1 # when identical minute
    _timeline = np.cumsum(_timeline[:-1])
    _timeline[:360] = _timeline[:360] + _timeline[-360:] # normalize to 24h
    _timeline = _timeline[:-360]
    return _timeline


def _getFZG(_TN):
    _FZGTable = pd.DataFrame(Visum.Net.TableDefinitions.ItemByKey("Teilnetze").TableEntries
                                     .GetMultipleAttributes(["TEILNETZGRUPPE", "FAHRZEUGE"], True), columns = ["TEILNETZGRUPPE", "FAHRZEUGE"])
    _FZG = _FZGTable[_FZGTable["TEILNETZGRUPPE"] == _TN]["FAHRZEUGE"].tolist()
    _FZG = _FZG[0].split(", ")
    _VehComb = pd.DataFrame(Visum.Net.VehicleCombinations.GetMultipleAttributes(["CODE", "RANG"]), columns = ["CODE", "RANG"])
    _VehComb = _VehComb[_VehComb["CODE"].isin(_FZG)]
    _VehComb = _VehComb.sort_values(by = "RANG")
    return _VehComb

def _getGebiete(_TN):
    _GebieteTable = pd.DataFrame(Visum.Net.TableDefinitions.ItemByKey("Teilnetze").TableEntries
                                     .GetMultipleAttributes(["TEILNETZGRUPPE", "GEBIETE"], True), columns = ["TEILNETZGRUPPE", "GEBIETE"])
    _Gebiete = _GebieteTable[_GebieteTable["TEILNETZGRUPPE"] == _TN]["GEBIETE"].tolist()
    _Gebiete = _Gebiete[0].split(", ")
    return _Gebiete

def _getVJTable(_TN):
    attrList = [[r"LINEROUTE\LINE\TEILNETZGRUPPE", "LINENAME", "WOCHENTAGNO", "SAISONNO", "DEP", "ARR", r"MIN:VEHJOURNEYSECTIONS\VEHCOMB\CODE", r"MIN:VEHJOURNEYSECTIONS\VEHCOMB\RANG"],
                ["TEILNETZGRUPPE", "LINE", "TAGNO", "SAISONNO", "DEP", "ARR", "VEH", "VEHRANG"]]
    _VJTable = pd.DataFrame(Visum.Net.VehicleJourneys.GetMultipleAttributes(attrList[0], True), columns = attrList[1])
    _VJTable["DEP"], _VJTable["ARR"] = _VJTable["DEP"] / 60, _VJTable["ARR"] / 60
    attrList = ["NO", "MO", "DI", "MI", "DO", "FR", "SA", "SO"]
    WD = pd.DataFrame(Visum.Net.TableDefinitions.ItemByKey("Wochentage").TableEntries.GetMultipleAttributes(attrList), columns = attrList)
    WD = WD.rename(columns={"NO": "TAGNO"})
    S = pd.DataFrame(Visum.Net.TableDefinitions.ItemByKey("Saisons").TableEntries.GetMultipleAttributes(["NO","S", "F"]), columns = ["SAISONNO", "S", "F"])
    _VJTable = pd.merge(_VJTable, WD, on = "TAGNO", how = "inner")
    _VJTable = pd.merge(_VJTable, S, on = "SAISONNO", how = "inner")
    _VJTable = _VJTable[_VJTable["TEILNETZGRUPPE"] == _TN]
    return _VJTable

def _valuesToList(_timeline, _FZG, _AddListVisum, _s, _t):
    for val in ["Bedarf", "firstMin", "lastMin", "nMin", "ersetzt"]:
        addTolist = [f"{val} {_s} {_t}", val]
        for f in _FZG["CODE"].tolist():
            col = f"{f}new" if f"{f}new" in _timeline.columns else f"{f}"
            if val == "Bedarf":
                addTolist.append(_timeline[col].max())
            elif val == "firstMin":
                maxMinutes = _timeline.loc[_timeline[col].idxmax(), "minute"]
                addTolist.append(time.strftime("%H:%M", time.gmtime(maxMinutes * 60)))
            elif val == "lastMin":
                maxMinutes = _timeline.loc[_timeline[col].iloc[::-1].idxmax(), "minute"]
                addTolist.append(time.strftime("%H:%M", time.gmtime(maxMinutes * 60)))
            elif val == "nMin":
                addTolist.append((_timeline[col] == _timeline[col].max()).sum())
            elif val == "ersetzt":
                if f"{f}new" in _timeline.columns:
                    addTolist.append(_timeline[f"{f}"].max() - _timeline[f"{f}new"].max())
                else:
                    addTolist.append(0)
            else:
                pass
        _AddListVisum.loc[len(_AddListVisum)] = addTolist 
    return _AddListVisum

for TN in Teilnetze:
    Gebiete = _getGebiete(TN)
    VJTable = _getVJTable(TN)
    FZG = _getFZG(TN)
    calculateFZG_Lines(FZG, VJTable, Saisons, Tage, Gebiete)
    calculateFZG_TN(FZG, VJTable, Saisons, Tage, TN)
