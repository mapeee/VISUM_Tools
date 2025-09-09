# -*- coding: utf-8 -*-
"""
Created on Fri Aug 29 14:56:48 2025

@author: peter
"""
import pandas as pd
from VisumPy.helpers import SetMulti


Tage = ["MO", "DI", "MI", "DO", "FR", "SA", "SO"]
Saisons = ["S", "F"]
Teilnetze = ["PI 1-3"]


def calculateKMSTD(_TN, _MengenTable, _Tage):
    LinesTable = _getLinesTable(_TN)
    _MengenTable["KENNZAHL2"] = _MengenTable["KENNZAHL"].str.split("_", n=1).str[1]
    
    AddList = pd.DataFrame(columns=["KENNZAHL", "KENNZAHL2", "KM", "STD"])
    for index, row in _MengenTable.iterrows():
        if row["KENNZAHL"][:2] != "P_":
            value_KM = 0
            value_STD = 0
            for t in _Tage:
                NAME = row["KENNZAHL"]
                NAME = NAME.replace("_F_", f"_F{t}_")
                NAME = NAME.replace("_S_", f"_S{t}_")
                value_KM += LinesTable[f"{NAME}KM"].sum()
                value_STD += LinesTable[f"{NAME}STD"].sum()
            AddList.loc[len(AddList)] = [row["KENNZAHL"], row["KENNZAHL2"], value_KM, value_STD]
        else:
            value_KM = AddList.loc[AddList["KENNZAHL2"] == row["KENNZAHL2"], "KM"].sum()
            value_STD = AddList.loc[AddList["KENNZAHL2"] == row["KENNZAHL2"], "STD"].sum()
            AddList.loc[len(AddList)] = [row["KENNZAHL"], row["KENNZAHL2"], value_KM, value_STD]

    SetMulti(Visum.Net.TableDefinitions.ItemByKey(f"{_TN} Mengen").TableEntries, "KM", AddList["KM"].tolist(), True)
    SetMulti(Visum.Net.TableDefinitions.ItemByKey(f"{_TN} Mengen").TableEntries, "STD", AddList["STD"].tolist(), True)

def _getLinesTable(_TN):
    attrList = []
    for i in Visum.Net.Lines.Attributes.GetAll:
        if i.IsUserDefined:
            attrList.append(i.ID)
    
    _LinesTable = pd.DataFrame(Visum.Net.Lines.GetMultipleAttributes(attrList, True), columns = attrList)
    _LinesTable = _LinesTable[_LinesTable["TEILNETZGRUPPE"] == _TN]
    return _LinesTable


for TN in Teilnetze:
    MengenTable = pd.DataFrame(Visum.Net.TableDefinitions.ItemByKey(f"{TN} Mengen").TableEntries.GetMultipleAttributes(["KENNZAHL"]), columns = ["KENNZAHL"])
    calculateKMSTD(TN, MengenTable, Tage)