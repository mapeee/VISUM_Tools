# -*- coding: utf-8 -*-
"""
Created on Fri Aug 29 14:56:48 2025

@author: peter
"""
import pandas as pd
from VisumPy.helpers import SetMulti

Gebiete = ["TN", "FHH", "PI", "SE"]
FZG = ["SB", "GB"]
Saisons = ["S", "F"]
Teilnetze = ["PI 1-3"]


def calculateKMSTD(_TN, _MengenTable):
    LinesTable = _getLinesTable(_TN, _MengenTable)
    AddList = pd.DataFrame(columns=["KENNZAHL", "KM", "STD"])
    for row in _MengenTable.TableEntries.GetMultiAttValues("KENNZAHL"):
        if "TN_" == row[1][:3]:
            value_KM = LinesTable.filter(like=f"{row[1][3:]}_KM(AP)").sum().sum()
            value_STD = LinesTable.filter(like=f"{row[1][3:]}_STD(AP)").sum().sum()
        else:
            value_KM = LinesTable[f"{row[1]}_KM(AP)"].sum()
            value_STD = LinesTable[f"{row[1]}_STD(AP)"].sum()
        AddList.loc[len(AddList)] = [row[1], value_KM, value_STD]

    SetMulti(Visum.Net.TableDefinitions.ItemByKey(f"{_TN} Mengen").TableEntries, "KM", AddList["KM"].tolist(), True)
    SetMulti(Visum.Net.TableDefinitions.ItemByKey(f"{_TN} Mengen").TableEntries, "STD", AddList["STD"].tolist(), True)

def _getLinesTable(_TN, _MengenTable):
    attrList = [f"{x[1]}_{suffix}(AP)" for x in _MengenTable.TableEntries.GetMultiAttValues("KENNZAHL") if "TN_" not in x[1] for suffix in ("KM", "STD")]
    attrList.append("TEILNETZGRUPPE")
    _LinesTable = pd.DataFrame(Visum.Net.Lines.GetMultipleAttributes(attrList, True), columns = attrList)
    _LinesTable = _LinesTable[_LinesTable["TEILNETZGRUPPE"] == _TN]
    return _LinesTable


for TN in Teilnetze:
    MengenTable = Visum.Net.TableDefinitions.ItemByKey(f"{TN} Mengen")
    MengenTable.TableEntries.RemoveAll()
    attrList = []
    for g in Gebiete:
        for s in Saisons:
            for f in FZG:
                attrList.append(f"{g}_{s}_{f}")
    MengenTable.AddMultiTableEntries(range(1, len(attrList) + 1))
    SetMulti(MengenTable.TableEntries, "KENNZAHL", attrList)
    calculateKMSTD(TN, MengenTable)