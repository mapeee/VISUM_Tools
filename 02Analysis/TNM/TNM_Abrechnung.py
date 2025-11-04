# -*- coding: utf-8 -*-
"""
Created on Mon Sep  1 09:47:03 2025

@author: peter
"""

import pandas as pd
from VisumPy.helpers import SetMulti

Teilnetze = ["PI 1-3"]
Parameter = pd.DataFrame(Visum.Net.TableDefinitions.ItemByKey("Parameter").TableEntries.GetMultipleAttributes(["KENNZAHL", "WERT"]), columns = ["KENNZAHL", "WERT"])


def calculateAbrechnung(_TN, _Gebiete, _Parameter):
    AbrechnungTable = _getAbrechnungTable(_TN)
    LinesTable = _getLinesTable(_TN)
    MengenTable = pd.DataFrame(Visum.Net.TableDefinitions.ItemByKey(f"{_TN} Mengen").TableEntries.GetMultipleAttributes(["KENNZAHL","KM","STD"]))
    MengenTable.columns = ["KENNZAHL", "KM", "STD"]
    MengenTable["KENNZAHL2"] = MengenTable["KENNZAHL"].str.replace(r"_[^_]+$", "", regex=True)
    
    FZG = _getFZG(_TN)
    FZGtable = pd.DataFrame(Visum.Net.TableDefinitions.ItemByKey(f"{_TN} Fahrzeugbedarf").TableEntries.GetMultipleAttributes(["TYP"] + FZG))
    FZGtable.columns = ["TYP"] + FZG
    FZGtable = FZGtable[FZGtable["TYP"] == "Bedarf"]
    for f in FZG:
        FZGtable[f] = FZGtable[f].astype(int)
    
    Schulwochen = int(_Parameter[_Parameter["KENNZAHL"] == "Schulwochen"]["WERT"].iloc[0])
    Ferienwochen = int(_Parameter[_Parameter["KENNZAHL"] == "Ferienwochen"]["WERT"].iloc[0])
    TA_FZG = _Parameter[_Parameter["KENNZAHL"] == "TA_FZG"]["WERT"].iloc[0]
    
    
    for index, row in AbrechnungTable.iterrows():
        if row["KENNZAHL"][-3:] == "_KM":
            FZG = row["KENNZAHL"].split("_")[0]
            g_sum = 0
            for g in _Gebiete:
                S = MengenTable[MengenTable["KENNZAHL"]==f"{g}_S_{FZG}"]["KM"].iloc[0] * Schulwochen
                F = MengenTable[MengenTable["KENNZAHL"]==f"{g}_F_{FZG}"]["KM"].iloc[0] * Ferienwochen
                AbrechnungTable.at[index, f"{g}_MENGEN"] = S + F
                g_sum += S + F
            AbrechnungTable.at[index, "GESAMT_MENGEN"] = g_sum
        elif row["KENNZAHL"] == "STD":
            g_sum = 0
            for g in _Gebiete:
                S = MengenTable[MengenTable["KENNZAHL2"]==f"{g}_S"]["STD"].sum() * Schulwochen
                F = MengenTable[MengenTable["KENNZAHL2"]==f"{g}_F"]["STD"].sum() * Ferienwochen
                AbrechnungTable.at[index, f"{g}_MENGEN"] = S + F
                g_sum += S + F
            AbrechnungTable.at[index, "GESAMT_MENGEN"] = g_sum
        else:
            AbrechnungTable.at[index, "GESAMT_MENGEN"] = FZGtable[row["KENNZAHL"]].max()
            for g in _Gebiete:
                pattern = rf"^{g}.*{row['KENNZAHL']}_{TA_FZG}.*$"
                m_cols = LinesTable.loc[:, LinesTable.columns.str.match(pattern)]
                col_sums = m_cols.sum(axis=0)
                AbrechnungTable.at[index, f"{g}_MENGEN"] = col_sums.max()
        
    _editAbrechnungTable(_TN, _Gebiete, AbrechnungTable)
    

def _editAbrechnungTable(_TN, _Gebiete, _AbrechnungTable):
    for g in _Gebiete:
        SetMulti(Visum.Net.TableDefinitions.ItemByKey(f"{_TN} Abrechnung").TableEntries, f"{g}_MENGEN", _AbrechnungTable[f"{g}_MENGEN"].tolist(), True)
    SetMulti(Visum.Net.TableDefinitions.ItemByKey(f"{_TN} Abrechnung").TableEntries, "GESAMT_MENGEN", _AbrechnungTable["GESAMT_MENGEN"].tolist(), True)
    

def _getAbrechnungTable(_TN):
    attrList = []
    for i in Visum.Net.TableDefinitions.ItemByKey(f"{_TN} Abrechnung").TableEntries.Attributes.GetAll:
        if i.IsUserDefined and i.Editable and i.ID != "VERRECHNUNGSSATZ":
            attrList.append(i.ID)
    _AbrechnungTable = pd.DataFrame(Visum.Net.TableDefinitions.ItemByKey(f"{_TN} Abrechnung").TableEntries
                                     .GetMultipleAttributes(attrList, True), columns = attrList)
    return _AbrechnungTable


def _getFZG(_TN):
    _FZGTable = pd.DataFrame(Visum.Net.TableDefinitions.ItemByKey("Teilnetze").TableEntries
                                     .GetMultipleAttributes(["TEILNETZGRUPPE", "FAHRZEUGE"], True), columns = ["TEILNETZGRUPPE", "FAHRZEUGE"])
    _FZG = _FZGTable[_FZGTable["TEILNETZGRUPPE"] == _TN]["FAHRZEUGE"].tolist()
    _FZG = _FZG[0].split(", ")

    return _FZG


def _getGebiete(_TN):
    _GebieteTable = pd.DataFrame(Visum.Net.TableDefinitions.ItemByKey("Teilnetze").TableEntries
                                     .GetMultipleAttributes(["TEILNETZGRUPPE", "GEBIETE"], True), columns = ["TEILNETZGRUPPE", "GEBIETE"])
    _Gebiete = _GebieteTable[_GebieteTable["TEILNETZGRUPPE"] == _TN]["GEBIETE"].tolist()
    _Gebiete = _Gebiete[0].split(", ")
    return _Gebiete

def _getLinesTable(_TN):
    attrList = [f"{x.ID}({suffix})" for x in Visum.Net.Lines.Attributes.GetAll 
                if any(tag in x.ID for tag in ("M1", "M2")) 
                for suffix in ("Mo", "Di", "Mi", "Do", "Fr", "Sa", "So")]
    attrList.append("TEILNETZGRUPPE")
    _LinesTable = pd.DataFrame(Visum.Net.Lines.GetMultipleAttributes(attrList, True), columns = attrList)
    _LinesTable = _LinesTable[_LinesTable["TEILNETZGRUPPE"] == _TN]
    return _LinesTable

for TN in Teilnetze:
    Gebiete = _getGebiete(TN)
    calculateAbrechnung(TN, Gebiete, Parameter)