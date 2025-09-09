# -*- coding: utf-8 -*-
"""
Created on Thu Aug 28 20:29:11 2025

@author: peter
"""

import pandas as pd
from VisumPy.helpers import SetMulti

Gebiete = ["FHH", "PI", "SE"]
Saisons = ["S", "F"]
Tage = ["MO", "DI", "MI", "DO", "FR", "SA", "SO"]
Einheiten = ["KM", "STD"]
FZG = ["SB", "GB"]

PTD = pd.DataFrame(Visum.Net.TerritoryPuTDetails.GetMultipleAttributes(["TERRITORYCODE", "LINENAME", r"SERVICEKM(AP)",r"SERVICETIME(AP)", 
                                                                        "VEHCOMBCODE", r"VEHJOURNEY\WOCHENTAGNO", r"VEHJOURNEY\SAISONNO"], True))
PTD.columns = ["Gebiet", "LINE", "KM", "STD", "FZG", "TAGNO", "SAISONNO"]
PTD["STD"] = PTD["STD"] / 60 / 60

WD = pd.DataFrame(Visum.Net.TableDefinitions.ItemByKey("Wochentage").TableEntries.GetMultipleAttributes(["NO", "MO", "DI", "MI", "DO", "FR", "SA", "SO"]))
WD.columns = ["TAGNO", "MO", "DI", "MI", "DO", "FR", "SA", "SO"]

S = pd.DataFrame(Visum.Net.TableDefinitions.ItemByKey("Saisons").TableEntries.GetMultipleAttributes(["NO","S", "F"]))
S.columns = ["SAISONNO", "S", "F"]

PTD = pd.merge(PTD, WD, on="TAGNO", how="inner")
PTD = pd.merge(PTD, S, on="SAISONNO", how="inner")


for s in Saisons:
    PTDs = PTD[PTD[s] == 1]
    for t in Tage:
        PTDt = PTDs[PTDs[t] == 1]
        for f in FZG:
            PTDf = PTDt[PTDt["FZG"] == f]
            result = (PTDf.groupby("LINE", as_index=False).agg({"STD": "sum"}))
            NAME = f"{s}{t}_{f}STD"
            edict = dict(zip(result["LINE"], result["STD"]))
            SetMulti(Visum.Net.Lines, NAME, [edict.get(line, 0) for _, line in Visum.Net.Lines.GetMultiAttValues("NAME",True)], True)
            for g in Gebiete:
                PTDg = PTDf[PTDf["Gebiet"] == g]
                result = (PTDg.groupby("LINE", as_index=False).agg({"KM": "sum", "STD": "sum"}))
                for e in Einheiten:
                    NAME = f"{g}_{s}{t}_{f}{e}"
                    edict = dict(zip(result["LINE"], result[e]))
                    SetMulti(Visum.Net.Lines, NAME, [edict.get(line, 0) for _, line in Visum.Net.Lines.GetMultiAttValues("NAME",True)], True)