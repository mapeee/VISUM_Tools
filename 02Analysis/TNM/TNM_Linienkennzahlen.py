# -*- coding: utf-8 -*-
"""
Created on Thu Aug 28 20:29:11 2025
@author: peter
"""

import pandas as pd
from VisumPy.helpers import SetMulti
from utils.TNM_Checks import check_TN
from utils.TNM_UDA import create_UDA


def get_territories():
    df_t_put = pd.DataFrame(Visum.Net.TerritoryPuTDetails.GetMultipleAttributes([r"TERRITORYCODE"], True),
                            columns = ["CODE"])
    t_unique = df_t_put['CODE'].drop_duplicates().values.tolist()
    return t_unique

def get_vehicleCombinations():
    df_vj = pd.DataFrame(Visum.Net.VehicleJourneySections.GetMultipleAttributes([r"VEHCOMB\CODE"], True), columns = ["FZG"])
    vc_unique = df_vj['FZG'].drop_duplicates().tolist()
    return vc_unique

###Checks######
check_TN(Visum)
###############

###Create UDAs###
create_UDA(Visum)
#################

Gebiete = get_territories()
Saisons = ["S", "F"]
Tage = ["AP", "Mo", "Di", "Mi", "Do", "Fr", "Sa", "So"]
Einheiten = ["KM", "STD"]
FZG = get_vehicleCombinations()

PTD = pd.DataFrame(Visum.Net.TerritoryPuTDetails.GetMultipleAttributes(["TERRITORYCODE", "LINENAME", "VEHCOMBCODE",
                                                                        "SERVICEKM(AP)", "SERVICETIME(AP)",
                                                                        "SERVICEKM(MO)", "SERVICEKM(DI)",
                                                                        "SERVICEKM(MI)", "SERVICEKM(DO)",
                                                                        "SERVICEKM(FR)", "SERVICEKM(SA)",
                                                                        "SERVICEKM(SO)",
                                                                        "SERVICETIME(MO)", "SERVICETIME(DI)",
                                                                        "SERVICETIME(MI)", "SERVICETIME(DO)",
                                                                        "SERVICETIME(FR)", "SERVICETIME(SA)",
                                                                        "SERVICETIME(SO)",
                                                                        r"VEHJOURNEY\MIN:VEHJOURNEYSECTIONS\S",
                                                                        r"VEHJOURNEY\MIN:VEHJOURNEYSECTIONS\F"], True))
PTD.columns = ["Gebiet", "LINE", "FZG", "APKM", "APSTD", "MoKM", "DiKM", "MiKM", "DoKM", "FrKM", "SaKM", "SoKM",
               "MoSTD", "DiSTD", "MiSTD", "DoSTD", "FrSTD", "SaSTD", "SoSTD", "S", "F"]
PTD[["APSTD", "MoSTD", "DiSTD", "MiSTD", "DoSTD", "FrSTD", "SaSTD", "SoSTD"]] /= 3600

for s in Saisons:
    PTDs = PTD[PTD[s] == 1]
    for f in FZG:
        PTDf = PTDs[PTDs["FZG"] == f]
        for t in Tage:
            Kenzahl = f"{s}_{f}_STD({t})"
            edict = PTDf.groupby("LINE")[f"{t}STD"].sum().to_dict()
            SetMulti(Visum.Net.Lines, Kenzahl, [edict.get(line, 0) for _, line in Visum.Net.Lines.GetMultiAttValues("NAME",True)], True)
            for g in Gebiete:
                PTDg = PTDf[PTDf["Gebiet"] == g]
                for e in Einheiten:
                    Kenzahl = f"{g}_{s}_{f}_{e}({t})"
                    edict = PTDg.groupby("LINE")[f"{t}{e}"].sum().to_dict()
                    SetMulti(Visum.Net.Lines, Kenzahl, [edict.get(line, 0) for _, line in Visum.Net.Lines.GetMultiAttValues("NAME",True)], True)
                    
Visum.Log(20480, "Linienkennzahlen: berechnet")