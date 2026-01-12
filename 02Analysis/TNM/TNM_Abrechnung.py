# -*- coding: utf-8 -*-
"""
Bereite Abrechnung auf Ebene FzgKombinationen auf
FahrzeugKombinationen (Teilnetze)

Erstellt: 09.01.2026
@author: mape
"""

import numpy as np
import pandas as pd
from VisumPy.helpers import SetMulti
from create.TNM_UDA import create_UDA_FZGAbr
from utils.TNM_Checks import check_timeIntervalls, check_BDA, check_vehiclejournyeSections


# import win32com.client.dynamic
# Visum = win32com.client.dynamic.Dispatch("Visum.Visum.25")
# Visum.IO.loadversion(r"C:\Users\peter\hvv.de\B-GR-Bereich B - Planungstool hvv\VisumTest\VisumTest_neu.ver")

# Parameter
S = 40
F = 12
M = "M1"


def main(Visum):
    if not _checks(Visum):
        return  False
    TN = Visum.Net.AttValue("TN")
    Gebiete = _get_territories(Visum)
    Einheiten = [["KM", "Fahrplankilometer"], ["STD", "Fahrplanstunden"], ["FZG", "Fahrzeugbedarf"]]
    FZG = _get_vehicleCombinations(Visum)
    create_UDA_FZGAbr(Visum, TN, Gebiete, Einheiten)
    
    LinesTable = _get_LinesTable(Visum, Gebiete, Einheiten, FZG, M)
    FZGKomb(Visum, TN, LinesTable, Gebiete, Einheiten, FZG, S, F, M)

def FZGKomb(Visum, _TN, _LinesTable, _Gebiete, _Einheiten, _FZG, _S, _F, _M):
    df_vc = pd.DataFrame(Visum.Net.VehicleCombinations.GetMultipleAttributes(["CODE"], True), columns = ["FZG"])
    for g, gname in _Gebiete:
        for e, ename in _Einheiten:
            data = []
            for f in df_vc["FZG"]:
                if f not in _FZG:
                    total  = 0
                elif e in ["KM", "STD"]:
                    total = ((_LinesTable[f"{g}_S_{f}_{e}(AP)"] * _S) +
                             (_LinesTable[f"{g}_F_{f}_{e}(AP)"] * _F)).sum()
                else:
                    cols = [c for c in _LinesTable.columns if "_S_" in c and _M in c and f in c and g in c]
                    max_S = _LinesTable[cols].sum().max()
                    cols = [c for c in _LinesTable.columns if "_F_" in c and _M in c and f in c and g in c]
                    max_F = _LinesTable[cols].sum().max()
                    total = max(max_S, max_F)
                data.append(total)
            SetMulti(Visum.Net.VehicleCombinations, f"{_TN}_{g}_{e}", data)

def _checks(Visum):
    if Visum.Net.VehicleJourneySections.CountActive == 0:
        Visum.Log(12288, "Keine aktiven FahrpanfahrtAbschnitte vorhanden")
        return False
    if not check_BDA(Visum, Visum.Net, "Network", "TN"):
        return False
    TN = Visum.Net.AttValue("TN")
    if not check_BDA(Visum, Visum.Net.Lines, "Lines", "TN"):
        return False
    if not check_BDA(Visum, Visum.Net.VehicleCombinations, "VehicleCombinations", f"{TN}_FZG_METN"):
        return False
    if not check_timeIntervalls(Visum):
        return False
    if not check_vehiclejournyeSections(Visum, True):
        return False
    return True

def _get_territories(Visum):
    df_t_put = pd.DataFrame(Visum.Net.VehicleJourneyItems.GetMultipleAttributes([r"TIMEPROFILEITEM\LINEROUTEITEM\NODE\MINACTIVE:CONTAININGTERRITORIES\CODE",
                                                                                  r"TIMEPROFILEITEM\LINEROUTEITEM\NODE\MINACTIVE:CONTAININGTERRITORIES\NAME"], True),
                            columns = ["CODE", "NAME"])
    t_unique = df_t_put[['CODE', "NAME"]].drop_duplicates().values.tolist()
    return t_unique

def _get_vehicleCombinations(Visum):
    df_vc = pd.DataFrame(Visum.Net.VehicleJourneySections.GetMultipleAttributes([r"VEHCOMB\CODE"], True),
                            columns = ["FZG"])
    vc_unique = df_vc['FZG'].drop_duplicates().values.tolist()
    return vc_unique
    
def _get_LinesTable(Visum, _Gebiete, _Einheiten, _FZG, _M):
    attrList = []
    for g, _ in _Gebiete:
        for e, _ in _Einheiten:
            for f in _FZG:
                times = ["Mo", "Di", "Mi", "Do", "Fr", "Sa", "So"] if e == "FZG" else ["AP"]
                e_val = _M if e == "FZG" else e
                for t in times:
                    attrList.extend([f"{g}_S_{f}_{e_val}({t})", f"{g}_F_{f}_{e_val}({t})",])
    _LinesTable = pd.DataFrame(Visum.Net.Lines.GetMultipleAttributes(attrList, True), columns = attrList)
    return _LinesTable
    
if __name__ == "__main__":           
    main(Visum)