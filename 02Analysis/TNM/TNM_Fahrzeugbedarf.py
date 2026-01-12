# -*- coding: utf-8 -*-
"""
Ermittle die Fahrzeugbedarfe
Linien, FahrzeugKombinationen (Teilnetze)

Erstellt: 01.09.2025
@author: mape
"""

import numpy as np
import pandas as pd
from VisumPy.helpers import SetMulti
from create.TNM_UDA import create_UDA_FZGBedarf, _create_UDA_FZG
from utils.TNM_Checks import check_timeIntervalls, check_BDA, check_vehiclejournyeSections


# import win32com.client.dynamic
# Visum = win32com.client.dynamic.Dispatch("Visum.Visum.25")
# Visum.IO.loadversion(r"C:\Users\peter\hvv.de\B-GR-Bereich B - Planungstool hvv\VisumTest\VisumTest_neu.ver")

# Parameter
AbAnOverlap = False

def main(Visum, _AbAnOverlap):
    if not _checks(Visum):
        return  False
    if not create_UDA_FZGBedarf(Visum):
        return False
    TN = Visum.Net.AttValue("TN")
    _create_UDA_FZG(Visum, f"{TN}_FZG_METN", f"TNM_{TN}", [1, 0], f"{TN} Fahrzeugbedarf gem. METN", None)
    Gebiete = _get_territories(Visum)
    Saisons = ["S", "F"]
    Tage = ["AP", "Mo", "Di", "Mi", "Do", "Fr", "Sa", "So"]
    FZG = _get_vehicleCombinations(Visum)
    
    VJTable = _get_VJTable(Visum, TN)
    
    FZG_Lines(Visum, VJTable, Saisons, Tage, FZG, Gebiete, _AbAnOverlap)
    FZG_Veh(Visum, VJTable, Saisons, Tage, FZG, TN, _AbAnOverlap)
    
    Visum.Log(20480, "Fahrzeugbedarfe: ermittelt")


def FZG_Lines(Visum, _VJTable, _Saisons, _Tage, _FZG, _Gebiete, _AbAnOverlap):
    for s in _Saisons:
        for t in _Tage:
            # if (AP)? # ggf maxBedarf je Fahrzeug über alle Wochentage
            AddListVisum = pd.DataFrame()
            for line in np.sort(_VJTable["LINE"].unique()):
                FZG_line = _calculate_FZG_Lines(_VJTable, _FZG, s, t, _AbAnOverlap, line)
                AddListVisum = pd.concat([AddListVisum, FZG_line], ignore_index=True)
            AddListVisum.fillna(0, inplace=True)
            AddListVisum.drop(columns=["LINENAME"], inplace=True)
            for col_name, col_data in AddListVisum.items():
                col_nameAP = col_name.replace(t, "AP")
                dfAP = pd.Series([x[1] for x in Visum.Net.Lines.GetMultiAttValues(col_nameAP, True)])
                dfAP = np.maximum(dfAP, col_data)
                SetMulti(Visum.Net.Lines, col_name, col_data.tolist(), True)
                SetMulti(Visum.Net.Lines, col_nameAP, dfAP.tolist(), True)
                
def FZG_Veh(Visum, _VJTable, _Saisons, _Tage, _FZG, _TN, _AbAnOverlap):
    Veh_df = pd.DataFrame()
    for s in _Saisons:
        for t in _Tage:
            if t == "AP":
                continue
            FZG_FK = _calculate_FZG_Veh(_VJTable, _FZG, s, t, _AbAnOverlap)
            Veh_df[FZG_FK.columns] = FZG_FK
    col_data = []
    df_vc = pd.DataFrame(Visum.Net.VehicleCombinations.GetMultipleAttributes(["CODE"], True), columns = ["FZG"])
    for f in df_vc["FZG"]:
        if Veh_df.columns.str.startswith(f"{f}_").any():
            cols = [c for c in Veh_df.columns if "_S_" in c and f in c]
            max_S = Veh_df[cols].sum().max()
            cols = [c for c in Veh_df.columns if "_F_" in c and f in c]
            max_F = Veh_df[cols].sum().max()
            fmax = max(max_S, max_F)
        else:
            fmax = 0
        col_data.append(fmax)
        SetMulti(Visum.Net.VehicleCombinations, f"{_TN}_FZG_METN", col_data)
        
def _build_timeline(_VJ, _AbAnOverlap):
    n_minutes = 24*60
    df_timeline = pd.DataFrame({"time": range(0, n_minutes)})
    df_timeline["ALL"] = 0
    df_timeline["FREI"] = 0
    VEHunique = _VJ["VEH"].drop_duplicates().tolist()
    df_timeline[VEHunique] = 0
    for dep, arr, veh in zip(_VJ["DEP"], _VJ["ARR"], _VJ["VEH"]):
        if _AbAnOverlap:
            df_timeline.loc[dep:arr, "ALL"] += 1
            df_timeline.loc[dep:arr, veh] += 1
        else:
            df_timeline.loc[dep:arr - 1, "ALL"] += 1 # if identical DEP and ARR not overlapping
            df_timeline.loc[dep:arr - 1, veh] += 1 # if identical DEP and ARR not overlapping
    return df_timeline

def _calculate_FZG_Lines(_VJTable, _FZG, _s, _t, _AbAnOverlap, _line):
    LineList = pd.DataFrame(columns=["LINENAME"])
    LineList.loc[len(LineList)] = [_line]
    _VJTable = _VJTable[(_VJTable["LINE"] == _line) & (_VJTable[_s] == 1) & (_VJTable["Tag"] == _t)]
    timeline = _build_timeline(_VJTable, _AbAnOverlap)
    for f in _FZG["FZG"].unique().tolist():
        if not f in timeline.columns:
            continue
        timeline[f"{f}max"] = timeline[f].max()
    timeline["FZGMAX"] = timeline.filter(like="max").sum(axis=1, skipna=True)
    for _f in _FZG["FZG"].unique().tolist():
        if not _f in timeline.columns:
            continue
        LineList[f"{_s}_{_f}_M1({_t})"] = timeline["ALL"].max() / timeline["FZGMAX"] * timeline[f"{_f}max"]
        m2_sum = LineList.filter(like="M2").sum(axis=1, skipna=True)
        LineList[f"{_s}_{_f}_M2({_t})"] = timeline[f"{_f}max"].clip(upper=timeline["ALL"].max()-m2_sum) 
    return LineList

def _calculate_FZG_Veh(_VJTable, _FZG, _s, _t, _AbAnOverlap):
    _VJTable = _VJTable[(_VJTable[_s] == 1) & (_VJTable["Tag"] == _t)]
    timeline = _build_timeline(_VJTable, _AbAnOverlap)
    Rang0 = 0
    for f, Rang1 in zip(_FZG["FZG"], _FZG["RANG"]):
        if not f in timeline.columns:
            continue
        if Rang0 <= Rang1:
            timeline[f"{f}ersatz"] = timeline[["FREI", f]].min(axis=1).where(timeline["FREI"] > 0, 0)
            timeline[[f, "FREI"]] = timeline[[f, "FREI"]].sub(timeline[f"{f}ersatz"], axis=0)
        timeline[f"{f}max"] = timeline[f].max()
        timeline["FREI"] = timeline["FREI"] + (timeline[f"{f}max"] - timeline[f])
        Rang0 = Rang1
    df_veh = timeline.filter(like="max").copy()
    df_veh.drop(df_veh.index[1:], inplace=True)
    df_veh.rename(columns=lambda c: c.replace("max", f"_{_s}_({_t})"), inplace=True)
    return df_veh

def _checks(Visum):
    if Visum.Net.VehicleJourneySections.CountActive == 0:
        Visum.Log(12288, "Keine aktiven FahrpanfahrtAbschnitte vorhanden")
        return False
    if not check_BDA(Visum, Visum.Net.VehicleJourneySections, "FahrpanfahrtAbschnitte", "S"):
        return False
    if not check_BDA(Visum, Visum.Net.VehicleJourneySections, "FahrpanfahrtAbschnitte", "F"):
        return False
    if not check_BDA(Visum, Visum.Net, "Network", "TN"):
        return False
    if not check_BDA(Visum, Visum.Net.Lines, "Lines", "TN"):
        return False
    if not check_BDA(Visum, Visum.Net.VehicleCombinations, "FahrzeugKombinationen", "RANG"):
        return False
    if not check_timeIntervalls(Visum):
        return False
    if not check_vehiclejournyeSections(Visum, True):
        return False
    return True

def _get_territories(Visum):
    df_t_put = pd.DataFrame(Visum.Net.VehicleJourneyItems.GetMultipleAttributes([r"TIMEPROFILEITEM\LINEROUTEITEM\NODE\MINACTIVE:CONTAININGTERRITORIES\CODE"], True),
                            columns = ["CODE"])
    t_unique = df_t_put['CODE'].drop_duplicates().values.tolist()
    return t_unique

def _get_vehicleCombinations(Visum):
    df_vc = pd.DataFrame(Visum.Net.VehicleJourneySections.GetMultipleAttributes([r"VEHCOMB\CODE", r"VEHCOMB\RANG"], True), columns = ["FZG", "RANG"])
    df_vc = df_vc[['FZG', "RANG"]].drop_duplicates().reset_index(drop=True)
    df_vc = df_vc.astype({"FZG": "string", "RANG": int})
    df_vc.sort_values(by="RANG", inplace=True)
    return df_vc

def _get_VJTable(Visum, _TN):
    next_day = {
    "Mo": "Di",
    "Di": "Mi",
    "Mi": "Do",
    "Do": "Fr",
    "Fr": "Sa",
    "Sa": "So",
    "So": "Mo"}
    attrList = [[r"VEHJOURNEY\LINENAME",r"VEHJOURNEY\NAME", "DEP", "ARR", r"VEHCOMB\CODE", r"VEHCOMB\RANG",
                 "S", "F",
                 r"ISVALID(Mo)", r"ISVALID(Di)", r"ISVALID(Mi)",
                 r"ISVALID(Do)", r"ISVALID(Fr)", r"ISVALID(Sa)",
                 r"ISVALID(So)"],
                ["LINE", "NAME", "DEP", "ARR", "VEH", "VEHRANG", "S", "F", "Mo", "Di", "Mi", "Do", "Fr", "Sa", "So"]]
    _VJTable = pd.DataFrame(Visum.Net.VehicleJourneySections.GetMultipleAttributes(attrList[0], True), columns = attrList[1])
    _VJTable["DEP"], _VJTable["ARR"] = _VJTable["DEP"] / 60, _VJTable["ARR"] / 60
    _VJTableDay = pd.DataFrame()
    for i, day in enumerate(["Mo", "Di", "Mi", "Do", "Fr", "Sa", "So"]):
        df_day = _VJTable.loc[_VJTable[day] == 1].copy()
        # Dupliziere alle Fahrten, die über den Tageswechsel gehen
        df_day["is_nextday"] = False
        mask = (df_day["DEP"] < 1440) & (df_day["ARR"] >= 1440)
        df_dup = df_day[mask].copy()
        df_dup["is_nextday"] = True
        df_day = pd.concat([df_day, df_dup], ignore_index=True)
        # Passe Ankunft und Abfahrt der duplizierten Fahrten an
        df_day.loc[(df_day["DEP"] < 1440) & (df_day["ARR"] >= 1440) & (df_day["is_nextday"] == False), ["ARR", "Tag"]] = [1439, day]
        df_day.loc[(df_day["DEP"] < 1440) & (df_day["ARR"] >= 1440) & (df_day["is_nextday"] == True), ["DEP", "Tag"]] = [0, day]
        df_day.loc[df_day["is_nextday"], "ARR"] -= 1440
        # Schiebe Fahrten nach 24 Uhr in das 0-24 Schema aber in den nächsten Tag
        mask = df_day["DEP"] >= 1440
        df_day.loc[mask, ["DEP", "ARR"]] -= 1440
        df_day.loc[mask, "is_nextday"] = True
        # Ordner jeder Fahrt/duplizierten Fahrt den Fahrtag zu
        df_day["Tag"] = np.where(df_day["is_nextday"] == False, day, next_day[day])
        _VJTableDay = pd.concat([_VJTableDay, df_day], ignore_index=True)
        
    _VJTableDay.drop(columns=["Mo", "Di", "Mi", "Do", "Fr", "Sa", "So", "is_nextday"], inplace=True)
    _VJTableDay[["LINE", "NAME", "VEH", "Tag"]] = _VJTableDay[["LINE", "NAME", "VEH", "Tag"]].astype("string")
    _VJTableDay["VEHRANG"] = _VJTableDay["VEHRANG"].astype(int)
    return _VJTableDay


if __name__ == "__main__":           
    main(Visum, AbAnOverlap)
