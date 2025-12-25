#!/usr/bin/env python3
"""
Erstelle die für die LKA notwendigen BDA
Linien

Erstellt: 28.08.2025
@author: mape
"""

import pandas as pd

def create_UDA(Visum):
    Gebiete = _get_territories(Visum)
    Saisons = ["S", "F"]
    FZG = _get_vehicleCombinations(Visum)
    Einheiten = ["KM", "STD"]
    
    # UDG
    UDGs = Visum.Net.UserDefinedGroups.GetMultiAttValues("NAME")
    for g, gname in Gebiete:
        if not any(f"TNM_{g}" in UDG for _, UDG in UDGs):
            Visum.Net.AddUserDefinedGroup(f"TNM_{g}", f"TNM {gname}")
        if not any(f"TNM_{g}_FZG" in UDG for _, UDG in UDGs):
            Visum.Net.AddUserDefinedGroup(f"TNM_{g}_FZG", f"TNM {gname} (Fahrzeugbedarf)")
    
    # UDA Vehicles
    for s in Saisons:
        for f in FZG:
            _createUDA(Visum, f"{s}_{f}_M1", "TNM_FZG", [2, 6]) # Fahrzuegbedarf je Linie: Methode 1
            _createUDA(Visum, f"{s}_{f}_M2", "TNM_FZG", [1, 0]) # Fahrzuegbedarf je Linie: Methode 2
            _createUDA(Visum, f"{s}_{f}_STD", "TNM_FZG", [2, 6]) # Stunden je Fahrzeug je Linie
            for g, gname in Gebiete:
                _createUDA(Visum, f"{g}_{s}_{f}_M1", f"TNM_{g}_FZG", [2, 6]) # Fahrzuegbedarf je Linie: Methode 1
                _createUDA(Visum, f"{g}_{s}_{f}_M2", f"TNM_{g}_FZG", [1, 0]) # Fahrzuegbedarf je Linie: Methode 2
                for e in Einheiten:
                    _createUDA(Visum, f"{g}_{s}_{f}_{e}", f"TNM_{g}", [2, 6])
                    
    Visum.Log(20480, "TNM-Rechenattribute (BDA): erzeugt")

def _createUDA(Visum, _name, _group, _type):
    if not Visum.Net.Lines.AttrExists(_name):
        Visum.Net.Lines.AddUserDefinedAttribute(_name, _name, _name, _type[0], _type[1], False, 0, None, 0, SubAttr = 'TISET_1')
        UDA = Visum.Net.Lines.Attributes.ItemByKey(_name)
        UDA.Comment = "TNM-Rechenattribut"
        UDA.UserDefinedGroup = _group
    
def _get_territories(Visum):
    df_t_put = pd.DataFrame(Visum.Net.TerritoryPuTDetails.GetMultipleAttributes([r"TERRITORYCODE", "TERRITORYNAME"], True),
                            columns = ["CODE", "NAME"])
    t_unique = df_t_put[['CODE', "NAME"]].drop_duplicates().values.tolist()
    return t_unique

def _get_vehicleCombinations(Visum):
    df_vj = pd.DataFrame(Visum.Net.VehicleJourneySections.GetMultipleAttributes([r"VEHCOMB\CODE"], True), columns = ["FZG"])
    vc_unique = df_vj['FZG'].drop_duplicates().tolist()
    return vc_unique


            