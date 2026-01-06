#!/usr/bin/env python3
"""
Erstelle die für die LKA notwendigen BDA
Linien, FahrzuegKombinationen

Erstellt: 28.08.2025
@author: mape
"""

import pandas as pd

def create_UDA_LKZ(Visum):
    Gebiete, Saisons, FZG, Einheiten = _get_Values(Visum)
    _check_UDG(Visum, Gebiete)
    for s, sname in Saisons:
        for f, fname in FZG:
            _create_UDA(Visum, f"{s}_{f}_STD", "TNM_FZG", [2, 6], f"{sname} {fname} Fahrplanstunden") # Stunden je Fahrzeug je Linie
            for g, gname in Gebiete:
                for e, ename in Einheiten:
                    _create_UDA(Visum, f"{g}_{s}_{f}_{e}", f"TNM_{g}", [2, 6], f"{gname} {sname} {fname} {ename}")
                    
    Visum.Log(20480, "TNM-Rechenattribute für LKZ: erzeugt")
    
def create_UDA_Mengen(Visum):
    Gebiete, Saisons, FZG, Einheiten = _get_Values(Visum)
    _check_UDG(Visum, Gebiete)
    for s, sname in Saisons:
        for f, fname in FZG:
            _create_UDA(Visum, f"{s}_{f}_M1", "TNM_FZG", [2, 6], f"{sname} {fname} Fahrzeugbedarf Methode 1") # Fahrzuegbedarf je Linie: Methode 1
            _create_UDA(Visum, f"{s}_{f}_M2", "TNM_FZG", [1, 0], f"{sname} {fname} Fahrzeugbedarf Methode 2") # Fahrzuegbedarf je Linie: Methode 2
            for g, gname in Gebiete:
                _create_UDA(Visum, f"{g}_{s}_{f}_M1", f"TNM_{g}_FZG", [2, 6], f"{gname} {sname} {fname} Fahrzeugbedarf Methode 1") # Fahrzuegbedarf je Linie: Methode 1
                _create_UDA(Visum, f"{g}_{s}_{f}_M2", f"TNM_{g}_FZG", [1, 0], f"{gname} {sname} {fname} Fahrzeugbedarf Methode 2") # Fahrzuegbedarf je Linie: Methode 2
                    
    Visum.Log(20480, "TNM-Rechenattribute für Mengen: erzeugt")

def _check_UDG(Visum, _Gebiete):
    UDGs = Visum.Net.UserDefinedGroups.GetMultiAttValues("NAME")
    for g, gname in _Gebiete:
        if not any(f"TNM_{g}" in UDG for _, UDG in UDGs):
            Visum.Net.AddUserDefinedGroup(f"TNM_{g}", f"TNM {gname}")
        if not any(f"TNM_{g}_FZG" in UDG for _, UDG in UDGs):
            Visum.Net.AddUserDefinedGroup(f"TNM_{g}_FZG", f"TNM {gname} (Fahrzeugbedarf)")
    

def _create_UDA(Visum, _name, _group, _type, _comment):
    if not Visum.Net.Lines.AttrExists(_name):
        Visum.Net.Lines.AddUserDefinedAttribute(_name, _name, _name, _type[0], _type[1], False, 0, None, 0, SubAttr = 'TISET_1')
        UDA = Visum.Net.Lines.Attributes.ItemByKey(_name)
        UDA.Comment = f"{_comment} | TNM-Rechenattribut"
        UDA.UserDefinedGroup = _group
        
def _get_Values(Visum):
    _Gebiete = _get_territories(Visum)
    _Saisons = [["S", "Schule"], ["F", "Ferien"]]
    _FZG = _get_vehicleCombinations(Visum)
    _Einheiten = [["KM", "Fahrplankilometer"], ["STD", "Fahrplanstunden"]]
    return _Gebiete, _Saisons, _FZG, _Einheiten
    
def _get_territories(Visum):
    df_t_put = pd.DataFrame(Visum.Net.TerritoryPuTDetails.GetMultipleAttributes([r"TERRITORYCODE", "TERRITORYNAME"], True),
                            columns = ["CODE", "NAME"])
    t_unique = df_t_put[['CODE', "NAME"]].drop_duplicates().values.tolist()
    return t_unique

def _get_vehicleCombinations(Visum):
    df_vj = pd.DataFrame(Visum.Net.VehicleJourneySections.GetMultipleAttributes([r"VEHCOMB\CODE", r"VEHCOMB\NAME"], True),
                         columns = ["CODE", "NAME"])
    vc_unique = df_vj[['CODE', 'NAME']].drop_duplicates().values.tolist()
    return vc_unique
           