#!/usr/bin/env python3
"""
Erstelle die für die LKA notwendigen BDA
Linien, FahrzuegKombinationen

Erstellt: 28.08.2025
@author: mape
"""

import pandas as pd

def create_UDA_STDKM(Visum):
    Gebiete, Saisons, FZG, Einheiten = _get_Values(Visum)
    _check_UDG(Visum, Gebiete)
    for s, sname in Saisons:
        for f, fname in FZG:
            _create_UDA_Lines(Visum, f"{s}_{f}_STD", "TNM_FZG", [2, 6], f"{sname} {fname} Fahrplanstunden") # Stunden je Fahrzeug je Linie
            _create_UDA_Lines(Visum, f"{s}_{f}_KM", "TNM_FZG", [2, 6], f"{sname} {fname} Fahrplankilometer") # Kilometer je Fahrzeug je Linie
            for g, gname in Gebiete:
                for e, ename in Einheiten:
                    _create_UDA_Lines(Visum, f"{g}_{s}_{f}_{e}", f"TNM_{g}", [2, 6], f"{gname} {sname} {fname} {ename}")
                    
    Visum.Log(20480, "TNM-Rechenattribute für Linien STD und KM: erzeugt")
    
def create_UDA_FZGBedarf(Visum):
    Gebiete, Saisons, FZG, Einheiten = _get_Values(Visum)
    _check_UDG(Visum, Gebiete)
    for s, sname in Saisons:
        for f, fname in FZG:
            _create_UDA_Lines(Visum, f"{s}_{f}_M1", "TNM_FZG", [2, 6], f"{sname} {fname} Fahrzeugbedarf Methode 1") # Fahrzuegbedarf je Linie: Methode 1
            _create_UDA_Lines(Visum, f"{s}_{f}_M2", "TNM_FZG", [1, 0], f"{sname} {fname} Fahrzeugbedarf Methode 2") # Fahrzuegbedarf je Linie: Methode 2
            for g, gname in Gebiete:
                if not _create_UDA_Lines(Visum, f"{g}_{s}_{f}_M1", f"TNM_{g}_FZG", [2, 6], f"{gname} {sname} {fname} Fahrzeugbedarf Methode 1", True): # Fahrzuegbedarf je Linie: Methode 1
                    return False
                if not _create_UDA_Lines(Visum, f"{g}_{s}_{f}_M2", f"TNM_{g}_FZG", [1, 0], f"{gname} {sname} {fname} Fahrzeugbedarf Methode 2", True): # Fahrzuegbedarf je Linie: Methode 2
                    return False
                    
    Visum.Log(20480, "TNM-Rechenattribute für Fahrzeugbedarfe: erzeugt")
    return True

def create_UDA_FZGAbr(Visum, TN, Gebiete, Einheiten):
    _check_UDG(Visum, None, TN)
    Kosten_AT = []
    for e, ename in Einheiten:
        areas = []
        for g, gname in Gebiete:
            _create_UDA_FZG(Visum, f"{TN}_{g}_{e}", f"TNM_{TN}", [2, 6], f"{TN} {gname} {ename}", None)
            areas.append(f"[{TN}_{g}_{e}]")
        formular = " + ".join(areas)
        _create_UDA_FZG(Visum, f"{TN}_{e}", f"TNM_{TN}", [2, 6], f"{TN} {ename}", formular)
        areas = []
        for g, gname in Gebiete:
            _create_UDA_FZG(Visum, f"{TN}_{g}_{e}_ANTEIL", f"TNM_{TN}", [2, 6], f"{TN} {gname} Anteil {ename} in %", f"{TN}_{e}")
            e_dict = {"KM": "KOSTENSATZKMSERVICEGES",
                      "STD": "KOSTENSATZSTDSERVICEGES",
                      "FZG": "KOSTENSATZFZGEINHEITGES(AP)"}
            e_val = f"{e}_METN" if e == "FZG" else e
            _create_UDA_FZG(Visum, f"{TN}_{g}_{e}_ATKOSTEN", f"TNM_{TN}", [2, 6], f"{TN} {gname} ATKOSTEN {ename}", f"[{TN}_{g}_{e}_ANTEIL] * [{TN}_{e_val}] * [{e_dict[e]}]")
            areas.append(f"[{TN}_{g}_{e}_ATKOSTEN]")
        formular = " + ".join(areas)
        _create_UDA_FZG(Visum, f"{TN}_{e}_KOSTEN", f"TNM_{TN}", [2, 6], f"{TN} KOSTEN {ename}", formular)
        Kosten_AT.append(f"[{TN}_{e}_KOSTEN]")
    formular = " + ".join(Kosten_AT)
    _create_UDA_FZG(Visum, f"{TN}_GESAMTKOSTEN", f"TNM_{TN}", [2, 6], f"{TN} GESAMTKOSTEN", formular)
    Visum.Log(20480, "TNM-Rechenattribute für Abrechnung (Fahrzeuge): erzeugt")

def _check_UDG(Visum, _Gebiete = None, _TN = None):
    UDGs = Visum.Net.UserDefinedGroups.GetMultiAttValues("NAME")
    if _TN:
        if not any(f"TNM_{_TN}" in UDG for _, UDG in UDGs):
            Visum.Net.AddUserDefinedGroup(f"TNM_{_TN}", f"Teilnetz {_TN}")
        return
    for g, gname in _Gebiete:
        if not any(f"TNM_{g}" in UDG for _, UDG in UDGs):
            Visum.Net.AddUserDefinedGroup(f"TNM_{g}", f"TNM {gname}")
        if not any(f"TNM_{g}_FZG" in UDG for _, UDG in UDGs):
            Visum.Net.AddUserDefinedGroup(f"TNM_{g}_FZG", f"TNM {gname} (Fahrzeugbedarf)")
    
def _create_UDA_Lines(Visum, _name, _group, _type, _comment, _formular = None):
    if Visum.Net.Lines.AttrExists(_name):
        return True
    if _formular:
        _g, _s, _f, _m = _name.split("_")
        f = [f"{_g}_{_s}_{_f}_STD(Mo)", f"{_s}_{_f}_STD(Mo)", f"{_s}_{_f}_{_m}(Mo)"]
        for i in f:
            if not Visum.Net.Lines.AttrExists(i):
                Visum.Log(12288, "Formel-BDA {_name} kann nicht erstellt werden: BDA {i} felt")
                return False
        f = [s.replace("Mo", "") for s in f]
        Visum.Net.Lines.AddUserDefinedAttribute(_name, _name, _name, _type[0], _type[1], Formula =f"([{f[0]}]/[{f[1]}])*[{f[2]}] ", SubAttr = 'TISET_1') 
    else:
        Visum.Net.Lines.AddUserDefinedAttribute(_name, _name, _name, _type[0], _type[1], False, 0, None, 0, SubAttr = 'TISET_1')
    UDA = Visum.Net.Lines.Attributes.ItemByKey(_name)
    UDA.Comment = f"{_comment} | TNM-Rechenattribut"
    UDA.UserDefinedGroup = _group
    return True
      
def _create_UDA_FZG(Visum, _name, _group, _type, _comment, _formular):
    if Visum.Net.VehicleCombinations.AttrExists(_name):
        return True
    if _formular and "_ANTEIL" in _name:
        f = _name.replace("_ANTEIL", "")
        Visum.Net.VehicleCombinations.AddUserDefinedAttribute(_name, _name, _name,  _type[0], _type[1], Formula = f"[{f}] / [{_formular}]")
    elif _formular:
        Visum.Net.VehicleCombinations.AddUserDefinedAttribute(_name, _name, _name,  _type[0], _type[1], Formula = _formular)
    else:
        Visum.Net.VehicleCombinations.AddUserDefinedAttribute(_name, _name, _name, _type[0], _type[1], False, 0, None, 0)
    UDA = Visum.Net.VehicleCombinations.Attributes.ItemByKey(_name)
    UDA.Comment = f"{_comment} | TNM-Rechenattribut"
    UDA.UserDefinedGroup = _group
    return True
    
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
           