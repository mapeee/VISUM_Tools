# -*- coding: utf-8 -*-
"""
Created on Thu Aug 28 20:14:40 2025

@author: peter
"""    
Gebiete = ["FHH", "PI", "SE"]
Saisons = ["S", "F"]
FZG = ["SB", "GB"]
Einheiten = ["KM", "STD"]


def createUDA(_name, _group, _type):
    if not Visum.Net.Lines.AttrExists(_name):
        Visum.Net.Lines.AddUserDefinedAttribute(_name, _name, _name, _type[0], _type[1], False, 0, None, 0, SubAttr = 'TISET_1')
        UDA = Visum.Net.Lines.Attributes.ItemByKey(_name)
        UDA.UserDefinedGroup = _group
    
# UDG
UDGs = Visum.Net.UserDefinedGroups.GetMultiAttValues("NAME")
for g in Gebiete:
    if not any(f"TNM_{g}" in UDG for _, UDG in UDGs):
        Visum.Net.AddUserDefinedGroup(f"TNM_{g}")
    if not any(f"TNM_{g}_FZG" in UDG for _, UDG in UDGs):
        Visum.Net.AddUserDefinedGroup(f"TNM_{g}_FZG")

# UDA Vehicles
for s in Saisons:
    for f in FZG:
        createUDA(f"{s}_{f}_M1", "TNM_FZG", [2, 6]) # Fahrzuegbedarf je Linie: Methode 1
        createUDA(f"{s}_{f}_M2", "TNM_FZG", [1, 0]) # Fahrzuegbedarf je Linie: Methode 2
        createUDA(f"{s}_{f}_STD", "TNM_FZG", [2, 6]) # Stunden je Fahrzeug je Linie
        for g in Gebiete:
            createUDA(f"{g}_{s}_{f}_M1", f"TNM_{g}_FZG", [2, 6]) # Fahrzuegbedarf je Linie: Methode 1
            createUDA(f"{g}_{s}_{f}_M2", f"TNM_{g}_FZG", [1, 0]) # Fahrzuegbedarf je Linie: Methode 2
            for e in Einheiten:
                createUDA(f"{g}_{s}_{f}_{e}", f"TNM_{g}", [2, 6])
            