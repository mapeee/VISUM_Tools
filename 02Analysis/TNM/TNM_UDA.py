# -*- coding: utf-8 -*-
"""
Created on Thu Aug 28 20:14:40 2025

@author: peter
"""    
Gebiete = ["FHH", "PI", "SE"]
Saisons = ["S", "F"]
Tage = ["MO", "DI", "MI", "DO", "FR", "SA", "SO"]
FZG = ["SB", "GB"]
Einheiten = ["KM", "STD"]


def createUDA(_Name, area, area2 = False):
    if not Visum.Net.Lines.AttrExists(_Name):
        if area == True:
            Visum.Net.Lines.AddUserDefinedAttribute(_Name, _Name, _Name, 2, 6, False, 0, None, 0)
            UDA = Visum.Net.Lines.Attributes.ItemByKey(_Name)
            UDA.UserDefinedGroup = f"TNM_{g}"
        else:
            if _Name[-3:] == "STD" or _Name[-2:] == "M1": Visum.Net.Lines.AddUserDefinedAttribute(_Name, _Name, _Name, 2, 6, False, 0, None, 0)
            else: Visum.Net.Lines.AddUserDefinedAttribute(_Name, _Name, _Name, 1, 0, False, 0, None, 0)
            UDA = Visum.Net.Lines.Attributes.ItemByKey(_Name)
            if area2:
                UDA.UserDefinedGroup = f"TNM_{area2}_FZG"
            else:
                UDA.UserDefinedGroup = f"TNM_FZG"
    
# UDG
UDGs = Visum.Net.UserDefinedGroups.GetMultiAttValues("NAME")
for g in Gebiete:
    if not any(f"TNM_{g}" in UDG for _, UDG in UDGs):
        Visum.Net.AddUserDefinedGroup(f"TNM_{g}")

# UDA Areas
for g in Gebiete:
    for s in Saisons:
        for t in Tage:
            for f in FZG:
                createUDA(f"{g}_{s}{t}_{f}M1", False, g) # Fahrzuegbedarf je Linie: Methode 1
                createUDA(f"{g}_{s}{t}_{f}M2", False, g) # Fahrzuegbedarf je Linie: Methode 2
                for e in Einheiten:
                    createUDA(f"{g}_{s}{t}_{f}{e}", True)

# UDA Vehicles
for s in Saisons:
    for t in Tage:
        for f in FZG:
            createUDA(f"{s}{t}_{f}M1", False) # Fahrzuegbedarf je Linie: Methode 1
            createUDA(f"{s}{t}_{f}M2", False) # Fahrzuegbedarf je Linie: Methode 2
            createUDA(f"{s}{t}_{f}STD", False) # Stunden je Fahrzeug je Linie