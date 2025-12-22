# -*- coding: utf-8 -*-
"""
Created on Thu Dec 18 09:33:42 2025

@author: peter
"""

import pandas as pd

def check_TN(Visum):
    check_attributes(Visum)
    check_lines(Visum)
    check_territories(Visum)
    check_vehicleCombinations(Visum)
    check_vehiclejournyeSections(Visum)
    check_timeIntervalls(Visum)
    Visum.Log(20480, "Checks: abgeschlossen")
    
def check_attributes(Visum):
    if not Visum.Net.Lines.AttrExists("TN"):
        Visum.Log(12288, "Linien: BDA 'TN' fehlt")
    if not Visum.Net.VehicleJourneySections.AttrExists("SAISON"):
        Visum.Log(12288, "FahrpanfahrtAbschnitte: BDA 'SAISON' fehlt")
    if not Visum.Net.VehicleCombinations.AttrExists("RANG"):
        Visum.Log(12288, "FahrzeugKombinationen: BDA 'RANG' fehlt")

def check_lines(Visum):
    df_l = pd.DataFrame(Visum.Net.Lines.GetMultipleAttributes(["TN"], False),
                               columns = ["TN"])
    df_l = df_l[df_l['TN'] != 'ohne']
    df_l = df_l[['TN']].drop_duplicates()
    df_l.rename(columns={'TN': 'CODE'}, inplace=True)
    _general_checks(Visum, "Linien", df_l)
    
def check_EFW_tables(Visum):
    pass
    #Check auf Kommentare in richtigem Format
    #Check auf richtige Gruppe (EFW)
    
def check_territories(Visum):
    df_t = pd.DataFrame(Visum.Net.Territories.GetMultipleAttributes(["CODE", "TYPENO"], True),
                               columns = ["CODE", "TYPENO"])
    df_t = df_t[df_t['TYPENO'].eq(1)]
    _general_checks(Visum, "Gebiete", df_t)
    
def check_timeIntervalls(Visum):
    if Visum.Net.TimeIntervalSets.Count == 0:
        Visum.Log(12288, "Das Netz enthält keine Zeitintervallmengen")
    TI = Visum.Net.TimeIntervalSets.GetAll
    if TI[0].TimeIntervals.Count != 7:
        Visum.Log(12288, f"Die Zeitintervallmenge {TI[0].AttValue('NAME')} enthält keine 7 Wochentage")
        
def check_vehicleCombinations(Visum):
    if Visum.Net.VehicleCombinations.Count == 0:
        Visum.Log(12288, "Das Netz hat keine FahrzeugKombinationen")
        return
    df_vc = pd.DataFrame(Visum.Net.VehicleCombinations.GetMultipleAttributes(["CODE"], False),
                               columns = ["CODE"])
    _general_checks(Visum, "FahrzeugKombinationen", df_vc)
    
def check_vehiclejournyeSections(Visum):
    df_vj = pd.DataFrame(Visum.Net.VehicleJourneySections.GetMultipleAttributes([r"VEHCOMB\CODE", "SAISON"], True),
                               columns = ["FZG", "SAISON"])
    if df_vj['FZG'].isna().any():
        Visum.Log(12288, f"FahrpanfahrtAbschnitte ohne Fahrzeug: {df_vj['FZG'].isna().sum()}")
    df_vj = df_vj[['SAISON']].drop_duplicates()
    if not df_vj['SAISON'].isin(["S", "F", "S+F"]).all():
        Visum.Log(12288, "BDA 'SAISON' enthält ungültige Werte (gültig: 'S', 'F', 'S+F')")

def _general_checks(Visum, element, df_check):
    if df_check['CODE'].str.contains(r'[ÖöÄäÜüß#-]', regex=True, na=False).any():
        Visum.Log(12288, f"{element}: Es gibt Sonderzeichen (#, ü, ä etc.) in CODES")
    if (df_check["CODE"].isna() | (df_check["CODE"].str.strip() == '')).any():
        Visum.Log(12288, f"{element}: Es gibt Leerwerte in CODES")
    if df_check['CODE'].duplicated().any():
        for row in df_check[df_check['CODE'].duplicated()].itertuples(index=False):
            Visum.Log(12288, f"{element}: Duplikat(e) in CODES: {row.CODE}")
