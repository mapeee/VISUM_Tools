# -*- coding: utf-8 -*-
"""
Führe die Abrechnung Ebene der Fahrzeugkombinationen durch
Ebenen:
    - aktives Teilnetz 
    - Fahrzeugkombinationen
    - Linien

Erstellt: 09.01.2026
@author: mape
Version: 0.8
"""

import pandas as pd
from VisumPy.helpers import SetMulti
from create.TNM_BDA import bda_fahrzeugkombinationen
from utils.TNM_Checks import check_zeitintervalle, check_bda, check_fahrplanfahrtabschnitte


# Parameter
SW = 40 # Schulwochen je Jahr
FW = 12 # Ferienwochen je Jahr
M = "M1" # Methode Fahrzeugbedarf (M1 oder M2)


def main(Visum):
    if not _checks(Visum):
        return  False
    TN = Visum.Net.AttValue("TN")
    Gebiete = _gebiete(Visum)
    Einheiten = [["KM", "Fahrplankilometer"], ["STD", "Fahrplanstunden"], ["FZG", "Fahrzeugbedarf"]]
    FZG = _fahrzeugkombinationen(Visum)
    # Erstelle die für die TN-Abrechnung relevanten Formel-BDA
    bda_fahrzeugkombinationen(Visum, TN, Gebiete, Einheiten)
    
    LinienTabelle = _linien(Visum, Gebiete, Einheiten, FZG, M)
    FZGKomb(Visum, TN, LinienTabelle, Gebiete, Einheiten, FZG, SW, FW, M)
    Visum.Log(20480, f"Abrechnung für {TN}: durchgeführt")
    return True


def FZGKomb(Visum, _TN, _LinienTabelle, _Gebiete, _Einheiten, _FZG, _SW, _FW, _M):
    '''Berechne die anfallenden Mengen (FPLKM, FPLSTD, FZGBEdarf):
        - je Fahrzeugkombination
        - je Gebiet (Kreis)
        Die weiteren Berechnungen erfolgen über Formel-BDA
    
    Parameters
    ----------
    Visum : Visum-Instanz
    _TN : string
        Teilnetz-CODE
    _LinienTabelle : pandas dataframe
        dataframe mit notwendigen BDA auf Ebene der aktiven Linien (Linien im TN)
    _Gebiete : Liste
        Liste der aktiven Gebiete mit Spalten (CODE und NAME)
    _Einheiten : Liste
        Liste mit Einheiten (Mengen): FPLKM, FPLSTD, FZG
    _FZG : Liste
        Liste mit unterschiedlichen Fahrzeugkombinationen
    _SW : int
        Schulwochen je Jahr
    _FW : int
        Ferienwochen je Jahr
    _M : string
        Verfahren zur Berechnung Fahrzeugbedarf (Methode 1 (M1) oder Methode 2 (M2))

    Returns
    -------
    None.

    '''
    # Iteration über alle Fahrzeugkombinationen in Visum
    df_fahrzeugkombinationen = pd.DataFrame(Visum.Net.VehicleCombinations.GetMultipleAttributes(["CODE"]), columns = ["FZG"])
    for g, _ in _Gebiete:
        for e, _ in _Einheiten:
            daten_g_e = []
            for f in df_fahrzeugkombinationen["FZG"]:
                # Wenn FZG nicht im TN vorhanden, dann Wert 0
                if f not in _FZG:
                    wert_g_e_f  = 0
                # Wenn ermittlung FPLKM und FPLSTD, dann Multiplikation mit Jahreswochen je Saison
                elif e in ["KM", "STD"]:
                    wert_g_e_f = ((_LinienTabelle[f"{g}_S_{f}_{e}(AP)"] * _SW) +
                             (_LinienTabelle[f"{g}_F_{f}_{e}(AP)"] * _FW)).sum()
                # Sonst Fahrzeugbedarf: Maximalbedarf an einem Tag in der Schulzeit und in den Ferien.
                # Dann Maximalbedarf aus Maximalbedarf aus Schulzeit und Ferien
                else:
                    spalten_s = [spalte for spalte in _LinienTabelle.columns if "_S_" in spalte and _M in spalte and f in spalte and g in spalte]
                    max_SW = _LinienTabelle[spalten_s].sum().max()
                    spalten_f = [spalte for spalte in _LinienTabelle.columns if "_F_" in spalte and _M in spalte and f in spalte and g in spalte]
                    max_FW = _LinienTabelle[spalten_f].sum().max()
                    wert_g_e_f = max(max_SW, max_FW)
                daten_g_e.append(wert_g_e_f)
            SetMulti(Visum.Net.VehicleCombinations, f"{_TN}_{g}_{e}", daten_g_e)


def _checks(Visum):
    '''Führe vorab verschiedene Prüfroutinen durch
        - alle notwendigen BDA vorhanden?
        - alle notwendigen Zeitintervallmengen und Zeitintervalle vorhanden?
        - Alle aktiven Fahrplanfahrtabschnitte mit notwendigen Attributen versehen (Fahrzeugkombinationen, Saison etc.)?
    '''
    if not check_bda(Visum, Visum.Net, "Network", "TN"):
        return False
    TN = Visum.Net.AttValue("TN")
    if not check_bda(Visum, Visum.Net.Lines, "Lines", "TN"):
        return False
    if not check_bda(Visum, Visum.Net.VehicleCombinations, "VehicleCombinations", f"{TN}_FZG_METN"):
        return False
    if not check_zeitintervalle(Visum):
        return False
    if not check_fahrplanfahrtabschnitte(Visum, True):
        return False
    return True


def _fahrzeugkombinationen(Visum):
    df_fahrplanfahrtabschnitte = pd.DataFrame(Visum.Net.VehicleJourneySections.GetMultipleAttributes([r"VEHCOMB\CODE"], True),
                            columns = ["FZG"])
    fzgkombinationen_unique = df_fahrplanfahrtabschnitte['FZG'].drop_duplicates().values.tolist()
    return fzgkombinationen_unique
    

def _gebiete(Visum):
    _df_oevgebietedetail = pd.DataFrame(Visum.Net.VehicleJourneyItems.GetMultipleAttributes([r"TIMEPROFILEITEM\LINEROUTEITEM\NODE\MINACTIVE:CONTAININGTERRITORIES\CODE",
                                                                                  r"TIMEPROFILEITEM\LINEROUTEITEM\NODE\MINACTIVE:CONTAININGTERRITORIES\NAME"], True),
                            columns = ["CODE", "NAME"])
    gebiete_unique = _df_oevgebietedetail[['CODE', "NAME"]].drop_duplicates().values.tolist()
    return gebiete_unique


def _linien(Visum, _Gebiete, _Einheiten, _FZG, _M):
    '''
    Parameters
    ----------
    Visum : Visum-Instanz
        DESCRIPTION.
    _Gebiete : Liste
        Liste der aktiven Gebiete mit Spalten (CODE und NAME)
    _Einheiten : Liste
        Liste mit Einheiten (Mengen): FPLKM, FPLSTD, FZG
    _FZG : Liste
        Liste mit unterschiedlichen Fahrzeugkombinationen
    _M : string
        Verfahren zur Berechnung Fahrzeugbedarf (Methode 1 (M1) oder Methode 2 (M2))

    Returns
    -------
    _df_linien : pandas dataframe
        dataframe mit allen relevanten BDA
    '''
    attrList = []
    for g, _ in _Gebiete:
        for e, _ in _Einheiten:
            for f in _FZG:
                tage = ["Mo", "Di", "Mi", "Do", "Fr", "Sa", "So"] if e == "FZG" else ["AP"]
                # Wenn Werte von Fahrzeugen (FZG) dann Referenzierung über Methode der Berechnung der Fahrzeugbedarfe
                e_wert = _M if e == "FZG" else e
                for t in tage:
                    attrList.extend([f"{g}_S_{f}_{e_wert}({t})", f"{g}_F_{f}_{e_wert}({t})",])
    _df_linien = pd.DataFrame(Visum.Net.Lines.GetMultipleAttributes(attrList, True), columns = attrList)
    return _df_linien
    

if __name__ == "__main__":           
    main(Visum)