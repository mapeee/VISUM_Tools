# -*- coding: utf-8 -*-
"""
Ermittle die Fahrplankilometer und Fahrplanstunden je Gebiet
Ebenen:
    - Linien
    - Gebiete-ÖV-Detail
    - Fahrplanfahrtabschnitte

Erstellt: 28.08.2025
@author: mape
Version: 0.8
"""

import pandas as pd
from VisumPy.helpers import SetMulti
from create.TNM_BDA import bda_linien
from utils.TNM_Checks import check_zeitintervalle, check_bda, check_gebiete, check_fahrplanfahrtabschnitte


def main(Visum):
    '''Generiere die Fahrplankilometer (FPLKM) und Fahrplanstunden (FPLSTD) aus der List 'Gebiete-ÖV-Detail'
    Generiere FPLKM und FPLSTD je:
        - Saison (S, F)
        - Fahrzeugkombination
        - je Tag (inkl. Gesamtwoche aus Analyseperiode (AP))
        - je Gebiet (Kreis)
    
    '''
    if not _checks(Visum):
        return  False
    bda_linien(Visum, True)
    Gebiete = _gebiete(Visum)
    Saisons = ["S", "F"]
    Tage = ["AP", "Mo", "Di", "Mi", "Do", "Fr", "Sa", "So"]
    Einheiten = ["KM", "STD"]
    FZG = _fahrzeugkombinationen(Visum)

    df_oevgebietedetail = pd.DataFrame(Visum.Net.TerritoryPuTDetails.GetMultipleAttributes(["TERRITORYCODE", "LINENAME", "VEHCOMBCODE",
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
    df_oevgebietedetail.columns = ["Gebiet", "LINE", "FZG", "APKM", "APSTD", "MoKM", "DiKM", "MiKM", "DoKM", "FrKM", "SaKM", "SoKM",
                   "MoSTD", "DiSTD", "MiSTD", "DoSTD", "FrSTD", "SaSTD", "SoSTD", "S", "F"]
    # Umrechnung von Sekunden auf Stunden (60*60)
    df_oevgebietedetail[["APSTD", "MoSTD", "DiSTD", "MiSTD", "DoSTD", "FrSTD", "SaSTD", "SoSTD"]] /= 3600

    for s in Saisons:
        df_oev_s = df_oevgebietedetail[df_oevgebietedetail[s] == 1]
        for f in FZG:
            df_oev_s_f = df_oev_s[df_oev_s["FZG"] == f]
            for t in Tage:
                # Gesamt FPLKM und FPLSTD je Saison, Fahrzeugkombination und Tag (über alle Kreise)
                KennzahlSTD, KennzahlKM = f"{s}_{f}_STD({t})", f"{s}_{f}_KM({t})"
                edictSTD, edictKM = df_oev_s_f.groupby("LINE")[f"{t}STD"].sum().to_dict(), df_oev_s_f.groupby("LINE")[f"{t}KM"].sum().to_dict()
                SetMulti(Visum.Net.Lines, KennzahlSTD, [edictSTD.get(line, 0) for _, line in Visum.Net.Lines.GetMultiAttValues("NAME",True)], True)
                # erstelle Liste in Linienreihenfolge passend zur Reihenfolge in Visum
                SetMulti(Visum.Net.Lines, KennzahlKM, [edictKM.get(line, 0) for _, line in Visum.Net.Lines.GetMultiAttValues("NAME",True)], True)
                for g in Gebiete:
                    # FPLKM und FPLSTD je Gebiet (Kreis) je Saison, Fahrzeugkombination und Tag
                    df_oev_s_f_g = df_oev_s_f[df_oev_s_f["Gebiet"] == g]
                    for e in Einheiten:
                        Kennzahl = f"{g}_{s}_{f}_{e}({t})"
                        edict = df_oev_s_f_g.groupby("LINE")[f"{t}{e}"].sum().to_dict()
                        SetMulti(Visum.Net.Lines, Kennzahl, [edict.get(line, 0) for _, line in Visum.Net.Lines.GetMultiAttValues("NAME",True)], True)
                        
    Visum.Log(20480, "Linienkennzahlen: berechnet")
    return True
    

def _checks(Visum):
    '''Führe vor der Berechnung die folgenden Checks durch:
        - Alle notwendigen BDA vorhanden?
        - Fahrplanfahrtabschnitte nur aus einem Teilnetz aktiv?
        - Haben alle aktiven Fahrplanfahrtabschnitte eine Fahrzeugkombination zugewiesen?
        - Sind die korrekten Gebiete aktiv und enthält der CODE keine Sonderzeichen?
        - Sind die korrekten Zeitintervallmengen und Zeitintervalle aktiv?
    '''
    if not check_fahrplanfahrtabschnitte(Visum):
        return False
    if not check_gebiete(Visum):
        return False
    if not check_bda(Visum, Visum.Net.VehicleJourneySections, "FahrpanfahrtAbschnitte", "S"):
        return False
    if not check_bda(Visum, Visum.Net.VehicleJourneySections, "FahrpanfahrtAbschnitte", "F"):
        return False
    if not check_zeitintervalle(Visum):
        return False
    return True


def _fahrzeugkombinationen(Visum):
    df_fahrplanfahrtabschnitte = pd.DataFrame(Visum.Net.VehicleJourneySections.GetMultipleAttributes([r"VEHCOMB\CODE"], True), columns = ["FZG"])
    fzgkombinationen_unique = df_fahrplanfahrtabschnitte['FZG'].drop_duplicates().tolist()
    return fzgkombinationen_unique


def _gebiete(Visum):
    _df_oevgebietedetail = pd.DataFrame(Visum.Net.TerritoryPuTDetails.GetMultipleAttributes([r"TERRITORYCODE"], True),
                            columns = ["CODE"])
    gebiete_unique = _df_oevgebietedetail['CODE'].drop_duplicates().values.tolist()
    return gebiete_unique


if __name__ == "__main__":           
    main(Visum)
