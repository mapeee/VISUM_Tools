#!/usr/bin/env python3
"""
Prüfe die TNM Visum-Datei auf verschiedene Fehler.
Ebenen:
    - Fahrplanfahrtabschnitte
    - Fahrzeugkombinationen
    - Gebiete
    - Linien
    - Netzattribute
    - Zeitintervalle
    - Zeitintervallmengen
    

Erstellt: 18.12.2025
@author: mape
Version: 0.8
"""

import re
import pandas as pd


def check_tn(Visum):
    '''Führe vordefnierte Checks bei einem Direktaufruf dieser Funktionen durch.
    Direktaufruf wenn __name__ == "__main__" (Skriptende).
    '''
    check_net(Visum)
    check_linien(Visum)
    check_gebiete(Visum)
    check_fahrzeugkombinationen(Visum)
    check_fahrplanfahrtabschnitte(Visum)
    check_zeitintervalle(Visum)
    Visum.Log(20480, "Checks: Abgeschlossen")

    
def check_bda(Visum, Net, Type, BDA):
    '''Prüft, ob BDA vorhanden ist.
    Parameter
    ----------
    Visum : Visum Instanz
    Net : Visum Element
        z.B. Visum.Net.Lines
    Type : string
        Typ des Visum-Elements (z.B. 'Linien').
    BDA : string
        Name des BDA, dessen Existens geprüft werden soll.

    Returns
    -------
    bool
    '''
    if not Net.AttrExists(BDA):
        Visum.Log(12288, f"{Type}: BDA '{BDA}' fehlt")
        return False
    return True


def check_fahrplanfahrtabschnitte(Visum, TN=None):
    '''Führe folgende Checks für alle aktiven FahrplanfahrtAbschnitte durch:
        - notwendige BDA vorhanden?
        - FahrplanfahrtAbschnitte aktiv?
        - TN bei aktivem Teilnetz:
            - Nur FahrplanfahrtAbschnitte aus einem Teilnetz
            - Keine FahrplanfahrtAbschnitte ohne Teilnetz
        - FahrplanfahrtAbschnitte ohne zugeordneter Fahrzeugkombination?
        - BDA für Saison (Schule/Ferien) richtig belegt?
    '''
    for i in ["SAISON", "S", "F"]:
        if not check_bda(Visum, Visum.Net.VehicleJourneySections, "FahrpanfahrtAbschnitte", i):
            return False
    if Visum.Net.VehicleJourneySections.CountActive == 0:
        Visum.Log(12288, "Keine aktiven FahrpanfahrtAbschnitte vorhanden")
        return False
    df_fplfahrtabschnitte = pd.DataFrame(Visum.Net.VehicleJourneySections.GetMultipleAttributes(
                        [r"VEHCOMB\CODE", "SAISON", r"VEHJOURNEY\LINEROUTE\LINE\TN"], True),
                        columns = ["FZG", "SAISON", "TN"]
                        )
    # Nur bei Berechnungen für ein Teilnetz (Fahrzeugbedarf, Abrechnung)
    if TN:
        if df_fplfahrtabschnitte["TN"].nunique() > 1:
            Visum.Log(12288, "Aktive FahrpanfahrtAbschnitte aus mehr als einem Teilnetz")
            return False
        if df_fplfahrtabschnitte["TN"].unique()[0] == "ohne":
            Visum.Log(12288, "Nur FahrpanfahrtAbschnitte ohne Teilnetz aktiv")
            return False
    if df_fplfahrtabschnitte['FZG'].isna().any():
        Visum.Log(12288, f"Aktive FahrpanfahrtAbschnitte ohne Fahrzeug: {df_fplfahrtabschnitte['FZG'].isna().sum()}")
    df_fplfahrtabschnitte = df_fplfahrtabschnitte[['SAISON']].drop_duplicates()
    if not df_fplfahrtabschnitte['SAISON'].isin(["S", "F", "S+F"]).all():
        Visum.Log(12288, "BDA 'SAISON' enthält ungültige Werte (gültig: 'S', 'F', 'S+F')")
    return True 


def check_fahrzeugkombinationen(Visum):
    '''Führe folgende Checks für alle Fahrzeugkombinationen durch:
        - notwendige BDA vorhanden?
        - Fahrzeugkombinationen vorhanden?
        - Verwendung der zugelassenen Zeichen in Attribut CODE?
    '''
    if not check_bda(Visum, Visum.Net.VehicleCombinations, "Fahrzeugkombinationen", "RANG"):
        return False
    if Visum.Net.VehicleCombinations.Count == 0:
        Visum.Log(12288, "Das Netz enthält keine Fahrzeugkombinationen")
        return False
    df_fahrzeugkomb = pd.DataFrame(Visum.Net.VehicleCombinations.GetMultipleAttributes(["CODE"], False),
                               columns = ["CODE"])
    if not _check_zeichen(Visum, "Fahrzeugkombinationen", df_fahrzeugkomb):
        return False
    return True
    

def check_gebiete(Visum):
    '''Führe folgende Checks für alle aktiven Gebiete(Kreise) durch:
        - Sind Gebiete von Typ 1 aktiv (Kreise)?
        - Verwendung der zugelassenen Zeichen in Attribut CODE?
    '''
    df_gebiete = pd.DataFrame(Visum.Net.Territories.GetMultipleAttributes(["CODE", "TYPENO"], True),
                               columns = ["CODE", "TYPENO"])
    df_gebiete = df_gebiete[df_gebiete['TYPENO'].eq(1)]
    if len(df_gebiete) == 0:
        Visum.Log(12288, "Keine aktiven Gebiete von Typ 1 vorhanden")
        return False
    if not _check_zeichen(Visum, "Gebiete", df_gebiete):
        return False
    return True
    

def check_linien(Visum):
    '''Führe folgende Checks für alle Linien durch:
        - notwendige BDA vorhanden?
        - Verwendung der zugelassenen Zeichen in Attribut CODE?
    '''
    if not check_bda(Visum, Visum.Net.Lines, "Linien", "TN"):
        return False
    df_linien = pd.DataFrame(Visum.Net.Lines.GetMultipleAttributes(["TN"], False),
                               columns = ["TN"])
    df_linien = df_linien[df_linien['TN'] != 'ohne']
    df_linien = df_linien[['TN']].drop_duplicates()
    # rename, damit passend zur Schreibeweise in _check_zeichen
    df_linien.rename(columns={'TN': 'CODE'}, inplace=True)
    # Prüfe auf Sonderzeichen, Leerwerte etc.
    if not _check_zeichen(Visum, "Linien", df_linien):
        return False
    return True
    

def check_net(Visum):
    '''Führe folgende Checks für Netzattribute durch:
        - notwendige BDA vorhanden?
        - Verwendung der zugelassenen Zeichen in Attribut CODE?
    ''' 
    if not all((
        check_bda(Visum, Visum.Net, "Network", "TN"),
        check_bda(Visum, Visum.Net, "Network", "ANABOVERLAP"),
        check_bda(Visum, Visum.Net, "Network", "M_FAHRZEUGBEDARF"),
        check_bda(Visum, Visum.Net, "Network", "F_WOCHEN"),
        check_bda(Visum, Visum.Net, "Network", "S_WOCHEN"),
    )):
        return False

    TN = Visum.Net.AttValue("TN")
    if re.search(r'[ÖöÄäÜüß#\- ]', TN):
        Visum.Log(12288, "Netzattribute: Es gibt Sonderzeichen (#, ü, ä, -, ' ' etc.) im BDA TN")
    return True
    

def check_zeitintervalle(Visum):
    '''Führe folgende Checks für Zeitbezüge und Zeitintervallmengen durch:
        - Ist ein Wochenkalender eingestellt?
        - Enthält das Netz Zeitintervallmengen?
        - Ist eine Zeitintervallmenge aktiv?
        - Ist eine Zeitintervallmenge 'TNM_Tage' im Netz vorhanden?
        - Heißt die aktive Zeitintervallmenge 'TNM_Tage'?
        - Enthält die Zeitintervallmange 'TNM_Tage' 7 Zeitintervalle (eines je Wochentag)?
        - Entspricht die Gesamtdauer der Zeitintervalle einer Woche?
        - Sind die korrekten Tagescodes in den Zeitintervallen vorhanden?
    '''
    if Visum.Net.CalendarPeriod.AttValue("TYPE") != "CALENDARPERIODWEEK":
        Visum.Log(12288, "Das ÖV-Angebot basiert nicht auf einem Wochenkalender")
        return False
    df_zeitintervalmengen = pd.DataFrame(Visum.Net.TimeIntervalSets.GetMultipleAttributes(["NO", "CODE", "ISANALYSISTIMEINTERVALSET"], False),
                                                                                      columns = ["NO", "CODE", "AKTIV"])
    if len(df_zeitintervalmengen) == 0:
        Visum.Log(12288, "Das Netz enthält keine Zeitintervallmenge")
        return False
    if df_zeitintervalmengen["AKTIV"].sum() == 0:
        Visum.Log(12288, "Im Netz ist keine Zeitintervallmenge aktiv")
        return False
    if not "TNM_Tage" in df_zeitintervalmengen["CODE"].values:
        Visum.Log(12288, "'TNM_Tage' ist nicht als Zeitintervallmenge vorhanden")
        return False
    if df_zeitintervalmengen.loc[df_zeitintervalmengen["AKTIV"] == 1, "CODE"].iloc[0] != "TNM_Tage":
        Visum.Log(12288, "'TNM_Tage' ist nicht die aktive Zeitintervallmenge")
        return False
    TNM_Zeitintervall = df_zeitintervalmengen.loc[df_zeitintervalmengen["AKTIV"] == 1]
    zeitintervalle = Visum.Net.TimeIntervalSets.ItemByKey(TNM_Zeitintervall["NO"][0])
    if zeitintervalle.TimeIntervals.Count != 7:
        Visum.Log(12288, f"Die Zeitintervallmenge {zeitintervalle.AttValue('NAME')} enthält keine 7 Wochentage")
        return False
    df_zeitintervalle = pd.DataFrame(zeitintervalle.TimeIntervals.GetMultipleAttributes(["CODE","DURATION"]),
                                     columns = ["CODE","DURATION"])
    if (df_zeitintervalle["DURATION"] != 86400).any(): # 86400 sind die täglichen Sekunden (24*60 = 1440*60 = 86400)
        Visum.Log(12288, "Einzelne Zeitintervalle sind nicht 1.440 Minuten lang")
        return False
    tagcodes = {"Mo", "Di", "Mi", "Do", "Fr", "Sa", "So"}
    fehlende_tagcodes = tagcodes - set(df_zeitintervalle["CODE"])
    if len(fehlende_tagcodes) > 0:
        Visum.Log(12288, f"Folgende Tage fehlen als Zeitintervall: {fehlende_tagcodes}")
        return False
    return True


def _check_zeichen(Visum, element, df_check):
    '''Prüfe, die CODEs der Fahrzeugkombinationen und Gebiete sowie das BDA TN der Linien auf:
        - Sonderzeiche
        - Leerwerte
        - Duplikate
    Parameter
    ----------
    Visum : Visum Instanz
    element : string
        Elementtyp, für den Zeichen geprüft werden.
    df_check : pandas DataFrame
        Enthält df mit allen zu prüfenden strings in Spalte CODE

    Returns
    -------
    bool
    '''
    if df_check['CODE'].str.contains(r'[ÖöÄäÜüß#\- ]', regex=True, na=False).any():
        meldungen = {"Linien": "BDA TN"}
        attr = meldungen.get(element, "CODES")
        Visum.Log(12288, f"{element}: Es gibt Sonderzeichen (#, ü, ä, -, ' ' etc.) im {attr}")
        return False
    if (df_check["CODE"].isna() | (df_check["CODE"].str.strip() == '')).any():
        Visum.Log(12288, f"{element}: Es gibt Leerwerte in CODES")
        return False
    if df_check['CODE'].duplicated().any():
        for row in df_check[df_check['CODE'].duplicated()].itertuples(index=False):
            Visum.Log(12288, f"{element}: Duplikat(e) in CODES: {row.CODE}")
            return False
    return True


if __name__ == "__main__":           
    check_tn(Visum)
