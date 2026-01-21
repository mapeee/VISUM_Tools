#!/usr/bin/env python3
"""
Erstelle die für die LKE notwendigen BDA
Ebenen:
    - Fahrzeugkombinationen
    - Linien

Erstellt: 28.08.2025
@author: mape
Version: 0.8
"""

import pandas as pd


def bda_fahrzeugkombinationen(Visum, TN, Gebiete, Einheiten):
    '''Erstelle auf Ebene der Fahrzeugkombinationen die Abrechnungs-BDA
    Erstelle die BDA in einer bestimmten Reihenfolge, da Formel-BDA das Vorhandensein
    anderer BDA voraussetzen.
    
    Parameters
    ----------
    Visum : Visum Instanz
    TN : string
        TN-CODE
    Gebiete : liste
        CODE und NAME der Gebiete
    Einheiten : liste
        CODE und NAME der Mengeneinheiten.

    Returns
    -------
    None.
    '''
    check_bdg(Visum, [[TN, False]])
    Kosten_AT = []
    for e, ename in Einheiten:
        kreise = []
        # Menge je Gebiet (Kreis) je Mengenart
        for g, gname in Gebiete:
            erstelle_bda_fahrzeugkombinationen(Visum, f"{TN}_{g}_{e}", f"TNM_{TN}", [2, 6], f"{TN} {gname} {ename}")
            kreise.append(f"[{TN}_{g}_{e}]")
        formel = " + ".join(kreise)
        # Gesamtmenge je Mengenart (über alle Gebiete/Kreise)
        erstelle_bda_fahrzeugkombinationen(Visum, f"{TN}_{e}", f"TNM_{TN}", [2, 6], f"{TN} {ename}", formel)
        kreise = []
        # Anteil je Gebiet je Mengenart an Gesamtmenge
        for g, gname in Gebiete:
            erstelle_bda_fahrzeugkombinationen(Visum, f"{TN}_{g}_{e}_ANTEIL", f"TNM_{TN}", [2, 6], f"{TN} {gname} Anteil {ename} in %", f"[{TN}_{g}_{e}] / [{TN}_{e}]")
            # Kosten je Gebiet (AT) je Mengenart
            e_dict = {"KM": "KOSTENSATZKMSERVICEGES",
                      "STD": "KOSTENSATZSTDSERVICEGES",
                      "FZG": "KOSTENSATZFZGEINHEITGES(AP)"}
            e_val = f"{e}_METN" if e == "FZG" else e
            erstelle_bda_fahrzeugkombinationen(Visum, f"{TN}_{g}_{e}_ATKOSTEN", f"TNM_{TN}", [2, 6], f"{TN} {gname} ATKOSTEN {ename}", f"[{TN}_{g}_{e}_ANTEIL] * [{TN}_{e_val}] * [{e_dict[e]}]")
            kreise.append(f"[{TN}_{g}_{e}_ATKOSTEN]")
        formel = " + ".join(kreise)
        # Kosten je Mengenart (über alle Gebiete/Kreise)
        erstelle_bda_fahrzeugkombinationen(Visum, f"{TN}_{e}_KOSTEN", f"TNM_{TN}", [2, 6], f"{TN} KOSTEN {ename}", formel)
        Kosten_AT.append(f"[{TN}_{e}_KOSTEN]")
    formel = " + ".join(Kosten_AT)
    # Gesamtkosten
    erstelle_bda_fahrzeugkombinationen(Visum, f"{TN}_GESAMTKOSTEN", f"TNM_{TN}", [2, 6], f"{TN} GESAMTKOSTEN", formel)
    Visum.Log(20480, "TNM-Rechenattribute für Abrechnung (Fahrzeuge): erzeugt")
    

def bda_linien(Visum, stdkm=None):
    '''Erstelle auf Ebene von Linien BDA differenziert nach:
        - Gebiete der Teilnetze
        - Saisons (S, F)
        - Einheiten (Mengenart): Stunden, Kilometer, Fahrzeugbedarf
        - Fahrzeugkombinationen der Teilnetze
    stdkm : bool
        Für Stunden/Kilometer oder Fahrzeugbedarf (M1, M2)
    '''
    Gebiete, Saisons, FZG = _werte(Visum)
    check_bdg(Visum, Gebiete)
    check_bdg (Visum, [["FZG", False]]) # Check for BDG "TNM_FZG"
    if stdkm:
        Einheiten_txt = "STD und KM"
        Einheiten = [["KM", "Fahrplankilometer", "", False], ["STD", "Fahrplanstunden", "", False]]
    else:
        Einheiten_txt = "Fahrzeugbedarf"
        Einheiten = [["M1", "Fahrzeugbedarf Methode 1", "_FZG", True], ["M2", "Fahrzeugbedarf Methode 2","_FZG", True]]
    for s, sname in Saisons:
        for f, fname in FZG:
            for e, ename, _, __ in Einheiten:
                erstelle_bda_linien(Visum, f"{s}_{f}_{e}", "TNM_FZG", [2, 6], f"{sname} {fname} {ename}")
            for g, gname in Gebiete:
                for e, ename, gr, formel in Einheiten:
                    if not erstelle_bda_linien(Visum, f"{g}_{s}_{f}_{e}", f"TNM_{g}{gr}", [2, 6], f"{gname} {sname} {fname} {ename}", formel):
                        return False
    
    Visum.Log(20480, f"TNM-Rechenattribute der Linien erzeugt: {Einheiten_txt}")
    return True


def check_bdg(Visum, _Gebiete):
    '''Prüfe, ob BDG (benutzerdefinierte Gruppen) vorhanden sind oder lege sie an.
    _Gebiete : Liste mit Gebieten (CODE, NAME) oder Teilnetz-CODE
        Wenn Teilnetz-CODE, dann zweiter Wert False
    '''
    BDGs = Visum.Net.UserDefinedGroups.GetMultiAttValues("NAME")
    for g, gname in _Gebiete:
        # if False, g ist CODE von Teilnetz (TNM_Fahrzeugbedarf)
        if not gname:
            if not any(f"TNM_{g}" in BDG for _, BDG in BDGs):
                Visum.Net.AddUserDefinedGroup(f"TNM_{g}", f"Teilnetz {g}")
            return
        if not any(f"TNM_{g}" in BDG for _, BDG in BDGs):
            Visum.Net.AddUserDefinedGroup(f"TNM_{g}", f"TNM {gname}")
        if not any(f"TNM_{g}_FZG" in BDG for _, BDG in BDGs):
            Visum.Net.AddUserDefinedGroup(f"TNM_{g}_FZG", f"TNM {gname} (Fahrzeugbedarf)")
    

def erstelle_bda_fahrzeugkombinationen(Visum, _name, _group, _type, _comment, _formel=None):
    '''
    Parameters
    ----------
    Visum : Visum-Instanz
    _name : string
        Name des neuen BDA
    _group : string
        BDG des neuen BDA
    _type : list[int, int]
        list[0] = Werttyp
        list[1] = Anzahl Dezimalstellen
    _comment : string
        Kommentar des neuen BDA
    _formel : string
        zu erstellende Formel

    Returns
    -------
    bool
        Kann nicht falsch sein, da alle in den Formeln notwendigen BDA vorher erstellt oder abgefragt (_METN) werden.
    '''
    if Visum.Net.VehicleCombinations.AttrExists(_name):
        return True
    if _formel:
        Visum.Net.VehicleCombinations.AddUserDefinedAttribute(_name, _name, _name, _type[0], _type[1], Formula = _formel)
    else:
        Visum.Net.VehicleCombinations.AddUserDefinedAttribute(_name, _name, _name, _type[0], _type[1], False, 0, None, 0)
    bda = Visum.Net.VehicleCombinations.Attributes.ItemByKey(_name)
    bda.Comment = f"{_comment} | TNM-Rechenattribut"
    bda.UserDefinedGroup = _group
    return True


def erstelle_bda_linien(Visum, _name, _group, _type, _comment, _formel=None):
    '''Erstellung von BDA zur Speicherung von Leistungsmengen.
    Erstelle diese BDA nur, wenn sie noch nicht vorhanden sind.
    Ordne alle BDA einer BDG zur Strukturierung zu.
    
    Parameters
    ----------
    _name : string
        Name des neuen BDA
    _group : string
        BDG des neuen BDA
    _type : list[int, int]
        list[0] = Werttyp
        list[1] = Anzahl Dezimalstellen
    _comment : string
        Kommentar des neuen BDA
    _formel : bool, optional
        Erstelle als Formel-BDA ja oder nein

    Returns
    -------
    bool
        Wenn False, dann konnte Formel-BDA nicht erstellt werden, weil ein anderes BDA fehlt.
    '''
    if Visum.Net.Lines.AttrExists(_name):
        return True
    # Formel nur für Fahrzeugbedarfe auf Linienebene, da sich der territoriale Anteil aus dem Bedarf je Linie und dem 
    # territorialen Anteil der Stunden ergibt.
    if _formel:
        _g, _s, _f, _m = _name.split("_")
        # prüfe, ob die notwendigen BDA mit Bezug zur richtigen Zeitintervallmenge bestehen
        # (Mo) steht stellvertretend für den Abgleich mit Zeitintervallmenge
        attr = [f"{_g}_{_s}_{_f}_STD(Mo)", f"{_s}_{_f}_STD(Mo)", f"{_s}_{_f}_{_m}(Mo)"]
        for i in attr:
            if not Visum.Net.Lines.AttrExists(i):
                Visum.Log(12288, "Formel-BDA {_name} kann nicht erstellt werden: BDA {i} felt")
                return False
        attr = [s.replace("Mo", "") for s in attr]
        formel = f"([{attr[0]}] / [{attr[1]}]) * [{attr[2]}]"
        Visum.Net.Lines.AddUserDefinedAttribute(_name, _name, _name, _type[0], _type[1], Formula = formel, SubAttr = 'TISET_1') 
    else:
        Visum.Net.Lines.AddUserDefinedAttribute(_name, _name, _name, _type[0], _type[1], False, 0, None, 0, SubAttr = 'TISET_1')
    bda = Visum.Net.Lines.Attributes.ItemByKey(_name)
    bda.Comment = f"{_comment} | TNM-Rechenattribut"
    bda.UserDefinedGroup = _group
    return True
      

def _werte(Visum):
    _Gebiete = _gebiete(Visum)
    _Saisons = [["S", "Schule"], ["F", "Ferien"]]
    _FZG = _fahrzeugkombinationen(Visum)
    return _Gebiete, _Saisons, _FZG


def _fahrzeugkombinationen(Visum):
    df_fahrplanfahrtabschnitte = pd.DataFrame(Visum.Net.VehicleJourneySections.GetMultipleAttributes([r"VEHCOMB\CODE", r"VEHCOMB\NAME"], True),
                         columns = ["CODE", "NAME"])
    fzgkombinationen_unique = df_fahrplanfahrtabschnitte[['CODE', 'NAME']].drop_duplicates().values.tolist()
    return fzgkombinationen_unique

    
def _gebiete(Visum):
    df_oevgebietedetail = pd.DataFrame(Visum.Net.TerritoryPuTDetails.GetMultipleAttributes([r"TERRITORYCODE", "TERRITORYNAME"], True),
                            columns = ["CODE", "NAME"])
    gebiete_unique = df_oevgebietedetail[['CODE', "NAME"]].drop_duplicates().values.tolist()
    return gebiete_unique
         