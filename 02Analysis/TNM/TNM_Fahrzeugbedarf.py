# -*- coding: utf-8 -*-
"""
Berechne die Fahrzeugbedarfe auf Ebene der Teilnetz und Linien
Ebenen:
    - Linien
    - Fahrzeugkombinationen
    - Teilnetze
    - Fahrplanfahrtabschnitte

Erstellt: 01.09.2025
@author: mape
Version: 0.8
"""

import numpy as np
import pandas as pd
from VisumPy.helpers import SetMulti
from create.TNM_BDA import bda_linien, erstelle_bda_fahrzeugkombinationen, check_bdg
from utils.TNM_Checks import check_zeitintervalle, check_bda, check_fahrplanfahrtabschnitte


def main(Visum):
    '''Berechne Fahrzeugkombinationen auf Ebene der Teilnetze und Linien
    - Führe vorab unterschiedliche Checks durch
    - generiere TNM-Rechenattribute auf Linienebene
    - Erstelle fehlende TNM-Rechenattribut für METN-Bedarf für Fahrzeugkombinationen
    '''
    if not _checks(Visum):
        return  False
    if not bda_linien(Visum):
        return False
    # Parameter
    AnAbOverlap = Visum.Net.AttValue("ANABOVERLAP")
    TN = Visum.Net.AttValue("TN")
    
    check_bdg(Visum, [[TN, False]])
    erstelle_bda_fahrzeugkombinationen(Visum, f"{TN}_FZG_METN", f"TNM_{TN}", [1, 0], f"{TN} Fahrzeugbedarf gem. METN")
    Gebiete = _gebiete(Visum)
    Saisons = ["S", "F"]
    Tage = ["AP", "Mo", "Di", "Mi", "Do", "Fr", "Sa", "So"]
    FZG = _fahrzeugkombinationen(Visum)
    
    df_Fahrplanfahrten = _fahrplanfahrten(Visum, TN)
    
    FZG_Linien(Visum, df_Fahrplanfahrten, Saisons, Tage, FZG, Gebiete, AnAbOverlap)
    FZG_TN(Visum, df_Fahrplanfahrten, Saisons, Tage, FZG, TN, AnAbOverlap)
    
    Visum.Log(20480, "Fahrzeugbedarfe: berechnet")
    return True


def FZG_Linien(Visum, _FPLTabelle, _Saisons, _Tage, _FZG, _Gebiete, _AnAbOverlap):
    '''
    Parameters
    ----------
    Visum : Visum-Instanz
    _FPLTabelle : pandas DataFrame
        Beinhaltet alle Fahrplanfahrten mit Fahrzeug (inkl. RANG), Abafahrt, Ankunft, Saison und Verkehrstag
    _Saisons : Liste
        Liste mit gültigen Saisons ['S', 'F']
    _Tage : Liste
        Liste mit gültigen Verkehrstagen. Inkl. AP für Analyseperiode (Gesamtwoche)
    _FZG : pandas DataFrame
        Alle Fahrzeugkombinationen im aktiven TN mit Fahrzeug-CODE und RANG
    _Gebiete : Liste
        Liste mit den Gebietes-CODES der Kreise im aktiven TN
    _AnAbOverlap : bool (True oder False)
        True, wenn identische Akunfts- und Abfahrtszeit als Überlappung gelten

    Returns
    -------
    None
    Schreibe die Fahrzeugbedarfe direkt in Visum
    '''
    # Iteration übe ralle Saisons und Tage
    for s in _Saisons:
        for t in _Tage:
            # if (AP)? # ggf maxBedarf je Fahrzeug über alle Wochentage
            AddListeVisum = pd.DataFrame()
            # Iteration über alle Linien in TN die an s und t Fahrtenangebot aufweisen
            for linie in np.sort(_FPLTabelle["LINE"].unique()):
                FZG_linie = _berechne_FZG_Linien(_FPLTabelle, linie, _FZG, s, t, _AnAbOverlap)
                AddListeVisum = pd.concat([AddListeVisum, FZG_linie], ignore_index=True)
            AddListeVisum.fillna(0, inplace=True)
            AddListeVisum.drop(columns=["LINENAME"], inplace=True)
            # iteriere über alle Spaltennamen in Liste über alle Fahrzeugtypen je Linie
            for spalte_name, spalte_daten in AddListeVisum.items():
                spalte_AP = spalte_name.replace(t, "AP") # z.B. F_ESB_M1(MO) --> F_ESB_M1(AP)
                df_AP = pd.Series([x[1] for x in Visum.Net.Lines.GetMultiAttValues(spalte_AP, True)])
                df_AP = np.maximum(df_AP, spalte_daten) #Abgleich, ob neuer Tagesweg dem Maximum entspricht oder nicht. Falls ja, dann in AP
                SetMulti(Visum.Net.Lines, spalte_name, spalte_daten.tolist(), True)
                SetMulti(Visum.Net.Lines, spalte_AP, df_AP.tolist(), True)
              
                
def FZG_TN(Visum, _FPLTabelle, _Saisons, _Tage, _FZG, _TN, _AnAbOverlap):
    '''
    Parameters
    ----------
    Visum : Visum-Instanz
    _FPLTabelle : pandas DataFrame
        Beinhaltet alle Fahrplanfahrten mit Fahrzeug (inkl. RANG), Abafahrt, Ankunft, Saison und Verkehrstag
    _Saisons : Liste
        Liste mit gültigen Saisons ['S', 'F']
    _Tage : Liste
        Liste mit gültigen Verkehrstagen. Inkl. AP für Analyseperiode (Gesamtwoche)
    _FZG : pandas DataFrame
        Alle Fahrzeugkombinationen im aktiven TN mit Fahrzeug-CODE und RANG
    _TN : string
        Teilnetz-CODE
    _AnAbOverlap : bool (True oder False)
        True, wenn identische Akunfts- und Abfahrtszeit als Überlappung gelten

    Returns
    -------
    None
    Schreibe die Fahrzeugbedarfe direkt in Visum
    '''
    # Erstelle DataFrame mit Bedarfen je Fahrzeugtyp über alle Saisons und Tage
    FZG_METN = pd.DataFrame()
    for s in _Saisons:
        for t in _Tage:
            if t == "AP": # nicht relevant auf Ebene Teilnetz
                continue
            FZG_FK = _berechne_FZG_TN(_FPLTabelle, _FZG, s, t, _AnAbOverlap)
            FZG_METN[FZG_FK.columns] = FZG_FK # Füge Spalten mit Namen aus Saison, Tag, Fahrzeugtyp hinzu
    spalte_daten = []
    for _, f in Visum.Net.VehicleCombinations.GetMultiAttValues("CODE"):
        if FZG_METN.columns.str.startswith(f"{f}_").any():
            # Schreibt alle Spalten, in denen Fahrzeug f und _S_ für Normalwoche vorkommen
            cols = [c for c in FZG_METN.columns if "_S_" in c and c.startswith(f"{f}_")]
            # Maximum über alle Wochentage
            max_S = FZG_METN[cols].sum().max() # sum() wird benötigt, obwohl es nur eine Zeile gibt
            cols = [c for c in FZG_METN.columns if "_F_" in c and c.startswith(f"{f}_")]
            max_F = FZG_METN[cols].sum().max()
            # Maximum über Saison (je Saison Maximum am Wochetag, somit Gesamtmaximum)
            fmax = max(max_S, max_F)
        else: # Wenn Fahrzeugkombination im TN nicht enthalten ist
            fmax = 0
        spalte_daten.append(fmax)
        SetMulti(Visum.Net.VehicleCombinations, f"{_TN}_FZG_METN", spalte_daten)


def _berechne_FZG_Linien(_FPLTabelle, _linie, _FZG, _s, _t, _AnAbOverlap):
    '''
    Parameters
    ----------
    _FPLTabelle : pandas DataFrame
        Beinhaltet alle Fahrplanfahrten mit Fahrzeug (inkl. RANG), Abafahrt, Ankunft, Saison und Verkehrstag
    _linie: string
        Linienkürzel
    _FZG : pandas DataFrame
        Alle Fahrzeugkombinationen im aktiven TN mit Fahrzeug-CODE und RANG
    _s : string
        Saison: 'S' oder 'F'
    _t : string
        kürzel Wochentag inkl. AP
    _AnAbOverlap : bool (True oder False)
        True, wenn identische Akunfts- und Abfahrtszeit als Überlappung gelten

    Returns
    -------
    LinienListe
        Fahrzeugbedarf über alle Fahrzeuge in Saison, an Tag und für Linie (_linie)
    '''
    # Erstelle Linienliste mit Spalte Linienname und befülle erste Zeile mit Linienkürzel (_linie)
    LinienListe = pd.DataFrame(columns=["LINENAME"])
    LinienListe.loc[len(LinienListe)] = [_linie]
    # Filter Fahrplanfahrten nach Linie, Saison und Tag
    _FPLTabelle = _FPLTabelle[(_FPLTabelle["LINE"] == _linie) & (_FPLTabelle[_s] == 1) & (_FPLTabelle["Tag"] == _t)]
    timeline = _erstelle_timeline(_FPLTabelle, _AnAbOverlap) # baue Timeline mit Fahrzeugbedarf je Fahrzeugtyp und Tagesminute (1440)
    # Iteration über alle Fahrzeuge (diese liegen nach Rang sortiert vor)
    for f in _FZG["FZG"].unique().tolist():
        if not f in timeline.columns: # Falls Fahrzeug zwar in Teilnetz aber nicht in gefiltern Fahrplandaten vorhanden
            continue
        # Neue Spalte mit Maxwert über Gesamttag (1440 Minuten) für Fahrzeug f
        timeline[f"{f}max"] = timeline[f].max()
    # maximaler Gesamtfahrzeugebedarf
    timeline["FZGMAX"] = timeline.filter(like="max").sum(axis=1, skipna=True)
    # Iteration über alle Fahrzeuge (diese liegen nach Rang sortiert vor)
    for _f in _FZG["FZG"].unique().tolist():
        if not _f in timeline.columns:
            continue
        # Methode 1 ((Gesamtbedarf Linie / Summe Maximalbedarf alle Fahrzeug)e * Maximalbedarf Fahrzeug)
        LinienListe[f"{_s}_{_f}_M1({_t})"] = timeline["ALL"].max() / timeline["FZGMAX"] * timeline[f"{_f}max"]
        # Methode 2 (Aufsteigend nach Rangfolge:
            # - Berechne Maximalbedarf je Fahrzeug und setze ziehe diesen vom Gesamtbedarf der Linie ab
            # - Wenn nichts mehr abzuziehen ist, gibt es auch keinen Bedarf
        m2_sum = LinienListe.filter(like="M2").sum(axis=1, skipna=True)
        LinienListe[f"{_s}_{_f}_M2({_t})"] = timeline[f"{_f}max"].clip(upper=timeline["ALL"].max()-m2_sum) 
    return LinienListe


def _berechne_FZG_TN(_FPLTabelle, _FZG, _s, _t, _AnAbOverlap):
    '''
    Der Fahrzeugbedarf auf Ebene der Teilnetze ergibt sich wie folgt:
        Aufsteigend (nach Rang von klein zu groß) wird der maximale Fahrzeugbedarf je Tag und Saison ermittelt
        Fahrzeuge dürfen andere Fahrzeuge ersetzen, wenn sie einen höheren oder gleichen Rang aufweisen
        Es wird also für jede Tagesminute ermittelt, wie viele Fahrzeuge benötigt werden und wie viele Fahrzeuge mit
        einem kleineren Rang frei sind und diese ersetzen können.
    
    Parameters
    ----------
    _FPLTabelle : pandas DataFrame
        Beinhaltet alle Fahrplanfahrten mit Fahrzeug (inkl. RANG), Abafahrt, Ankunft, Saison und Verkehrstag
    _FZG : pandas DataFrame
        Alle Fahrzeugkombinationen im aktiven TN mit Fahrzeug-CODE und RANG
    _s : string
        Saison: 'S' oder 'F'
    _t : string
        kürzel Wochentag inkl. AP
    _AnAbOverlap : bool (True oder False)
        True, wenn identische Akunfts- und Abfahrtszeit als Überlappung gelten

    Returns
    -------
    LinienListe
        Fahrzeugbedarf über alle Fahrzeuge in Saison, an Tag und für Linie (_linie)
    '''
    # Filter Fahrplanfahrten nach Saison und Tag
    _FPLTabelle = _FPLTabelle[(_FPLTabelle[_s] == 1) & (_FPLTabelle["Tag"] == _t)]
    timeline = _erstelle_timeline(_FPLTabelle, _AnAbOverlap)
    timeline["FREI"] = 0 # Spalte mit freien Fahrzeugen je Tagesminute
    Rang0 = 0
    for f, Rang1 in zip(_FZG["FZG"], _FZG["RANG"]):
        if not f in timeline.columns: # Falls Fahrzeug zwar in Teilnetz aber nicht in gefiltern Fahrplandaten vorhanden 
            continue
        # prüft, ob Rang aus vorheriger Iteration kleiner oder gleich Rang aus aktueller Iteration
        # Wenn WAHR, dann dürfen Fahrzeuge aus aktueller Iteration durch freie Fahrzeuge ersetzt werden
        # Muss immer WAHR sein, da Fahrzeuge nach RANG sortiert (ggf. IF-Bedingung entfernen)
        if Rang0 <= Rang1:
            # in neuer Spalte wird eingetragen, wie viele Fahrzeuge durch freie Fahrzeuge zur Tagesminute ersetzt werden
            # Wenn Fahrzeuge zur Tagesminute FREI sind, wird minimum aus FREI und f als Ersatz genommen
            # Es könne nur Fahrzeuge ersetzen die FREI sind aber auch nicht mehr, als benötigt werden.
            timeline[f"{f}ersatz"] = timeline[["FREI", f]].min(axis=1).where(timeline["FREI"] > 0, 0)
            # Zieht zeilenweise (Tagesminuten) den Fahrzeugersatz von den freien und benötigten Fahrzeugen ab.
            # Anschließend sind weniger Fahrzeuge FREI und es werden weniger (oder keine) vom Typ f benötigt (fneu).
            timeline[[f"{f}neu", "FREI"]] = timeline[[f, "FREI"]].sub(timeline[f"{f}ersatz"], axis=0)
        timeline[f"{f}max"] = timeline[f"{f}neu"].max() # Maximalbedarf nach Ersatz von Fahrzeug f
        # Fur alle tagesminmuten (Zeilen): Freie Fahrzeuge = Maximalbedarf f über Tag im Bedarf zu dieser Minute
        timeline["FREI"] = timeline["FREI"] + (timeline[f"{f}max"] - timeline[f"{f}neu"])
        Rang0 = Rang1
    # Als Export werden nur die Spalten mit den Maximalbedarfen je Fahrzeugtyp benötigt
    if _s == "S" and _t == "Mo": # ermöglichst späteres Einfügen in Excel
        timeline.to_clipboard(index=False)
    TNListe = timeline.filter(like="max").copy()
    TNListe.drop(TNListe.index[1:], inplace=True) # Lösche alle Zeilen außer der ersten, da Maximalbedarfe über Tag immer identisch
    TNListe.rename(columns=lambda c: c.replace("max", f"_{_s}_({_t})"), inplace=True)
    return TNListe


def _checks(Visum):
    '''Führe unterschiedliche checks durch:
        - BDAs 'S', 'F', 'TN', 'RANG' vorhanden?
        - Zeitintervallmenge und Zeitintervalle korrekt definiert?
        - Nur Fahrten aus einem Teilnetz aktiv?
        - Alle Fahrten mit Fahrzeugkombinationen belegt?
        - Saisons korrekt vergeben ('S', 'S+F' oder 'F')?
    '''
    return all((
    check_bda(Visum, Visum.Net, "Network", "TN"),
    check_bda(Visum, Visum.Net, "Network", "ANABOVERLAP"),
    check_bda(Visum, Visum.Net.Lines, "Lines", "TN"),
    check_bda(Visum, Visum.Net.VehicleCombinations, "FahrzeugKombinationen", "RANG"),
    check_zeitintervalle(Visum),
    check_fahrplanfahrtabschnitte(Visum, True),
    ))


def _erstelle_timeline(_FPL_linie, _AnAbOverlap):
    '''
    Parameters
    ----------
    _FPL_linie : pandas DataFrame
        Beinhaltet alle Fahrplanfahrten mit Fahrzeug (inkl. RANG), Abafahrt, Ankunft, Saison und Verkehrstag der Linie
    _AnAbOverlap : bool (True oder False)
        True, wenn identische Akunfts- und Abfahrtszeit als Überlappung gelten

    Returns
    -------
    df_timeline: pandas DataFrame
        Fahrzeugbedarf je Tagesminute (1440) insgesamt und je Fahrzeugtyp
    '''
    df_timeline = pd.DataFrame({"minute": range(0, 1440)}) # erstelle timeline für 1440 Minuten am Tag (24 Stunden * 60 Minuten)
    df_timeline["time"] = pd.to_datetime(df_timeline["minute"], unit="m").dt.strftime("%H:%M")
    df_timeline["ALL"] = 0
    FZGunique = _FPL_linie["FZG"].drop_duplicates().tolist() # Unique Fahrzeuge in den Fahrplandaten
    df_timeline[FZGunique] = 0 # lege Spalten mit initialem Fahrzeugbedarf je Fahrzeug mit 0 an
    # iteration über die drei Spalten aus Fahrplandaten, Wert jeweils aus Zeile
    # erhöhe in Timeline jeweils den Zähler je Fahrzeug um eins, wenn betreffenden Zeitintervall (slice) Fahrten vorhanden
    for dep, arr, fzg in zip(_FPL_linie["DEP"], _FPL_linie["ARR"], _FPL_linie["FZG"]):
        if _AnAbOverlap:
            df_timeline.loc[dep:arr, "ALL"] += 1
            df_timeline.loc[dep:arr, fzg] += 1
        else:
            df_timeline.loc[dep:arr - 1, "ALL"] += 1 # if identical DEP and ARR not overlapping
            df_timeline.loc[dep:arr - 1, fzg] += 1 # if identical DEP and ARR not overlapping
    return df_timeline


def _fahrplanfahrten(Visum, _TN):
    '''
    Parameters
    ----------
    Visum : Visum-Instanz
    _TN : String
        CODE vom aktiven Teilnetz.

    Returns
    -------
    _VJTableDay : pandas DataFrame
        DataFrame mit den relevanten Attributen von Fahrplanfahrten über alle Verkehrstage und Saisons
        
        Es werden alle Fahrplanfahrten über den Tagwechsel kopiert und am Tagwechsel gebrochen.
        So entstehen zwei Fahrten, eine bis 24:00 Uhr und eine am Folgetag ab 00:00 Uhr
        Fahrten, die nach 24:00 Uhr beginnen, werden komplett in den Folgetag geschoeben und Abfahrt und Ankunft um 24 Stunden vorverlegt.
    '''
    
    _next_tag = {
    "Mo": "Di",
    "Di": "Mi",
    "Mi": "Do",
    "Do": "Fr",
    "Fr": "Sa",
    "Sa": "So",
    "So": "Mo"}
    _attrListe = [[r"VEHJOURNEY\LINENAME",r"VEHJOURNEY\NAME", "DEP", "ARR", r"VEHCOMB\CODE", r"VEHCOMB\RANG",
                 "S", "F",
                 r"ISVALID(Mo)", r"ISVALID(Di)", r"ISVALID(Mi)",
                 r"ISVALID(Do)", r"ISVALID(Fr)", r"ISVALID(Sa)",
                 r"ISVALID(So)"],
                ["LINE", "NAME", "DEP", "ARR", "FZG", "FZGRANG", "S", "F", "Mo", "Di", "Mi", "Do", "Fr", "Sa", "So"]]
    _FPLTabelle = pd.DataFrame(Visum.Net.VehicleJourneySections.GetMultipleAttributes(_attrListe[0], True), columns = _attrListe[1])
    # Wandle Sekundenwerte aus Visum in Minutenwerte um
    _FPLTabelle["DEP"], _FPLTabelle["ARR"] = _FPLTabelle["DEP"] / 60, _FPLTabelle["ARR"] / 60
    _FPLTabelleWoche = pd.DataFrame()
    # iteriere über alle Wochentage
    for tag in ["Mo", "Di", "Mi", "Do", "Fr", "Sa", "So"]:
        # nimm nur Fahrplanfahrten, die an dem jeweiligen Tag aktiv sind
        df_tag = _FPLTabelle.loc[_FPLTabelle[tag] == 1].copy()
        # Dupliziere alle Fahrten, die über den Tageswechsel gehen
        # Filter erst diese Fahrten, kopiere sind anschließend und fülle das Schlüsselfeld mit True
        df_tag["is_nextday"] = False
        maske = (df_tag["DEP"] < 1440) & (df_tag["ARR"] >= 1440) # Filter Fahrten über Tageswechsel (1440 Minuten = 24 Stunden)
        df_tagwechsel = df_tag[maske].copy() # Kopiere diese Fahrten
        df_tagwechsel["is_nextday"] = True
        df_tag = pd.concat([df_tag, df_tagwechsel], ignore_index=True) # Füge die duplizierten Fahrten ans Ende der Fahrtenliste an
        # Passe Ankunft und Abfahrt der duplizierten Fahrten an
        df_tag.loc[(df_tag["DEP"] < 1440) & (df_tag["ARR"] >= 1440) & (df_tag["is_nextday"] == False), ["ARR", "Tag"]] = [1440, tag] # erste Tag, Ankunft 23:59
        df_tag.loc[(df_tag["DEP"] < 1440) & (df_tag["ARR"] >= 1440) & (df_tag["is_nextday"] == True), ["DEP", "Tag"]] = [0, tag] # zweter Tag, Abfahrt 00:00
        df_tag.loc[df_tag["is_nextday"], "ARR"] -= 1440 # zweiter Tag, Ankunft -24 Stunden
        # Schiebe Fahrten nach 24 Uhr in das 0-24 Schema aber gültig am nächsten Tag
        maske = df_tag["DEP"] >= 1440
        df_tag.loc[maske, ["DEP", "ARR"]] -= 1440
        df_tag.loc[maske, "is_nextday"] = True
        # Ordner jeder Fahrt/duplizierten Fahrt den korrekten Fahrtag zu
        df_tag["Tag"] = np.where(df_tag["is_nextday"] == False, tag, _next_tag[tag]) # wenn nächster Tag = True, dann nächster Tag aus Liste
        _FPLTabelleWoche = pd.concat([_FPLTabelleWoche, df_tag], ignore_index=True)
        
    _FPLTabelleWoche.drop(columns=["Mo", "Di", "Mi", "Do", "Fr", "Sa", "So", "is_nextday"], inplace=True)
    # Lege die Feldwerte fest um spätere Fehler zu vermeiden (wenn dtype = object)
    _FPLTabelleWoche[["LINE", "NAME", "FZG", "Tag"]] = _FPLTabelleWoche[["LINE", "NAME", "FZG", "Tag"]].astype("string")
    _FPLTabelleWoche["FZGRANG"] = _FPLTabelleWoche["FZGRANG"].astype(int)
    return _FPLTabelleWoche


def _fahrzeugkombinationen(Visum):
    '''
    Parameters
    ----------
    Visum : Visum-Instanz

    Returns
    -------
    df_fahrplanfahrtabschnitte : pandas DataFrame
        DataFrame mit unterschiedlichen Fahrzeugkombinationen und dem Rang sortiert nach Rang (aufsteigend)
    '''
    # Nutze nur Fahrzeugkombinationen die auf aktiven Fahrplanfahrtabschnitten verwendet werden
    df_fahrplanfahrtabschnitte = pd.DataFrame(Visum.Net.VehicleJourneySections.GetMultipleAttributes([r"VEHCOMB\CODE", r"VEHCOMB\RANG"], True), columns = ["FZG", "RANG"])
    df_fahrplanfahrtabschnitte = df_fahrplanfahrtabschnitte[['FZG', "RANG"]].drop_duplicates().reset_index(drop=True)
    # Lege die Feldwerte fest um spätere Fehler zu vermeiden (wenn dtype = object)
    df_fahrplanfahrtabschnitte = df_fahrplanfahrtabschnitte.astype({"FZG": "string", "RANG": int})
    # Bringe in Reihenfolge nach BDA RANG (aufsteigend) für spätere Berechnungen
    df_fahrplanfahrtabschnitte.sort_values(by="RANG", inplace=True)
    return df_fahrplanfahrtabschnitte


def _gebiete(Visum):
    _df_oevgebietedetail = pd.DataFrame(Visum.Net.TerritoryPuTDetails.GetMultipleAttributes([r"TERRITORYCODE"], True),
                            columns = ["CODE"])
    gebiete_unique = _df_oevgebietedetail['CODE'].drop_duplicates().values.tolist()
    return gebiete_unique


if __name__ == "__main__":           
    main(Visum)
