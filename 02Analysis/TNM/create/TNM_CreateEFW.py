#!/usr/bin/env python3
"""
Erzeuge eine leere BDT (benutzerdefinierte Tabelle) für EFW (Entfernungswerk) des aktiven Teilnetzes

Erstellt: 25.12.2025
@author: mape
Version: 0.9
"""


def erstelle_bdt(Visum):
    '''Erzeuge eine leere BDT EFW für das aktive Teilnetz (TN)
    '''
    if not Visum.Net.AttrExists("TN"):
        Visum.Log(12288, "Netzattribut: BDA 'TN' fehlt")
        return False
    TN = Visum.Net.AttValue("TN")
    BDT = Visum.Net.TableDefinitions.GetMultiAttValues("NAME")
    if any(f"{TN} EFW" in item for item in BDT):
        Visum.Log(12288, f"BDT existiert schon: {TN} EFW")
        return False
    EFW = Visum.Net.AddTableDefinition(f"{TN} EFW")
    EFW.SetAttValue("GROUP","EFW")
    EFW.SetAttValue("COMMENT",f"Entfernungswerk TN {TN}")
    erstelle_bda(Visum, EFW)
    Visum.Log(20480, f"BDT erstellt: {TN} EFW")
    return True

def erstelle_bda(Visum, _EFW):
    '''Erstelle die für die EFW-Tabelle notwendigen BDA
    '''
    BDGs = Visum.Net.UserDefinedGroups.GetMultiAttValues("NAME")
    if not any("EFW" in BDG for _, BDG in BDGs):
        Visum.Net.AddUserDefinedGroup("EFW", "Entfernungswerk")
    for BDA in ["Linie", "VonHPName", "NachHPName"]:
        _EFW.TableEntries.AddUserDefinedAttribute(BDA, BDA, BDA, 5)
        BDA = _EFW.TableEntries.Attributes.ItemByKey(BDA)
        BDA.UserDefinedGroup = "EFW"
        BDA.MaxStringLen = 100
    
    _EFW.TableEntries.AddUserDefinedAttribute("METER", "Meter", "Meter", 1)
    BDA = _EFW.TableEntries.Attributes.ItemByKey("METER")
    BDA.UserDefinedGroup = "EFW"
    # BDA VORHANDEN ist WAHR, wenn die Zeile aus dem EFW im Netz verwendet wird.
    _EFW.TableEntries.AddUserDefinedAttribute(ID = "VORHANDEN", ShortName = "Vorhanden", LongName = "Vorhanden", VT = 9, \
                                              Formula = r"TableLookup(LINIENROUTENELEMENT LV; LV[HALTEPUNKT\NAME]=[VONHPNAME]&LV[NAECHSTROUTEPKT\HALTEPUNKT\NAME]=[NACHHPNAME]&LV[LINNAME]=[LINIE];LV[LINEROUTEID])")
    BDA = _EFW.TableEntries.Attributes.ItemByKey("VORHANDEN")
    BDA.UserDefinedGroup = "EFW"

if __name__ == "__main__":           
    erstelle_bdt(Visum)
