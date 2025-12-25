#!/usr/bin/env python3
"""
Erzeuge eine leere BDT (benutzerdefinierte Tabelle) für Entfernungswerk des aktiven Teilnetzes

Erstellt: 25.12.2025
@author: mape
"""

def CreateBDT(Visum, TN):
    BDT = Visum.Net.TableDefinitions.GetMultiAttValues("NAME")
    if any(f"{TN}_EFW" in item for item in BDT):
        Visum.Log(12288, f"BDT existiert schon: {TN}_EFW")
        return
    EFW = Visum.Net.AddTableDefinition(f"{TN} EFW")
    EFW.SetAttValue("GROUP","EFW")
    EFW.SetAttValue("COMMENT",f"Entfernungswerk TN {TN}")
    CreateBDA(Visum, EFW)
    Visum.Log(20480, f"BDT erstellt: {TN} EFW")

def CreateBDA(Visum, _EFW):
    for BDA in ["Linie", "VonHPName", "NachHPName"]:
        _EFW.TableEntries.AddUserDefinedAttribute(BDA, BDA, BDA, 5)
        BDA = _EFW.TableEntries.Attributes.ItemByKey(BDA)
        BDA.UserDefinedGroup = "EFW"
        BDA.MaxStringLen = 100
    
    _EFW.TableEntries.AddUserDefinedAttribute("METER", "Meter", "Meter", 1)
    BDA = _EFW.TableEntries.Attributes.ItemByKey("METER")
    BDA.UserDefinedGroup = "EFW"

    _EFW.TableEntries.AddUserDefinedAttribute(ID = "VORHANDEN", ShortName = "Vorhanden", LongName = "Vorhanden", VT = 9, \
    Formula = r"TableLookup(LINIENROUTENELEMENT LV; LV[HALTEPUNKT\NAME]=[VONHPNAME]&LV[NAECHSTROUTEPKT\HALTEPUNKT\NAME]=[NACHHPNAME]&LV[LINNAME]=[LINIE];LV[LINEROUTEID])")
    BDA = _EFW.TableEntries.Attributes.ItemByKey("VORHANDEN")
    BDA.UserDefinedGroup = "EFW"


TN = Visum.Net.AttValue("TN")
CreateBDT(Visum, TN)
