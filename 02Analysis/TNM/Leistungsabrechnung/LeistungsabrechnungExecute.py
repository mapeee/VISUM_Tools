# -*- coding: utf-8 -*-

from VisumPy.AddIn import AddIn
_ = AddIn.gettext


def createEFW(Visum, TN):
    '''
    Erzeuge eine leere BDT EFW für das aktive Teilnetz (TN)
    '''
    EFW = Visum.Net.AddTableDefinition(f"{TN} EFW")
    EFW.SetAttValue("GROUP","EFW")
    EFW.SetAttValue("COMMENT",f"Entfernungswerk TN {TN}")
    _erstelle_bda(Visum, EFW)
    return True, _("EFW of SN: created")


def delSNUDA(Visum, TN):
    for i in Visum.Net.VehicleCombinations.Attributes.GetAll:
      if "TNM-Rechenattribut" in i.Comment and "_GESAMT" in i.Name and i.Editable == False:
        Visum.Net.VehicleCombinations.DeleteUserDefinedAttribute(i.ID)
    
    for i in Visum.Net.VehicleCombinations.Attributes.GetAll:
      if "TNM-Rechenattribut" in i.Comment and "_KOSTEN" in i.Name and i.Editable == False:
        Visum.Net.VehicleCombinations.DeleteUserDefinedAttribute(i.ID)
    
    for i in Visum.Net.VehicleCombinations.Attributes.GetAll:
      if "TNM-Rechenattribut" in i.Comment and "_ATKOSTEN" in i.Name and i.Editable == False:
        Visum.Net.VehicleCombinations.DeleteUserDefinedAttribute(i.ID)
    
    for i in Visum.Net.VehicleCombinations.Attributes.GetAll:
      if "TNM-Rechenattribut" in i.Comment and "_ANTEIL" in i.Name and i.Editable == False:
        Visum.Net.VehicleCombinations.DeleteUserDefinedAttribute(i.ID)
    
    for i in Visum.Net.VehicleCombinations.Attributes.GetAll:
      if "TNM-Rechenattribut" in i.Comment and i.Editable == False:
        Visum.Net.VehicleCombinations.DeleteUserDefinedAttribute(i.ID)
    
    for i in Visum.Net.VehicleCombinations.Attributes.GetAll:
      if "TNM-Rechenattribut" in i.Comment:
        Visum.Net.VehicleCombinations.DeleteUserDefinedAttribute(i.ID)
    
    return _("TNM-Attributes (SN): deleted")


def delLinesUDA(Visum):
    for i in Visum.Net.Lines.Attributes.GetAll:
      if "TNM-Rechenattribut" in i.Comment and i.Editable == False:
        Visum.Net.Lines.DeleteUserDefinedAttribute(i.ID)
    
    for i in Visum.Net.Lines.Attributes.GetAll:
      if "TNM-Rechenattribut" in i.Comment:
        Visum.Net.Lines.DeleteUserDefinedAttribute(i.ID)
    
    return _("TNM-Attributes (Lines): deleted")


def _erstelle_bda(Visum, _EFW):
    '''
    Erstelle die für die EFW-Tabelle notwendigen BDA
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
                                              Formula = r"TableLookup(LINIENROUTENELEMENT LV; LV[HALTEPUNKT\NAME]=[VONHPNAME]&LV[NAECHSTROUTEPKT\HALTEPUNKT\NAME]=[NACHHPNAME];LV[LINEROUTEID])")
    BDA = _EFW.TableEntries.Attributes.ItemByKey("VORHANDEN")
    BDA.UserDefinedGroup = "EFW"