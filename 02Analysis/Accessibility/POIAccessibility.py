#-------------------------------------------------------------------------------
# Name:        Calculate POI-Accessibility in PTV Visum
# Author:      mape
# Created:     03/06/2025
# Copyright:   (c) mape 2025
# Licence:     GNU GENERAL PUBLIC LICENSE
# #---------------------

# # Parameters
# ConTable = ["E_Anbindungen"]
# ConTableFields = ["ID_ZENSUS", "ID_HSTBER", "TWALK"] #ID POI, ID Stop Area, Impedance
# POIValues = [23, "ID"]
# startMinutes = 5


import pandas as pd
from VisumPy.helpers import SetMulti

pd.set_option("future.no_silent_downcasting", True)

def mergeTables(_ConUDT, _StopAreaTable, _startMinutes):
    ConStopTable = pd.merge(_ConUDT, _StopAreaTable, left_on = "stopareano", right_on = "no", how = "inner")
    ConStopTable["isominutes"] = ConStopTable["impedance"] + ConStopTable["minutes"]
    min_idx = ConStopTable.groupby("idpoi")["isominutes"].idxmin()
    
    _minConStopTable = ConStopTable.loc[min_idx].reset_index(drop=True)
    _minConStopTable["isominutes"] = _minConStopTable["isominutes"] + _startMinutes
    
    return _minConStopTable

def POITable(_Visum, _POIValues, _minConStopTable):
    _POITable = Visum.Net.POICategories.ItemByKey(int(_POIValues[0]))
    _POITable = pd.DataFrame(_POITable.POIs.GetMultipleAttributes([_POIValues[1]], True))
    _POITable.columns = [_POIValues[1]]
    _POIConStopTable = pd.merge(_POITable, _minConStopTable, left_on = _POIValues[1], right_on = "idpoi", how = "left")
    _POIConStopTable["isominutes"] = _POIConStopTable["isominutes"].fillna(999)
    
    _POIConStopTableList = _POIConStopTable["isominutes"].tolist()
    SetMulti(_Visum.Net.POICategories.ItemByKey(int(_POIValues[0])).POIs, "ISOTIME", _POIConStopTableList, True)
    
def UDAPOI(_POIValues):
    UDAName = "ISOTIME"
    if Visum.Net.POICategories.ItemByKey(int(_POIValues[0])).POIs.AttrExists(UDAName):
        return
    POI = Visum.Net.POICategories.ItemByKey(23)
    POI.POIs.AddUserDefinedAttribute(UDAName, UDAName, UDAName, 2, 2, Ignored=False, MinVal=None, MaxVal=None, DefVal=999)
    uda = Visum.Net.POICategories.ItemByKey(23).POIs.Attributes.ItemByKey(UDAName)
    uda.Comment = "TravelTime from PT-Isochrones and connection table"
    uda.UserDefinedGroup = "Erreichbarkeiten"

def UDG():
    if not any(name == "Erreichbarkeiten" for _, name in Visum.Net.UserDefinedGroups.GetMultiAttValues("Name")):
        Visum.Net.AddUserDefinedGroup("Erreichbarkeiten", "Erreichbarkeitsindikatoren und Zwischenschritte")

def Tables(Con, ConFields):
    # Connection Table
    _ConUDT = Visum.Net.TableDefinitions.ItemByKey(Con[0])
    _ConUDT = pd.DataFrame(_ConUDT.TableEntries.GetMultipleAttributes(ConFields, True))
    _ConUDT.columns = ["idpoi", "stopareano", "impedance"]
    # StopAereas Table
    _StopAreaTable = pd.DataFrame(Visum.Net.StopAreas.GetMultipleAttributes(["NO", "ISOCTIMEPUT", "ISOCTRANSFERSPUT"], True))
    _StopAreaTable.columns = ["no", "seconds", "transfers"]
    _StopAreaTable["minutes"] = _StopAreaTable["seconds"] / 60
    _StopAreaTable = _StopAreaTable[_StopAreaTable["minutes"] < 300]
    
    return _ConUDT, _StopAreaTable


# Execution
UDG()
UDAPOI(POIValues)
ConUDT, StopAreaTable = Tables(ConTable, ConTableFields)
minConStopTable = mergeTables(ConUDT, StopAreaTable, startMinutes)
POITable(Visum, POIValues, minConStopTable)