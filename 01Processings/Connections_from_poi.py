#-------------------------------------------------------------------------------
# Name:        Create connections from poi
# Author:      mape
# Created:     23/06/2025
# Copyright:   (c) mape 2025
# Licence:     GNU GENERAL PUBLIC LICENSE
# #---------------------

import pandas as pd
from VisumPy.helpers import SetMulti


def addConnections(_selCon, _oldCon, _tWalk):
    _meter, _seconds = [], []
    _nPrT = 0
    for index, row in _selCon.iterrows():
        if Visum.Net.Connectors.ExistsByKey(row["STOPAREA"], row["ZONE"]):
            Visum.Log(16384, f"PrT-Connection from Zone {int(row['ZONE'])} to Node {int(row['STOPAREA'])} just existis!")
            _nPrT += 1
            continue
        _con = Visum.Net.AddConnector(row["ZONE"], row["STOPAREA"])
        _meter.extend([row["avg_meter"] / 1000] * 2)
        _seconds.extend([_round30(row["avg_meter"] / (_tWalk / 3.6))] * 2)
    
    # Edit Attributes
    SetMulti(Visum.Net.Connectors, "TSYSSET", ["OEVFUSS"]*Visum.Net.Connectors.CountActive, True)
    SetMulti(Visum.Net.Connectors, "LENGTH", _meter, True)
    SetMulti(Visum.Net.Connectors, "T0_TSYS(OEVFUSS)", _seconds, True)
    
    # Check for ZONES without new PT-Connections
    _misingZones = int(_oldCon["ZONENO"].nunique() - _selCon["ZONE"].nunique())
    if _misingZones > 0:
        Visum.Log(16384, f"New PT-Connections for {_misingZones} Zone(s) missing!")
        for index, row in _oldCon[~_oldCon["ZONENO"].isin(_selCon["ZONE"])].iterrows():
            if Visum.Net.Connectors.ExistsByKey(row["NODENO"], row["ZONENO"]):
                _con = Visum.Net.Connectors.DestItemByKey(row["NODENO"], row["ZONENO"])
            else:
                _con = Visum.Net.AddConnector(row["ZONENO"], row["NODENO"])      
            _con.SetAttValue("TSYSSET", row["TSYSSET"])
            _con.SetAttValue("LENGTH", row["LENGTH"])
            _con.SetAttValue("T0_TSYS(OEVFUSS)", row["T0_TSYS(OEVFUSS)"])

    Visum.Log(20480, f"{len(_selCon)-_nPrT} new PT-Connectors added")
    return

def delConnections(_ZoneFilter = None):
    if not _ZoneFilter:
        return
    Visum.Filters.InitAll()
    conFilter = Visum.Filters.ConnectorFilter()
    conFilter.AddCondition("OP_NONE", False, "ZONENO", 7,_ZoneFilter[0])
    conFilter.AddCondition("OP_AND", False, "ZONENO", 5, _ZoneFilter[1])
    conFilter.AddCondition("OP_AND", False, "TSYSSET", 14, "OEVFUSS")
    _n_cons = int(Visum.Net.Connectors.CountActive / 2)
    _oldCons = pd.DataFrame(Visum.Net.Connectors.GetMultipleAttributes(["ZONENO", "NODENO", "TSYSSET", "LENGTH", "T0_TSYS(OEVFUSS)"], True))
    _oldCons.columns = ["ZONENO", "NODENO", "TSYSSET", "LENGTH", "T0_TSYS(OEVFUSS)"]
    Visum.Net.Connectors.RemoveAll()
    
    Visum.Log(20480, f"{_n_cons} existing PT-Connectors deleted")
    return _oldCons

def poi2Center(_attrCon):
    _x, _y = [], []
    df_zone = _attrCon.groupby(["ZONE"], as_index=False).agg({
        "pot_con": "sum", "pot_x": "sum", "pot_y": "sum"})
    df_zone["avg_x"] = df_zone["pot_x"] / df_zone["pot_con"]
    df_zone["avg_y"] = df_zone["pot_y"] / df_zone["pot_con"]
    
    for i in Visum.Net.Zones.GetMultipleAttributes(["NO", "XCOORD", "YCOORD"], False):
        if len(df_zone[df_zone["ZONE"]==i[0]]) == 0:
            _x.append(i[1])
            _y.append(i[2])
            continue
        if df_zone[df_zone["ZONE"]==i[0]]["pot_x"].iloc[0] == 0:
            _x.append(i[1])
            _y.append(i[2])
            continue
        _x.append(df_zone[df_zone["ZONE"]==i[0]]["avg_x"].iloc[0])
        _y.append(df_zone[df_zone["ZONE"]==i[0]]["avg_y"].iloc[0])
        
    SetMulti(Visum.Net.Zones, "XCOORD", _x, False)
    SetMulti(Visum.Net.Zones, "YCOORD", _y, False)    
    
    Visum.Log(20480, f"Centers of {len(df_zone)} Zones shifted")
    return

def poi2udt(_POICat, _UDTName):
    # poi
    df_poi = pd.DataFrame(Visum.Net.POICategories.ItemByKey(_POICat).POIs.GetMultipleAttributes(
        ["NO", "NAME", "CODE", "POTENTIAL_CON", "ZONE", "XCOORD", "YCOORD"]))
    df_poi.columns = ["NO", "NAME", "CODE", "POTENTIAL", "ZONE", "X", "Y"]
    df_poi["NAME"] = df_poi["NAME"].astype("string")
    df_poi["CODE"] = df_poi["CODE"].astype("string")
    
    # udt GIS meter
    udt = Visum.Net.TableDefinitions.ItemByKey(_UDTName)
    df_udt = pd.DataFrame(udt.TableEntries.GetMultipleAttributes(["NO_POTENTIAL"]))
    df_udt.columns = ["NO"]
    
    # merge
    df_merged = pd.merge(df_udt, df_poi, on="NO", how="left")
    for i in ["NAME", "CODE", "POTENTIAL", "ZONE", "X", "Y"]:
        my_list = df_merged[i].tolist()
        SetMulti(udt.TableEntries, i, my_list, True)
    
    Visum.Log(20480, "Attributes merged")
    return

def StopAreas2Connect(_attrCon):
    
    # at least 15% of potential
    _attrCon = _attrCon[(_attrCon['pot_con'] / _attrCon['pot_max']) >= 0.15]
    
    # at Least 120% potential of neutral StopArea (TYPE 8)
    max8 = _attrCon[_attrCon["TYPENO"]==8]
    max8 = max8.rename(columns={'pot_con': 'pot_8'})
    _attrCon = pd.merge(_attrCon, max8[['pot_8', 'ZONE', 'STOPNO']], on = ['ZONE', 'STOPNO'], how = "left")
    _attrCon['pot_8'] = _attrCon['pot_8'].fillna(1)
    _attrCon = _attrCon[((_attrCon['pot_con'] / _attrCon['pot_8']) >= 1.2) | (_attrCon['TYPENO'] == 8)]
    
    # avoid multiple cons to same lines
    _linesPot = _attrCon.explode('LINES')
    _linesPot = _linesPot.groupby(['ZONE', 'LINES'], as_index = False)['pot_con'].max()
    _attrCon["pot_lines"] = _attrCon.apply(_pot2Lines, _lines2stoparea = _linesPot, axis=1)
    _attrCon = _attrCon[_attrCon["pot_lines"] == 1]
    
    Visum.Log(20480, "StopAreas to be connected have been identified")
    return _attrCon

def weights2Connections(_UDTName, _pot_fact, _ZoneFilter = None):
    udt = Visum.Net.TableDefinitions.ItemByKey(_UDTName)
    df_udt = pd.DataFrame(udt.TableEntries.GetMultipleAttributes(
        ["NO", "CODE", "STOPAREA", "METER", "POTENTIAL", "ZONE", "X", "Y"]))
    df_udt.columns = ["NO", "CODE", "STOPAREA", "METER", "POTENTIAL", "ZONE", "X", "Y"]
    df_udt["CODE"] = df_udt["CODE"].astype("string")
    if _ZoneFilter:
        df_udt = df_udt[df_udt["ZONE"].between(_ZoneFilter[0], _ZoneFilter[1])]
        
    # weights
    df_udt["pot_con"] = df_udt["POTENTIAL"] * df_udt["CODE"].map(_pot_fact)
    df_udt["pot_zone"] = df_udt.groupby("ZONE")["pot_con"].transform("sum")
    df_udt["pot_con"] = ((1 - df_udt["METER"] / df_udt["METER"].max())**3) * df_udt["pot_con"]
    df_udt["pot_meter"] = df_udt["METER"] * df_udt["pot_con"]
    df_udt["pot_x"] = df_udt["X"] * df_udt["pot_con"]
    df_udt["pot_y"] = df_udt["Y"] * df_udt["pot_con"]
    df_udt = df_udt.groupby(["ZONE", "STOPAREA"], as_index=False).agg({
        "pot_con": "sum", "pot_zone": "first", "pot_meter": "sum", "pot_x": "sum", "pot_y": "sum"})
    df_udt['pot_max'] = df_udt.groupby('ZONE')['pot_con'].transform('max')
    df_udt["avg_meter"] = df_udt["pot_meter"] / df_udt["pot_con"]
        
    # StopAreas    
    df_stoparea = pd.DataFrame(Visum.Net.StopAreas.GetMultipleAttributes(
        ["NO", "STOPNO", "TYPENO", r"MIN:STOPPOINTS\DISTINCT:SERVINGVEHJOURNEYS\LINENAME"]))
    df_stoparea.columns = ["NO", "STOPNO", "TYPENO", "LINES"]
    df_stoparea = df_stoparea[~((df_stoparea['LINES'] == "") & (df_stoparea['TYPENO'] != 8))] # no stops without lines and TypNo != 8
    df_stoparea['LINES'] = df_stoparea['LINES'].str.split('|')
    
    # merge
    df_merged = pd.merge(df_udt, df_stoparea, left_on = "STOPAREA", right_on = "NO", how = "left")
    
    Visum.Log(20480, "Connections weighted and ready for calculations")
    return df_merged

def _pot2Lines(data, _lines2stoparea):
    if data["TYPENO"] == 8:
        return 1
    _df_lines = _lines2stoparea[_lines2stoparea["ZONE"]==data["ZONE"]]
    for i in data["LINES"]:
        if data["pot_con"] >= _df_lines.loc[(_df_lines['ZONE'] == data["ZONE"]) & (_df_lines['LINES'] == i), 'pot_con'].max() * 0.9:
            return 1 
    return 0

def _round30(n):
    return round(n / 30) * 30

Visum.Filters.InitAll()
if calpoi2udt:
    poi2udt(POICat, UDTName)
attrCon = weights2Connections(UDTName, pot_fact, ZoneFilter)
selCon = StopAreas2Connect(attrCon)
oldCon = delConnections(ZoneFilter)
addConnections(selCon, oldCon, tWalk)
if calpoi2Center:
    poi2Center(attrCon)
Visum.Log(20480, "Finished")
