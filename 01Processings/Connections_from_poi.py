#-------------------------------------------------------------------------------
# Name:        Create connections from poi
# Author:      mape
# Created:     23/06/2025
# Copyright:   (c) mape 2025
# Licence:     GNU GENERAL PUBLIC LICENSE
# #---------------------

import pandas as pd
from sklearn.cluster import DBSCAN
from VisumPy.helpers import SetMulti


def addConnections(_selCon, _tWalk, _shares = False):
    _meter, _seconds, _weights = [], [], []
    _nPrT = 0
    Visum.Net.Connectors.SetPassive()
    for index, row in _selCon.iterrows():
        if Visum.Net.Connectors.ExistsByKey(row["STOPAREA"], row["ZONE"]):
            Visum.Log(16384, f"PrT-Connection from Zone {int(row['ZONE'])} to Node {int(row['STOPAREA'])} just existis!")
            _nPrT += 1
            continue
        _con = Visum.Net.AddConnector(row["ZONE"], row["STOPAREA"])
        _con.Active = True
        _meter.extend([row["avg_meter"] / 1000] * 2)
        if _shares:
            _seconds.extend([240] * 2)
            _weights.extend([row["weights"] * 100] * 2)
        else:
            _seconds.extend([_round30(row["avg_meter"] / (_tWalk / 3.6))] * 2)
            _weights.extend([1] * 2)
    
    # Edit Attributes
    SetMulti(Visum.Net.Connectors, "TSYSSET", ["OEVFUSS"] * Visum.Net.Connectors.CountActive, True)
    SetMulti(Visum.Net.Connectors, "LENGTH", _meter, True)
    SetMulti(Visum.Net.Connectors, "T0_TSYS(OEVFUSS)", _seconds, True)
    SetMulti(Visum.Net.Connectors, "WEIGHT(PUT)", _weights, True)
    Visum.Net.Connectors.SetActive()
    
    if _shares: SetMulti(Visum.Net.Zones, "SHAREPUT", [1] * Visum.Net.Zones.CountActive, True)
    else: SetMulti(Visum.Net.Zones, "SHAREPUT", [0] * Visum.Net.Zones.CountActive, True)
    
    Visum.Log(20480, f"{(len(_selCon)-_nPrT) * 2} new PT-Connectors added")
    return

def delConnections(_selCon):
    Cons = Visum.Net.Connectors.GetAll
    _n_cons = 0
    for con in Cons[::2]: # every second for reverse direction
        if not con.AttValue("ZONENO") in selCon["ZONE"].values:
            continue
        if con.AttValue("TSYSSET") != "OEVFUSS":
            continue
        if con.AttValue("TYPENO") == 5:
            continue
        Visum.Net.RemoveConnector(con)
        _n_cons+=2

    Visum.Log(20480, f"{_n_cons} existing PT-Connectors deleted")
    return _n_cons

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

def StopAreas2Connect(_attrCon, _shares = False):
    
    # at least 15% of max potential
    _attrCon = _attrCon[(_attrCon['pot_con'] / _attrCon['pot_max']) >= 0.15]
    
    # at Least 120% potential of neutral StopArea (TYPE 8)
    max8 = _attrCon[_attrCon["TYPENO"]==8]
    max8 = max8.rename(columns={'pot_con': 'pot_8'})
    _attrCon = pd.merge(_attrCon, max8[['pot_8', 'ZONE', 'STOPNO', 'cluster']], on = ['ZONE', 'STOPNO', 'cluster'], how = "left")
    _attrCon['pot_8'] = _attrCon['pot_8'].fillna(1)
    _attrCon = _attrCon[((_attrCon['pot_con'] / _attrCon['pot_8']) >= 1.2) | (_attrCon['TYPENO'] == 8)]
    
    # at least 20% of max vehicle journeys
    if _shares:
        _attrCon['nVJmax'] = _attrCon.groupby(['ZONE', "cluster"])['nVJ'].transform('max')
        _attrCon = _attrCon[(_attrCon['nVJ'] / _attrCon['nVJmax']) >= 0.2]
        _attrCon['pot_con_sum'] = _attrCon.groupby(['ZONE', "cluster"])['pot_con'].transform('sum')
        _attrCon["weights"] = (_attrCon["cluster_pot_sum"] / _attrCon["cluster_pot_uniqueSum"]) * (_attrCon["pot_con"] / _attrCon["pot_con_sum"])
        _attrCon["avg_meter_sum"] = _attrCon["avg_meter"] * _attrCon["weights"]
        _attrCon = _attrCon.groupby(["ZONE", "STOPAREA"], as_index=False).agg({'weights': 'sum', 'avg_meter_sum': 'sum'})
        _attrCon["avg_meter"] = _attrCon['avg_meter_sum'] / _attrCon['weights']
    
    # avoid multiple cons to same lines
    if not _shares:
        _linesPot = _attrCon.explode('LINES')
        _linesPot = _linesPot.groupby(['ZONE', 'LINES'], as_index = False)['pot_con'].max()
        _attrCon["pot_lines"] = _attrCon.apply(_pot2Lines, _lines2stoparea = _linesPot, axis=1)
        _attrCon = _attrCon[_attrCon["pot_lines"] == 1]
    
    Visum.Log(20480, "StopAreas to be connected have been identified")
    return _attrCon

def weights2Connections(_UDTName, _pot_fact, _shares = False, _maxDist = 1500):
    _zones = pd.DataFrame(Visum.Net.Zones.GetMultipleAttributes(["NO"], True))
    _zones.columns = ["NO"]
    Visum.Log(20480, f"Creating new PT-Connectors for {len(_zones)} Zones")
    
    udt = Visum.Net.TableDefinitions.ItemByKey(_UDTName)
    df_udt = pd.DataFrame(udt.TableEntries.GetMultipleAttributes(
        ["NO", "CODE", "STOPAREA", "METER", "POTENTIAL", "ZONE", "X", "Y"]))
    df_udt.columns = ["NO", "CODE", "STOPAREA", "METER", "POTENTIAL", "ZONE", "X", "Y"]
    df_udt = df_udt[df_udt["ZONE"].isin(_zones["NO"])]
    
    # weights and cluster
    df_udt["pot_sum"] = df_udt["POTENTIAL"] * df_udt["CODE"].map(_pot_fact)
    df_udt["pot_con"] = ((1 - df_udt["METER"] / _maxDist)**3) * df_udt["pot_sum"]
    df_udt["pot_meter"] = df_udt["METER"] * df_udt["pot_con"]
    df_udt["pot_x"] = df_udt["X"] * df_udt["pot_con"]
    df_udt["pot_y"] = df_udt["Y"] * df_udt["pot_con"]
    aggr = {'pot_con': 'sum', 'pot_meter': 'sum', 'pot_x': 'sum', 'pot_y': 'sum'}
    df_udt["cluster"] = 1
    
    if _shares:
        df_udt = df_udt.groupby('ZONE', group_keys = False).apply(lambda z: _clusterXY(z, z.name),include_groups = False)
        df_udt["cluster_pot_sum"] = df_udt.groupby(['ZONE', "cluster"])['pot_sum'].transform('sum')
        df_udt["cluster_pot_max"] = df_udt.groupby('ZONE')['cluster_pot_sum'].transform('max')
        df_udt = df_udt[(df_udt["cluster_pot_max"] != 0) & (df_udt["cluster_pot_sum"] / df_udt["cluster_pot_max"] >= 0.1) & (df_udt["cluster"] != -1)]
        df_udt["cluster_n"] = df_udt.groupby('ZONE')['cluster'].transform('nunique')
        # df_udt["cluster_pot_uniqueSum"] = df_udt["cluster_pot_sum"].unique().sum()
        df_udt["cluster_pot_uniqueSum"] = df_udt.groupby("ZONE")["cluster_pot_sum"].transform(lambda x: x.unique().sum())
        aggr.update({'cluster_pot_sum': 'first', 'cluster_pot_max': 'first', "cluster_pot_uniqueSum": "first", 'cluster_n': 'first'})
        
    df_udt = df_udt.groupby(["ZONE", "STOPAREA", "cluster"], as_index=False).agg(aggr)
    df_udt['pot_max'] = df_udt.groupby(['ZONE', "cluster"])['pot_con'].transform('max')
    df_udt["avg_meter"] = df_udt["pot_meter"] / df_udt["pot_con"]
        
    # StopAreas    
    df_stoparea = pd.DataFrame(Visum.Net.StopAreas.GetMultipleAttributes(
        ["NO", "STOPNO", "TYPENO", r"MIN:STOPPOINTS\DISTINCT:SERVINGVEHJOURNEYS\LINENAME",
         r"STOP\DISTINCT:STOPAREAS\DISTINCT:STOPPOINTS\DISTINCT:SERVINGVEHJOURNEYS\LINENAME",
         r"SUM:STOPPOINTS\COUNT:SERVINGVEHJOURNEYS"], False))
    df_stoparea.columns = ["NO", "STOPNO", "TYPENO", "LINES", "LINES_STOP", "nVJ"]
    df_stoparea.loc[df_stoparea['TYPENO'] == 8, 'LINES'] = df_stoparea['LINES_STOP']
    df_stoparea.drop('LINES_STOP', axis = 1, inplace = True)
    df_stoparea = df_stoparea[~((df_stoparea['LINES'] == "") & (df_stoparea['TYPENO'] != 8))] # no stops without lines and TypNo != 8
    df_stoparea['LINES'] = df_stoparea['LINES'].str.split('|')
    df_stoparea['nLines'] = df_stoparea['LINES'].str.len()
    
    # merge
    df_merged = pd.merge(df_udt, df_stoparea, left_on = "STOPAREA", right_on = "NO", how = "left")
    
    Visum.Log(20480, "Connections weighted and ready for calculations")
    return df_merged

def _clusterXY(_dfXY, _zone, _eps = 150, _min_samples = 2):
    _dfXY["ZONE"] = _zone
    XY = _dfXY[['X', 'Y']]
    model = DBSCAN(eps = _eps, min_samples = _min_samples)
    _dfXY['cluster'] = model.fit_predict(XY)
    return _dfXY

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

if calpoi2udt:
    poi2udt(POICat, UDTName)
attrCon = weights2Connections(UDTName, pot_fact, shares)
selCon = StopAreas2Connect(attrCon, shares)
delConnections(selCon)
addConnections(selCon, tWalk, shares)
poi2Center(attrCon)
Visum.Log(20480, "Finished")
