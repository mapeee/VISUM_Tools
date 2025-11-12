# -*- coding: utf-8 -*-
"""
Created on Wed Oct 29 12:55:50 2025

@author: peter
"""
import numpy as np
import pandas as pd
from shapely import wkt

from VisumPy import helpers
from VisumPy.helpers import SetMulti


def haversine_km(lon1, lat1, lon2, lat2):
    R = 6371.0088  # mean Earth radius in km
    lon1, lat1, lon2, lat2 = map(np.radians, (lon1, lat1, lon2, lat2))
    dlon = lon2 - lon1
    dlat = lat2 - lat1
    a = np.sin(dlat/2.0)**2 + np.cos(lat1) * np.cos(lat2) * np.sin(dlon/2.0)**2
    return 2 * R * np.arcsin(np.sqrt(a))


# Skim-Matrix / StopAreas
dfIso = pd.DataFrame(helpers.GetMatrixRaw(Visum, 63, intoMat=None))
# dfIso = Visum.Net.Matrices.ItemByKey(63).GetValuesFloat()
# dfIso = pd.DataFrame(dfIso)
stops = pd.DataFrame(Visum.Net.StopAreas.GetMultipleAttributes(["NO", r"NODE\MIN:CONTAININGZONES\LAGE"]), columns = ["NO","area"])
stops["NO"] = stops["NO"].astype("int32")
dfIso.columns = stops["NO"].tolist()
dfIso.index = stops["NO"].tolist()
stops = stops[stops["area"] <= area]
dfIso = dfIso.stack().reset_index()
dfIso.columns = ["fromStop", "toStop", "time"]
dfIso = dfIso[dfIso["time"] <= maxTime]
dfIso[["fromStop", "toStop"]] = dfIso[["fromStop", "toStop"]].astype("int32")
dfIso["time"] = pd.to_numeric(dfIso["time"], downcast="float")

# Access / Egress
con = Visum.Net.TableDefinitions.ItemByKey("E_Anbindungen").TableEntries.GetMultipleAttributes(["ID_ZENSUS", "EW","TWALK","ID_HSTBER"], False)
con = pd.DataFrame(con, columns=["ID_ZENSUS", "EW","TWALK","ID_StopArea"])
con[["ID_ZENSUS", "EW", "ID_StopArea"]] = con[["ID_ZENSUS", "EW", "ID_StopArea"]].astype("int32")
con["TWALK"] = pd.to_numeric(con["TWALK"], downcast="float")
con = pd.merge(con, stops, left_on="ID_StopArea", right_on="NO", how="inner", sort=False)
con = con[con["TWALK"] < conTime]
con.drop(columns=["NO", "area"], inplace=True)
if SAQ: # add WKTXYWGS84
    POI = pd.DataFrame(Visum.Net.POICategories.ItemByKey(23).POIs.GetMultipleAttributes(["ID", "WKTLOCWGS84"], True), columns=["ID", "WKTXY"])
    POI["geometry"] = POI["WKTXY"].apply(wkt.loads)
    POI["X"] = POI["geometry"].apply(lambda p: p.x)
    POI["Y"] = POI["geometry"].apply(lambda p: p.y)
    con = con.merge(POI[["ID", "X", "Y"]], left_on="ID_ZENSUS", right_on="ID", how="left", sort=False)
    con.drop(columns=["ID"], inplace=True)
    con[["X", "Y"]] = con[["X", "Y"]].astype("float32")
con = con.sort_values("ID_ZENSUS")

# OD-impedence / Chunks
groups = con["ID_ZENSUS"].unique()
n_chunks = 100
group_chunks = np.array_split(groups, n_chunks)
chunks = [con[con["ID_ZENSUS"].isin(g)] for g in group_chunks]
results = []
for chunk in chunks:
    zuIso = pd.merge(chunk, dfIso, left_on="ID_StopArea", right_on="fromStop", how="inner", sort=False)
    zuIso["time"] = zuIso["time"] + zuIso["TWALK"]
    zuIso = zuIso[zuIso["time"] < maxTime]
    zuIso = zuIso[zuIso["time"] == zuIso.groupby(["ID_ZENSUS", "toStop"])["time"].transform("min")]
    zuIso.drop(columns=["ID_StopArea", "TWALK", "EW"], inplace=True)
    zuIso[["ID_ZENSUS", "fromStop", "toStop"]] = zuIso[["ID_ZENSUS", "fromStop", "toStop"]].astype("int32")

    zuabIso = pd.merge(zuIso, con, left_on="toStop", right_on="ID_StopArea", how="inner", sort=False)
    zuabIso["time"] = zuabIso["time"]+zuabIso["TWALK"]
    zuabIso.drop(columns=["TWALK", "ID_StopArea", "fromStop", "toStop"], inplace=True)
    zuabIso = zuabIso[zuabIso["time"] < maxTime]
    zuabIso = zuabIso[zuabIso["time"] == zuabIso.groupby(["ID_ZENSUS_x", "ID_ZENSUS_y"])["time"].transform("min")]
    results.append(zuabIso)
    del zuIso, zuabIso
zuabIso = pd.concat(results, ignore_index=True)

# Indicator
if SAQ:
    zuabIso["ll"] = haversine_km(
        zuabIso["X_x"].to_numpy(),
        zuabIso["Y_x"].to_numpy(),
        zuabIso["X_y"].to_numpy(),
        zuabIso["Y_y"].to_numpy()
    ).astype("float32")
    zuabIso.drop(columns=["X_x", "Y_x", "X_y", "Y_y"], inplace=True)
    zuabIso["llkmh"] = zuabIso["ll"] / (zuabIso["time"] / 60)
    
    zuabIso["SAQ"] = "F"
    zuabIso["A"] = 1/(0.19*pow(zuabIso["ll"],-0.5)+0.0031)
    zuabIso.loc[(zuabIso["llkmh"] > zuabIso["A"]) & (zuabIso["SAQ"] == "F"), "SAQ"] = "A"
    zuabIso["B"] = 1/(0.22*pow(zuabIso["ll"],-0.5)+0.0037)
    zuabIso.loc[(zuabIso["llkmh"] > zuabIso["B"]) & (zuabIso["SAQ"] == "F"), "SAQ"] = "B"
    zuabIso["C"] = 1/(0.26*pow(zuabIso["ll"],-0.5)+0.0044)
    zuabIso.loc[(zuabIso["llkmh"] > zuabIso["C"]) & (zuabIso["SAQ"] == "F"), "SAQ"] = "C"
    zuabIso["D"] = 1/(0.32*pow(zuabIso["ll"],-0.5)+0.0052)
    zuabIso.loc[(zuabIso["llkmh"] > zuabIso["D"]) & (zuabIso["SAQ"] == "F"), "SAQ"] = "D"
    zuabIso["E"] = 1/(0.40*pow(zuabIso["ll"],-0.5)+0.0063)
    zuabIso.loc[(zuabIso["llkmh"] > zuabIso["E"]) & (zuabIso["SAQ"] == "F"), "SAQ"] = "E"
    
    results = zuabIso.loc[zuabIso["SAQ"] == "A"].groupby("ID_ZENSUS_x", as_index=False)["EW"].sum()

else:
    results = zuabIso.groupby("ID_ZENSUS_x", as_index=False)["EW"].sum()

# Add to POI
POITable = Visum.Net.POICategories.ItemByKey(23)
POITable = pd.DataFrame(POITable.POIs.GetMultipleAttributes(["ID", "VQ"], True), columns=["ID", "VQ"])
POITable[["ID", "VQ"]] = POITable[["ID", "VQ"]].astype("int32")
POITable = pd.merge(POITable, results, left_on="ID", right_on="ID_ZENSUS_x", how="left").fillna(0)
POITableList = POITable["EW"].tolist()
SetMulti(Visum.Net.POICategories.ItemByKey(23).POIs, "VQ", POITableList, False)

for name in dir():
    if not name.startswith("_"):
        del globals()[name]
import gc; gc.collect()