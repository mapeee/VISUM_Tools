#-------------------------------------------------------------------------------
# Name:        Calculate PT Quality Classes in PTV Visum on Stop-level
# Author:      mape
# Created:     25/03/2025
# Copyright:   (c) mape 2025
# Licence:     GNU GENERAL PUBLIC LICENSE
# #---------------------

import numpy as np
import os
from osgeo import ogr, osr
import pandas as pd
from pathlib import Path
import tempfile
from VisumPy.helpers import SetMulti

pd.set_option('future.no_silent_downcasting', True)

def StopCategories(_Visum):
    _Visum.Log(20480, "Calculate Stop categories for %s Stop(s)" % _Visum.Net.Stops.CountActive)
    fromTime1, toTime1 = (int(fromHour1) * 60 * 60), (int(toHour1) * 60 * 60)
    fromTime2, toTime2 = (int(fromHour2) * 60 * 60), (int(toHour2) * 60 * 60)
    mainlines = {
        "StopType1": {'IC': 1, 'ICE': 1, 'ICE(S)': 1, 'Nightjet': 1,
               'RE': 2, 'RB': 2,
               'S-Bahn': 3, 'AKN': 3,
               'U-Bahn': 3,
               '': 10},
        "StopType2": {'': 10},
        "StopType3": {'XpressBus': 5,
               'Metrobus': 6, 'Nachtbus': 6, 'Bus': 6, 'Schiff': 6,
               '': 10}
        }
    mainlines_all = {**mainlines["StopType1"], **mainlines["StopType2"], **mainlines["StopType3"]}
    
    VJI = pd.DataFrame(_Visum.Net.VehicleJourneyItems.GetMultipleAttributes(
        ["VEHJOURNEYNO", "Dep",r"TIMEPROFILEITEM\LINEROUTEITEM\STOPPOINT\STOPAREA\STOPNO" , r"VEHJOURNEY\LINEROUTE\LINE\MAINLINENAME"], True))
    VJI.columns = ["VJNO", "Dep", "StopNo", "Mainline"]
    VJI = VJI[VJI["Dep"].notna()]
    # Filter out rows where StopNo is equal to the previous row's StopNo (to Departures at same Stop)
    VJI = VJI[~(
    (VJI['StopNo'] == VJI['StopNo'].shift(1)) & 
    (VJI['VJNO'] == VJI['VJNO'].shift(1)))]
    
    if fromTime2 == toTime2:
        VJI = VJI[(VJI["Dep"] >= fromTime1) & (VJI["Dep"] <= toTime1)]
    else:
        VJI = VJI[
        ((VJI["Dep"] >= fromTime1) & (VJI["Dep"] <= toTime1)) |
        ((VJI["Dep"] >= fromTime2) & (VJI["Dep"] <= toTime2))
        ]
    VJI = VJI.reset_index(drop=True)
    VJI["StopType"] = VJI["Mainline"].apply(lambda x: _get_stop_type(x, mainlines))
    
    # open corresponding Stops
    _StopsDF = pd.DataFrame(_Visum.Net.Stops.GetMultipleAttributes(
        ["NO", "Name", r"DISTINCTACTIVE:STOPAREAS\DISTINCTACTIVE:STOPPOINTS\DISTINCTACTIVE:SERVINGVEHJOURNEYS\LINEROUTE\LINE\MAINLINENAME",
         "XCOORD", "YCOORD"], True))
    _StopsDF.columns = ["StopNo", "StopName", "MAINLINES", "X", "Y"]
    
    for i in [["StopType1", "HKATT1"], ["StopType2", "HKATT2"], ["StopType3", "HKATT3"], ["all", "HKAT"]]:
        _Stops = _StopsDF 
    
        # Get StopType
        if i[0] == "all": _Stops[i[0]] = _Stops['MAINLINES'].apply(lambda x: min(mainlines_all.get(e, 10) for e in x.split(',')))
        else: _Stops[i[0]] = _Stops['MAINLINES'].apply(lambda x: min(mainlines[i[0]].get(e, 10) for e in x.split(',')))
        
        # count StopDepartures in VHI for each strop
        if i[0] == "all": StopCounts = VJI["StopNo"].value_counts().reset_index()
        else: StopCounts = VJI[VJI["StopType"] == i[0]]["StopNo"].value_counts().reset_index()
        StopCounts.columns = ["StopNo", "DepNo"]
        _Stops = _Stops.merge(StopCounts, on="StopNo", how="left")
        _Stops["DepNo"] = _Stops["DepNo"].fillna(0).astype(int)
        _Stops["Hour"] = _Stops["DepNo"] / (int(toHour1) - int(fromHour1) + int(toHour2) - int(fromHour2))
        _Stops["Hour"] = _Stops["Hour"].round(0) #round departures
        
        # Stop categories from StopType and departures  in PTV Visum
        conditions = [
            (_Stops["Hour"] >= 24) & (_Stops[i[0]] < 4),
            (_Stops["Hour"] >= 24) & (_Stops[i[0]] == 4),
            (_Stops["Hour"] >= 24) & (_Stops[i[0]] < 10),
            (_Stops["Hour"] >= 12) & (_Stops[i[0]] < 4),
            (_Stops["Hour"] >= 12) & (_Stops[i[0]] == 4),
            (_Stops["Hour"] >= 12) & (_Stops[i[0]] < 10),
            (_Stops["Hour"] >= 6) & (_Stops[i[0]] < 4),
            (_Stops["Hour"] >= 6) & (_Stops[i[0]] == 4),
            (_Stops["Hour"] >= 6) & (_Stops[i[0]] < 10),
            (_Stops["Hour"] >= 4) & (_Stops[i[0]] < 4),
            (_Stops["Hour"] >= 4) & (_Stops[i[0]] == 4),
            (_Stops["Hour"] >= 4) & (_Stops[i[0]] < 10),
            (_Stops["Hour"] >= 2) & (_Stops[i[0]] < 4),
            (_Stops["Hour"] >= 2) & (_Stops[i[0]] == 4),
            (_Stops["Hour"] >= 2) & (_Stops[i[0]] < 10),
            (_Stops["Hour"] >= 1) & (_Stops[i[0]] < 4),
            (_Stops["Hour"] >= 1) & (_Stops[i[0]] == 4),
            (_Stops["Hour"] >= 1) & (_Stops[i[0]] < 10)
            ]
    
        choices = [
            "I", "I", "II", 
            "I", "II", "III",
            "II", "III", "IV",
            "III", "IV", "V",
            "IV", "V", "VI",
            "V", "VI", "VII",
            ]
        
        _Stops[i[1]] = np.select(conditions, choices, default = "X")
        
        # to Visum UDA 'HKAT'
        PTClass = _Stops[i[1]].tolist()
        SetMulti(_Visum.Net.Stops, i[1], PTClass, True)
    
    _Stops_FHH = _stopcat_fhh(_Visum)
    _Stops = _Stops.merge(_Stops_FHH[['StopNo', 'HKAT_FHH']], on='StopNo', how='left')
    
    _Visum.Log(20480, "Stop categories calculated")
    return _Stops

def checks(_Visum):
    if _Visum.Net.Stops.CountActive == 0:
        _Visum.Log(12288, "No active stops!")
        return False
    return True

def CreateShape(_Visum):
    spatial_ref = osr.SpatialReference()
    spatial_ref.ImportFromEPSG(25832)
    
    temp_dir = tempfile.mkdtemp()
    shapefile_path = os.path.join(temp_dir, "buffered_polygons.shp")
    _Visum.Log(20480, f"Temporary shapefile path: {shapefile_path}")
    
    driver = ogr.GetDriverByName("ESRI Shapefile")
    if os.path.exists(shapefile_path):
        driver.DeleteDataSource(shapefile_path)
    _data_source = driver.CreateDataSource(shapefile_path)
    _Shape = _data_source.CreateLayer("buffered_polygons", spatial_ref, ogr.wkbPolygon)
    
    # Add some attribute fields
    _Shape.CreateField(ogr.FieldDefn("StopNo", ogr.OFTInteger))
    _Shape.CreateField(ogr.FieldDefn("StopName", ogr.OFTString))
    _Shape.CreateField(ogr.FieldDefn("Category", ogr.OFTString))
    _Shape.CreateField(ogr.FieldDefn("Class", ogr.OFTString))
    _Shape.CreateField(ogr.FieldDefn("Distance", ogr.OFTInteger))
    
    _Visum.Log(20480, "Shape created (temporary)")
    return shapefile_path, _data_source, _Shape

def CreatePolygons(_Visum ,_Shape, _stops, _data_source):
    _Visum.Log(20480, "Start creating polygons")
    
    cases = {
        "I": ["300m:A", "400m:A", "600m:B", "1000m:C", "1500m:D"],
        "II": ["300m:A", "400m:B", "600m:C", "1000m:D", "1500m:E"],
        "III": ["300m:B", "400m:C", "600m:D", "1000m:E", "1500m:F"],
        "IV": ["300m:C", "400m:D", "600m:E", "1000m:F", "1500m:G"],
        "V": ["300m:D", "400m:E", "600m:F", "1000m:G"],
        "VI": ["300m:E", "400m:F", "600m:G"],
        "VII": ["300m:F", "400m:G"],
    }
    
    for category, distances in cases.items():
        stops_cat = _stops[_stops["HKAT_FHH"] == category]
        stops_cat = list(zip(stops_cat["StopNo"].astype(int), stops_cat["StopName"], stops_cat["X"], stops_cat["Y"], stops_cat["DepNo"]))
        
        for StopNo, StopName, x, y, Dep in stops_cat:
            point = ogr.Geometry(ogr.wkbPoint)
            point.AddPoint(x, y)
            point.AssignSpatialReference(_Shape.GetSpatialRef())
        
            for distance_class in distances:
                distance = int(distance_class.split("m:")[0])
                PTClass = distance_class.split("m:")[1]
            
                # buffer
                buffered_polygon = point.Buffer(distance)
                
                feature_def = _Shape.GetLayerDefn()
                feature = ogr.Feature(feature_def)
                feature.SetGeometry(buffered_polygon)
                feature.SetField("Category", f"StopNo: {StopNo} - Cat: {category} - Dep: {Dep} - Scenario: {Scenario}")
                feature.SetField("Class", PTClass)
                feature.SetField("Distance", distance)
                feature.SetField("StopNo", StopNo)
                feature.SetField("StopName", f"{StopName} - {distance}m")
        
                # Add to layer
                _Shape.CreateFeature(feature)
                feature = None  # Free memory

    _data_source.FlushCache()
    _Visum.Log(20480, "Polygons created and saved")         
    return _Shape
    
def ImportShapePOI(_Visum, _Shape):
    _Visum.Log(20480, "Importing shape to Visum-POI")
    if deloldPOI:
        _Visum.Net.POICategories.ItemByKey(40).POIs.RemoveAll()
    
    ShapeImport = _Visum.IO.CreateImportShapeFilePara()
    ShapeImport.AddAttributeAllocation("StopName", "Name")
    ShapeImport.AddAttributeAllocation("Class", "Code")
    ShapeImport.AddAttributeAllocation("Category", "Comment")
    ShapeImport.ObjectType = 9 # import as POI
    ShapeImport.SetAttValue("POIKEY", 40) # POI Category 40
    
    _Visum.IO.ImportShapefile(_Shape, ShapeImport)
    
def _get_stop_type(mainline, mainlines):
    for stop_type in ["StopType1", "StopType2", "StopType3"]:
        value = mainlines[stop_type].get(mainline)
        if value is not None:
            return stop_type
    return "not found"  # default 

def _stopcat_fhh(_Visum):
    _StopsDF = pd.DataFrame(_Visum.Net.Stops.GetMultipleAttributes(
        ["NO", "HKAT", "HKATT1", "HKATT2", "HKATT3"], True))
    _StopsDF.columns = ["StopNo", "HKAT", "HKATT1", "HKATT2", "HKATT3"]
    
    # Roman numeral conversion dictionaries
    roman_to_int = {'I':1, 'II':2, 'III':3, 'IV':4, 'V':5, 'VI':6, 'VII':7, 'VIII':8, 'IX':9, 'X':10}
    int_to_roman = {v: k for k, v in roman_to_int.items()}
    # Convert to integers
    _StopsDF_int = _StopsDF.replace(roman_to_int)
    # Calculate new column
    _StopsDF['HKAT_FHH'] = (_StopsDF_int[['HKATT1', 'HKATT2', 'HKATT3']].min(axis=1) - 1).clip(lower=1)
    _StopsDF['HKAT_FHH'] = _StopsDF_int[['HKAT']].join(_StopsDF['HKAT_FHH']).max(axis=1)
    # # Convert back to Roman
    _StopsDF['HKAT_FHH'] = _StopsDF['HKAT_FHH'].map(int_to_roman)

    _stopcat_fhh = _StopsDF['HKAT_FHH'].tolist()
    SetMulti(_Visum.Net.Stops, 'HKAT_FHH', _stopcat_fhh, True)
    
    return _StopsDF
    
## Calculations ##
Visum.Log(20480, "Starting...")

if not checks(Visum):
    raise Exception("")
Stops = StopCategories(Visum)
ShapePath, DataSource, ShapeDef = CreateShape(Visum)
ShapePolygons = CreatePolygons(Visum, ShapeDef, Stops, DataSource)
ImportShapePOI(Visum, ShapePath)

Visum.Log(20480, "Open GraphicParameters to display accessibility measures")
gpa_path = Path.cwd() / "Grafikparameter" / "OEV_Gueteklassen.gpa"
Visum.Net.GraphicParameters.Open(gpa_path)

Visum.Log(20480, "Finished!")
