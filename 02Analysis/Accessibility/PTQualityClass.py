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

def Checks(_Visum):
    if _Visum.Net.Stops.CountActive == 0:
        _Visum.Log(12288, "No active stops!")
        return False
    if False in [int(sublist[1]) > int(sublist[0]) for sublist in intervals]:
        _Visum.Log(12288, "End time must be after start time!")
        return False
    if not all(intervals[i][1] < intervals[i+1][0] for i in range(len(intervals) - 1)):
        _Visum.Log(12288, "Time intervals are overlapping!")
        return False
    return True

def ClipIntersectShape(_Visum, _ShapePath, _clip, _intersect, _ClipShape, _InterShape):
    import geopandas as gpd
    _ShapeGpd = gpd.read_file(_ShapePath)
    
    if _clip:
        _Visum.Log(20480, "Start clipping shapes")
        _clipShapeGpd = gpd.read_file(_ClipShape)
        _clipShapeGpd = _clipShapeGpd.to_crs(_ShapeGpd.crs)
        _shape_path = Path(_ShapePath)
        cliped_shape_path = _shape_path.with_name("polygons_clip.shp")
        clipped = gpd.clip(_ShapeGpd, _clipShapeGpd)
        _Visum.Log(20480, "Shapes clipped")
        if not _intersect:
            clipped.to_file(cliped_shape_path)
            return cliped_shape_path

    if _intersect:
        _Visum.Log(20480, "Start intersecting shapes")
        _interShapeGpd = gpd.read_file(_InterShape)
        if _clip: _ShapeGpd = clipped
        else:
            _ShapeGpd = gpd.read_file(_ShapePath)
            _interShapeGpd = _interShapeGpd.to_crs(_ShapeGpd.crs)
        _shape_path = Path(_ShapePath)
        intersected_shape_path = _shape_path.with_name("polygons_intersect.shp")
        intersection = gpd.overlay(_interShapeGpd, _ShapeGpd, how="intersection") 
        intersection['NO'] = range(1, len(intersection) + 1)
        intersection.to_file(intersected_shape_path)
        
        _Visum.Log(20480, "Shapes intersected")
        return intersected_shape_path
    
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
    
    categories = {
        "I": ["300m:A", "400m:A", "600m:B", "1000m:C", "1500m:D"],
        "II": ["300m:A", "400m:B", "600m:C", "1000m:D", "1500m:E"],
        "III": ["300m:B", "400m:C", "600m:D", "1000m:E", "1500m:F"],
        "IV": ["300m:C", "400m:D", "600m:E", "1000m:F", "1500m:G"],
        "V": ["300m:D", "400m:E", "600m:F", "1000m:G", "1500m:I"],
        "VI": ["300m:E", "400m:F", "600m:G", "1500m:I"],
        "VII": ["300m:F", "400m:G", "1500m:I"],
    }
    
    for category, distances in categories.items():
        stops_cat = _stops[_stops[HKAT] == category]
        stops_cat = list(zip(stops_cat["StopNo"].astype(int), stops_cat["StopName"], stops_cat["X"], stops_cat["Y"], stops_cat["DepHour"].astype(int)))
        
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
                feature.SetField("Category", f"StopNo: {StopNo} - Cat: {category} - DepHour: {Dep} - Scenario: {Scenario}")
                feature.SetField("Class", PTClass)
                feature.SetField("Distance", distance)
                feature.SetField("StopNo", StopNo)
                feature.SetField("StopName", f"{StopName} - {distance}m")
        
                # Add to layer
                _Shape.CreateFeature(feature)
                feature = None  # Free memory

    _data_source.FlushCache()
    _Visum.Log(20480, "Polygons created and saved")         
    
def ImportShapePOI(_Visum, _Shape):
    _Visum.Log(20480, "Importing shapes to Visum-POI")
    if deloldPOI:
        _Visum.Net.POICategories.ItemByKey(40).POIs.RemoveAll()
    
    ShapeImport = _Visum.IO.CreateImportShapeFilePara()
    ShapeImport.AddAttributeAllocation("StopName", "Name")
    ShapeImport.AddAttributeAllocation("Class", "Code")
    ShapeImport.AddAttributeAllocation("Category", "Comment")
    ShapeImport.ObjectType = 9 # import as POI
    ShapeImport.SetAttValue("POIKEY", 40) # POI Category 40
    
    _Visum.IO.ImportShapefile(_Shape, ShapeImport)

def StopCategories(_Visum):
    _Visum.Log(20480, "Calculate Stop categories for %s Stop(s)" % _Visum.Net.Stops.CountActive)
    mainlinesInput = {
        "StopType1": {'IC', 'ICE', 'ICE(S)', 'Nightjet', 'RE', 'RB', 'S-Bahn', 'AKN', 'U-Bahn', ''},
        "StopType2": {''},
        "StopType3": {'XpressBus', 'Metrobus', 'Nachtbus', 'Bus', 'Schiff', ''}
        }
    
    stop_type_values = {
        "StopType1": 1,
        "StopType2": 2,
        "StopType3": 3
        }
    
    mainlines = {}
    for stop_type, services in mainlinesInput.items():
        value = stop_type_values.get(stop_type, 10)  # default to 0 if not found
        mainlines[stop_type] = {
            service: 10 if service == '' else value
            for service in services
        }

    mainlines_all = {**mainlines["StopType1"], **mainlines["StopType2"], **mainlines["StopType3"]}
    
    VJI = pd.DataFrame(_Visum.Net.VehicleJourneyItems.GetMultipleAttributes(
        ["VEHJOURNEYNO", "INDEX", "Dep",r"TIMEPROFILEITEM\LINEROUTEITEM\STOPPOINT\STOPAREA\STOPNO" ,
         r"VEHJOURNEY\LINEROUTE\LINE\MAINLINENAME", r"COUNT:COUPLEDVEHJOURNEYITEMS"], True))
    VJI.columns = ["VJNO", "INDEX", "Dep", "StopNo", "Mainline", "Coupled"]
    VJI = VJI[VJI["Dep"].notna()]
    
    # Filter out rows where StopNo is equal to the previous row's StopNo (two Departures at same Stop)
    VJI = VJI[~(
    (VJI['StopNo'] == VJI['StopNo'].shift(1)) & 
    (VJI['VJNO'] == VJI['VJNO'].shift(1)))]
    
    # Coupled sections and line ends
    VJI["nDep"] = 1 / VJI["Coupled"] # Coupled sections reducing the weight of departures
    if LineEnd: VJI.loc[VJI["INDEX"] == 1, "nDep"] *= 2 # Count first index *2 for missing arrivals
    
    # Selecting VehJour in time intervals
    scaled_intervals = [[start * 60 * 60, end * 60 * 60] for start, end in intervals] + [[start * 60 * 60 + 86400, end * 60 * 60 + 86400] for start, end in intervals]
    VJI = VJI[VJI["Dep"].apply(lambda x: any(start <= x <= end for start, end in scaled_intervals))]
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
        
        # count StopDepartures in VHI for each stop
        if i[0] == "all": StopCounts = VJI.groupby("StopNo", as_index = False)["nDep"].sum()
        else: StopCounts = VJI[VJI["StopType"] == i[0]].groupby("StopNo", as_index = False)["nDep"].sum()
        StopCounts.columns = ["StopNo", "DepNo"]
        _Stops = _Stops.merge(StopCounts, on="StopNo", how="left")
        _Stops["DepNo"] = _Stops["DepNo"].fillna(0).astype(int)
        _Stops["DepHour"] = _Stops["DepNo"] / sum(end - start for start, end in intervals)
        _Stops["DepHour"] = _Stops["DepHour"].round(0) #round departures
        
        # Stop categories from StopType and departures  in PTV Visum
        conditions = [
            (_Stops["DepHour"] >= 24) & (_Stops[i[0]] == 1),
            (_Stops["DepHour"] >= 24) & (_Stops[i[0]] == 2),
            (_Stops["DepHour"] >= 24) & (_Stops[i[0]] == 3),
            (_Stops["DepHour"] >= 12) & (_Stops[i[0]] == 1),
            (_Stops["DepHour"] >= 12) & (_Stops[i[0]] == 2),
            (_Stops["DepHour"] >= 12) & (_Stops[i[0]] == 3),
            (_Stops["DepHour"] >= 6) & (_Stops[i[0]] == 1),
            (_Stops["DepHour"] >= 6) & (_Stops[i[0]] == 2),
            (_Stops["DepHour"] >= 6) & (_Stops[i[0]] == 3),
            (_Stops["DepHour"] >= 4) & (_Stops[i[0]] == 1),
            (_Stops["DepHour"] >= 4) & (_Stops[i[0]] == 2),
            (_Stops["DepHour"] >= 4) & (_Stops[i[0]] == 3),
            (_Stops["DepHour"] >= 2) & (_Stops[i[0]] == 1),
            (_Stops["DepHour"] >= 2) & (_Stops[i[0]] == 2),
            (_Stops["DepHour"] >= 2) & (_Stops[i[0]] == 3),
            (_Stops["DepHour"] >= 1) & (_Stops[i[0]] == 1),
            (_Stops["DepHour"] >= 1) & (_Stops[i[0]] == 2),
            (_Stops["DepHour"] >= 1) & (_Stops[i[0]] == 3)
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

if not Checks(Visum):
    raise Exception("")
Stops = StopCategories(Visum)
ShapePath, DataSource, ShapeDef = CreateShape(Visum)
CreatePolygons(Visum, ShapeDef, Stops, DataSource)
if clip or intersect:
    ShapePath = ClipIntersectShape(Visum, ShapePath, clip, intersect, clipShape, intersectShape)
ImportShapePOI(Visum, ShapePath)

Visum.Log(20480, "Open GraphicParameters to display accessibility measures")
gpa_path = Path.cwd() / "Grafikparameter" / "OEV_Gueteklassen.gpa"
Visum.Net.GraphicParameters.Open(gpa_path)

Visum.Log(20480, "Finished!")
