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

# > Finalisierung Skript (auch in Doku)
# > adding StopNo and StopName to new UDA or No only or with DepNumber

def StopCategories(_Visum):
    _Visum.Log(20480, "Calculate Stop categories for %s Stop(s)" % _Visum.Net.Stops.CountActive)
    fromTime, toTime = (8 * 60 * 60), (18 * 60 * 60)
    
    VJI = pd.DataFrame(_Visum.Net.VehicleJourneyItems.GetMultipleAttributes(
        ["Dep",r"TIMEPROFILEITEM\LINEROUTEITEM\STOPPOINT\STOPAREA\STOPNO"], True))
    VJI.columns = ["Dep", "StopNo"]
    VJI = VJI[VJI["Dep"].notna()]
    VJI = VJI[(VJI["Dep"] >= fromTime) & (VJI["Dep"] <= toTime)]
    VJI = VJI.reset_index(drop=True)
    
    # open corresponding Stops
    _Stops = pd.DataFrame(_Visum.Net.Stops.GetMultipleAttributes(
        ["NO", "Name", r"MIN:STOPAREAS\MIN:STOPPOINTS\ANGEBOT", "XCOORD", "YCOORD"], True))
    _Stops.columns = ["StopNo", "StopName", "PTLevel", "X", "Y"]
    
    # count StopDepartures in VHI for each strop
    StopCounts = VJI["StopNo"].value_counts().reset_index()
    StopCounts.columns = ["StopNo", "Dep0818"]
    _Stops = _Stops.merge(StopCounts, on="StopNo", how="left")
    _Stops["Dep0818"] = _Stops["Dep0818"].fillna(0).astype(int)
    _Stops["Hour"] = _Stops["Dep0818"] / 10
    
    # Stop categories in PTV Visum
    conditions = [
        (_Stops["Hour"] > 24) & (_Stops["PTLevel"] < 5),
        (_Stops["Hour"] > 24),
        (_Stops["Hour"] > 12) & (_Stops["PTLevel"] < 4),
        (_Stops["Hour"] > 12) & (_Stops["PTLevel"] == 4),
        (_Stops["Hour"] > 12),
        (_Stops["Hour"] > 6) & (_Stops["PTLevel"] < 4),
        (_Stops["Hour"] > 6) & (_Stops["PTLevel"] == 4),
        (_Stops["Hour"] > 6),
        (_Stops["Hour"] > 4) & (_Stops["PTLevel"] < 4),
        (_Stops["Hour"] > 4) & (_Stops["PTLevel"] == 4),
        (_Stops["Hour"] > 4),
        (_Stops["Hour"] > 2) & (_Stops["PTLevel"] < 4),
        (_Stops["Hour"] > 2) & (_Stops["PTLevel"] == 4),
        (_Stops["Hour"] > 2),
        (_Stops["Hour"] > 1) & (_Stops["PTLevel"] < 4),
        (_Stops["Hour"] > 1) & (_Stops["PTLevel"] == 4),
        (_Stops["Hour"] > 1)
        ]

    choices = [
        "I", "II", 
        "I", "II", "III",
        "II", "III", "IV",
        "III", "IV", "V",
        "IV", "V", "VI",
        "V", "VI", "VII",
        ]
    
    _Stops["HKAT"] = np.select(conditions, choices, default="VII")
    
    # to Visum UDA 'HKAT'
    PTClass = _Stops["HKAT"].tolist()
    SetMulti(_Visum.Net.Stops, "HKAT", PTClass, True)
    
    _Visum.Log(20480, "Stop categories calculated")
    return _Stops

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
        "I": ["300m:A", "500m:A", "750m:B", "1000m:C"],
        "II": ["300m:A", "500m:B", "750m:C", "1000m:D"],
        "III": ["300m:B", "500m:C", "750m:D", "1000m:E"],
        "IV": ["300m:C", "500m:D", "750m:E", "1000m:F"],
        "V": ["300m:D", "500m:E", "750m:F"],
        "VI": ["300m:E", "500m:F"],
        "VII": ["300m:F"],
    }
    
    for category, distances in cases.items():
        stops_cat = _stops[_stops["HKAT"] == category]
        stops_cat = list(zip(stops_cat["StopNo"].astype(int), stops_cat["StopName"], stops_cat["X"], stops_cat["Y"], stops_cat["Dep0818"]))
        
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
    
    ShapeImport = _Visum.IO.CreateImportShapeFilePara()
    ShapeImport.AddAttributeAllocation("StopName", "Name")
    ShapeImport.AddAttributeAllocation("Class", "Code")
    ShapeImport.AddAttributeAllocation("Category", "Comment")
    ShapeImport.ObjectType = 9 # import as POI
    ShapeImport.SetAttValue("POIKEY", 40) # POI Category 40
    
    _Visum.IO.ImportShapefile(_Shape, ShapeImport)
    
    
## Calculations ##
Visum.Log(20480, "Starting...")

Stops = StopCategories(Visum)
ShapePath, DataSource, ShapeDef = CreateShape(Visum)
ShapePolygons = CreatePolygons(Visum, ShapeDef, Stops, DataSource)
ImportShapePOI(Visum, ShapePath)

Visum.Log(20480, "Open GraphicParameters to display accessibility measures")
gpa_path = Path.cwd() / "Grafikparameter" / "OEV_Gueteklassen.gpa"
Visum.Net.GraphicParameters.Open(gpa_path)

Visum.Log(20480, "Finished!")
