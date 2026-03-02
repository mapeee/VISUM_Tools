# -*- coding: utf-8 -*-
import numpy as np
import os
from osgeo import ogr, osr
import pandas as pd
from pathlib import Path
import tempfile
import sys
import wx
from VisumPy.helpers import SetMulti
from VisumPy.AddIn import AddIn, AddInState, AddInParameter
_ = AddIn.gettext


def Run(param):
    createUDA(Visum)
    Stops = StopCategories(Visum)
    GeoJSONPath, DataSource, GeoJSONDef = CreateGeoJSON(Visum)
    CreatePolygons(Visum, GeoJSONDef, Stops, DataSource)
    if param["clip"]:
        GeoJSONPath = ClipGeoJSON(Visum, GeoJSONPath)
    ImportGeoJSONPOI(Visum, GeoJSONPath)
    Visum.Graphic.Redraw()
    
    Visum.Log(20480,_("PT quality levels calculated"))
    
def ClipGeoJSON(Visum, _GeoJSONPath):
    '''
    Schneide bestehende GeoJSON-Datei auf Basis gelieferter Geometrien aus.
    Speicher die ausgeschnittenen Polygone unter neuem Pfad und liefere diesen zurück.
    '''
    clipGeoJSON = param["clipfiles"]
    import geopandas as gpd
    _GeoJSONGpd = gpd.read_file(_GeoJSONPath)
    for i in clipGeoJSON:
        _clipGeoJSONGpd = gpd.read_file(i)
        _clipGeoJSONGpd = _clipGeoJSONGpd[['geometry']]
        _clipGeoJSONGpd = _clipGeoJSONGpd.to_crs(_GeoJSONGpd.crs)
        _GeoJSONGpd = gpd.clip(_GeoJSONGpd, _clipGeoJSONGpd)
    
    _GeoJSON_path = Path(_GeoJSONPath)
    cliped_GeoJSON_path = _GeoJSON_path.with_name("polygons_clip.geojson")
    _GeoJSONGpd.to_file(cliped_GeoJSON_path, driver="GeoJSON")
    Visum.Log(20480, _("Polygons cut out"))
    return cliped_GeoJSON_path

def CreateGeoJSON(Visum):  
    '''
    Erstelle die in PTV Visum zu importierende GeoJSON als temporäres File (geht nach dem Herunterfahren verloren)
    '''
    temp_dir = tempfile.mkdtemp()
    geojson_path  = os.path.join(temp_dir, "buffered_polygons.geojson")
    Visum.Log(20480, _("Temporary GeoJSON path: %s") %(geojson_path))
    
    driver = ogr.GetDriverByName("GeoJSON")
    spatial_ref = osr.SpatialReference()
    spatial_ref.ImportFromEPSG(25832) # EPSG:25832 (ETRS89 / UTM zone 32N)
    if os.path.exists(geojson_path):
        driver.DeleteDataSource(geojson_path)
    _data_source = driver.CreateDataSource(geojson_path)
    _GeoJSON = _data_source.CreateLayer("buffered_polygons", srs=spatial_ref, geom_type=ogr.wkbPolygon, options=["RFC7946=YES"])
    
    # Add some attribute fields
    _GeoJSON.CreateField(ogr.FieldDefn("StopNo", ogr.OFTInteger))
    _GeoJSON.CreateField(ogr.FieldDefn("StopName", ogr.OFTString))
    _GeoJSON.CreateField(ogr.FieldDefn("Category", ogr.OFTString))
    _GeoJSON.CreateField(ogr.FieldDefn("Scenario", ogr.OFTString))
    _GeoJSON.CreateField(ogr.FieldDefn("Class", ogr.OFTString))
    _GeoJSON.CreateField(ogr.FieldDefn("Distance", ogr.OFTInteger))
    
    _GeoJSON.SyncToDisk()
    _data_source.FlushCache()
    return geojson_path, _data_source, _GeoJSON

def CreatePolygons(Visum ,_GeoJSON, _stops, _data_source):
    HKAT = ["HKAT", "HKAT_FHH"][param["bt"]]
    scenario = param["sc"]
    # Zuordung HstKategorie zur Entfernungsklasse (Luflinie)
    categories = {
        "I": ["300m:A", "400m:A", "600m:B", "1000m:C", "1500m:D"],
        "II": ["300m:A", "400m:B", "600m:C", "1000m:D", "1500m:E"],
        "III": ["300m:B", "400m:C", "600m:D", "1000m:E", "1500m:F"],
        "IV": ["300m:C", "400m:D", "600m:E", "1000m:F", "1500m:G"],
        "V": ["300m:D", "400m:E", "600m:F", "1000m:G", "1500m:I"],
        "VI": ["300m:E", "400m:F", "600m:G", "1500m:I"],
        "VII": ["300m:F", "400m:G", "1500m:I"],
    }
    
    polygon_count = 0
    for category, distances in categories.items():
        stops_cat = _stops[_stops[HKAT] == category]
        stops_cat = list(zip(stops_cat["STOPNO"].astype(int), stops_cat["STOPNAME"], stops_cat["X"], stops_cat["Y"], stops_cat["DepHour"].astype(int)))
        for StopNo, StopName, x, y, Dep in stops_cat:
            point = ogr.Geometry(ogr.wkbPoint)
            point.AddPoint(x, y)
            point.AssignSpatialReference(_GeoJSON.GetSpatialRef())
        
            for distance_class in distances:
                distance = int(distance_class.split("m:")[0]) # nutze split "m:", um m nicht abtrennen zu müssen
                PTClass = distance_class.split("m:")[1]
                # buffer
                feature_def = _GeoJSON.GetLayerDefn()
                feature = ogr.Feature(feature_def)
                buffered_polygon = point.Buffer(distance)
                feature.SetGeometry(buffered_polygon)
                feature.SetField("Category", f"StopNo: {StopNo} - HstKat: {category} - AbStunde: {Dep}")
                feature.SetField("Scenario", scenario)
                feature.SetField("Class", PTClass)
                feature.SetField("Distance", distance)
                feature.SetField("StopNo", StopNo)
                feature.SetField("StopName", f"{StopName} - {distance}m")
                # Add to file
                _GeoJSON.CreateFeature(feature)
                feature = None  # Free memory
                polygon_count+=1
                
    _data_source.FlushCache()
    Visum.Log(20480, _("%s Polygons created") %(str(polygon_count)))
    
def ImportGeoJSONPOI(Visum, _GeoJSON):
    '''
    Importiere Polygone nach PTV Visum
    '''
    poi = param["poi"]
    poiNO = [i.AttValue("NO") for i in Visum.Net.POICategories.GetAll][poi] # get Number of POI category
    deloldPOI = param["poidel"]
    
    if deloldPOI:
        Visum.Net.POICategories.ItemByKey(poiNO).POIs.RemoveAll()
    
    GeoJSONImport = Visum.IO.CreateImportGeoJSONPara()
    GeoJSONImport.AddAttributeAllocation("StopName", "Name")
    GeoJSONImport.AddAttributeAllocation("Class", "Code")
    GeoJSONImport.AddAttributeAllocation("Category", "Comment")
    GeoJSONImport.AddAttributeAllocation("Scenario", "Szenario")
    GeoJSONImport.ObjectType = 9 # import as POI
    GeoJSONImport.SetAttValue("POIKEY", poiNO)
    Visum.IO.ImportGeoJSON(_GeoJSON, GeoJSONImport)

def StopCategories(Visum):
    '''

    '''
    Visum.Log(20480, _("Calculate stop categories for %s stops") %(str(Visum.Net.Stops.CountActive)))
    intervals = param["ti"]
    lineend = param["le"]
    weekday = param["wd"]
    dict_scml = param["scml"]

    # chose only VJ on valid days (valid day or daily (all))
    if Visum.Net.CalendarPeriod.AttValue("TYPE") == "CALENDARPERIODWEEK":
        day = ["ISVALID(MO)", "ISVALID(TU)", "ISVALID(WE)", "ISVALID(TH)", "ISVALID(FR)", "ISVALID(SA)", "ISVALID(SU)"][weekday]
    else:
        day = "ISVALID(1)"
    
    VJI = pd.DataFrame(Visum.Net.VehicleJourneyItems.GetMultipleAttributes(
        ["VEHJOURNEYNO", "INDEX", "Dep",r"TIMEPROFILEITEM\LINEROUTEITEM\STOPPOINT\STOPAREA\STOPNO",
         r"VEHJOURNEY\LINEROUTE\LINE\MAINLINENAME", r"COUNT:COUPLEDVEHJOURNEYITEMS", day], True))
    VJI.columns = ["VJNO", "INDEX", "DEP", "STOPNO", "MAINLINE", "CHAINED", "DAY"]
    VJI = VJI[VJI["DEP"].notna()]
    VJI = VJI[VJI["DAY"] == 1] # only valid days
    
    # Filter out rows where STOPNO is equal to the previous row's STOPNO (two Departures at same Stop)
    VJI = VJI[~(
    (VJI['STOPNO'] == VJI['STOPNO'].shift(1)) & 
    (VJI['VJNO'] == VJI['VJNO'].shift(1)))]
    
    # Coupled sections and line ends
    VJI["nDEP"] = 1 / VJI["CHAINED"] # Coupled sections reducing the weight of departures
    if lineend: VJI.loc[VJI["INDEX"] == 1, "nDEP"] *= 2 # Count first index *2 for missing arrivals (not = 2 for chained VJ)
    
    # Selecting VehJour in time intervals; double the intervals for the next morning (also for weekcalendar because also trips after 24:00)
    scaled_intervals = [[start, end] for start, end in intervals] + [[start + 86400, end + 86400] for start, end in intervals] # 86400 seconds a day
    VJI = VJI[VJI["DEP"].apply(lambda x: any(start <= x <= end for start, end in scaled_intervals))]
    VJI = VJI.reset_index(drop=True)

    # Adding StopTypes to VJI
    VJI = VJI.merge(dict_scml[["MAINLINE", "STOPTYPE"]], on="MAINLINE", how="left")

    # open corresponding Stops
    StopsDF = pd.DataFrame(Visum.Net.Stops.GetMultipleAttributes(
        ["NO", "NAME", r"DISTINCTACTIVE:STOPAREAS\DISTINCTACTIVE:STOPPOINTS\DISTINCTACTIVE:SERVINGVEHJOURNEYS\LINEROUTE\LINE\MAINLINENAME",
         "XCOORD", "YCOORD"], True))
    StopsDF.columns = ["STOPNO", "STOPNAME", "MAINLINES", "X", "Y"]
    
    # all for later difference between HKAT in different stoptype-level
    for i in [["StopType1", 1, "HKAT1"], ["StopType2", 2, "HKAT2"], ["StopType3", 3, "HKAT3"], ["all", None, "HKAT"]]:
        _Stops = StopsDF.copy()
        # Get StopType
        if i[0] == "all":
            df_mainline = dict_scml
        else:
            df_mainline = dict_scml[dict_scml["STOPTYPE"] == i[1]]
        df_mainline = df_mainline.set_index("MAINLINE")["STOPTYPE"].to_dict()
        _Stops[i[0]] = _Stops['MAINLINES'].apply(lambda x: min(df_mainline.get(e, 10) for e in x.split(',')))

        # count StopDepartures in VHI for each stop
        if i[0] == "all": StopCounts = VJI.groupby("STOPNO", as_index = False)["nDEP"].sum()
        else: StopCounts = VJI[VJI["STOPTYPE"] == i[1]].groupby("STOPNO", as_index = False)["nDEP"].sum()
        StopCounts.columns = ["STOPNO", "DepNo"]
        _Stops = _Stops.merge(StopCounts, on="STOPNO", how="left")
        _Stops["DepNo"] = _Stops["DepNo"].fillna(0).astype(int)
        _Stops["DepHour"] = _Stops["DepNo"] / sum(end/60/60 - start/60/60 for start, end in intervals) # only use single interval and not scaled ones (each time interval twice)
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
        
        _Stops[i[2]] = np.select(conditions, choices, default = "X")
        
        # to Visum UDA 'HKAT'
        PTClass = _Stops[i[2]].tolist()
        SetMulti(Visum.Net.Stops, i[2], PTClass, True)
    
    _Stops_FHH = _stopcat_fhh(Visum)
    _Stops = _Stops.merge(_Stops_FHH[['STOPNO', 'HKAT_FHH']], on='STOPNO', how='left')

    return _Stops

def createUDA(Visum):
    '''
    Erstelle die noch nicht vorhanden und benötigten BDA.
    '''
    poi = param["poi"]
    poiNO = [i.AttValue("NO") for i in Visum.Net.POICategories.GetAll][poi]
    n = 0
    for e, i in enumerate(["HKAT", "HKAT_FHH", "HKAT1", "HKAT2", "HKAT3"]):
        if Visum.Net.Stops.AttrExists(i):
            continue
        if e == 0:
            Visum.Net.Stops.AddUserDefinedAttribute(i, "Haltestellenkategorie", "Haltestellenkategorie", 5)
            uda = Visum.Net.Stops.Attributes.ItemByKey(i)
            uda.Comment = _("Stop category (PT quality levels)")
        elif e == 1:
            Visum.Net.Stops.AddUserDefinedAttribute(i, "Haltestellenkategorie in FHH", "Haltestellenkategorie in FHH", 5)
            uda = Visum.Net.Stops.Attributes.ItemByKey(i)
            uda.Comment = _("Stop category in FHH (PT quality levels)")
        else:
            Visum.Net.Stops.AddUserDefinedAttribute(i, f"Haltestellenkategorie HstTyp {i[-1]}", f"Haltestellenkategorie HstTyp {i[-1]}", 5)
            uda = Visum.Net.Stops.Attributes.ItemByKey(i)
            uda.Comment = _("Stop category for Stop type %s (PT quality levels)" %(i[-1]))
        uda.MaxStringLen = 4
        uda.StringValueDefault = "X"
        n+=1
    if not Visum.Net.POICategories.ItemByKey(poiNO).POIs.AttrExists("Szenario"):
        Visum.Net.POICategories.ItemByKey(poiNO).POIs.AddUserDefinedAttribute("Szenario", "Szenario", "Szenario", 5)
        uda = Visum.Net.Stops.Attributes.ItemByKey(i)
        uda.Comment = _("PT Qualities: Scenario")
        uda.MaxStringLen = 30
        uda.StringValueDefault = "X"
        n+=1
    if n > 0:
        Visum.Log(20480,_("%s UDA added") %(str(n)))
        
def _stopcat_fhh(Visum):
    _StopsDF = pd.DataFrame(Visum.Net.Stops.GetMultipleAttributes(
        ["NO", "HKAT", "HKAT1", "HKAT2", "HKAT3"], True))
    _StopsDF.columns = ["STOPNO", "HKAT", "HKAT1", "HKAT2", "HKAT3"]
    
    # Roman numeral conversion dictionaries
    roman_to_int = {'I':1, 'II':2, 'III':3, 'IV':4, 'V':5, 'VI':6, 'VII':7, 'VIII':8, 'IX':9, 'X':10}
    int_to_roman = {v: k for k, v in roman_to_int.items()}
    # Convert to integers
    _StopsDF_int = _StopsDF.replace(roman_to_int)
    # Calculate new column
    # in FHH: HKAT is best from HKAT1 to HKAT3 minus 3 and HKAT over all
    _StopsDF['HKAT_FHH'] = (_StopsDF_int[['HKAT1', 'HKAT2', 'HKAT3']].min(axis=1) - 1).clip(lower=1)
    _StopsDF['HKAT_FHH'] = _StopsDF_int[['HKAT']].join(_StopsDF['HKAT_FHH']).max(axis=1)
    # # Convert back to Roman
    _StopsDF['HKAT_FHH'] = _StopsDF['HKAT_FHH'].map(int_to_roman)

    _stopcat_fhh = _StopsDF['HKAT_FHH'].tolist()
    SetMulti(Visum.Net.Stops, 'HKAT_FHH', _stopcat_fhh, True)
    
    return _StopsDF


if len(sys.argv) > 1:
    addIn = AddIn()
else:
    addIn = AddIn(Visum)

if addIn.IsInDebugMode:
    app = wx.PySimpleApp(0)
    Visum = addIn.VISUM
    addInParam = AddInParameter(addIn, None)
else:
    addInParam = AddInParameter(addIn, Parameter)

if addIn.State != AddInState.OK:
    addIn.ReportMessage(addIn.ErrorObjects[0].ErrorMessage)
else:
    try:
        defaultParam = {"" : False}
        param = addInParam.Check(True, defaultParam)
        Run(param)
    except:
        addIn.HandleException(addIn.TemplateText.MainApplicationError)
