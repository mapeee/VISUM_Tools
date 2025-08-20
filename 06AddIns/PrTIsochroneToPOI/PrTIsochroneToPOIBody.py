# -*- coding: utf-8 -*-
from datetime import datetime
import json
import openrouteservice as ors
import os
from shapely import wkt
import sys
import tempfile
import wx
from VisumPy.AddIn import AddIn, AddInState, AddInParameter
_ = AddIn.gettext


def Run(param):
    _locations, _nos, _names = _GetLocations()
    _range = list(map(int, param["Range"].split(",")))
    _POICat = next(num for name, num in Visum.Net.POICategories.GetMultipleAttributes(["NAME","NO"]) if name == param["POICat"])
    _mode = {_("Car") : "driving-car", _("Walk") : "foot-walking", _("Bike") : "cycling-regular", _("Pedelec") : "cycling-electric"}[param["Mode"]]
    _imp =  {_("Distance"): "distance", _("Time") : "time"}[param["Imp"]]
    if _imp == "time":
        _range = [sec * 60 for sec in _range]

    # ==== ORS request ====
    client = ors.Client(key = "eyJvcmciOiI1YjNjZTM1OTc4NTExMTAwMDFjZjYyNDgiLCJpZCI6ImQ5OTgwMjY4YjVkZDRlODJiZTFjMTQ5ZGZjOTFkMjAzIiwiaCI6Im11cm11cjY0In0=")
    payload = {
      "locations" : _locations,
      "range" : _range,
      "attributes" : ["total_pop"],
      "range_type" : _imp,
      "smoothing" : 0}
    if _mode != "driving-car":
        payload["options"] = {"avoid_features": ["ferries"]}
    
    iso = client.request(url = f"/v2/isochrones/{_mode}", post_json=payload)
    
    # ==== Save to file ====
    Visum.Log(20480, _("Importing shapes to Visum-POI"))
    temp_dir = tempfile.mkdtemp()
    json_path = os.path.join(temp_dir, "isochrone.geojson")
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(iso, f, ensure_ascii=False, indent=2)
     
    # ==== Import to Visum ====
    JSONImport = Visum.IO.CreateImportGeoJSONPara()
    JSONImport.AddAttributeAllocation("value", "CODE")
    JSONImport.AddAttributeAllocation("group_index", "NAME")
    JSONImport.AddAttributeAllocation("total_pop", "COMMENT")
    JSONImport.ObjectType = 9 # import as POI
    JSONImport.SetAttValue("POIKEY", _POICat)
    Visum.IO.ImportGeoJSON(json_path, JSONImport)
    
    # ==== Edit POIs ====
    POIs = Visum.Net.POICategories.ItemByKey(_POICat).POIs.GetAll[-(len(_range) * len(_locations)):]
    for poi in POIs:
        _group = int(poi.AttValue("NAME"))
        _value = int(float(poi.AttValue("CODE")))
        _total_pop = int(float(poi.AttValue("COMMENT")))
        poi.SetAttValue("NAME", _names[_group])
        poi.SetAttValue("CODE", _value)
        _date = datetime.now().strftime("%d.%m.%Y")
        poi.SetAttValue("COMMENT", f"mode: {param['Mode']} | imp: {param['Imp']} | pop: {_total_pop} | date: {_date}")
    
    # ==== End ====
    Visum.Net.GraphicParameters.POIs.SetAttValue("DRAWSURFACES", 0)
    Visum.Net.GraphicParameters.POIs.SetAttValue("DRAWSURFACES", 1)
    Visum.Log(20480, _("Finished"))

def _GetLocations():
    _xy, _no, _name = [], [], []
    if Visum.Net.Nodes.CountActive <= 5:
        for i in Visum.Net.Nodes.GetMultipleAttributes(["NO", "NAME", "WKTLOCWGS84"], True):
            geom = wkt.loads(i[2])
            _xy.append(list(geom.coords[0]))
            _no.append(i[0])
            _name.append(i[1])
        return _xy, _no, _name
    for i in Visum.Net.Marking.GetAll:
        geom = wkt.loads(i.AttValue("WKTLOCWGS84"))
        _xy.append(list(geom.coords[0]))
        _no.append(i.AttValue("NO"))
        _name.append(i.AttValue("NAME"))
    return _xy, _no, _name

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
        defaultParam = {"POICat" : 0}
        param = addInParam.Check(True, defaultParam)
        Run(param)
    except:
        addIn.HandleException(addIn.TemplateText.MainApplicationError)
