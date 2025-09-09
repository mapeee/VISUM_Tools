# -*- coding: utf-8 -*-

import pandas as pd
import sys
import wx
from VisumPy.AddIn import AddIn, AddInState, AddInParameter
_ = AddIn.gettext

from MapFrame import PlotFrame
    

def Run():    
    boxes, xmin, xmax, ymax = CreateDict()
    frame = PlotFrame(None, boxes, xlim = (xmin, xmax), ylim = (0, ymax + 2))
    frame.Show()
    
    Visum.Log(20480, _("Map PT Journeys"))

def CreateDict():
    attrList = [["LINENAME", r"LINEROUTE\LINE\MAINLINENAME", "DEP", "ARR", "DURATION", "NAME", r"MIN:VEHJOURNEYSECTIONS\VEHCOMBNO"],
                ["LINENAME", "MAINLINE", "DEP", "ARR", "DURATION", "NAME", "VEH"]]
    VJTable = pd.DataFrame(Visum.Net.VehicleJourneys.GetMultipleAttributes(attrList[0], True), columns = attrList[1])
    VJTable.sort_values(by="DEP", ascending = True, inplace = True)
    VJTable["RowCount"] = _getRowCount(VJTable)
    
    # box values
    VJTable["x0"] = VJTable["DEP"] / 60 / 60
    VJTable["y0"] = VJTable["RowCount"] + 0.03
    VJTable["w"] = VJTable["DURATION"] / 60 / 60
    VJTable["h"] = 0.94
    VJTable["fill_alpha"] = 0.4
    VJTable = _setColors(VJTable)

    _boxes = VJTable.to_dict(orient = "records")
    _xmin = int(VJTable["DEP"].min() / 60 / 60)
    _xmax = int(VJTable["ARR"].max() / 60 / 60) + 1
    _ymax = VJTable["RowCount"].max()
    
    return _boxes, _xmin, _xmax, _ymax

def _getRowCount(_VJTable):
    lst = [[0, 0]]
    rowcount = []
    
    for _, row in _VJTable.iterrows():
        dep, arr = row["DEP"], row["ARR"]
        replaced = False
        for i, (arrl, rowl) in enumerate(lst):
            if dep >= arrl:
                lst[i] = [arr, lst[i][1]]   # replace
                replaced = True
                rowcount.append(lst[i][1])
                break
        if not replaced:
            lst.append([arr, lst[i][1] + 1])  # add as new element
            rowcount.append(lst[i][1] + 1)
    return rowcount

def _setColors(_VJTable):
    colors = {
    "AKN": ["darkorange", "darkorange", "black"],
    "Bus": ["peru", "peru", "black"],
    "Metrobus": ["blueviolet", "blueviolet", "black"],
    "Nachtbus": ["midnightblue", "midnightblue", "black"],
    "RB": ["black", "black", "black"],
    "RE": ["black", "black", "black"],
    "S-Bahn": ["lawngreen", "lawngreen", "black"],
    "Schiff": ["cyan", "cyan", "black"],
    "U-Bahn": ["blue", "blue", "black"],
    "XpressBus": ["lime", "lime", "black"],
    "IC": ["red", "red", "black"],
    "ICE": ["red", "red", "black"],
    "ICE(S)": ["red", "red", "black"]}
       
    colorsdf = pd.DataFrame.from_dict(colors, orient = "index", columns = ["edgecolor", "facecolor", "textcolor"])
    colorsdf["MAINLINE"] = colorsdf.index
    colorsdf = colorsdf.reset_index(drop=True)
    
    _VJTable = pd.merge(_VJTable, colorsdf, on = "MAINLINE", how = "left")
    _VJTable[["edgecolor", "facecolor", "textcolor"]] = _VJTable[["edgecolor", "facecolor", "textcolor"]].fillna("silver")
    
    return _VJTable

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
        Run()
    except:
        addIn.HandleException(addIn.TemplateText.MainApplicationError)
