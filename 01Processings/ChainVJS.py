# -*- coding: cp1252 -*-
#!/usr/bin/python

#-------------------------------------------------------------------------------
# Name:        Chain different VehicleJourneySections from specific Line at specific StopPoint
# Author:      mape
# Created:     27/05/2026
# Copyright:   (c) mape 2026
# Licence:     GNU GENERAL PUBLIC LICENSE
# #-------------------------------------------------------------------------------


'''
To calculate in PTV Visum
'''


# import json
# import os
import pandas as pd
# import win32com.client.dynamic


## Input ##
Line = "RE83"
StopNo = 8000237




# directory = os.path.dirname(os.path.abspath(__file__))
# with open(os.path.join(directory ,"config.json"), "r") as json_file:
#     data = json.load(json_file)



def chainVJS(_Visum, fromVJSdf, toVJSdf, dayIndex):
    fromVJS = fromVJSdf["VEHJOURNEYNO"]
    fromVJSNO = fromVJSdf["NO"]
    fromVJS = _Visum.Net.VehicleJourneySections.ItemByKey(fromVJS, fromVJSNO)
    toVJS = toVJSdf["VEHJOURNEYNO"]
    toVJSNO = toVJSdf["NO"]
    toVJS = _Visum.Net.VehicleJourneySections.ItemByKey(toVJS, toVJSNO)
    
    fromVJS.SetChainedUpSection(toVJS, dayIndex, True, True)


def dataVJS(_Visum, _Line, _StopNo):
    dfVJS = pd.DataFrame(_Visum.Net.VehicleJourneySections.GetMultipleAttributes(["NO","VEHJOURNEYNO",r"VEHJOURNEY\DIRECTIONCODE",
                                                                              r"VEHJOURNEY\LINENAME","NUMMONDAYS","NUMTUESDAYS",
                                                                              "NUMWEDNESDAYS","NUMTHURSDAYS","NUMFRIDAYS",
                                                                              "NUMSATURDAYS","NUMSUNDAYS","DEP","ARR",
                                                                              "FROMSTOPPOINTNO","TOSTOPPOINTNO"], True), 
                      columns = ["NO","VEHJOURNEYNO","DIRECT","LINE","MO","TUE","WED","THU","FRI","SAT","SUN","DEP","ARR","FROM","TO"])


    float_cols = dfVJS.select_dtypes(include='float64').columns
    dfVJS[float_cols] = dfVJS[float_cols].astype('int64')
    dfVJS = dfVJS[dfVJS['LINE'] == _Line]
    dfVJS = dfVJS[(dfVJS['FROM'] == _StopNo) | (dfVJS['TO'] == _StopNo)]
    return dfVJS


def identifier(_Visum, _VJSline, _StopNo):
    '''
    Identify VJS to chain
    '''
    for i in ["H", "R"]:
        _VJSdirect = _VJSline[_VJSline['DIRECT'] == i].copy()
        _VJSdirect['TIME'] = _VJSdirect['DEP'] # create Column 'Time' with timeslot for chainings
        _VJSdirect.loc[_VJSdirect['TO'] == _StopNo, 'TIME'] = _VJSdirect['ARR'] # relevant time is ARR-Time, when VJS ends at StopPoint
        _VJSdirect = _VJSdirect.sort_values('TIME')
        for idx, row in _VJSdirect.iterrows():
            if row["TO"] != _StopNo:
                continue
            for _dayIndex, day in enumerate(["MO","TUE","WED","THU","FRI","SAT","SUN"]):
                if row[day] == 0:
                    continue
                times = row["TIME"]
                matches = _VJSdirect[(_VJSdirect['TIME'].between(times, times+1200)) & 
                                     (_VJSdirect['FROM'] == _StopNo) & 
                                     (_VJSdirect[day] == 1)]
                if not matches.empty:
                    matches = matches.iloc[0]
                    chainVJS(_Visum, row, matches, _dayIndex+1)


# def VisumOpen(Net):
#     _Visum = win32com.client.dynamic.Dispatch("Visum.Visum.26")
#     _Visum.IO.loadversion(Net)
#     _Visum.Filters.InitAll()
#     return _Visum


# Visum = VisumOpen(data["Network"])
VJSline = dataVJS(Visum, Line, StopNo)
identifier(Visum, VJSline, StopNo)






