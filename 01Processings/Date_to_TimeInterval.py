# -*- coding: utf-8 -*-
"""
Created on Wed Jan 29 14:07:35 2025

@author: peter
"""
import win32com.client.dynamic
from datetime import datetime, timedelta


Visum = win32com.client.dynamic.Dispatch("Visum.Visum.25")
Visum.IO.loadversion(r"")
Visum.Filters.InitAll()

years = [2024]


def generate_date_strings(_year):
    start_date = datetime(_year, 1, 1)
    date_list = []
    for i in range(365 + (1 if (_year % 4 == 0 and (_year % 100 != 0 or _year % 400 == 0)) else 0)):
        current_date = start_date + timedelta(days=i)
        date_string1 = current_date.strftime("%b_%a_%d")
        date_string2 = current_date.strftime("%d.%m.%Y")
        date_list.append([date_string1, date_string2])
    return date_list


maxinterval = max(Visum.Net.TimeIntervalSets.GetMultiAttValues("NO"))[1] + 1

for e, year in enumerate(years):
    TimeIntNo = int(maxinterval + e)
    TimeIntSet = Visum.Net.AddTimeIntervalSet(TimeIntNo)
    TimeIntSet.SetAttValue("CODE", str(year))
    TimeIntSet.SetAttValue("Name", "AFZ")

    date_strings = generate_date_strings(year)
    for i, date in enumerate(date_strings):
        TimeInt = TimeIntSet.AddTimeInterval(date[1], i, i + 1, DayIndex=1)
        TimeInt.SetAttValue("Name", date[0])
    
 
    for UDA in ["FahrtNr_AFZ", "Einsteiger_AFZ", "Aussteiger_AFZ", "FzgNr_AFZ"]:
        Visum.Net.VehicleJourneyItems.AddUserDefinedAttribute(UDA, UDA, UDA, 1)
    
    

    