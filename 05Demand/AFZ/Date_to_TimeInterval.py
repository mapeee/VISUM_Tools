# -*- coding: cp1252 -*-
#!/usr/bin/python

#-------------------------------------------------------------------------------
# Name:        Script to prepare a Visum file for storing AFZ data with daily resolution.
# Author:      mape
# Created:     29/01/2025
# Copyright:   (c) mape 2025
# Licence:     GNU GENERAL PUBLIC LICENSE
#-------------------------------------------------------------------------------

from datetime import datetime, timedelta
import json
from pathlib import Path
import os
import sys
import win32com.client.dynamic

directory = Path(os.path.dirname(os.path.abspath(sys.argv[0])))
with open(os.path.join(directory ,"config.json"), "r") as json_file:
    data = json.load(json_file)

# generate string for each day a year
def generate_date_strings(_year):
    start_date = datetime(_year, 1, 1)
    date_list = []
    for i in range(365 + (1 if (_year % 4 == 0 and (_year % 100 != 0 or _year % 400 == 0)) else 0)):
        current_date = start_date + timedelta(days=i)
        date_string1 = current_date.strftime("%b_%a_%d")
        date_string2 = current_date.strftime("%d.%m.%Y")
        day_number = current_date.day
        month_number = current_date.month
        date_list.append([date_string1, date_string2, day_number, month_number])
    return date_list


# load Visum
Visum = win32com.client.dynamic.Dispatch("Visum.Visum.25")
Visum.IO.loadversion(data["Network"])
Visum.Filters.InitAll()

# parameter
years = [2024] # list containing years to enter
maxTimeIntervalNo = max(Visum.Net.TimeIntervalSets.GetMultiAttValues("NO"))[1] + 1


for e, year in enumerate(years):
    TimeIntNo = int(maxTimeIntervalNo + e)
    TimeIntSet = Visum.Net.AddTimeIntervalSet(TimeIntNo)
    TimeIntSet.SetAttValue("CODE", str(year))
    TimeIntSet.SetAttValue("Name", "AFZ")

    # create time interval for each string / day a year
    date_strings = generate_date_strings(year)
    for i, date in enumerate(date_strings):
        TimeInt = TimeIntSet.AddTimeInterval(date[1], date[2] + (date[3] * 3600), date[2] + 1  + (date[3] * 3600), DayIndex = 1)
        TimeInt.SetAttValue("Name", date[0])
    
    # create additional user defined attributes to organize AFZ data
    for UDA in [f"FahrtNr_AFZ_{year}", f"Einsteiger_AFZ_{year}", f"Aussteiger_AFZ_{year}", f"FzgNr_AFZ_{year}", f"Belastung_ab_AFZ_{year}"]:
        if Visum.Net.VehicleJourneyItems.AttrExists(UDA): continue
        Visum.Net.VehicleJourneyItems.AddUserDefinedAttribute(UDA, UDA, UDA, 1, 0, False, "", "", "", 0, False, "", f"TISET_{TimeIntNo}", True)
        UDG = Visum.Net.VehicleJourneyItems.Attributes.ItemByKey(UDA)
        UDG.UserDefinedGroup = "AFZ"
    
    UDA = f"Saldo_AFZ_{year}"
    if not Visum.Net.VehicleJourneys.AttrExists(UDA):
        Visum.Net.VehicleJourneys.AddUserDefinedAttribute(UDA, UDA, UDA, 1)
        UDG = Visum.Net.VehicleJourneys.Attributes.ItemByKey(UDA)
        UDG.UserDefinedGroup = "AFZ"
        
        
    
    

    