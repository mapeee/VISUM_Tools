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


# Function to generate date strings for a given year
def generate_date_strings(year):
    start_date = datetime(year, 1, 1)
    date_list = []
    for i in range(365 + (1 if (year % 4 == 0 and (year % 100 != 0 or year % 400 == 0)) else 0)):
        current_date = start_date + timedelta(days=i)
        date_string1 = current_date.strftime("%b_%a_%d")
        date_string2 = current_date.strftime("%d.%m.%Y")
        date_list.append([date_string1, date_string2])
    return date_list

# Example usage
year = 2024 # Change the year as needed 
date_strings = generate_date_strings(year)

# Print the first 10 dates as a sample
TimeIntSet = Visum.Net.AddTimeIntervalSet(3)
TimeIntSet.SetAttValue("CODE", str(year))
TimeIntSet.SetAttValue("Name", "AFZ")

for i, date in enumerate(date_strings):
    TimeInt = TimeIntSet.AddTimeInterval(date[0], i, i + 1, DayIndex=1)
    TimeInt.SetAttValue("Name", date[1])
    