#-------------------------------------------------------------------------------
# Name:        Add and delete Network Elements used for PTV Visum accessibility measures
# Author:      mape
# Created:     25/03/2025
# Copyright:   (c) mape 2025
# Licence:     GNU GENERAL PUBLIC LICENSE
# #---------------------

# UDG (User defined groups)
UDGS = {
    "Erreichbarkeiten": {"Description": "Erreichbarkeitsindikatoren und Zwischenschritte"}
    }

# UDA (User defined attributes)
UDAS = {
    "Stops": {
        "HKAT": {
            "ID": "HKAT",
            "Name": "Haltestellenkategorie",
            "VT": 5,
            "DefVal": "X",
            "Comment": "Haltestellenkategorie (OEV-Gueteklasse)"},
        "HKAT_FHH": {
            "ID": "HKAT_FHH",
            "Name": "Haltestellenkategorie in FHH",
            "VT": 5,
            "DefVal": "X",
            "Comment": "Haltestellenkategorie für Haltestellentyp 1 (OEV-Gueteklasse)"},
        "HKATT1": {
            "ID": "HKATT1",
            "Name": "Haltestellenkategorie HstTyp 1",
            "VT": 5,
            "DefVal": "X",
            "Comment": "Haltestellenkategorie für Haltestellentyp 1 (OEV-Gueteklasse)"},
        "HKATT2": {
            "ID": "HKATT2",
            "Name": "Haltestellenkategorie HstTyp 2",
            "VT": 5,
            "DefVal": "X",
            "Comment": "Haltestellenkategorie für Haltestellentyp 2 (OEV-Gueteklasse)"},
        "HKATT3": {
            "ID": "HKATT3",
            "Name": "Haltestellenkategorie HstTyp 3",
            "VT": 5,
            "DefVal": "X",
            "Comment": "Haltestellenkategorie für Haltestellentyp 3 (OEV-Gueteklasse)"}
            },
    "StopAreas": {
        "CC_SNAMUTS": {
            "ID": "Cc_SNAMUTS",
            "Name": "Closeness centrality",
            "VT": 2,
            "DefVal": 0,
            "Comment": "Closeness centrality (SNAMUTS)"},
        "DC_SNAMUTS": {
            "ID": "Dc_SNAMUTS",
            "Name": "Degree centrality",
            "VT": 2,
            "DefVal": 0,
            "Comment": "Degree centrality (SNAMUTS)"}
                }
    }

# POI Categories
POICats = {
    4: {"Code": "Erreichbarkeiten",
        "Name": "E Erreichbarkeiten",
        "Parent": 0},
    40: {"Code": "OEV-Gueteklassen",
        "Name": "E OEV-Gueteklassen",
        "Parent": 4},
    }

# Functions
def UDG(_UDGS):
    for UDG, attr in _UDGS.items():
        if any(name == UDG for _, name in Visum.Net.UserDefinedGroups.GetMultiAttValues("Name")):
            if Add == False:
                Visum.Net.UserDefinedGroups.GetFilteredSet(f'[NAME]="{UDG}"').RemoveAll()
                Visum.Log(20480, f"UserDefinedGroup: '{UDG}' deleted")
            else:
                continue
        else:
            if Add == True:
                Visum.Net.AddUserDefinedGroup(UDG, attr["Description"])
                Visum.Log(20480, f"UserDefinedGroup: '{UDG}' added")

def UDA(_UDAS):
    for UDA, attr in _UDAS.items():
        if UDA == "Stops":
            for _key, subdict in _UDAS["Stops"].items():
                if Add == False and Visum.Net.Stops.AttrExists(subdict["ID"]):
                    Visum.Net.Stops.DeleteUserDefinedAttribute(subdict["ID"])
                    Visum.Log(20480, f"Stops-UserDefinedAttribute: '{subdict['ID']}' deleted")
                    continue
                if Add == True and Visum.Net.Stops.AttrExists(subdict["ID"]) == False:
                    Visum.Net.Stops.AddUserDefinedAttribute(subdict["ID"], subdict["Name"], subdict["Name"], subdict["VT"], DefVal = subdict["DefVal"])
                    UDA_Visum = Visum.Net.Stops.Attributes.ItemByKey(subdict["ID"])
                    UDA_Visum.UserDefinedGroup = "Erreichbarkeiten"
                    UDA_Visum.Comment = subdict["Comment"]
                    Visum.Log(20480, f"Stops-UserDefinedAttribute: '{subdict['ID']}' added")
        if UDA == "StopAreas":
            for _key, subdict in _UDAS["StopAreas"].items():
                if Add == False and Visum.Net.StopAreas.AttrExists(subdict["ID"]):
                    Visum.Net.StopAreas.DeleteUserDefinedAttribute(subdict["ID"])
                    Visum.Log(20480, f"StopAreas-UserDefinedAttribute: '{subdict['ID']}' deleted")
                    continue
                if Add == True and Visum.Net.StopAreas.AttrExists(subdict["ID"]) == False:
                    Visum.Net.StopAreas.AddUserDefinedAttribute(subdict["ID"], subdict["Name"], subdict["Name"], subdict["VT"], DefVal = subdict["DefVal"])
                    UDA_Visum = Visum.Net.StopAreas.Attributes.ItemByKey(subdict["ID"])
                    UDA_Visum.UserDefinedGroup = "Erreichbarkeiten"
                    UDA_Visum.Comment = subdict["Comment"]
                    Visum.Log(20480, f"StopAreas-UserDefinedAttribute: '{subdict['ID']}' added")
                
def POI(_POICats):
    for POICat, attr in _POICats.items():
        if Add == False and any(Cat == POICat for _, Cat in Visum.Net.POICategories.GetMultiAttValues("NO", False)):
            Visum.Net.RemovePOICategory(Visum.Net.POICategories.ItemByKey(POICat))
            Visum.Log(20480, f"POI-Category {POICat} ('{attr['Code']}') deleted")
            continue
        if Add == True and any(Cat == POICat for _, Cat in Visum.Net.POICategories.GetMultiAttValues("NO", False)) == False:
            POIAdd = Visum.Net.AddPOICategory()
            POIAdd.SetAttValue("NO", POICat)
            POIAdd.SetAttValue("CODE", attr["Code"])
            POIAdd.SetAttValue("NAME", attr["Name"])
            if attr["Parent"] != 0:
                POIAdd.SetAttValue("PARENTCATNO", attr["Parent"])
            Visum.Log(20480, f"POI-Category {POICat} ('{attr['Code']}') added")

# Process
UDG(UDGS)
UDA(UDAS)
POI(POICats)