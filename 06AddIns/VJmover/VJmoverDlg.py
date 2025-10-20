#!/usr/bin/env python
import sys
import os
import pandas as pd
import wx
from VisumPy.helpers import SetMulti
from VisumPy.AddIn import AddIn, AddInState, AddInParameter
_ = AddIn.gettext

class InfoFrame(wx.Frame):
    def __init__(self, title):
        super(InfoFrame, self).__init__(None, id=-1, title=title, style=wx.CAPTION | wx.STAY_ON_TOP,
                                        size=(190, 190))
        self.Centre()        
        img = wx.Image(addIn.DirectoryPath +'logo.png',wx.BITMAP_TYPE_ANY)
        img = img.Scale(55,30,wx.IMAGE_QUALITY_BOX_AVERAGE)
        img = img.ConvertToBitmap()
        png = wx.StaticBitmap(self, -1, img, (0, 0))
        self.button = wx.Button(self, -1, _("OK"))
        self.Bind(wx.EVT_BUTTON, self.__OnOK, self.button)
        self.SetBackgroundColour(wx.Colour(wx.NullColour))
        self.label = wx.StaticText(self,label= _("hvv GmbH"))
        sizer = wx.BoxSizer(wx.VERTICAL)
        sizer.Add(png,0,wx.LEFT,5)
        sizer.AddSpacer(10)
        sizer.Add(self.label,0,wx.LEFT,10)
        sizer.AddSpacer(2)
        sizer.Add(wx.StaticText(self,label= _("Marcus Peter")),0,wx.LEFT,10)
        sizer.Add(wx.StaticText(self,label= _("24.09.2025")),0,wx.LEFT,10)
        sizer.AddSpacer(2)
        sizer.Add(wx.StaticText(self,label= _("Version 1.0: Initial")),0,wx.LEFT,10)
        sizer.Add(wx.StaticText(self,label= _("Version 1.1: Duplicate VJ")),0,wx.LEFT,10)
        sizer.AddSpacer(10)
        sizer.Add(self.button,0,wx.ALIGN_CENTER,5)
        sizer.AddSpacer(5)
        self.SetSizer(sizer)
        font = wx.Font(8, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.BOLD)
        self.label.SetFont(font)
        self.Show()

    def __OnOK(self, _event):
        self.Close(True)

class MyDialog(wx.Dialog):
    def __init__(self, parent, title):
        super(MyDialog, self).__init__(parent, id=-1, title=title, size=(400, 300),
                                       style=wx.CAPTION | wx.STAY_ON_TOP)

        self.__InitUI()

    def __InitUI(self):   
        self.button_tt = wx.Button(self, -1, _("Open TimeTable"))
        self.toggle_vals = wx.ToggleButton(self, -1, _("Time"), name = "values")
        self.toggle_stops = wx.ToggleButton(self, -1, _(" "), name = "route")
        self.toggle_vals.Bind(wx.EVT_TOGGLEBUTTON, self._on_toggle)
        self.toggle_stops.Bind(wx.EVT_TOGGLEBUTTON, self._on_toggle)        
        self.button_1 = wx.Button(self, -1, _("+1"), name = "1")
        self.button_5 = wx.Button(self, -1, _("+5"), name = "5")
        self.button_15 = wx.Button(self, -1, _("+15"), name = "15")
        self.button_1m = wx.Button(self, -1, _("-1"), name = "-1")
        self.button_5m = wx.Button(self, -1, _("-5"), name = "-5")
        self.button_15m = wx.Button(self, -1, _("-15"), name = "-15")
        self.button_dupli = wx.Button(self, -1, _('Duplicate'))
        self.button_help = wx.Button(self, -1, _('Help'))
        self.button_info = wx.Button(self, -1, _('Info'))
        self.button_exit = wx.Button(self, wx.ID_CANCEL, _('Close'))

        self.Bind(wx.EVT_BUTTON, self.OnTT, self.button_tt)
        self.Bind(wx.EVT_BUTTON, self.OnMove, self.button_1)
        self.Bind(wx.EVT_BUTTON, self.OnMove, self.button_5)
        self.Bind(wx.EVT_BUTTON, self.OnMove, self.button_15)
        self.Bind(wx.EVT_BUTTON, self.OnMove, self.button_1m)
        self.Bind(wx.EVT_BUTTON, self.OnMove, self.button_5m)
        self.Bind(wx.EVT_BUTTON, self.OnMove, self.button_15m)
        self.Bind(wx.EVT_BUTTON, self.OnDupli, self.button_dupli)
        self.Bind(wx.EVT_BUTTON, self.OnHelp, self.button_help)
        self.Bind(wx.EVT_BUTTON, self.OnInfo, self.button_info)
        self.Bind(wx.EVT_BUTTON, self.OnExit, self.button_exit)
        
        self.tt = False
        self.__do_layout()
        
    def __do_layout(self):
        self.button_1m.SetForegroundColour(wx.Colour(255, 0, 0))
        self.button_5m.SetForegroundColour(wx.Colour(255, 0, 0))
        self.button_15m.SetForegroundColour(wx.Colour(255, 0, 0))
        self.button_dupli.SetForegroundColour(wx.Colour(0, 100, 0))
        sb_exe = wx.StaticBox(self, -1)
        sb_exe.SetFont(wx.Font(10, wx.DEFAULT, wx.NORMAL, wx.BOLD))
        sbSizer_exe = wx.StaticBoxSizer(sb_exe, wx.VERTICAL)
        sbSizer_exe.SetMinSize((200, 100))
        gbSizer = wx.GridBagSizer(hgap=5, vgap=5)
        gbSizer.Add(self.toggle_vals, (0, 0), wx.DefaultSpan, wx.ALL, 5)
        gbSizer.Add(self.toggle_stops, (0, 1), wx.DefaultSpan, wx.ALL, 5)
        gbSizer.Add(self.button_1, (1, 0), wx.DefaultSpan, wx.ALL, 5)
        gbSizer.Add(self.button_5, (2, 0), wx.DefaultSpan, wx.ALL, 5)
        gbSizer.Add(self.button_15, (3, 0), wx.DefaultSpan, wx.ALL, 5)
        gbSizer.Add(self.button_1m, (1, 1), wx.DefaultSpan, wx.ALL, 5)
        gbSizer.Add(self.button_5m, (2, 1), wx.DefaultSpan, wx.ALL, 5)
        gbSizer.Add(self.button_15m, (3, 1), wx.DefaultSpan, wx.ALL, 5)
        sbSizer_exe.Add(self.button_tt, flag = wx.ALIGN_CENTER | wx.TOP, border = 10)
        sbSizer_exe.Add(gbSizer, flag = wx.ALIGN_CENTER | wx.TOP, border = 10)
        sbSizer_exe.AddSpacer(5)
        sbSizer_exe.Add(self.button_dupli, flag = wx.ALIGN_CENTER | wx.TOP, border = 10)
        sbSizer_exe.AddSpacer(15)
        sbSizer_exe.Add(self.button_help, flag = wx.ALIGN_CENTER | wx.TOP, border = 10)
        sbSizer_exe.Add(self.button_info, flag = wx.ALIGN_CENTER | wx.TOP, border = 10)
        sbSizer_exe.Add(self.button_exit, flag = wx.ALIGN_CENTER | wx.TOP, border = 10)
        vbox = wx.BoxSizer(wx.VERTICAL)    
        vbox.Add(sbSizer_exe, proportion = 1, flag = wx.EXPAND | wx.LEFT | wx.RIGHT, border = 10)
        vbox.AddSpacer(10)

        self.SetSizerAndFit(vbox)
        self.Layout()
        self.Centre()
        
    def _on_toggle(self, event):
        button = event.GetEventObject()
        if button.GetName() == "route":
            self.toggle_vals.SetLabel(_("Route"))
            self.toggle_vals.SetValue(True)
        if button.GetValue():
            if button.GetName() == "route":
                button.SetLabel(_("Stops From"))
            else:
                button.SetLabel(_("Route"))
                self.toggle_stops.SetLabel(_("Stops From"))
                self.toggle_stops.SetValue(True)
        else:
            if button.GetName() == "route":
                button.SetLabel(_("Stops To"))
            else:
                button.SetLabel(_("Time"))
                self.toggle_stops.SetLabel(_(" "))
                self.toggle_stops.SetValue(False)

    def OnDupli(self, event):
        if not _setVJactive():
            return
        if not self.tt:
            addIn.ReportMessage(_("Re-open TimeTable via button!"))
            return False
        df_VJ = pd.DataFrame(Visum.Net.VehicleJourneys.GetMultipleAttributes(["NO", "TIMEPROFILEID", "DEP", "FROMTPROFITEMINDEX",
                                                                              "TOTPROFITEMINDEX", "OPERATORNO",
                                                                              r"MIN:VEHJOURNEYSECTIONS\VEHCOMBNO",
                                                                              r"MIN:VEHJOURNEYSECTIONS\VALIDDAYSNO",
                                                                              r"MIN:VEHJOURNEYSECTIONS\OPERATINGPERIODNO"], True),
                             columns = ["NO", "TIMEPROFILEID", "DEP", "FROMTPROFITEMINDEX", r"TOTPROFITEMINDEX", "OPERATORNO",
                                        "VEHCOMBNO", "VALIDDAYSNO", "OPERATINGPERIODNO"])
        df_VJ.sort_values(by="DEP", inplace=True)
        for index, row in df_VJ.iterrows():
            df_VJNO = pd.DataFrame(Visum.Net.VehicleJourneys.GetMultipleAttributes(["NO"]), columns = ["NO"])
            VJNO = _getFreeNO(df_VJNO, row["NO"])
            VJ = Visum.Net.AddVehicleJourney(VJNO, row["TIMEPROFILEID"])
            VJ.SetAttValue("FROMTPROFITEMINDEX", row["FROMTPROFITEMINDEX"])
            VJ.SetAttValue("TOTPROFITEMINDEX", row["TOTPROFITEMINDEX"])
            VJ.SetAttValue("DEP", row["DEP"])
            VJ.SetAttValue("OPERATORNO", row["OPERATORNO"])
            for VJS in VJ.VehicleJourneySections.GetAll:
                VJS.SetAttValue("VEHCOMBNO", row["VEHCOMBNO"])
                VJS.SetAttValue("VALIDDAYSNO", row["VALIDDAYSNO"])
                VJS.SetAttValue("OPERATINGPERIODNO", row["OPERATINGPERIODNO"])
        
        Visum.Net.VehicleJourneys.SetActive()
        addInParam.SaveParameter(None)
        self.tt.SortByDepartureAtCommonStops()
        
    def OnExit(self, _event):
        if not addIn.IsInDebugMode:
            Terminated.set()
        self.Destroy()
    
    def OnHelp(self, _event):
        try:
            os.startfile(addIn.DirectoryPath + _("HelpVJmover.htm"))
        except:
            addIn.HandleException()             
    def OnInfo(self, _event):
        title = _("Info")
        InfoFrame(title=title)    
        
    def OnMove(self, event):
        if not _setVJactive():
            return
            
        if self.toggle_vals.GetValue() is False: # minutes
            attr = "DEP"
        elif self.toggle_stops.GetValue() is True: # stops from
            attr = "FROMTPROFITEMINDEX"
        else: # stops to
            attr = "TOTPROFITEMINDEX"
        movings = int(event.GetEventObject().GetName())
        if self.toggle_vals.GetValue() is False: # minutes
            movings = movings*60
        
        df_VJ = pd.DataFrame(Visum.Net.VehicleJourneys.GetMultipleAttributes(["DEP", "FROMTPROFITEMINDEX", "TOTPROFITEMINDEX",r"TIMEPROFILE\ENDTIMEPROFILEITEM\INDEX"], True), 
                             columns = ["DEP", "FROMTPROFITEMINDEX", "TOTPROFITEMINDEX", "ENDINDEX"])

        if self.toggle_vals.GetValue() is False: # minutes
            if (df_VJ[attr] + movings < 0).any():
                addIn.ReportMessage(_("Start before 0!"))
                return
        elif self.toggle_stops.GetValue() is True:
            if (df_VJ[attr] + movings < 1).any():
                addIn.ReportMessage(_("Start before Index 1!"))
                return
            if (df_VJ["FROMTPROFITEMINDEX"] + movings >= df_VJ["TOTPROFITEMINDEX"]).any():
                addIn.ReportMessage(_("Start after finish!"))
                return
        else:
            if (df_VJ["TOTPROFITEMINDEX"] + movings <= df_VJ["FROMTPROFITEMINDEX"]).any():
                addIn.ReportMessage(_("Finish before start!"))
                return
            if (df_VJ["TOTPROFITEMINDEX"] + movings > df_VJ["ENDINDEX"]).any():
                addIn.ReportMessage(_("Finish after last stop!"))
                return

        df_VJ[attr] = df_VJ[attr] + movings
        moverValues = df_VJ[attr].tolist()
        
        SetMulti(Visum.Net.VehicleJourneys, attr, moverValues, True)
        Visum.Net.VehicleJourneys.SetActive()
        addInParam.SaveParameter(None)
        
    def OnTT(self, event):
        if Visum.Workbench.IsTabularTimetableRunning():
            if not self.tt:
                addIn.ReportMessage(_("Close and re-open TimeTable via Button!"))
            else:
                addIn.ReportMessage(_("Tabluar TimeTable just open!"))
            return False
        self.tt = Visum.Workbench.CreateTabularTimetable
        self.tt.Show(True)

def CheckNetwork():
    if 0 in [Visum.Net.Nodes.Count,Visum.Net.Links.Count]:
        addIn.ReportMessage(_("Current Visum Version has no VehicleJourneys!"))
        if not addIn.IsInDebugMode:
            Terminated.set()
        return False
    return True

def _checkMarking():
    if Visum.Net.Marking.ObjectType != 20:
        addIn.ReportMessage(_("No VehicleJourneys marked!"))
        return False
    if Visum.Net.Marking.Count > 200:
        addIn.ReportMessage(_("To many VehicleJourneys marked!"))
        return False
    if Visum.Net.Marking.Count == 0:
        addIn.ReportMessage(_("No VehicleJourney marked!"))
        return False
    return True

def _getFreeNO(df, last):
    nums = set(df["NO"].astype(int))
    max_num = max(nums) if nums else last
    for i in range(int(last) + 1, int(max_num) + 2):
        if i not in nums:
            return i

def _setVJactive():
    if not _checkMarking():
        return False
    Visum.Net.VehicleJourneys.SetPassive()
    df_Active = pd.DataFrame(Visum.Net.VehicleJourneys.GetMultipleAttributes(["NO", "ISINSELECTION"]), columns = ["NO", "ISINSELECTION"])
    for VJ in Visum.Net.Marking.GetAll:
        NO_VJ = VJ.AttValue("NO")
        df_Active.loc[df_Active["NO"] == NO_VJ, "ISINSELECTION"] = 1
    activeVJ = df_Active["ISINSELECTION"].tolist()
    SetMulti(Visum.Net.VehicleJourneys, "ISINSELECTION", activeVJ)
    return True

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
        wx.InitAllImageHandlers()
        if CheckNetwork():
            dialog_1 = MyDialog(None, _("VJ Mover"))
            app.SetTopWindow(dialog_1)
            dialog_1.ShowModal()
            if addIn.IsInDebugMode:
                app.MainLoop()
    except:
        addIn.HandleException(addIn.TemplateText.MainApplicationError)
        if not addIn.IsInDebugMode:
            Terminated.set()
