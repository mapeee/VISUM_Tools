#!/usr/bin/env python
import wx
import sys
import os
from VisumPy.AddIn import AddIn, AddInState, AddInParameter
_ = AddIn.gettext

class MyDialog(wx.Dialog):
    def __init__(self, parent, title):
        super(MyDialog, self).__init__(parent, id=-1, title=title, size=(400, 300), style=wx.CAPTION | wx.STAY_ON_TOP)

        self.__InitUI()

    def __InitUI(self):
        self.textctrl_Range = wx.TextCtrl(self, -1, "...", size=(150, -1))
        
        self.combo_POICat = wx.ComboBox(self, -1, "...")
        self.combo_Mode = wx.ComboBox(self, -1, "...")
        self.combo_Imp = wx.ComboBox(self, -1, "...")
        
        self.button_ok = wx.Button(self, -1, _("OK"))
        self.button_init = wx.Button(self, -1, _("Init"))
        self.button_help = wx.Button(self, -1, _('Help'))
        self.button_exit = wx.Button(self, wx.ID_CANCEL, _('Cancel'))

        self.combo_Imp.Bind(wx.EVT_COMBOBOX, self.OnComboImpChoice)
        self.Bind(wx.EVT_BUTTON, self.OnOK, self.button_ok)
        self.Bind(wx.EVT_BUTTON, self.OnInit, self.button_init)
        self.Bind(wx.EVT_BUTTON, self.OnHelp, self.button_help)
        self.Bind(wx.EVT_BUTTON, self.OnExit, self.button_exit)

        self.__do_layout()
        self.__set_properties()
        self.__do_comboChoice()
        
        defaultParam = {"POICat" : "...", "Imp" : "...", "Mode" : "...", "Range" : "..."}
        param = addInParam.Check(False, defaultParam)   
        
    def __do_comboChoice(self):
        self.combo_POICat.Clear()
        self.combo_POICat.Append([i.AttValue("Name") for i in Visum.Net.POICategories.GetAll])
        self.combo_POICat.SetValue("...")
        
        self.combo_Mode.Clear()
        self.combo_Mode.Append([_("Car"), _("Walk"), _("Bike"), _("Pedelec")])
        self.combo_Mode.SetValue(_("Walk"))
        
        self.combo_Imp.Clear()
        self.combo_Imp.Append([_("Distance"), _("Time")])
        self.combo_Imp.SetValue(_("Distance"))
        
        self.textctrl_Range.SetValue(_("Ranges in meter (sep ',')"))
        
    def __set_properties(self):
        _version = "1.0"
        self.SetTitle(_("PrT-IsoChrones To POI %s") %_version)
        
    def __do_layout(self):  
        sb_pref = wx.StaticBox(self, -1, _("Preferences"))
        sb_pref.SetFont(wx.Font(8, wx.DEFAULT, wx.NORMAL, wx.BOLD))
        sbSizer_pref = wx.StaticBoxSizer(sb_pref, wx.VERTICAL)
        sbSizer_pref.AddSpacer(10)
        sbSizer_pref.Add(self.combo_POICat, flag = wx.ALIGN_CENTER | wx.ALL, border = 2)
        sbSizer_pref.Add(self.combo_Mode, flag = wx.ALIGN_CENTER | wx.ALL, border = 2)
        sbSizer_pref.Add(self.combo_Imp, flag = wx.ALIGN_CENTER | wx.ALL, border = 2)
        sbSizer_pref.Add(self.textctrl_Range, flag = wx.ALIGN_CENTER | wx.ALL, border = 2)
        
        sb_end = wx.StaticBox(self, -1, _("End"))
        sb_end.SetFont(wx.Font(8, wx.DEFAULT, wx.NORMAL, wx.BOLD))
        sbSizer_end = wx.StaticBoxSizer(sb_end, wx.VERTICAL)
        sbSizer_end.SetMinSize((200, 50))
        sbSizer_end.AddSpacer(10)
        sbSizer_end.Add(self.button_ok, flag = wx.ALIGN_CENTER | wx.ALL, border = 2)
        sbSizer_end.Add(self.button_init, flag = wx.ALIGN_CENTER | wx.ALL, border = 2)
        sbSizer_end.AddSpacer(10)
        sbSizer_end.Add(self.button_help, flag = wx.ALIGN_CENTER | wx.ALL, border = 2)
        sbSizer_end.Add(self.button_exit, flag = wx.ALIGN_CENTER | wx.ALL, border = 2)
    
        vbox = wx.BoxSizer(wx.VERTICAL)
        vbox.Add(sbSizer_pref, proportion = 0, flag = wx.EXPAND | wx.LEFT | wx.RIGHT, border = 10)
        vbox.AddSpacer(10)
        vbox.Add(sbSizer_end, proportion = 1, flag = wx.EXPAND | wx.LEFT | wx.RIGHT, border = 10)
        vbox.AddSpacer(10)

        self.SetSizerAndFit(vbox)
        self.Layout()
        self.Centre()
    
    def OnComboImpChoice(self,event):
        _imp =  {_("Distance"): "distance", _("Time") : "time"}[self.combo_Imp.GetValue()]
        if _imp == "distance":
            self.textctrl_Range.SetValue(_("Ranges in meter (sep ',')"))
        else:
            self.textctrl_Range.SetValue(_("Ranges in minutes (sep ',')"))
        return
    
    def OnExit(self,event):
        if not addIn.IsInDebugMode:
            Terminated.set()
        self.Destroy()
        
    def OnHelp(self,event):
        try:
            os.startfile(addIn.DirectoryPath + _("HelpPrTIsochroneToPOI.htm"))
        except:
            addIn.HandleException()  
            
    def OnInit(self,event):
        self.__do_comboChoice()
        
    def OnOK(self,event):
        param, paramOK = self.setParameter()
        if not _TestLocations():
            return
        if not _TestRanges(param["Range"]):
            return
        
        if not paramOK or param["POICat"] == "...":
            addIn.ReportMessage(_("Please select all values!"))
            return
        
        addInParam.SaveParameter(param)
        
        self.OnExit(None)
            
    def setParameter(self):
        param = dict()
        try:
            param["POICat"] = self.combo_POICat.GetValue()
            param["Mode"] = self.combo_Mode.GetValue()
            param["Imp"] = self.combo_Imp.GetValue()
            param["Range"] = self.textctrl_Range.GetValue()
            return param, True
        except:
            addIn.HandleException(_("PrTIsoChroneToPOI, value error: "))
            return param, False

def CheckNetwork():
    if 0 in [Visum.Net.Nodes.Count, Visum.Net.POICategories.Count]:
        addIn.ReportMessage(_("Current Visum Version has no Nodes and/or POI-Categories!"))
        if not addIn.IsInDebugMode:
            Terminated.set()
        return False
    return True

def _TestLocations():
    if Visum.Net.Nodes.CountActive <= 5:
        return True
    if Visum.Net.Marking.ObjectType != 1:
        addIn.ReportMessage(_("Please mark Node(s)!"))
        return False
    _markings = Visum.Net.Marking.Count
    if _markings == 0:
        addIn.ReportMessage(_("No nodes selected!"))
        return False
    if _markings > 5:
        addIn.ReportMessage(_("%s nodes selected, 5 are max!") %str(_markings))
        return False
    return True

def _TestRanges(_param):
    try:
        list(map(int, _param.split(",")))
    except:
        addIn.ReportMessage(_("Wrong format for ranges!"))
        return False
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
            dialog_1 = MyDialog(None, "")
            app.SetTopWindow(dialog_1)
            dialog_1.ShowModal()
            if addIn.IsInDebugMode:
                app.MainLoop()
    except:
        addIn.HandleException(addIn.TemplateText.MainApplicationError)
        if not addIn.IsInDebugMode:
            Terminated.set()
