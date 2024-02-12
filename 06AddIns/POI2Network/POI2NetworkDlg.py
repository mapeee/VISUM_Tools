#!/usr/bin/env python
import wx
import sys
import os
import numpy as np
from VisumPy.AddIn import AddIn, AddInState, AddInParameter
_ = AddIn.gettext

class InfoFrame(wx.Frame):
    def __init__(self, title):
        super(InfoFrame, self).__init__(None, id=-1, title=title, style=wx.CAPTION | wx.STAY_ON_TOP, size=(180, 180))
        self.Centre()
        
        self.button = wx.Button(self, -1, _("OK"))
        self.Bind(wx.EVT_BUTTON, self.__OnOK, self.button)
        self.SetBackgroundColour(wx.Colour(wx.NullColour))
        
        self.label = wx.StaticText(self,label= _("hvv GmbH"))
        
        sizer = wx.BoxSizer(wx.VERTICAL)
        sizer.AddSpacer(10)
        sizer.Add(self.label,0,wx.LEFT,10)
        sizer.AddSpacer(2)
        sizer.Add(wx.StaticText(self,label= _("Marcus Peter")),0,wx.LEFT,10)
        sizer.Add(wx.StaticText(self,label= _("05.02.2024")),0,wx.LEFT,10)
        sizer.AddSpacer(2)
        sizer.Add(wx.StaticText(self,label= _("Version 1.0: Initial")),0,wx.LEFT,10)
        sizer.Add(wx.StaticText(self,label= _("Version 1.1: Initial Values")),0,wx.LEFT,10)
        sizer.Add(wx.StaticText(self,label= _("Version 1.2: Replace existing")),0,wx.LEFT,10)
        sizer.AddSpacer(20)
        sizer.Add(self.button,0,wx.ALIGN_CENTER,5)
        sizer.AddSpacer(10)
        self.SetSizer(sizer)
        
        font = wx.Font(8, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.BOLD)
        self.label.SetFont(font)

        self.Show()

    def __OnOK(self,event):
        self.Close(True)

class MyDialog(wx.Dialog):
    def __init__(self, parent, title):
        super(MyDialog, self).__init__(parent, id=-1, title=title, style=wx.CAPTION | wx.STAY_ON_TOP)

        self.__InitUI()
        self.Centre()

    def __InitUI(self):
        self.label1 = wx.StaticText(self, -1, _("POI - Category"))
        self.label2 = wx.StaticText(self, -1, _("POI - Attribute"))
        self.label3 = wx.StaticText(self, -1, _("Network Type"))
        self.label4 = wx.StaticText(self, -1, _("Replace existing?"))
        
        self.button3 = wx.Button(self, -1, _("OK"))
        self.button1 = wx.Button(self, -1, _("Help"))
        self.button2 = wx.Button(self, -1, _('Info'))
        self.button4 = wx.Button(self, -1, _('Cancel'))
        
        self.checkbox1 = wx.CheckBox(self) 
        
        self.combo1 = wx.ComboBox(self, -1, "...")
        self.combo1.Bind(wx.EVT_COMBOBOX, self.OnComboChoice)
        self.combo2 = wx.ComboBox(self, -1, "...")
        self.combo3 = wx.ComboBox(self, -1, "...")
   
        self.Bind(wx.EVT_BUTTON, self.OnHelp, self.button1)
        self.Bind(wx.EVT_BUTTON, self.OnInfo, self.button2)
        self.Bind(wx.EVT_BUTTON, self.OnOK, self.button3)
        self.Bind(wx.EVT_BUTTON, self.OnExit, self.button4)
        self.Bind(wx.EVT_CLOSE, self.OnExit)

        
        self.__do_layout()
        self.__set_properties()
        self.__do_comboChoice()
        
        defaultParam = {"POICat" : "...", "POIAttr" : "...", "NetworkType" : None, "Replace" : False}
        param = addInParam.Check(False, defaultParam)
        
        try:self.__initWxControlValues(param)
        except: pass
        
    def __initWxControlValues(self, param):
        self.combo1.SetSelection([int(i.AttValue("NO")) for i in Visum.Net.POICategories.GetAll].index(param["POICat"]))
        attrList = self.OnComboChoice(None)
        self.combo2.SetSelection(attrList.index(param["POIAttr"]))
        self.combo3.SetSelection(["Link","Node","StopPoint","StopArea","Zone"].index(param["NetworkType"]))
        self.checkbox1.SetValue(param["Replace"])
        
    def __set_properties(self):
        self.SetTitle(_("POI to Network"))
        self.button2.SetDefault()
        
    def __do_layout(self):       
        vbox = wx.BoxSizer(wx.VERTICAL)
        
        box_1 = wx.GridBagSizer()
        box_1.Add(self.label1, (1,1), wx.DefaultSpan, wx.ALIGN_CENTER_VERTICAL, 10)
        box_1.Add(self.combo1, (1,3), (1,3), wx.EXPAND, 10)
        box_1.Add(self.label2, (3,1), wx.DefaultSpan, wx.ALIGN_CENTER_VERTICAL, 10)
        box_1.Add(self.combo2, (3,3), (1,3), wx.EXPAND, 10)
        box_1.Add(self.label3, (5,1), wx.DefaultSpan, wx.ALIGN_CENTER_VERTICAL, 10)
        box_1.Add(self.combo3, (5,3), (1,3), wx.EXPAND, 10)
        box_1.Add(self.label4, (7,1), wx.DefaultSpan, wx.ALIGN_CENTER_VERTICAL, 10)
        box_1.Add(self.checkbox1, (7,3), (1,3), wx.EXPAND, 10)

        box_2 = wx.GridBagSizer()
        box_2.Add(self.button1, (0,1), wx.DefaultSpan, wx.EXPAND, 0)
        box_2.Add(self.button2, (0,2), wx.DefaultSpan, wx.EXPAND, 0)
        box_2.Add(self.button3, (0,4), wx.DefaultSpan, wx.EXPAND, 0)
        box_2.Add(self.button4, (0,5), wx.DefaultSpan, wx.EXPAND, 10)

        vbox.Add(box_1,0,wx.LEFT)
        vbox.AddSpacer(30)
        vbox.Add(box_2,0,wx.RIGHT,10)
        vbox.AddSpacer(20)
        self.SetSizerAndFit(vbox)

        self.Layout()
        
    def __do_comboChoice(self):
        self.combo1.Clear()
        self.combo1.Append([i.AttValue("Name") for i in Visum.Net.POICategories.GetAll])
        self.combo1.SetValue("...")
        
        self.combo3.Clear()
        self.combo3.Append([_(i) for i in ["Link","Node","StopPoint","StopArea","Zone"]])
        self.combo3.SetValue("...")

    def OnExit(self,event):
        if not addIn.IsInDebugMode:
            Terminated.set()
        self.Destroy()
        
    def OnComboChoice(self,event):
        self.combo2.Clear()
        v = [i.AttValue("NO") for i in Visum.Net.POICategories.GetAll][self.combo1.GetSelection()]
        l = []
        for i in Visum.Net.POICategories.ItemByKey(v).POIs.Attributes.GetAll:
            if i.ValueType==1: l.append(i.ID)
        self.combo2.Append(l)
        self.combo2.SetValue("...")
        return l
      
    def OnInfo(self,event):
        title = _("Info")
        frame = InfoFrame(title=title)
      
    def OnHelp(self,event):
        try:
            os.startfile(addIn.DirectoryPath + _("HelpPOI2Network.htm"))
        except:
            addIn.HandleException()
    
    def OnOK(self,event):
        param, paramOK = self.setParameter()
        
        if not paramOK or param["POICat"] == "..." or param["POIAttr"] == "..." or param["NetworkType"] == "...":
            addIn.ReportMessage(_("Please select all values!"))
            return
        
        addInParam.SaveParameter(param)
        
        self.OnExit(None)
    
    def setParameter(self):
        param = dict()
        try:
            param["POICat"] = int([i.AttValue("NO") for i in Visum.Net.POICategories.GetAll][self.combo1.GetSelection()])
            param["POIAttr"] = self.combo2.GetValue()
            param["NetworkType"] = ["Link","Node","StopPoint","StopArea","Zone"][self.combo3.GetSelection()]
            param["Replace"] = self.checkbox1.GetValue()
            return param, True
        except:
            addIn.HandleException(_("POI -> Network, value error: "))
            return param, False
        
def CheckPOIs():
    POIs = Visum.Net.POICategories
    if len(POIs) == 0:
        addIn.ReportMessage(_("Current VISUM Version doesn't contain any POI! Please create a POI first!"))
        if not addIn.IsInDebugMode:
            Terminated.set()
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
        if CheckPOIs():
            dialog_1 = MyDialog(None, "")
            app.SetTopWindow(dialog_1)
            dialog_1.ShowModal()
            if addIn.IsInDebugMode:
                app.MainLoop()
    except:
        addIn.HandleException(addIn.TemplateText.MainApplicationError)
        if not addIn.IsInDebugMode:
            Terminated.set()
