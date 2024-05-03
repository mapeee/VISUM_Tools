# -*- coding: utf-8 -*-
#!/usr/bin/env python
import wx
import sys
import os
from VisumPy.AddIn import AddIn, AddInState, AddInParameter
_ = AddIn.gettext

class InfoFrame(wx.Frame):
    def __init__(self, title):
        super(InfoFrame, self).__init__(None, id=-1, title=title, style=wx.CAPTION | wx.STAY_ON_TOP, size=(190, 230))
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
        sizer.Add(wx.StaticText(self,label= _("16.02.2024")),0,wx.LEFT,10)
        sizer.AddSpacer(2)
        sizer.Add(wx.StaticText(self,label= _("Version 1.0: Initial")),0,wx.LEFT,10)
        sizer.Add(wx.StaticText(self,label= _("Version 1.1: Set Various")),0,wx.LEFT,10)
        sizer.AddSpacer(10)
        sizer.Add(self.button,0,wx.ALIGN_CENTER,5)
        sizer.AddSpacer(5)
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
        self.label_CheckCon = wx.StaticText(self, -1, _("Check Connectors"))
        self.label_CheckStops = wx.StaticText(self, -1, _("Check Stops"))
        self.label_CheckTSys = wx.StaticText(self, -1, _("Check TSys"))
        
        self.label_SetCon = wx.StaticText(self, -1, _("Set Connector times"))
        self.label_SetPtt = wx.StaticText(self, -1, _("Set Perceived travel times"))
        self.label_SetFS = wx.StaticText(self, -1, _("Set Fare Systems"))
        self.label_SetV = wx.StaticText(self, -1, _("Set Various"))

        self.button2 = wx.Button(self, -1, _('Info'))
        self.button3 = wx.Button(self, -1, _("OK"))
        self.button4 = wx.Button(self, -1, _('Cancel'))
        self.button5 = wx.Button(self, -1, _('Select all'))
        self.button6 = wx.Button(self, -1, _('Deselect all'))
        self.button7 = wx.Button(self, -1, _("Help Checks"), size=(80,20))
        self.button8 = wx.Button(self, -1, _("Help Sets"), size=(110,20))
               
        self.checkbox_CheckCon = wx.CheckBox(self)
        self.checkbox_CheckStops = wx.CheckBox(self)
        self.checkbox_CheckTSys = wx.CheckBox(self)
        
        self.checkbox_SetCon = wx.CheckBox(self)
        self.checkbox_SetPtt = wx.CheckBox(self) 
        self.checkbox_SetFS = wx.CheckBox(self) 
        self.checkbox_SetV = wx.CheckBox(self) 
        
        self.Bind(wx.EVT_BUTTON, self.OnHelp, self.button7)
        self.Bind(wx.EVT_BUTTON, self.OnHelp, self.button8)
        self.Bind(wx.EVT_BUTTON, self.OnInfo, self.button2)
        self.Bind(wx.EVT_BUTTON, self.OnOK, self.button3)
        self.Bind(wx.EVT_BUTTON, self.OnExit, self.button4)
        self.Bind(wx.EVT_CLOSE, self.OnExit)
        self.Bind(wx.EVT_BUTTON, self.OnSelectAll, self.button5)
        self.Bind(wx.EVT_BUTTON, self.OnDeselectAll, self.button6)

        self.__do_layout()
        self.__set_properties()
        
        defaultParam = {"CheckTSys" : False, "CheckCon" : False, "CheckStops" : False,
                        "SetCon" : False, "SetPtt" : False, "SetFareSystem" : False}
        param = addInParam.Check(False, defaultParam)
   
        try: self.__initWxControlValues(param)
        except: pass
        
    def __initWxControlValues(self, param):
        self.checkbox_CheckCon.SetValue(param["CheckCon"])
        self.checkbox_CheckStops.SetValue(param["CheckStops"])
        self.checkbox_CheckTSys.SetValue(param["CheckTSys"])
        self.checkbox_SetCon.SetValue(param["SetCon"])
        self.checkbox_SetPtt.SetValue(param["SetPtt"])
        self.checkbox_SetFS.SetValue(param["SetFareSystem"])
        self.checkbox_SetV.SetValue(param["SetVarious"])
   
    def __set_properties(self):
        self.SetTitle(_("Standi Tools"))
        self.button3.SetDefault()
        
    def __do_layout(self):       
        vbox = wx.BoxSizer(wx.VERTICAL)
        sb1 = wx.StaticBox(self, -1, _("Checking")) 
        sbSizer1 = wx.StaticBoxSizer(sb1, wx.VERTICAL) 
        sb2 = wx.StaticBox(self, -1, _("Set Values")) 
        sbSizer2 = wx.StaticBoxSizer(sb2, wx.VERTICAL) 
        
        sb1.SetFont(wx.Font(9, wx.DEFAULT, wx.NORMAL, wx.BOLD))
        sb2.SetFont(wx.Font(9, wx.DEFAULT, wx.NORMAL, wx.BOLD))
        
        box_1 = wx.GridBagSizer()
        box_1.Add(self.label_CheckTSys, (0,0))
        box_1.Add(self.checkbox_CheckTSys, (0,1), wx.DefaultSpan, wx.LEFT, 5)
        box_1.Add(self.label_CheckCon, (1,0))
        box_1.Add(self.checkbox_CheckCon, (1,1), wx.DefaultSpan, wx.LEFT, 5)
        box_1.Add(self.label_CheckStops, (2,0))
        box_1.Add(self.checkbox_CheckStops, (2,1), wx.DefaultSpan, wx.LEFT, 5)
        box_1.Add(self.button7, (3,0), wx.DefaultSpan, wx.TOP,5)
        sbSizer1.Add(box_1, 0, wx.ALL|wx.CENTER, 5)
        
        box_2 = wx.GridBagSizer()
        box_2.Add(self.label_SetCon, (0,0))
        box_2.Add(self.checkbox_SetCon, (0,1), wx.DefaultSpan, wx.LEFT, 5)
        box_2.Add(self.label_SetFS, (1,0))
        box_2.Add(self.checkbox_SetFS, (1,1), wx.DefaultSpan, wx.LEFT, 5)
        box_2.Add(self.label_SetPtt, (2,0))
        box_2.Add(self.checkbox_SetPtt, (2,1), wx.DefaultSpan, wx.LEFT, 5)
        box_2.Add(self.label_SetV, (3,0))
        box_2.Add(self.checkbox_SetV, (3,1), wx.DefaultSpan, wx.LEFT, 5)
        box_2.Add(self.button8, (4,0), wx.DefaultSpan, wx.TOP,5)        
        sbSizer2.Add(box_2, 0, wx.ALL|wx.CENTER, 5)  

        box_3 = wx.GridBagSizer()
        box_3.Add(self.button5, (0,1), wx.DefaultSpan, wx.EXPAND, 0)
        box_3.Add(self.button6, (0,2), wx.DefaultSpan, wx.EXPAND, 10)
        box_3.Add(self.button2, (1,1), wx.DefaultSpan, wx.TOP|wx.EXPAND, 10)
        box_3.Add(self.button3, (1,2), wx.DefaultSpan, wx.TOP|wx.EXPAND, 10)
        box_3.Add(self.button4, (1,3), wx.DefaultSpan, wx.TOP|wx.EXPAND|wx.RIGHT, 10)

        vbox.Add(sbSizer1,0,wx.LEFT,10)
        vbox.AddSpacer(15)
        vbox.Add(sbSizer2,0,wx.LEFT,10)
        vbox.AddSpacer(15)
        vbox.Add(box_3,0,wx.CENTER,10)
        vbox.AddSpacer(5)
        self.SetSizerAndFit(vbox)

        self.Layout()

    def OnExit(self,event):
        if not addIn.IsInDebugMode:
            Terminated.set()
        self.Destroy()
      
    def OnInfo(self,event):
        title = _("Info")
        frame = InfoFrame(title=title)
      
    def OnHelp(self,event):
        if event.GetEventObject().GetLabel() == _("Help Checks"): path = "HelpStandiTools_Checks.htm"
        if event.GetEventObject().GetLabel() == _("Help Sets"): path = "HelpStandiTools_Sets.htm"
        try:
            os.startfile(addIn.DirectoryPath + _(path))
        except:
            addIn.HandleException()
    
    def OnOK(self,event):
        param, paramOK = self.setParameter()
        
        if not paramOK or True not in [param["SetFareSystem"],param["SetCon"],param["SetPtt"],param["SetVarious"],
                                       param["CheckTSys"],param["CheckCon"], param["CheckStops"]]:
            addIn.ReportMessage(_("Please select at least one procedure!"))
            return
        
        addInParam.SaveParameter(param)
        
        self.OnExit(None)
        
    def OnSelectAll(self,event):
        self.checkbox_CheckCon.SetValue(True)
        self.checkbox_CheckStops.SetValue(True)
        self.checkbox_CheckTSys.SetValue(True)
        self.checkbox_SetCon.SetValue(True)
        self.checkbox_SetPtt.SetValue(True)
        self.checkbox_SetFS.SetValue(True)
        self.checkbox_SetV.SetValue(True)
        
    def OnDeselectAll(self,event):
        self.checkbox_CheckCon.SetValue(False)
        self.checkbox_CheckStops.SetValue(False)
        self.checkbox_CheckTSys.SetValue(False)
        self.checkbox_SetCon.SetValue(False)
        self.checkbox_SetPtt.SetValue(False)
        self.checkbox_SetFS.SetValue(False)
        self.checkbox_SetV.SetValue(False)
    
    def setParameter(self):
        param = dict()
        try:
            param["CheckCon"] = self.checkbox_CheckCon.GetValue()
            param["CheckStops"] = self.checkbox_CheckStops.GetValue()
            param["CheckTSys"] = self.checkbox_CheckTSys.GetValue()
            param["SetCon"] = self.checkbox_SetCon.GetValue()
            param["SetPtt"] = self.checkbox_SetPtt.GetValue()
            param["SetFareSystem"] = self.checkbox_SetFS.GetValue()
            param["SetVarious"] = self.checkbox_SetV.GetValue()
            return param, True
        except:
            addIn.HandleException(_("Standi Tools, value error: "))
            return param, False

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
        dialog_1 = MyDialog(None, "")
        app.SetTopWindow(dialog_1)
        dialog_1.ShowModal()
        if addIn.IsInDebugMode:
            app.MainLoop()
    except:
        addIn.HandleException(addIn.TemplateText.MainApplicationError)
        if not addIn.IsInDebugMode:
            Terminated.set()