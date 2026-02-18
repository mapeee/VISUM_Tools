#!/usr/bin/env python
import wx
import sys
import os
from VisumPy.AddIn import AddIn, AddInState, AddInParameter
_ = AddIn.gettext

class InfoFrame(wx.Frame):
    def __init__(self, title):
        super(InfoFrame, self).__init__(None, id=-1, title=title, style=wx.CAPTION | wx.STAY_ON_TOP, size=(190, 190))
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
        sizer.Add(wx.StaticText(self,label= _("05.02.2026")),0,wx.LEFT,10)
        sizer.AddSpacer(2)
        sizer.Add(wx.StaticText(self,label= _("Version 0.9: Beta")),0,wx.LEFT,10)
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
        super(MyDialog, self).__init__(parent, id=-1, title=title, size=(400, 300), style=wx.CAPTION | wx.STAY_ON_TOP)

        self.__InitUI()

    def __InitUI(self):
        self.labelGebiete = wx.StaticText(self, -1, _("Add territories"))
        self.labelKalender = wx.StaticText(self, -1, _("Add weekly calendar"))
        self.labelFahrzeuge = wx.StaticText(self, -1, _("Add standard vehicles"))
        
        self.cbGebiete = wx.CheckBox(self, -1, label="")
        self.cbKalender = wx.CheckBox(self, -1, label="")
        self.cbFahrzeuge = wx.CheckBox(self, -1, label="")

        self.button_ok = wx.Button(self, -1, _("OK"))
        self.button_help = wx.Button(self, -1, _('Help'))
        self.button_info = wx.Button(self, -1, _('Info'))
        self.button_exit = wx.Button(self, wx.ID_CANCEL, _('Close'))

        self.Bind(wx.EVT_BUTTON, self.OnOK, self.button_ok)
        self.Bind(wx.EVT_BUTTON, self.OnHelp, self.button_help)
        self.Bind(wx.EVT_BUTTON, self.OnInfo, self.button_info)
        self.Bind(wx.EVT_BUTTON, self.OnExit, self.button_exit)

        self.__do_layout()
        self.__set_properties()
        
        defaultParam = {"Gebiete" : True, "Kalender" : True, "Fahrzeuge" : True}
        addInParam.Check(False, defaultParam)
        
    def __set_properties(self):
        font = self.labelGebiete.GetFont()
        font.MakeBold()
        font.SetPointSize(8)
        # self.labelGebiete.SetFont(font)
        # self.labelKalender.SetFont(font)
        # self.labelFahrzeuge.SetFont(font)
        
        self.cbGebiete.SetValue(True)
        self.cbKalender.SetValue(True)
        self.cbFahrzeuge.SetValue(True)
        if Visum.Net.VehicleCombinations.Count > 0:
            self.cbFahrzeuge.SetValue(False)
        
    def __do_layout(self):   
        sb_para = wx.StaticBox(self, -1, _("Parameters"))
        sb_para.SetFont(wx.Font(8, wx.DEFAULT, wx.NORMAL, wx.BOLD))
        sbSizer_para = wx.StaticBoxSizer(sb_para, wx.VERTICAL)
        sbSizer_para.SetMinSize((200, 50))
        grid_para = wx.FlexGridSizer(rows=0, cols=3, hgap=0, vgap=5)
        grid_para.Add(self.labelGebiete, 0, flag = wx.ALIGN_LEFT | wx.ALIGN_CENTER_VERTICAL)
        grid_para.Add((20,1))
        grid_para.Add(self.cbGebiete, 0)
        grid_para.Add(self.labelKalender, 0, flag = wx.ALIGN_LEFT | wx.ALIGN_CENTER_VERTICAL)
        grid_para.Add((20,1))
        grid_para.Add(self.cbKalender, 0)
        grid_para.Add(self.labelFahrzeuge, 0, flag = wx.ALIGN_LEFT | wx.ALIGN_CENTER_VERTICAL)
        grid_para.Add((20,1))
        grid_para.Add(self.cbFahrzeuge, 0)
        sbSizer_para.Add(grid_para, 1, wx.ALL | wx.ALIGN_CENTER, 10)
        
        sb_end = wx.StaticBox(self, -1, _("End"))
        sb_end.SetFont(wx.Font(8, wx.DEFAULT, wx.NORMAL, wx.BOLD))
        sbSizer_end = wx.StaticBoxSizer(sb_end, wx.VERTICAL)
        sbSizer_end.SetMinSize((200, 50))
        grid_end = wx.FlexGridSizer(rows=2, cols=2, hgap=5, vgap=5)
        grid_end.Add(self.button_ok, flag = wx.EXPAND)
        grid_end.Add(self.button_exit, flag = wx.EXPAND)
        grid_end.Add(self.button_help, flag = wx.EXPAND)
        grid_end.Add(self.button_info, flag = wx.EXPAND)
        sbSizer_end.Add(grid_end, 1, wx.ALL | wx.ALIGN_CENTER, 10)
        
        vbox = wx.BoxSizer(wx.VERTICAL)
        vbox.AddSpacer(10)
        vbox.Add(sbSizer_para, proportion = 0, flag = wx.EXPAND | wx.LEFT | wx.RIGHT, border = 10)
        vbox.AddSpacer(10)
        vbox.Add(sbSizer_end, proportion = 0, flag = wx.EXPAND | wx.LEFT | wx.RIGHT, border = 10)

        self.SetSizerAndFit(vbox)
        self.Layout()
        self.Centre()
        
    def OnExit(self,event):
        if not addIn.IsInDebugMode:
            Terminated.set()
        self.Destroy()
        
    def OnHelp(self,event):
        try:
            os.startfile(addIn.DirectoryPath + _("HelpErstelleTNMBasis.htm"))
        except:
            addIn.HandleException() 
            
    def OnInfo(self, event):
        title = _("Info")
        frame = InfoFrame(title=title)    
        
    def OnOK(self,event):
        param, paramOK = self.setParameter()
        if not paramOK:
            addIn.ReportMessage(_("Some error occurred."))
            return
        if param["Fahrzeuge"]:
            if not any(col2 == "Bus" for _, col2 in Visum.Net.TSystems.GetMultiAttValues("CODE")):
                addIn.ReportMessage(_("TSystem[CODE] 'Bus': missing"))
                return
        addInParam.SaveParameter(param)
        self.OnExit(None)
        
    def setParameter(self):
        param = dict()
        try:
            param["Gebiete"] = self.cbGebiete.GetValue()
            param["Kalender"] = self.cbKalender.GetValue()
            param["Fahrzeuge"] = self.cbFahrzeuge.GetValue()
            return param, True
        except:
            addIn.HandleException(_("Create TNM-Elements, value error: "))
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

        dialog_1 = MyDialog(None, _("Create TNM-Elements"))
        app.SetTopWindow(dialog_1)
        dialog_1.ShowModal()
        if addIn.IsInDebugMode:
            app.MainLoop()
    except:
        addIn.HandleException(addIn.TemplateText.MainApplicationError)
        if not addIn.IsInDebugMode:
            Terminated.set()
