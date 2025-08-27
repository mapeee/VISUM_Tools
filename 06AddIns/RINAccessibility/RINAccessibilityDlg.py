#!/usr/bin/env python
import os
import sys
import wx
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
        sizer.Add(wx.StaticText(self,label= _("03.05.2024")),0,wx.LEFT,10)
        sizer.AddSpacer(2)
        sizer.Add(wx.StaticText(self,label= _("Version 1.0: Initial")),0,wx.LEFT,10)
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
        super().__init__(parent, title=_("RIN Accessibility (beta)"), style=wx.CAPTION | wx.STAY_ON_TOP)

        self.__InitUI()
        self.Centre()

    def __InitUI(self):
        buttons_exe = {
            "PT SAQ": wx.Button(self, label=_("PT SAQ"), name="PT_SAQ")}
        
        buttons_con = {
            "Cancel": wx.Button(self, label=_('Cancel')),
            "Help": wx.Button(self, label=_("Help")),
            "Info": wx.Button(self, label=_('Info'))}
        
        buttons_exe["PT SAQ"].Bind(wx.EVT_BUTTON, self.OnSAQ)
        buttons_con["Help"].Bind(wx.EVT_BUTTON, self.OnHelp)
        buttons_con["Info"].Bind(wx.EVT_BUTTON, self.OnInfo)
        buttons_con["Cancel"].Bind(wx.EVT_BUTTON, self.OnExit)
        self.Bind(wx.EVT_CLOSE, self.OnExit)
        
        grid = wx.GridSizer(1, 3, 10, 10)
        for button in buttons_con.values():
            grid.Add(button, 0, wx.EXPAND)
        
        vbox = wx.BoxSizer(wx.VERTICAL)
        for button in buttons_exe.values():
            vbox.Add(button, 0, wx.ALL, 10)
        vbox.AddSpacer(15)
        vbox.Add(grid, 0, wx.ALL, 10)
        vbox.AddSpacer(15)
        
        self.SetSizerAndFit(vbox)
    
    def OnExit(self,event):
        if not addIn.IsInDebugMode:
            Terminated.set()
        self.Destroy()
        
    def OnInfo(self,event):
        InfoFrame(title=_("Info"))
      
    def OnHelp(self,event):
        try:
            os.startfile(addIn.DirectoryPath + _("HelpRINAccessibility.htm"))
        except:
            addIn.HandleException()
    
    def OnSAQ(self,event):
        param, paramOK = self.setParameter()
        param["Proce"] = event.GetEventObject().GetName()
        if not paramOK:
            addIn.ReportMessage(_("Please select at least one procedure!"))
            return

        addInParam.SaveParameter(param)
        
        self.OnExit(None)
    
    def setParameter(self):
        param = dict()
        try:
            if Visum.Net.Marking.Count > 0: param["Marking"] = [i.AttValue("NO") for i in Visum.Net.Marking.GetAll]
            else: param["Marking"] = []
            return param, True
        except:
            addIn.HandleException(_("RIN Accessibility, value error!"))
            return param, False

def CheckLines():
    StopAreas = Visum.Net.Lines
    if len(StopAreas) == 0:
        addIn.ReportMessage(_("Current VISUM Version doesn't contain any Line! Please create a Line first!"))
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
        if CheckLines():
            dialog = MyDialog(None, "")
            app.SetTopWindow(dialog)
            dialog.ShowModal()
            if addIn.IsInDebugMode:
                app.MainLoop()
    except:
        addIn.HandleException(addIn.TemplateText.MainApplicationError)
        if not addIn.IsInDebugMode:
            Terminated.set()
