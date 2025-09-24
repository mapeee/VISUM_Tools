#!/usr/bin/env python
import sys
import os
import wx
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
        self.toggle_vals = wx.ToggleButton(self, -1, _("Minutes"), name = "values")
        self.toggle_stops = wx.ToggleButton(self, -1, _(" "), name = "stops")
        self.toggle_vals.Bind(wx.EVT_TOGGLEBUTTON, self._on_toggle)
        self.toggle_stops.Bind(wx.EVT_TOGGLEBUTTON, self._on_toggle)        
        self.button_1 = wx.Button(self, -1, _("+1"), name = "1")
        self.button_5 = wx.Button(self, -1, _("+5"), name = "5")
        self.button_15 = wx.Button(self, -1, _("+15"), name = "15")
        self.button_1m = wx.Button(self, -1, _("-1"), name = "-1")
        self.button_5m = wx.Button(self, -1, _("-5"), name = "-5")
        self.button_15m = wx.Button(self, -1, _("-15"), name = "-15")
        self.button_help = wx.Button(self, -1, _('Help'))
        self.button_info = wx.Button(self, -1, _('Info'))
        self.button_exit = wx.Button(self, wx.ID_CANCEL, _('Close'))

        self.Bind(wx.EVT_BUTTON, self.OnMove, self.button_1)
        self.Bind(wx.EVT_BUTTON, self.OnMove, self.button_5)
        self.Bind(wx.EVT_BUTTON, self.OnMove, self.button_15)
        self.Bind(wx.EVT_BUTTON, self.OnMove, self.button_1m)
        self.Bind(wx.EVT_BUTTON, self.OnMove, self.button_5m)
        self.Bind(wx.EVT_BUTTON, self.OnMove, self.button_15m)
        self.Bind(wx.EVT_BUTTON, self.OnHelp, self.button_help)
        self.Bind(wx.EVT_BUTTON, self.OnInfo, self.button_info)
        self.Bind(wx.EVT_BUTTON, self.OnExit, self.button_exit)
        self.__do_layout()    
        
    def __do_layout(self):
        self.button_1m.SetForegroundColour(wx.Colour(255, 0, 0))
        self.button_5m.SetForegroundColour(wx.Colour(255, 0, 0))
        self.button_15m.SetForegroundColour(wx.Colour(255, 0, 0))
        sb_exe = wx.StaticBox(self, -1)
        sb_exe.SetFont(wx.Font(10, wx.DEFAULT, wx.NORMAL, wx.BOLD))
        sbSizer_exe = wx.StaticBoxSizer(sb_exe, wx.VERTICAL)
        sbSizer_exe.SetMinSize((200, 100))
        gbSizer = wx.GridBagSizer(hgap=5, vgap=5)
        gbSizer.Add(self.button_1, (0, 0), wx.DefaultSpan, wx.ALL, 5)
        gbSizer.Add(self.button_5, (1, 0), wx.DefaultSpan, wx.ALL, 5)
        gbSizer.Add(self.button_15, (2, 0), wx.DefaultSpan, wx.ALL, 5)
        gbSizer.Add(self.button_1m, (0, 1), wx.DefaultSpan, wx.ALL, 5)
        gbSizer.Add(self.button_5m, (1, 1), wx.DefaultSpan, wx.ALL, 5)
        gbSizer.Add(self.button_15m, (2, 1), wx.DefaultSpan, wx.ALL, 5)
        sbSizer_exe.Add(self.toggle_vals, flag = wx.ALIGN_CENTER | wx.TOP, border = 10)
        sbSizer_exe.Add(self.toggle_stops, flag = wx.ALIGN_CENTER | wx.TOP, border = 10)
        sbSizer_exe.Add(gbSizer, flag = wx.ALIGN_CENTER | wx.TOP, border = 10)
        sbSizer_exe.AddSpacer(20)
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
        if button.GetName() == "stops":
            self.toggle_vals.SetLabel(_("Stops"))
            self.toggle_vals.SetValue(True)
        if button.GetValue():
            if button.GetName() == "stops":
                button.SetLabel(_("Stops From"))
            else:
                button.SetLabel(_("Stops"))
                self.toggle_stops.SetLabel(_("Stops From"))
                self.toggle_stops.SetValue(True)
        else:
            if button.GetName() == "stops":
                button.SetLabel(_("Stops To"))
            else:
                button.SetLabel(_("Minutes"))
                self.toggle_stops.SetLabel(_(" "))
                self.toggle_stops.SetValue(False)
        
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
        if not _checkMarking():
            return
        movings = int(event.GetEventObject().GetName())
        for VJ in Visum.Net.Marking.GetAll:
            if self.toggle_vals.GetValue() is False:
                VJ.SetAttValue("DEP", VJ.AttValue("DEP") + (60 * movings))
            elif self.toggle_stops.GetValue() is True:
                if VJ.AttValue("FROMTPROFITEMINDEX") + movings < 1:
                    addIn.ReportMessage(_("Start before Index 1!"))
                    return
                if VJ.AttValue("FROMTPROFITEMINDEX") + movings >= VJ.AttValue("TOTPROFITEMINDEX"):
                    addIn.ReportMessage(_("Start after finish!"))
                    return
                VJ.SetAttValue("FROMTPROFITEMINDEX", VJ.AttValue("FROMTPROFITEMINDEX") + movings)
            else:
                if VJ.AttValue("TOTPROFITEMINDEX") + movings <= VJ.AttValue("FROMTPROFITEMINDEX"):
                    addIn.ReportMessage(_("Finish before start!"))
                    return
                if VJ.AttValue("TOTPROFITEMINDEX") + movings > VJ.AttValue(r"TIMEPROFILE\ENDTIMEPROFILEITEM\INDEX"):
                    addIn.ReportMessage(_("Finish after last stop!"))
                    return
                VJ.SetAttValue("TOTPROFITEMINDEX", VJ.AttValue("TOTPROFITEMINDEX") + movings)
        
        addInParam.SaveParameter(None)

  
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
    if Visum.Net.Marking.Count > 20:
        addIn.ReportMessage(_("To many VehicleJourneys marked!"))
        return False
    if Visum.Net.Marking.Count == 0:
        addIn.ReportMessage(_("No VehicleJourney marked!"))
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
            dialog_1 = MyDialog(None, _("VJ Mover"))
            app.SetTopWindow(dialog_1)
            dialog_1.ShowModal()
            if addIn.IsInDebugMode:
                app.MainLoop()
    except:
        addIn.HandleException(addIn.TemplateText.MainApplicationError)
        if not addIn.IsInDebugMode:
            Terminated.set()
