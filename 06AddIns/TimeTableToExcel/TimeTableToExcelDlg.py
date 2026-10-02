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
        sizer.Add(wx.StaticText(self,label= _("09.10.2025")),0,wx.LEFT,10)
        sizer.AddSpacer(2)
        sizer.Add(wx.StaticText(self,label= _("Version 0.8: beta")),0,wx.LEFT,10)
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
        self.label1 = wx.StaticText(self, -1, _("Active lines: "))
        self.label2 = wx.StaticText(self, -1, str(Visum.Net.Lines.CountActive))
        self.txt_PuTCon = wx.TextCtrl(self, style=wx.TE_READONLY, size =(150,20), value=r"C:\\")
        
        self.button_PuTCon = wx.Button(self, -1, _('Folder'), size=(80,20))
        self.button_export = wx.Button(self, -1, _("Export"))
        self.button_help = wx.Button(self, -1, _('Help'))
        self.button_info = wx.Button(self, -1, _('Info'))
        self.button_exit = wx.Button(self, wx.ID_CANCEL, _('Cancel'))

        self.Bind(wx.EVT_BUTTON, self.OnPuTCon, self.button_PuTCon)
        self.Bind(wx.EVT_BUTTON, self.OnExport, self.button_export)
        self.Bind(wx.EVT_BUTTON, self.OnHelp, self.button_help)
        self.Bind(wx.EVT_BUTTON, self.OnInfo, self.button_info)
        self.Bind(wx.EVT_BUTTON, self.OnExit, self.button_exit)

        self.__do_layout()
        self.__set_properties()
        
        defaultParam = {"PuTConPath" : False}
        param = addInParam.Check(False, defaultParam)
        
    def __set_properties(self):
        font = self.label2.GetFont()
        font.MakeBold()
        self.label2.SetFont(font)
        
    def __do_layout(self):   
        
        sb_exe = wx.StaticBox(self, -1)
        sb_exe.SetFont(wx.Font(10, wx.DEFAULT, wx.NORMAL, wx.BOLD))
        sbSizer_exe = wx.StaticBoxSizer(sb_exe, wx.VERTICAL)
        sbSizer_exe.SetMinSize((200, 100))
        
        sb_PuTCon = wx.StaticBox(self, -1, _("Properties")) 
        sbSizer_PuTCon = wx.StaticBoxSizer(sb_PuTCon, wx.VERTICAL)
        sb_PuTCon.SetFont(wx.Font(8, wx.DEFAULT, wx.NORMAL, wx.BOLD))
        
        box_PuTCon = wx.GridBagSizer()
        box_PuTCon.Add(self.txt_PuTCon, (0,0), wx.DefaultSpan, wx.TOP|wx.ALIGN_CENTER_VERTICAL,5)
        box_PuTCon.Add(self.button_PuTCon, (0,1),  wx.DefaultSpan, wx.TOP|wx.LEFT, 5)
        box_PuTCon.Add((0, 5), (1,0), (1,2))
        box_PuTCon.Add(self.label1, (2,0), wx.DefaultSpan, wx.TOP|wx.ALIGN_RIGHT,5)
        box_PuTCon.Add(self.label2, (2,1), wx.DefaultSpan, wx.TOP|wx.ALIGN_CENTER_VERTICAL,5)
        sbSizer_PuTCon.Add(box_PuTCon, 0, wx.ALL|wx.CENTER, 5) 
        
        sbSizer_exe.Add(sbSizer_PuTCon, 0, wx.ALIGN_CENTER | wx.TOP, 5)
        sbSizer_exe.Add(self.button_export, flag = wx.ALIGN_CENTER | wx.TOP, border = 10)
        sbSizer_exe.AddSpacer(10)
        sbSizer_exe.Add(self.button_help, flag = wx.ALIGN_CENTER | wx.TOP, border = 10)
        sbSizer_exe.Add(self.button_info, flag = wx.ALIGN_CENTER | wx.TOP, border = 10)
        sbSizer_exe.Add(self.button_exit, flag = wx.ALIGN_CENTER | wx.TOP, border = 10)
        
        vbox = wx.BoxSizer(wx.VERTICAL)
    
        vbox.Add(sbSizer_exe, proportion = 1, flag = wx.EXPAND | wx.LEFT | wx.RIGHT, border = 10)
        vbox.AddSpacer(10)

        self.SetSizerAndFit(vbox)
        self.Layout()
        self.Centre()
        
    def OnExit(self,event):
        if not addIn.IsInDebugMode:
            Terminated.set()
        self.Destroy()
        
    def OnExport(self,event):
        param, paramOK = self.setParameter()
        param["PuTConPath"] = self.txt_PuTCon.GetValue()
        if not os.path.isdir(os.path.split(param["PuTConPath"])[0]):
            self.txt_PuTCon.SetForegroundColour(wx.RED)
            addIn.ReportMessage(_("Export-Folder not existing!"))
            return
        addInParam.SaveParameter(param)
        self.OnExit(None)    
        
    def OnHelp(self,event):
        try:
            os.startfile(addIn.DirectoryPath + _("HelpTimeTableToExcel.htm"))
        except:
            addIn.HandleException() 
            
    def OnInfo(self, event):
        title = _("Info")
        frame = InfoFrame(title=title)    
        
    def OnPuTCon(self,event):
        dlg = wx.DirDialog(self, _("Choose your directory to export Timetables"),
                           style=wx.DD_DEFAULT_STYLE | wx.DD_NEW_DIR_BUTTON)
        if dlg.ShowModal() == wx.ID_OK:
            pathindir = dlg.GetPath()
            self.txt_PuTCon.SetValue(pathindir)
            self.txt_PuTCon.SetForegroundColour(wx.BLACK)
        dlg.Destroy()
        
    def setParameter(self):
        param = dict()
        
        try:
            param["PuTConPath"] = False
            return param, True
        except:
            return param, False

def CheckNetwork():
    if Visum.Net.VehicleJourneys.Count == 0:
        addIn.ReportMessage(_("Current Visum Version has no VehicleJourneys!"))
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
        if CheckNetwork():
            dialog_1 = MyDialog(None, _("TimeTable To Excel"))
            app.SetTopWindow(dialog_1)
            dialog_1.ShowModal()
            if addIn.IsInDebugMode:
                app.MainLoop()
    except:
        addIn.HandleException(addIn.TemplateText.MainApplicationError)
        if not addIn.IsInDebugMode:
            Terminated.set()
