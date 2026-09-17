#!/usr/bin/env python
import wx
import sys
import os
import ExcelTemplate as ET
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
        sizer.Add(wx.StaticText(self,label= _("01.09.2026")),0,wx.LEFT,10)
        sizer.AddSpacer(2)
        sizer.Add(wx.StaticText(self,label= _("Version 1.0.")),0,wx.LEFT,10)
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
        self.txt_Folder = wx.TextCtrl(self, style=wx.TE_READONLY, size =(150,20), value=r"c:\\")
        
        self.button_folder = wx.Button(self, -1, _('Folder'))
        self.button_import = wx.Button(self, -1, _("Import"))
        self.button_template = wx.Button(self, -1, _("Template"))
        self.button_help = wx.Button(self, -1, _('Help'))
        self.button_info = wx.Button(self, -1, _('Info'))
        self.button_close = wx.Button(self, wx.ID_CANCEL, _('Close'))

        self.Bind(wx.EVT_BUTTON, self.OnFolder, self.button_folder)
        self.Bind(wx.EVT_BUTTON, self.OnImport, self.button_import)
        self.Bind(wx.EVT_BUTTON, self.OnTemplate, self.button_template)
        self.Bind(wx.EVT_BUTTON, self.OnHelp, self.button_help)
        self.Bind(wx.EVT_BUTTON, self.OnInfo, self.button_info)
        self.Bind(wx.EVT_BUTTON, self.OnExit, self.button_close)

        self.__do_layout()
        self.__set_properties()
        
        defaultParam = {"ExcelFolder" : False}
        param = addInParam.Check(False, defaultParam)
        
    def __set_properties(self):
        self.button_import.SetFont(self.button_import.GetFont().Bold())
        
    def __do_layout(self):   
        sb_exe = wx.StaticBox(self, -1)
        sbSizer_exe = wx.StaticBoxSizer(sb_exe, wx.VERTICAL)
        sbSizer_exe.SetMinSize((200, 100))
                
        box_PuTCon = wx.GridBagSizer()
        box_PuTCon.Add(self.txt_Folder, (0,0), flag=wx.EXPAND | wx.TOP | wx.ALIGN_CENTER_VERTICAL, border = 5)
        box_PuTCon.AddGrowableCol(0)

        box_Import = wx.GridBagSizer()
        box_Import.Add(self.button_folder, (0,0), wx.DefaultSpan, wx.TOP|wx.ALIGN_CENTER_VERTICAL,5)
        box_Import.Add(self.button_template, (0,1), wx.DefaultSpan, wx.TOP|wx.ALIGN_CENTER_VERTICAL,5)
        box_Import.Add(self.button_import, (0,2), wx.DefaultSpan, wx.TOP|wx.ALIGN_CENTER_VERTICAL,5)

        sbSizer_exe.Add(box_PuTCon, 0, wx.EXPAND | wx.TOP, 5)
        sbSizer_exe.Add(box_Import, 0, wx.ALIGN_CENTER | wx.TOP, 5)
        sbSizer_exe.AddSpacer(30)
        sbSizer_exe.Add(self.button_help, flag = wx.ALIGN_CENTER | wx.TOP, border = 5)
        sbSizer_exe.Add(self.button_info, flag = wx.ALIGN_CENTER | wx.TOP, border = 5)
        sbSizer_exe.Add(self.button_close, flag = wx.ALIGN_CENTER | wx.TOP, border = 5)
        
        vbox = wx.BoxSizer(wx.VERTICAL)
    
        vbox.Add(sbSizer_exe, proportion = 1, flag = wx.EXPAND | wx.LEFT | wx.RIGHT, border = 10)
        vbox.AddSpacer(10)

        self.SetSizerAndFit(vbox)
        self.Layout()
        self.Centre()
        
    def OnExit(self, event):
        if not addIn.IsInDebugMode:
            Terminated.set()
        self.Destroy()
        
    def OnFolder(self,event):
        with wx.FileDialog(self, _("Choose your Excel Import File"),
                           wildcard="Excel files (*.xlsx;*.xls)|*.xlsx;*.xls",
                           style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST) as dlg:
            if dlg.ShowModal() == wx.ID_OK:
                self.txt_Folder.SetValue(dlg.GetPath())
                self.txt_Folder.SetForegroundColour(wx.BLACK)
        
    def OnImport(self, event):
        param, paramOK = self.setParameter()
        if not paramOK:
            return
        if param["ExcelFolder"] == r"c:\\":
            addIn.ReportMessage(_("Please select an import file"))
            return
        addInParam.SaveParameter(param)
        self.OnExit(None) 
        
    def OnHelp(self, event):
        try:
            os.startfile(addIn.DirectoryPath + _("HelpExcelToTimeTable.htm"))
        except:
            addIn.HandleException() 
            
    def OnInfo(self, event):
        title = _("Info")
        InfoFrame(title=title)    
        
    def OnTemplate(self, event):
        ET.open(Visum)
        self.OnExit(None) 
        
    def setParameter(self):
        param = dict()   
        try:
            param["ExcelFolder"] = self.txt_Folder.GetValue()
            return param, True
        except:
            return param, False

def CheckNetwork():
    if not Visum.Net.StopPoints.Count:
        addIn.ReportMessage(_("Current Visum Version has no Stops! Create Stops first!"))
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
            dialog_1 = MyDialog(None, _("Excel to TimeTable"))
            app.SetTopWindow(dialog_1)
            dialog_1.ShowModal()
            if addIn.IsInDebugMode:
                app.MainLoop()
    except:
        addIn.HandleException(addIn.TemplateText.MainApplicationError)
        if not addIn.IsInDebugMode:
            Terminated.set()
