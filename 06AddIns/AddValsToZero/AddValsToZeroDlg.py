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
        self.label1 = wx.StaticText(self, -1, _("Set all AddVals to 0?"))

        self.button_ok = wx.Button(self, -1, _("OK"))
        self.button_help = wx.Button(self, -1, _('Help'))
        self.button_exit = wx.Button(self, wx.ID_CANCEL, _('Cancel'))

        self.Bind(wx.EVT_BUTTON, self.OnOK, self.button_ok)
        self.Bind(wx.EVT_BUTTON, self.OnHelp, self.button_help)
        self.Bind(wx.EVT_BUTTON, self.OnExit, self.button_exit)

        self.__do_layout()
        self.__set_properties()
        
    def __set_properties(self):
        font = self.label1.GetFont()
        font.MakeBold()
        self.label1.SetFont(font)
        
    def __do_layout(self):   
        
        sb_exe = wx.StaticBox(self, -1)
        sb_exe.SetFont(wx.Font(10, wx.DEFAULT, wx.NORMAL, wx.BOLD))
        sbSizer_exe = wx.StaticBoxSizer(sb_exe, wx.VERTICAL)
        sbSizer_exe.SetMinSize((200, 100))
        
        sbSizer_exe.Add(self.label1, flag = wx.ALIGN_CENTER | wx.TOP, border = 10)
        sbSizer_exe.Add(self.button_ok, flag = wx.ALIGN_CENTER | wx.TOP, border = 10)
        sbSizer_exe.Add(self.button_help, flag = wx.ALIGN_CENTER | wx.TOP, border = 10)
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
        
    def OnHelp(self,event):
        try:
            os.startfile(addIn.DirectoryPath + _("HelpAddValsToZero.htm"))
        except:
            addIn.HandleException()  
        
    def OnOK(self,event):
        addInParam.SaveParameter({"AddVal" : 1})
        self.OnExit(None)

def CheckNetwork():
    if 0 in [Visum.Net.Nodes.Count,Visum.Net.Links.Count]:
        addIn.ReportMessage(_("Current Visum Version has nothing to set to 0! Create Network elements first!"))
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
            dialog_1 = MyDialog(None, "")
            app.SetTopWindow(dialog_1)
            dialog_1.ShowModal()
            if addIn.IsInDebugMode:
                app.MainLoop()
    except:
        addIn.HandleException(addIn.TemplateText.MainApplicationError)
        if not addIn.IsInDebugMode:
            Terminated.set()
