#!/usr/bin/env python
import wx
import sys
import os
from VisumPy.AddIn import AddIn, AddInState, AddInParameter
_ = AddIn.gettext

class MyDialog(wx.Dialog):
    def __init__(self, parent, title):
        super(MyDialog, self).__init__(parent, id=-1, title=title, style=wx.CAPTION | wx.STAY_ON_TOP)

        self.__InitUI()
        self.Centre()

    def __InitUI(self):
        self.label1 = wx.StaticText(self, -1, _("Set all AddVals to 0?"))

        self.button1 = wx.Button(self, -1, _("OK"))
        self.button2 = wx.Button(self, -1, _('Cancel'))
        
        self.Bind(wx.EVT_BUTTON, self.OnOK, self.button1)
        self.Bind(wx.EVT_BUTTON, self.OnExit, self.button2)
        self.Bind(wx.EVT_CLOSE, self.OnExit)

        self.__do_layout()
        self.__set_properties()
        
        param = addInParam.Check(False, {"AddVal" : "..."})        
        
    def __set_properties(self):
        self.SetTitle(_("AddVal To Zero"))
        self.button1.SetDefault()
        
        font = self.label1.GetFont()
        font.MakeBold()  # Make the font bold
        self.label1.SetFont(font)
        
    def __do_layout(self):       
        self.empty_cell = (10,10)
        
        vbox = wx.BoxSizer(wx.VERTICAL)
        
        box_1 = wx.GridBagSizer(5,5)
        box_1.Add(self.label1, (0,1), wx.DefaultSpan, wx.EXPAND, 0)
        box_1.Add(self.button1, (2,1), wx.DefaultSpan, wx.LEFT, 0)
        box_1.Add(self.button2, (3,1), wx.DefaultSpan, wx.LEFT, 0)
        box_1.Add(self.empty_cell, (3,2), wx.DefaultSpan, wx.EXPAND, 60)

        vbox.AddSpacer(10)
        vbox.Add(box_1,0,wx.LEFT)
        vbox.AddSpacer(10)
        self.SetSizerAndFit(vbox)

        self.Layout()
        
    def OnExit(self,event):
        if not addIn.IsInDebugMode:
            Terminated.set()
        self.Destroy()
        
    def OnOK(self,event):
        addInParam.SaveParameter({"AddVal" : 1})
        self.OnExit(None)
    
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
