# -*- coding: utf-8 -*-
#!/usr/bin/env python
import wx
import sys
import os
from VisumPy.AddIn import AddIn, AddInState, AddInParameter
_ = AddIn.gettext
# begin wxGlade: extracode
# end wxGlade

class InfoFrame(wx.Frame):
    def __init__(self, title):
        super(InfoFrame, self).__init__(None, id=-1, title=title, style=wx.CAPTION | wx.STAY_ON_TOP, size=(170, 200))
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
        sizer.Add(wx.StaticText(self,label= _("09.02.2024")),0,wx.LEFT,10)
        sizer.AddSpacer(2)
        sizer.Add(wx.StaticText(self,label= _("Version 1.0: Initial")),0,wx.LEFT,10)
        sizer.Add(wx.StaticText(self,label= _("Version 1.1: Deselect all")),0,wx.LEFT,10)
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
        self.label1 = wx.StaticText(self, -1, _("Matrices - Base case"))
        self.label2 = wx.StaticText(self, -1, _("Matrices - Planning case"))
        self.label3 = wx.StaticText(self, -1, _("Matrices - Shift"))
        self.label4 = wx.StaticText(self, -1, _("Matrices - Formular"))
        self.label5 = wx.StaticText(self, -1, _("Replace existing matrices?"))
        
        self.button1 = wx.Button(self, -1, _("Help"))
        self.button2 = wx.Button(self, -1, _('Info'))
        self.button3 = wx.Button(self, -1, _("OK"))
        self.button4 = wx.Button(self, -1, _('Cancel'))
        self.button5 = wx.Button(self, -1, _('Select all'))
        self.button6 = wx.Button(self, -1, _('Deselect all'))
        
        matrices_b = [_("Pkw_E_O : Car trips adults base case : 1"),
                      _("Lkw_GV : Truck rides : 3"),
                      _("OEV_E_O : PT rides adults base case : 10"),
                      _("OEV_S : PT rides students : 12"),
                      _("MIV_Widerstand_O : Car impedance base case : 35"),
                      _("OEV_Widerstand_O : PT impedance base case : 37"),
                      _("BH_Tag_O : PT headway base case : 39"),
                      _("BH_Spitze_O : PT headway peak hour base case : 41"),
                      _("OEV_Reisezeit_O : PT travel time base case : 43")]
        matrices_p = [_("Pkw_E_M : Car trips adults planning case : 2"),
                      _("OEV_E_M : PT rides adults planning case : 11"),
                      _("MIV_Widerstand_M : Car impedance planning case : 36"),
                      _("OEV_Widerstand_M : PT impedance planning case : 38"),
                      _("BH_Tag_M : PT headway planning case : 40"),
                      _("BH_Spitze_M : PT headway peak hour planning case : 42"),
                      _("OEV_Reisezeit_M : PT travel time planning case : 44"),
                      _("OEV_BefW_M : PT In-vehicle distance planning case : 60")]
        matrices_s = [_("verlOEV : Changes to PT rides : 20"),
                      _("indOEV : Induced PT rides : 21"),
                      _("verlMIV : Changes to MT rides : 22"),
                      _("verlPkw : Changes to Car trips : 23"),
                      _("indMIV : Induced MT rides : 24"),
                      _("indPkw : Induced Car trips : 25")]
        matrices_f = [_("MIV_E_O : MT rides base case : 4"),
                      _("MIV_E_M : MT rides planning case : 5"),
                      _("OEV_E_WidDiff : PT impedance difference adults : 61"),
                      _("OEV_S_WidDiff : PT impedance difference students : 62")]
        self.checkList1 = wx.CheckListBox(self, -1, choices=[_(i) for i in matrices_b])
        self.checkList2 = wx.CheckListBox(self, -1, choices=[_(i) for i in matrices_p])        
        self.checkList3 = wx.CheckListBox(self, -1, choices=[_(i) for i in matrices_s])
        self.checkList4 = wx.CheckListBox(self, -1, choices=[_(i) for i in matrices_f])
        for i in range(len(matrices_b)): self.checkList1.SetItemBackgroundColour(i,wx.Colour(220,220,220))
        for i in range(len(matrices_p)): self.checkList2.SetItemBackgroundColour(i,wx.Colour(193,250,193))
        for i in range(len(matrices_s)): self.checkList3.SetItemBackgroundColour(i,wx.Colour(250,227,240))
        
        self.checkbox1 = wx.CheckBox(self) 
        
        self.Bind(wx.EVT_BUTTON, self.OnHelp, self.button1)
        self.Bind(wx.EVT_BUTTON, self.OnInfo, self.button2)
        self.Bind(wx.EVT_BUTTON, self.OnOK, self.button3)
        self.Bind(wx.EVT_BUTTON, self.OnExit, self.button4)
        self.Bind(wx.EVT_CLOSE, self.OnExit)
        self.Bind(wx.EVT_BUTTON, self.OnSelectAll, self.button5)
        self.Bind(wx.EVT_BUTTON, self.OnDeselectAll, self.button6)

        self.__do_layout()
        self.__set_properties()
        
        defaultParam = {"Matrices_b" : "...", "Items_b" : "...", "Matrices_p" : "...",
                        "Items_p" : "...", "Matrices_s" : "...", "Items_s" : "...",
                        "Matrices_f" : "...", "Items_f" : "...", "Replace" : False}
        param = addInParam.Check(False, defaultParam)
   
        try: self.__initWxControlValues(param)
        except: pass
        
    def __initWxControlValues(self, param):
        self.checkList1.SetCheckedItems(param["Items_b"])
        self.checkList2.SetCheckedItems(param["Items_p"])
        self.checkList3.SetCheckedItems(param["Items_s"])
        self.checkList4.SetCheckedItems(param["Items_f"])
        self.checkbox1.SetValue(param["Replace"])
   
    def __set_properties(self):
        self.SetTitle(_("Create Standi Matrices"))
        self.button2.SetDefault()
        
    def __do_layout(self):       
        vbox = wx.BoxSizer(wx.VERTICAL)
        
        box_1 = wx.GridBagSizer()
        box_1.Add(self.label1, (1,1), wx.DefaultSpan, wx.ALIGN_CENTER_VERTICAL, 10)
        box_1.Add(self.checkList1, (1,3), (1,3), wx.EXPAND, 10)
        box_1.Add(self.label2, (3,1), wx.DefaultSpan, wx.ALIGN_CENTER_VERTICAL, 10)
        box_1.Add(self.checkList2, (3,3), (1,3), wx.EXPAND, 10)
        box_1.Add(self.label3, (5,1), wx.DefaultSpan, wx.ALIGN_CENTER_VERTICAL, 10)
        box_1.Add(self.checkList3, (5,3), (1,3), wx.EXPAND, 10)
        box_1.Add(self.label4, (7,1), wx.DefaultSpan, wx.ALIGN_CENTER_VERTICAL, 10)
        box_1.Add(self.checkList4, (7,3), (1,3), wx.EXPAND, 10)
        box_1.Add(self.label5, (9,1), wx.DefaultSpan, wx.ALIGN_CENTER_VERTICAL, 10)
        box_1.Add(self.checkbox1, (9,3), wx.DefaultSpan, wx.ALIGN_CENTER_VERTICAL, 10)

        box_2 = wx.GridBagSizer()
        box_2.Add(self.button1, (0,1), wx.DefaultSpan, wx.EXPAND, 0)
        box_2.Add(self.button2, (0,2), wx.DefaultSpan, wx.EXPAND, 0)
        box_2.Add(self.button3, (0,4), wx.DefaultSpan, wx.EXPAND, 0)
        box_2.Add(self.button4, (0,5), wx.DefaultSpan, wx.EXPAND, 10)
        box_2.Add(self.button5, (0,7), wx.DefaultSpan, wx.RIGHT, 0)
        box_2.Add(self.button6, (0,8), wx.DefaultSpan, wx.RIGHT, 10)

        vbox.Add(box_1,0,wx.RIGHT,10)
        vbox.AddSpacer(30)
        vbox.Add(box_2,0,wx.RIGHT,10)
        vbox.AddSpacer(20)
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
        try:
            os.startfile(addIn.DirectoryPath + _("HelpStandiMatrices.htm"))
        except:
            addIn.HandleException()
    
    def OnOK(self,event):
        param, paramOK = self.setParameter()
        
        if not paramOK or len(param["Matrices_b"]+param["Matrices_p"]+param["Matrices_s"]+param["Matrices_f"]) == 0:
            addIn.ReportMessage(_("Please select one matrice!"))
            return
        
        # wx.MessageBox(_('finished!'), 'Dialog', wx.OK | wx.ICON_INFORMATION)
        addInParam.SaveParameter(param)
        
        self.OnExit(None)
        
    def OnSelectAll(self,event):
        self.checkList1.SetCheckedItems(range(self.checkList1.GetCount()))
        self.checkList2.SetCheckedItems(range(self.checkList2.GetCount()))
        self.checkList3.SetCheckedItems(range(self.checkList3.GetCount()))
        self.checkList4.SetCheckedItems(range(self.checkList4.GetCount()))
        
    def OnDeselectAll(self,event):
        for i in range(self.checkList1.GetCount()):self.checkList1.Check(i,False)
        for i in range(self.checkList2.GetCount()):self.checkList2.Check(i,False)
        for i in range(self.checkList3.GetCount()):self.checkList3.Check(i,False)
        for i in range(self.checkList4.GetCount()):self.checkList4.Check(i,False)
    
    def setParameter(self):
        param = dict()
        try:
            param["Matrices_b"] = self.checkList1.GetCheckedStrings()
            param["Items_b"] = self.checkList1.GetCheckedItems()
            param["Matrices_p"] = self.checkList2.GetCheckedStrings()
            param["Items_p"] = self.checkList2.GetCheckedItems()
            param["Matrices_s"] = self.checkList3.GetCheckedStrings()
            param["Items_s"] = self.checkList3.GetCheckedItems()
            param["Matrices_f"] = self.checkList4.GetCheckedStrings()
            param["Items_f"] = self.checkList4.GetCheckedItems()
            param["Replace"] = self.checkbox1.GetValue()
            return param, True
        except:
            addIn.HandleException(_("Standi Matrices, value error: "))
            return param, False
        
def CheckZones():
    Zones = Visum.Net.Zones
    if len(Zones) == 0:
        addIn.ReportMessage(_("Current VISUM Version doesn't contain any Zone! Please create Zones first!"))
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
        if CheckZones():
            dialog_1 = MyDialog(None, "")
            app.SetTopWindow(dialog_1)
            dialog_1.ShowModal()
            if addIn.IsInDebugMode:
                app.MainLoop()
    except:
        addIn.HandleException(addIn.TemplateText.MainApplicationError)
        if not addIn.IsInDebugMode:
            Terminated.set()