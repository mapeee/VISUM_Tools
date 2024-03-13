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
        self.label1 = wx.StaticText(self, -1, _("Set existing to 0?"))

        self.button1 = wx.Button(self, -1, _("Help"))
        self.button2 = wx.Button(self, -1, _('Info'))
        self.button3 = wx.Button(self, -1, _("OK"))
        self.button4 = wx.Button(self, -1, _('Cancel'))
        self.button5 = wx.Button(self, -1, _('Select all'))
        self.button6 = wx.Button(self, -1, _('Deselect all'))
        
        UDA_m = [_("TSYS : SCHIENENBONUS : Rail Bonus"),
                      _("ZONE : LAGE : Location"),
                      _("ZONE : PARKEN : Parking Impedance"),
                      _("TURN : GRUNDBELASTUNG : Base volume"),
                      _("LINK : BUSSPUR : Bus lane"),
                      _("LINK : GRUNDBELASTUNG : Base volume"),
                      _("LINK : KLASSE : Track class"),
                      _("STOPPOINT : RAUSSTATTUNG : Penalty for equipment"),
                      _("LINEROUTEITEM : ABSAUFSCHLAG : absIncrease"),
                      _("LINEROUTEITEM : RELAUFSCHLAG : relIncrease")]
        UDA_a = [_("LINK : BEL_FV_MIT : Vol Long-distance PT plannig case"),
                      _("LINK : BEL_FV_OHNE : Vol Long-distance PT base case"),
                      _("LINK : BEL_MIV_MIT : Vol MT plannig case"),
                      _("LINK : BEL_MIV_OHNE : Vol MT base case"),
                      _("LINK : BEL_OEV_MIT : Vol PT planning case"),
                      _("LINK : BEL_OEV_OHNE : Vol PT base case")]
        self.checkList1 = wx.CheckListBox(self, -1, choices=[_(i) for i in UDA_m])
        self.checkList2 = wx.CheckListBox(self, -1, choices=[_(i) for i in UDA_a])        
        for i in range(len(UDA_m)): self.checkList1.SetItemBackgroundColour(i,[wx.Colour(255,227,221),
                                                                               wx.Colour(240,240,220),wx.Colour(240,240,220),
                                                                               wx.Colour(245,245,245),
                                                                               wx.Colour(193,250,193),
                                                                               wx.Colour(193,250,193),wx.Colour(193,250,193),
                                                                               wx.Colour(252,223,255),
                                                                               wx.Colour(216,228,255),wx.Colour(216,228,255)][i])
        for i in range(len(UDA_a)): self.checkList2.SetItemBackgroundColour(i,wx.Colour(193,250,193))
        
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
        
        defaultParam = {"UDA_m" : "...", "Items_m" : "...", "UDA_a" : "...",
                        "Items_a" : "...", "Replace" : False}
        param = addInParam.Check(False, defaultParam)
   
        try: self.__initWxControlValues(param)
        except: pass
        
    def __initWxControlValues(self, param):
        self.checkList1.SetCheckedItems(param["Items_m"])
        self.checkList2.SetCheckedItems(param["Items_a"])
        self.checkbox1.SetValue(param["Replace"])
   
    def __set_properties(self):
        self.SetTitle(_("Create Standi-UDA"))
        self.button2.SetDefault()
        
    def __do_layout(self):       
        vbox = wx.BoxSizer(wx.VERTICAL)
        sb1 = wx.StaticBox(self, -1, _("UDA - Manually")) 
        sbSizer1 = wx.StaticBoxSizer(sb1, wx.VERTICAL) 
        sb1.SetFont(wx.Font(9, wx.DEFAULT, wx.NORMAL, wx.BOLD))

        box_1 = wx.GridBagSizer()
        box_1.Add(self.checkList1, (0,1), (0,2), wx.EXPAND, 5)
        sbSizer1.Add(box_1, 0, wx.ALL|wx.CENTER, 5)
        
        sb2 = wx.StaticBox(self, -1, _("UDA - Automatically")) 
        sbSizer2 = wx.StaticBoxSizer(sb2, wx.VERTICAL)
        sb2.SetFont(wx.Font(9, wx.DEFAULT, wx.NORMAL, wx.BOLD))
        
        box_2 = wx.GridBagSizer()
        box_2.Add(self.checkList2, (0,1), (0,2))
        box_2.Add(self.label1, (1,1), (0,0), wx.EXPAND|wx.TOP, 5)
        box_2.Add(self.checkbox1, (1,2), (0,0), wx.LEFT|wx.TOP, 5)
        sbSizer2.Add(box_2, 0, wx.ALL|wx.CENTER, 5)  

        box_3 = wx.GridBagSizer()
        box_3.Add(self.button1, (0,1), wx.DefaultSpan, wx.EXPAND, 0)
        box_3.Add(self.button2, (0,2), wx.DefaultSpan, wx.EXPAND, 0)
        box_3.Add(self.button3, (0,4), wx.DefaultSpan, wx.EXPAND, 0)
        box_3.Add(self.button4, (0,5), wx.DefaultSpan, wx.EXPAND, 10)
        box_3.Add(self.button5, (0,7), wx.DefaultSpan, wx.RIGHT, 0)
        box_3.Add(self.button6, (0,8), wx.DefaultSpan, wx.RIGHT, 10)

        vbox.Add(sbSizer1,1,wx.ALL|wx.EXPAND,10)
        vbox.AddSpacer(10)
        vbox.Add(sbSizer2,0,wx.ALL|wx.EXPAND,10)
        vbox.AddSpacer(10)
        vbox.Add(box_3)
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
            os.startfile(addIn.DirectoryPath + _("HelpStandiUDA.htm"))
        except:
            addIn.HandleException()
    
    def OnOK(self,event):
        param, paramOK = self.setParameter()
        
        if len(param["UDA_m"]) > 0:
            wx.MessageBox(_("Please fill values into all Manually UDA!"), _("Attention!") ,wx.OK | wx.ICON_INFORMATION)
        
        if not paramOK or len(param["UDA_m"]+param["UDA_a"]) == 0:
            addIn.ReportMessage(_("Please select one UDA!"))
            return
        
        addInParam.SaveParameter(param)
        
        self.OnExit(None)
        
    def OnSelectAll(self,event):
        self.checkList1.SetCheckedItems(range(self.checkList1.GetCount()))
        self.checkList2.SetCheckedItems(range(self.checkList2.GetCount()))
        
    def OnDeselectAll(self,event):
        for i in range(self.checkList1.GetCount()):self.checkList1.Check(i,False)
        for i in range(self.checkList2.GetCount()):self.checkList2.Check(i,False)
    
    def setParameter(self):
        param = dict()
        try:
            param["UDA_m"] = self.checkList1.GetCheckedStrings()
            param["Items_m"] = self.checkList1.GetCheckedItems()
            param["UDA_a"] = self.checkList2.GetCheckedStrings()
            param["Items_a"] = self.checkList2.GetCheckedItems()
            param["Replace"] = self.checkbox1.GetValue()
            return param, True
        except:
            addIn.HandleException(_("Standi-UDA, value error: "))
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