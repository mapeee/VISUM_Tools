# -*- coding: utf-8 -*-
#!/usr/bin/env python
import wx
import sys
import os
from VisumPy.AddIn import AddIn, AddInState, AddInParameter
_ = AddIn.gettext

class InfoFrame(wx.Frame):
    def __init__(self, title):
        super(InfoFrame, self).__init__(None, id=-1, title=title, style=wx.CAPTION | wx.STAY_ON_TOP, size=(170, 200))
        self.Centre()
        
        img = wx.Image(addIn.DirectoryPath +'logo.png',wx.BITMAP_TYPE_ANY)
        img = img.Scale(55,30,wx.IMAGE_QUALITY_BOX_AVERAGE)
        img = img.ConvertToBitmap()
        png = wx.StaticBitmap(self, -1, img, (0, 0))
        
        self.button = wx.Button(self, -1, _("Close"))
        self.Bind(wx.EVT_BUTTON, self.__OnOK, self.button)
        self.SetBackgroundColour(wx.Colour(wx.NullColour))
        
        self.label = wx.StaticText(self,label= _("hvv GmbH"))
        
        sizer = wx.BoxSizer(wx.VERTICAL)
        sizer.Add(png,0,wx.LEFT,5)
        sizer.AddSpacer(10)
        sizer.Add(self.label,0,wx.LEFT,10)
        sizer.AddSpacer(2)
        sizer.Add(wx.StaticText(self,label= _("Marcus Peter")),0,wx.LEFT,10)
        sizer.Add(wx.StaticText(self,label= _("21.02.2024")),0,wx.LEFT,10)
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
        self.label_PrT_bc = wx.StaticText(self, -1, _("PrT"))
        self.label_PuT_bc = wx.StaticText(self, -1, _("PuT"))
        self.label_Val_bc = wx.StaticText(self, -1, _("Base case values"))
        self.label_Del_bc = wx.StaticText(self, -1, _("Delete interim steps"))
        self.label_PrT_pc = wx.StaticText(self, -1, _("PrT"))
        self.label_PrT_pcput = wx.StaticText(self, -1, _("PrT (+PT assignment)"))
        self.label_PuT_pc = wx.StaticText(self, -1, _("PuT"))
        self.label_PuT_pcprt = wx.StaticText(self, -1, _("PuT (+PrT iteration)"))
        self.label_Val_pc = wx.StaticText(self, -1, _("Planning case values"))
        self.label_Del_pc = wx.StaticText(self, -1, _("Delete interim steps"))
        self.label_PuTCon = wx.StaticText(self, -1, _("Folder to Import/Export"))

        self.button_GoVal_bc = wx.Button(self, -1, _("Go"), size=(40,17), name="GoVal_bc")
        self.button_GoDel_bc = wx.Button(self, -1, _("Go"), size=(40,17))
        self.button_GoVal_pc = wx.Button(self, -1, _("Go"), size=(40,17), name="GoVal_pc")
        self.button_GoDel_pc = wx.Button(self, -1, _("Go"), size=(40,17))
        self.button_SelectAll_bc = wx.Button(self, -1, _('Select all'))
        self.button_DeselectAll_bc = wx.Button(self, -1, _('Deselect all'))
        self.button_PuTCon = wx.Button(self, -1, _('Folder'), size=(80,20))

        self.button_GoPuT_op_bc = wx.Button(self, -1, _("PuT Service base case"), size=(180,30), name="GoPuT_op_bc")
        self.button_GoPuT_op_pc = wx.Button(self, -1, _("PuT Service planning case"), size=(180,30), name="GoPuT_op_pc")
        self.button_GoDel_NKV = wx.Button(self, -1, _("Delete components"), size=(135,40), name="GoDel_NKV") 
        self.button_GoImport_NKV = wx.Button(self, -1, _("Import components"), size=(135,40), name="GoImport_NKV")
        self.button_GoNKV = wx.Button(self, -1, _("Calculate NKV"), size=(180,30), name="GoNKV")
        
        self.button_ExeProc = wx.Button(self, -1, _("Execute procedures"), size=(130,25), name="ExeProc")
        self.button_ExePre = wx.Button(self, -1, _("Preset procedures"), size=(145,25), name="ExePre")
        self.button_Cancel = wx.Button(self, -1, _('Cancel'))
        self.button_Info = wx.Button(self, -1, _('Info'))
        self.button_Help = wx.Button(self, -1, _("Help"))
               
        self.checkbox_PrT_bc = wx.CheckBox(self, name="PrT_bc")
        self.checkbox_PuT_bc = wx.CheckBox(self, name="PuT_bc")
        self.checkbox_Val_bc = wx.CheckBox(self, name="Val_bc")
        self.checkbox_Del_bc = wx.CheckBox(self, name="Del_bc")
        self.checkbox_PrT_pc = wx.CheckBox(self, name="PrT_pc")
        self.checkbox_PrT_pcput = wx.CheckBox(self, name="PrT_pcput") 
        self.checkbox_PuT_pc = wx.CheckBox(self, name="PuT_pc")
        self.checkbox_PuT_pcprt = wx.CheckBox(self, name="PuT_pcprt")
        self.checkbox_Val_pc = wx.CheckBox(self, name="Val_pc")
        self.checkbox_Del_pc = wx.CheckBox(self, name="Del_pc")
        
        self.txt_PuTCon = wx.TextCtrl(self,style=wx.TE_READONLY,size =(250,20), value=r"C:\Daten\Verbindungen\verb.con")
        
        
        self.Bind(wx.EVT_BUTTON, self.OnSelectAll_bc, self.button_SelectAll_bc)
        self.Bind(wx.EVT_BUTTON, self.OnDeselectAll_bc, self.button_DeselectAll_bc)
        self.Bind(wx.EVT_BUTTON, self.OnCalVal, self.button_GoVal_bc)
        self.Bind(wx.EVT_BUTTON, self.OnCalVal, self.button_GoVal_pc)
        self.Bind(wx.EVT_BUTTON, self.OnDelSteps, self.button_GoDel_bc)
        self.Bind(wx.EVT_BUTTON, self.OnDelSteps, self.button_GoDel_pc)
        
        self.Bind(wx.EVT_BUTTON, self.OnNKV, self.button_GoPuT_op_bc)
        self.Bind(wx.EVT_BUTTON, self.OnNKV, self.button_GoPuT_op_pc)
        self.Bind(wx.EVT_BUTTON, self.OnNKV, self.button_GoDel_NKV)
        self.Bind(wx.EVT_BUTTON, self.OnNKV, self.button_GoImport_NKV)
        self.Bind(wx.EVT_BUTTON, self.OnNKV, self.button_GoNKV)
        
        self.Bind(wx.EVT_BUTTON, self.OnPuTCon, self.button_PuTCon)
        
        self.Bind(wx.EVT_BUTTON, self.OnHelp, self.button_Help)
        self.Bind(wx.EVT_BUTTON, self.OnInfo, self.button_Info)
        self.Bind(wx.EVT_BUTTON, self.OnExeProc, self.button_ExeProc)
        self.Bind(wx.EVT_BUTTON, self.OnExeProc, self.button_ExePre)
        self.Bind(wx.EVT_BUTTON, self.OnExit, self.button_Cancel)
        self.Bind(wx.EVT_CLOSE, self.OnExit)

        self.Bind(wx.EVT_CHECKBOX, self.OnCheck, self.checkbox_PrT_bc)
        self.Bind(wx.EVT_CHECKBOX, self.OnCheck, self.checkbox_PuT_bc)
        self.Bind(wx.EVT_CHECKBOX, self.OnCheck, self.checkbox_Val_bc)
        self.Bind(wx.EVT_CHECKBOX, self.OnCheck, self.checkbox_Del_bc)
        
        self.Bind(wx.EVT_CHECKBOX, self.OnCheck, self.checkbox_PrT_pc)
        self.Bind(wx.EVT_CHECKBOX, self.OnCheck, self.checkbox_PrT_pcput)
        self.Bind(wx.EVT_CHECKBOX, self.OnCheck, self.checkbox_PuT_pc)
        self.Bind(wx.EVT_CHECKBOX, self.OnCheck, self.checkbox_PuT_pcprt)
        self.Bind(wx.EVT_CHECKBOX, self.OnCheck, self.checkbox_Val_pc)
        self.Bind(wx.EVT_CHECKBOX, self.OnCheck, self.checkbox_Del_pc)

        self.__do_layout()
        self.__set_properties()
        
        defaultParam = {"PrT_bc" : False, "PuT_bc" : False, "Val_bc" : False, "Del" : False, "ExeProc" : False,
                        "PrT_pc" : False, "PrT_pcput" : False, "PuT_pc" : False, "PuT_pcprt" : False, "Val_pc" : False,
                        "PuTCon" : False, "PuTConPath" : False,
                        "NKV" : False}
        param = addInParam.Check(False, defaultParam)
   
    def __set_properties(self):
        self.SetTitle(_("Standi Procedures"))
        
    def __do_layout(self):         
        self.button_ExeProc.SetBackgroundColour((255, 230, 200, 255))
        self.button_GoPuT_op_bc.SetBackgroundColour((217, 221, 255))
        self.button_GoPuT_op_pc.SetBackgroundColour((217, 221, 255))
        self.button_GoDel_NKV.SetBackgroundColour((255, 198, 216))
        self.button_GoNKV.SetBackgroundColour((191, 255, 187))
        
        self.button_PuTCon.Disable() 
        self.txt_PuTCon.Disable()
        self.label_PuTCon.Disable()
        
        vbox = wx.BoxSizer(wx.VERTICAL)
        sb_exe = wx.StaticBox(self, -1, _("Execute")) 
        sbSizer_exe = wx.StaticBoxSizer(sb_exe, wx.VERTICAL)
        sb_bc = wx.StaticBox(self, -1, _("Base case")) 
        sbSizer_bc = wx.StaticBoxSizer(sb_bc, wx.VERTICAL) 
        sb_pc = wx.StaticBox(self, -1, _("Planning case")) 
        sbSizer_pc = wx.StaticBoxSizer(sb_pc, wx.VERTICAL) 
        sb_PuTCon = wx.StaticBox(self, -1, _("PuT Connections")) 
        sbSizer_PuTCon = wx.StaticBoxSizer(sb_PuTCon, wx.VERTICAL)
        sb_NKV = wx.StaticBox(self, -1, _("NKV"))
        sbSizer_NKV = wx.StaticBoxSizer(sb_NKV, wx.VERTICAL)
        
        sb_exe.SetFont(wx.Font(10, wx.DEFAULT, wx.NORMAL, wx.BOLD))
        sb_bc.SetFont(wx.Font(8, wx.DEFAULT, wx.NORMAL, wx.BOLD))
        sb_pc.SetFont(wx.Font(8, wx.DEFAULT, wx.NORMAL, wx.BOLD))
        sb_PuTCon.SetFont(wx.Font(8, wx.DEFAULT, wx.NORMAL, wx.BOLD))
        sb_NKV.SetFont(wx.Font(10, wx.DEFAULT, wx.NORMAL, wx.BOLD))
        
        self.label_PrT_bc.SetFont(wx.Font(9, wx.DEFAULT, wx.DEFAULT, wx.BOLD))
        self.label_PuT_bc.SetFont(wx.Font(9, wx.DEFAULT, wx.DEFAULT, wx.BOLD))
        self.label_PrT_pc.SetFont(wx.Font(9, wx.DEFAULT, wx.DEFAULT, wx.BOLD))
        self.label_PrT_pcput.SetFont(wx.Font(9, wx.DEFAULT, wx.DEFAULT, wx.BOLD))
        self.label_PuT_pc.SetFont(wx.Font(9, wx.DEFAULT, wx.DEFAULT, wx.BOLD))
        self.label_PuT_pcprt.SetFont(wx.Font(9, wx.DEFAULT, wx.DEFAULT, wx.BOLD))

        box_bc = wx.GridBagSizer()
        box_bc.Add(self.label_PrT_bc, (0,0), wx.DefaultSpan, wx.TOP,5)
        box_bc.Add(self.checkbox_PrT_bc, (0,1), wx.DefaultSpan, wx.LEFT|wx.TOP, 5)
        box_bc.Add(self.label_PuT_bc, (1,0), wx.DefaultSpan, wx.TOP,5)
        box_bc.Add(self.checkbox_PuT_bc, (1,1), wx.DefaultSpan, wx.LEFT|wx.TOP, 5)
        box_bc.Add(self.label_Val_bc, (2,0), wx.DefaultSpan, wx.TOP|wx.ALIGN_CENTER_VERTICAL,5)
        box_bc.Add(self.checkbox_Val_bc, (2,1), wx.DefaultSpan, wx.LEFT|wx.TOP|wx.RIGHT|wx.ALIGN_CENTER_VERTICAL,5)
        box_bc.Add(self.button_GoVal_bc, (2,2), wx.DefaultSpan, wx.TOP, 5)
        box_bc.Add(self.label_Del_bc, (3,0), wx.DefaultSpan, wx.TOP|wx.ALIGN_CENTER_VERTICAL,5)
        box_bc.Add(self.checkbox_Del_bc, (3,1), wx.DefaultSpan, wx.LEFT|wx.TOP|wx.RIGHT|wx.ALIGN_CENTER_VERTICAL,5)
        box_bc.Add(self.button_GoDel_bc, (3,2), wx.DefaultSpan, wx.TOP, 5)
        box_bc.Add(self.button_SelectAll_bc, (4,0), wx.DefaultSpan, wx.TOP|wx.EXPAND,5)
        box_bc.Add(self.button_DeselectAll_bc, (4,1), (0,2), wx.TOP|wx.EXPAND,5)
        sbSizer_bc.Add(box_bc, 0, wx.ALL|wx.CENTER, 5)
        
        box_pc = wx.GridBagSizer()
        box_pc.Add(self.label_PrT_pc, (0,0), wx.DefaultSpan, wx.TOP,5)
        box_pc.Add(self.checkbox_PrT_pc, (0,1), wx.DefaultSpan, wx.LEFT|wx.TOP,5)
        box_pc.Add(self.label_PrT_pcput, (1,0), wx.DefaultSpan, wx.TOP,5)
        box_pc.Add(self.checkbox_PrT_pcput, (1,1), wx.DefaultSpan, wx.LEFT|wx.TOP,5)
        box_pc.Add(self.label_PuT_pc, (2,0), wx.DefaultSpan, wx.TOP,5)
        box_pc.Add(self.checkbox_PuT_pc, (2,1), wx.DefaultSpan, wx.LEFT|wx.TOP,5)
        box_pc.Add(self.label_PuT_pcprt, (3,0), wx.DefaultSpan, wx.TOP,5)
        box_pc.Add(self.checkbox_PuT_pcprt, (3,1), wx.DefaultSpan, wx.LEFT|wx.TOP,5)
        box_pc.Add(self.label_Val_pc, (4,0), wx.DefaultSpan, wx.TOP|wx.ALIGN_CENTER_VERTICAL,5)
        box_pc.Add(self.checkbox_Val_pc, (4,1), wx.DefaultSpan, wx.LEFT|wx.TOP|wx.RIGHT|wx.ALIGN_CENTER_VERTICAL,5)
        box_pc.Add(self.button_GoVal_pc, (4,2), wx.DefaultSpan, wx.TOP, 5)
        box_pc.Add(self.label_Del_pc, (5,0), wx.DefaultSpan, wx.TOP|wx.ALIGN_CENTER_VERTICAL,5)
        box_pc.Add(self.checkbox_Del_pc, (5,1), wx.DefaultSpan, wx.LEFT|wx.TOP|wx.RIGHT|wx.ALIGN_CENTER_VERTICAL,5)
        box_pc.Add(self.button_GoDel_pc, (5,2), wx.DefaultSpan, wx.TOP, 5)
        sbSizer_pc.Add(box_pc, 0, wx.ALL|wx.CENTER, 5)  

        box_PuTCon = wx.GridBagSizer()
        box_PuTCon.Add(self.label_PuTCon, (0,0),(0,2), wx.TOP|wx.ALIGN_CENTER_VERTICAL,5)
        box_PuTCon.Add(self.button_PuTCon, (0,2),  wx.DefaultSpan, wx.TOP|wx.LEFT, 5)
        box_PuTCon.Add(self.txt_PuTCon, (1,0), (0,3), wx.TOP|wx.ALIGN_CENTER_VERTICAL,5)
        sbSizer_PuTCon.Add(box_PuTCon, 0, wx.ALL|wx.CENTER, 5) 
        
        box_Mid = wx.GridBagSizer()
        box_Mid.Add(self.button_ExeProc, (0,0), (0,2), wx.TOP|wx.EXPAND|wx.LEFT, 5)
        box_Mid.Add(self.button_ExePre, (0,3), wx.DefaultSpan, wx.TOP|wx.EXPAND|wx.RIGHT, 5)
        
        sbSizer_exe.Add(sbSizer_bc, 0, wx.ALL, 5)
        sbSizer_exe.Add(sbSizer_pc, 0, wx.ALL, 5)
        sbSizer_exe.Add(sbSizer_PuTCon, 0, wx.ALL, 5)
        sbSizer_exe.Add(box_Mid, 0, wx.ALL, 5)

        box_NKV = wx.GridBagSizer()
        box_NKV.Add(self.button_GoPuT_op_bc, (0,0), (0,2), wx.EXPAND|wx.TOP, 5)
        box_NKV.Add(self.button_GoPuT_op_pc, (1,0), (0,2), wx.EXPAND|wx.TOP, 5)
        box_NKV.Add(self.button_GoDel_NKV, (3,0),  wx.DefaultSpan, wx.TOP, 5)
        box_NKV.Add(self.button_GoImport_NKV, (3,1),  wx.DefaultSpan, wx.TOP, 5)
        box_NKV.Add(self.button_GoNKV, (5,0), (0,2), wx.EXPAND|wx.TOP, 5)
        sbSizer_NKV.Add(box_NKV, 0, wx.ALL|wx.CENTER, 5)

        box_Last = wx.GridBagSizer()
        box_Last.Add(self.button_Help, (0,0), wx.DefaultSpan, wx.TOP|wx.EXPAND, 5)
        box_Last.Add(self.button_Info, (0,1), wx.DefaultSpan, wx.TOP|wx.EXPAND, 5)
        box_Last.Add(self.button_Cancel, (0,3), wx.DefaultSpan, wx.TOP|wx.RIGHT, 5)

        vbox.Add(sbSizer_exe,0,wx.LEFT|wx.RIGHT,10)
        vbox.AddSpacer(20)
        vbox.Add(sbSizer_NKV,0,wx.LEFT|wx.RIGHT,10)
        vbox.AddSpacer(20)
        vbox.Add(box_Last,0,wx.LEFT,7)
        vbox.AddSpacer(5)
        self.SetSizerAndFit(vbox)

        self.Layout()
    
    def OnCalVal(self,event):
        param, paramOK = self.setParameter()
        for i in param: param[i]=False
        if event.GetEventObject().GetName() =="GoVal_bc": param["Val_bc"] = True
        else: param["Val_pc"] = True
        
        addInParam.SaveParameter(param)
        self.OnExit(None)
    
    def OnCheck(self,event):
        if event.IsChecked() and "_pc" in event.GetEventObject().GetName():
            self.checkbox_PrT_bc.SetValue(False)
            self.checkbox_PuT_bc.SetValue(False)
            self.checkbox_Val_bc.SetValue(False)
            self.checkbox_Del_bc.SetValue(False)
            if event.GetEventObject().GetName() not in ["Val_pc","Del_pc"]:
                self.checkbox_PrT_pc.SetValue(False)    
                self.checkbox_PrT_pcput.SetValue(False)
                self.checkbox_PuT_pc.SetValue(False)
                self.checkbox_PuT_pcprt.SetValue(False)
                self.checkbox_Val_pc.SetValue(True)
                if event.GetEventObject().GetName() =="PrT_pc":
                    self.checkbox_PrT_pc.SetValue(True)
                    self.button_PuTCon.Disable()
                    self.txt_PuTCon.Disable()
                    self.label_PuTCon.Disable()
                if event.GetEventObject().GetName() =="PrT_pcput":
                    self.checkbox_PrT_pcput.SetValue(True)
                    self.button_PuTCon.Disable()
                    self.txt_PuTCon.Disable()
                    self.label_PuTCon.Disable()
                if event.GetEventObject().GetName() =="PuT_pc": self.checkbox_PuT_pc.SetValue(True)
                if event.GetEventObject().GetName() =="PuT_pcprt": self.checkbox_PuT_pcprt.SetValue(True)
        if event.IsChecked() and "_bc" in event.GetEventObject().GetName():
            self.checkbox_PrT_pc.SetValue(False)    
            self.checkbox_PrT_pcput.SetValue(False)
            self.checkbox_PuT_pc.SetValue(False)
            self.checkbox_PuT_pcprt.SetValue(False)
            self.checkbox_Val_pc.SetValue(False)
            self.checkbox_Del_pc.SetValue(False)
            self.button_PuTCon.Disable()
            self.txt_PuTCon.Disable()
            self.label_PuTCon.Disable()
            if event.GetEventObject().GetName() in ["PrT_bc","PuT_bc"]: self.checkbox_Val_bc.SetValue(True)
        if event.IsChecked() and "PuT_pc" in event.GetEventObject().GetName():
            self.button_PuTCon.Enable()
            self.txt_PuTCon.Enable()
            self.label_PuTCon.Enable()
        if not event.IsChecked() and "PuT_pc" in event.GetEventObject().GetName():
            self.button_PuTCon.Disable()
            self.txt_PuTCon.Disable()
            self.label_PuTCon.Disable()
            
    def OnDelSteps(self,event):
        param, paramOK = self.setParameter()
        for i in param: param[i]=False
        param["Del"] = True
        
        addInParam.SaveParameter(param)
        self.OnExit(None)

    def OnDeselectAll_bc(self,event):
        self.checkbox_PrT_bc.SetValue(False)
        self.checkbox_PuT_bc.SetValue(False)
        self.checkbox_Val_bc.SetValue(False)
        self.checkbox_Del_bc.SetValue(False)
        
    def OnNKV(self,event):
        param, paramOK = self.setParameter()
        for i in param: param[i]=False
        param["NKV"] = event.GetEventObject().GetName()
        addInParam.SaveParameter(param)
        self.OnExit(None)
        
    def OnPuTCon(self,event):
        dlg = wx.DirDialog(self, _("Choose your directory to save PuT connections:"),
                           style=wx.DD_DEFAULT_STYLE | wx.DD_NEW_DIR_BUTTON)
        if dlg.ShowModal() == wx.ID_OK:
            pathindir = dlg.GetPath()+r"\verb.con"
            self.txt_PuTCon.SetValue(pathindir)
            self.txt_PuTCon.SetForegroundColour(wx.BLACK)
        dlg.Destroy()

    def OnExeProc(self,event):
        param, paramOK = self.setParameter()
        param["NKV"] = False
        
        if True in [param["PuT_pc"], param["PuT_pcprt"]]:
            param["PuTCon"] = True
            param["PuTConPath"] = self.txt_PuTCon.GetValue()
            if not os.path.isdir(os.path.split(param["PuTConPath"])[0]):
                self.txt_PuTCon.SetForegroundColour(wx.RED)
                addIn.ReportMessage(_("Folder for PuT Connections not existing!"))
                return
        
        if event.GetEventObject().GetName() == "ExeProc":
            param["ExeProc"] = True
            
        if event.GetEventObject().GetName() == "ExePre":
            param["Val_bc"] = False
            param["Val_pc"] = False
            param["Del"] = False
            self.checkbox_Val_bc.SetValue(False)
            self.checkbox_Val_pc.SetValue(False)
            self.checkbox_Del_bc.SetValue(False)
            if True not in [param["PrT_bc"],param["PuT_bc"],param["PrT_pc"],param["PrT_pcput"],param["PuT_pc"],param["PuT_pcprt"]]:
                addIn.ReportMessage(_("Preset procedures not possible for key values and deletion only!"))
                return
        
        if True in [param["PrT_pc"],param["PrT_pcput"],param["PuT_pc"],param["PuT_pcprt"],param["Val_pc"]]:
            if not self._check_pc(): return
        
        if not paramOK or True not in [param["PrT_bc"],param["PuT_bc"],param["Val_bc"],param["Del"],
                                       param["PrT_pc"],param["PrT_pcput"],param["PuT_pc"],param["PuT_pcprt"],param["Val_pc"]]:
            addIn.ReportMessage(_("Please select at least one procedure!"))
            return
        
        addInParam.SaveParameter(param)
        
        self.OnExit(None) 
   
    def OnExit(self,event):
        if not addIn.IsInDebugMode:
            Terminated.set()
        self.Destroy()
  
    def OnHelp(self,event):
        try:
            os.startfile(addIn.DirectoryPath + _("HelpStandiProcedures.htm"))
        except:
            addIn.HandleException()  
  
    def OnInfo(self,event):
        title = _("Info")
        frame = InfoFrame(title=title)
        
    def OnSelectAll_bc(self,event):
        self.checkbox_PrT_bc.SetValue(True)
        self.checkbox_PuT_bc.SetValue(True)
        self.checkbox_Val_bc.SetValue(True)
        self.checkbox_Del_bc.SetValue(True)
        
        self.checkbox_PrT_pc.SetValue(False)    
        self.checkbox_PrT_pcput.SetValue(False)
        self.checkbox_PuT_pc.SetValue(False)
        self.checkbox_PuT_pcprt.SetValue(False)
        self.checkbox_Val_pc.SetValue(False)
        self.checkbox_Del_pc.SetValue(False)
    
    def setParameter(self):
        param = dict()
        
        try:
            param["PrT_bc"] = self.checkbox_PrT_bc.GetValue()
            param["PuT_bc"] = self.checkbox_PuT_bc.GetValue()
            param["Val_bc"] = self.checkbox_Val_bc.GetValue()
            param["Del"] = max(self.checkbox_Del_bc.GetValue(), self.checkbox_Del_pc.GetValue())
            param["PrT_pc"] = self.checkbox_PrT_pc.GetValue()
            param["PrT_pcput"] = self.checkbox_PrT_pcput.GetValue()
            param["PuT_pc"] = self.checkbox_PuT_pc.GetValue()
            param["PuT_pcprt"] = self.checkbox_PuT_pcprt.GetValue()
            param["Val_pc"] = self.checkbox_Val_pc.GetValue()
            param["ExeProc"] = False
            param["PuTCon"] = False
            param["PuTConPath"] = False
            param["NKV"] = False
            return param, True
        except:
            addIn.HandleException(_("Standi Procedures, value error: "))
            return param, False
        
    def _check_pc(self):
        if Visum.Net.Matrices.ItemByKey(1).AttValue("SUM") == 0:
            addIn.ReportMessage(_("Car demand base case missing!"))
            return False
        if Visum.Net.Matrices.ItemByKey(10).AttValue("SUM") == 0:
            addIn.ReportMessage(_("PT demand base case missing!"))
            return False
        if Visum.Net.Matrices.ItemByKey(35).AttValue("SUM") == 0:
            addIn.ReportMessage(_("Car impedance base case missing!"))
            return False
        if Visum.Net.Matrices.ItemByKey(37).AttValue("SUM") == 0:
            addIn.ReportMessage(_("PT impedance base case missing!"))
            return False
        return True

def CheckNetwork():
    if 0 in [Visum.Net.Links.Count,Visum.Net.Zones.Count,Visum.Net.VehicleJourneys.Count]:
        addIn.ReportMessage(_("Current VISUM Version doesn't contain Links, Zones and PT Network! Create Network elements first!"))
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