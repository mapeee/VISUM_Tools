#!/usr/bin/env python
import wx
import sys
import os
from pathlib import Path
import subprocess
import LeistungsabrechnungExecute as LE
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
        sizer.Add(wx.StaticText(self,label= _("03.02.2026")),0,wx.LEFT,10)
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
        super(MyDialog, self).__init__(parent, id=-1, title=title, size=(400, 250), style=wx.CAPTION | wx.STAY_ON_TOP)

        self.__InitUI()

    def __InitUI(self):
        self.labelTN = wx.StaticText(self, -1, _("Active Subnetwork"))
        self.labelSW = wx.StaticText(self, -1, _("School weeks"))
        self.labelFW = wx.StaticText(self, -1, _("Holiday weeks"))
        self.labelFB = wx.StaticText(self, -1, _("Vehicle needs"))
        self.labelEFW = wx.StaticText(self, -1, _("EFW"))
        self.labelUEL = wx.StaticText(self, -1, _("Overlap"))
        self.comboTN = wx.ComboBox(self, -1, "...")
        self.comboFB = wx.ComboBox(self, -1, "...")
        self.spinSW = wx.SpinCtrl(self, -1, min=0, max=52)
        self.spinFW = wx.SpinCtrl(self, -1, min=0, max=52)
        self.cbUEL = wx.CheckBox(self, -1, label="")
        self.cbEFW = wx.CheckBox(self, -1, label="")

        self.button_createEFW = wx.Button(self, -1, _("Create EFW"), name = "createEFW")
        self.button_PerformanceStatement = wx.Button(self, -1, _("Performance statement"), name = "PerformanceStatement")
        self.button_setParameters = wx.Button(self, -1, _("Set Parameters"), name = "setParameters")
        self.button_openLAR = wx.Button(self, -1, _("Open invoice"), name = "openLAR")
      
        self.button_exportMETN = wx.Button(self, -1, _("Export METN-List"), name = "exportMETN")
        self.button_delUDALines = wx.Button(self, -1, _("Delete Lines-UDA"), name = "delLinesUDA")
        self.button_delUDASN = wx.Button(self, -1, _("Delete SN-UDA"), name = "delSNUDA")
        
        self.button_help = wx.Button(self, -1, _('Help'))
        self.button_info = wx.Button(self, -1, _('Info'))
        self.button_exit = wx.Button(self, wx.ID_CANCEL, _('Close'))
        
        self.Bind(wx.EVT_COMBOBOX, self._on_comboChoice, self.comboTN)
        
        self.Bind(wx.EVT_CHECKBOX, self._on_checkbox, self.cbEFW)
        
        self.Bind(wx.EVT_BUTTON, self.OnProc, self.button_createEFW)
        self.Bind(wx.EVT_BUTTON, self.OnExecute, self.button_PerformanceStatement)
        self.Bind(wx.EVT_BUTTON, self.OnExecute, self.button_setParameters)
        self.Bind(wx.EVT_BUTTON, self.OnExecute, self.button_openLAR)
        
        self.Bind(wx.EVT_BUTTON, self.OnProc, self.button_exportMETN)
        self.Bind(wx.EVT_BUTTON, self.OnProc, self.button_delUDALines)
        self.Bind(wx.EVT_BUTTON, self.OnProc, self.button_delUDASN)
        
        self.Bind(wx.EVT_BUTTON, self.OnHelp, self.button_help)
        self.Bind(wx.EVT_BUTTON, self.OnInfo, self.button_info)
        self.Bind(wx.EVT_BUTTON, self.OnExit, self.button_exit)

        self.__do_layout()
        self.__do_parameters()
        self.__do_comboChoice()
        self.__set_properties()
        
        defaultParam = {"TN" : False, "SW" : False, "FW" : False, "EFW" : False, "FB" : False,
                        "UEL" : False}
        addInParam.Check(False, defaultParam)
        
    def __do_comboChoice(self):
        self.comboTN.Clear()
        unique_vals = list(dict.fromkeys(x[1] for x in Visum.Net.Lines.GetMultiAttValues("TN", False)))
        unique_vals = [x for x in unique_vals if x not in ("", "ohne")]
        self.comboTN.Append(unique_vals)
        TN = Visum.Net.AttValue("TN")
        if TN not in unique_vals and len(unique_vals) > 0: # Nimm voreingestellte Belegung aus BDA TN, wenn nicht in Linien, dann erstes aus Linien
            TN = unique_vals[0]
        self.comboTN.SetValue(TN)
        
        self.comboFB.Clear()
        self.comboFB.Append(["M1", "M2"])
        FB = Visum.Net.AttValue("M_FAHRZEUGBEDARF")
        if FB not in ["M1", "M2"]:
            FB = "M1"
        self.comboFB.SetValue(FB)
            
        
    def __do_parameters(self):
        self.spinSW.SetValue(int(Visum.Net.AttValue("S_WOCHEN")))
        self.spinFW.SetValue(int(Visum.Net.AttValue("F_WOCHEN")))
        self.cbUEL.SetValue(bool(Visum.Net.AttValue("ANABOVERLAP")))
        self.cbEFW.SetValue(bool(Visum.Net.AttValue("EFW")))
        
    def __set_properties(self):
        font = self.labelTN.GetFont()
        font.MakeBold()
        font.SetPointSize(8)
        self.labelTN.SetFont(font)
        self.labelSW.SetFont(font)
        self.labelFW.SetFont(font)
        self.labelFB.SetFont(font)
        self.labelEFW.SetFont(font)
        self.labelUEL.SetFont(font)
        
        TN = self.comboTN.GetValue()
        BDT = Visum.Net.TableDefinitions.GetMultiAttValues("NAME")
        if any(f"{TN} EFW" in item for item in BDT):
            self.button_createEFW.Disable()
        else:
            self.cbEFW.SetValue(False)
        if not Visum.Net.VehicleCombinations.AttrExists(f"{TN}_GESAMTKOSTEN"):
            self.button_openLAR.Disable()
        if not Visum.UserPreferences.DocumentName.endswith(".ver"):
            self.button_PerformanceStatement.Disable() # nicht im Szenario-Management ausführen. Dort nur über die Szenarien ausführen.
        else:
            self.button_PerformanceStatement.SetBackgroundColour(wx.Colour(184, 255, 184))
        
        font = self.button_exportMETN.GetFont()
        font.SetStyle(wx.FONTSTYLE_ITALIC)
        self.button_exportMETN.SetFont(font)
        self.button_exportMETN.SetForegroundColour("grey")
        
    def __do_layout(self):      
        sb_para = wx.StaticBox(self, -1, _("Parameters"))
        sb_para.SetFont(wx.Font(8, wx.DEFAULT, wx.NORMAL, wx.BOLD))
        sbSizer_para = wx.StaticBoxSizer(sb_para, wx.VERTICAL)
        sbSizer_para.SetMinSize((200, 50))
        grid_para = wx.FlexGridSizer(rows=0, cols=2, hgap=0, vgap=5)
        grid_para.Add(self.labelTN, 0, flag = wx.ALIGN_LEFT | wx.ALIGN_CENTER_VERTICAL)
        grid_para.Add(self.comboTN, 0)
        grid_para.Add((0,0))
        grid_para.Add(self.button_createEFW, 0)
        grid_para.Add(wx.StaticLine(self), 0, wx.EXPAND | wx.TOP | wx.BOTTOM, 8)
        grid_para.Add(wx.StaticLine(self), 0, wx.EXPAND | wx.TOP | wx.BOTTOM, 8)
        grid_para.Add(self.labelSW, 0, flag = wx.ALIGN_LEFT | wx.ALIGN_CENTER_VERTICAL)
        grid_para.Add(self.spinSW, 0, flag = wx.ALIGN_LEFT | wx.ALIGN_CENTER_VERTICAL)
        grid_para.Add(self.labelFW, 0, flag = wx.ALIGN_LEFT | wx.ALIGN_CENTER_VERTICAL)
        grid_para.Add(self.spinFW, 0, flag = wx.ALIGN_LEFT | wx.ALIGN_CENTER_VERTICAL)
        grid_para.Add(wx.StaticLine(self), 0, wx.EXPAND | wx.TOP | wx.BOTTOM, 8)
        grid_para.Add(wx.StaticLine(self), 0, wx.EXPAND | wx.TOP | wx.BOTTOM, 8)
        grid_para.Add(self.labelEFW, 0, flag = wx.ALIGN_LEFT | wx.ALIGN_CENTER_VERTICAL)
        grid_para.Add(self.cbEFW, 0, flag = wx.ALIGN_LEFT | wx.ALIGN_CENTER_VERTICAL)
        grid_para.Add(wx.StaticLine(self), 0, wx.EXPAND | wx.TOP | wx.BOTTOM, 8)
        grid_para.Add(wx.StaticLine(self), 0, wx.EXPAND | wx.TOP | wx.BOTTOM, 8)
        grid_para.Add(self.labelFB, 0, flag = wx.ALIGN_LEFT | wx.ALIGN_CENTER_VERTICAL)
        grid_para.Add(self.comboFB, 0)
        grid_para.Add(self.labelUEL, 0, flag = wx.ALIGN_LEFT | wx.ALIGN_CENTER_VERTICAL)
        grid_para.Add(self.cbUEL, 0, flag = wx.ALIGN_LEFT | wx.ALIGN_CENTER_VERTICAL)
        grid_para.AddSpacer(5)
        grid_para.AddSpacer(5)
        grid_para.Add(wx.StaticLine(self), 0, wx.EXPAND | wx.TOP | wx.BOTTOM, 0, (1, 2))
        grid_para.Add(wx.StaticLine(self), 0, wx.EXPAND | wx.TOP | wx.BOTTOM, 0, (1, 2))
        grid_para.AddSpacer(15)
        grid_para.AddSpacer(15)
        grid_para.Add(self.button_PerformanceStatement, 0, flag = wx.EXPAND)
        grid_para.Add(self.button_setParameters, 0, flag = wx.EXPAND)
        grid_para.AddSpacer(15)
        grid_para.AddSpacer(15)
        sbSizer_para.Add(grid_para, 1, wx.ALL | wx.ALIGN_CENTER, 10)
        
        sb_pref = wx.StaticBox(self, -1, _("Procedures"))
        sb_pref.SetFont(wx.Font(8, wx.DEFAULT, wx.NORMAL, wx.BOLD))
        sbSizer_pref = wx.StaticBoxSizer(sb_pref, wx.VERTICAL)
        sbSizer_pref.SetMinSize((200, 50))
        grid_pref = wx.FlexGridSizer(rows=0, cols=2, hgap=5, vgap=5)
        grid_pref.Add(self.button_openLAR, 0, flag = wx.EXPAND)
        grid_pref.Add(self.button_exportMETN, 0, flag = wx.EXPAND)
        grid_pref.AddSpacer(2)
        grid_pref.AddSpacer(2)
        grid_pref.Add(self.button_delUDALines, 0, flag = wx.EXPAND)
        grid_pref.Add(self.button_delUDASN, 1, flag = wx.EXPAND)
        grid_pref.AddGrowableCol(1, 1)
        sbSizer_pref.Add(grid_pref, 1, wx.ALL | wx.ALIGN_CENTER, 10)
        
        sb_end = wx.StaticBox(self, -1, _("End"))
        sb_end.SetFont(wx.Font(8, wx.DEFAULT, wx.NORMAL, wx.BOLD))
        sbSizer_end = wx.StaticBoxSizer(sb_end, wx.VERTICAL)
        sbSizer_end.SetMinSize((200, 50))
        grid_end = wx.FlexGridSizer(rows=1, cols=3, hgap=5, vgap=5)
        grid_end.Add(self.button_exit, flag = wx.EXPAND)
        grid_end.Add(self.button_help, flag = wx.EXPAND)
        grid_end.Add(self.button_info, flag = wx.EXPAND)
        sbSizer_end.Add(grid_end, 1, wx.ALL | wx.ALIGN_CENTER, 10)
        
        vbox = wx.BoxSizer(wx.VERTICAL)
        vbox.AddSpacer(10)
        vbox.Add(sbSizer_para, proportion = 0, flag = wx.EXPAND | wx.LEFT | wx.RIGHT, border = 10)
        vbox.AddSpacer(10)
        vbox.Add(sbSizer_pref, proportion = 0, flag = wx.EXPAND | wx.LEFT | wx.RIGHT, border = 10)
        vbox.AddSpacer(10)
        vbox.Add(sbSizer_end, proportion = 1, flag = wx.EXPAND | wx.LEFT | wx.RIGHT, border = 10)
        vbox.AddSpacer(10)

        self.SetSizerAndFit(vbox)
        self.Layout()
        self.Centre()
        
    def _on_comboChoice(self,event):
        TN = self.comboTN.GetValue()
        BDT = Visum.Net.TableDefinitions.GetMultiAttValues("NAME")
        if any(f"{TN} EFW" in item for item in BDT):
            self.button_createEFW.Disable()
        else:
            self.button_createEFW.Enable()
            self.cbEFW.SetValue(False)
        if not Visum.Net.VehicleCombinations.AttrExists(f"{TN}_GESAMTKOSTEN"):
            self.button_openLAR.Disable()
        else:
            self.button_openLAR.Enable()
            
    def _on_checkbox(self,event):
        if not self.cbEFW.GetValue():
            return
        TN = self.comboTN.GetValue()
        BDT = Visum.Net.TableDefinitions.GetMultiAttValues("NAME")
        if not any(f"{TN} EFW" in item for item in BDT):
            addIn.ReportMessage(_("EFW of SN %s not available.") %(TN))
            self.cbEFW.SetValue(False)
        
    def OnExecute(self,event):
        param, paramOK = self.setParameter()
        if not paramOK:
            addIn.ReportMessage(_("Some error occurred."))
            return
        param["Proc"] = event.GetEventObject().GetName()
        addInParam.SaveParameter(param)
        self.OnExit(None)
        
    def OnExit(self,event):
        if not addIn.IsInDebugMode:
            Terminated.set()
        self.Destroy()
        
    def OnHelp(self,event):
        try:
            try:
                file = Path(addIn.DirectoryPath + _("HelpLeistungsabrechnung.htm")).resolve()
                url = file.as_uri() + "#_Toc221540567"
                subprocess.run(["cmd", "/c", "start", "", "msedge", url], check=False)
            except:
                os.startfile(addIn.DirectoryPath + _("HelpLeistungsabrechnung.htm"))
        except:
            addIn.HandleException() 
            
    def OnInfo(self, event):
        title = _("Info")
        frame = InfoFrame(title=title)    
        
    def OnProc(self,event):
        proc = True
        ProcName = event.GetEventObject().GetName()
        if ProcName == "createEFW":
            proc, message = LE.createEFW(Visum, self.comboTN.GetValue())
        elif ProcName == "exportMETN":
            message = _("Not yet implemented.")
        elif ProcName == "delSNUDA":
            message = LE.delSNUDA(Visum, self.comboTN.GetValue())
        elif ProcName == "delLinesUDA":
            message = LE.delLinesUDA(Visum)
        else:
            pass
            
        if not proc:
            addIn.ReportMessage(message)
        else:
            addIn.ReportMessage(message, 2)
            
    def setParameter(self):
        param = dict()
        try:
            param["TN"] = self.comboTN.GetValue()
            param["SW"] = self.spinSW.GetValue()
            param["FW"] = self.spinFW.GetValue()
            param["EFW"] = self.cbEFW.GetValue()
            param["FB"] = self.comboFB.GetValue()
            param["UEL"] = self.cbUEL.GetValue()
            return param, True
        except:
            addIn.HandleException(_("Performance statement, value error: "))
            return param, False

def CheckNetwork():
    def _error():
        if not addIn.IsInDebugMode:
            Terminated.set()
        return False
    if Visum.Net.VehicleJourneyItems.Count == 0:
        addIn.ReportMessage(_("Current Visum Version has no VehicleJourneyItems. Create PT-Network first."))
        return _error()
    for i in ["TN", "S_WOCHEN", "F_WOCHEN", "EFW", "ANABOVERLAP", "M_FAHRZEUGBEDARF"]:
        if not Visum.Net.AttrExists(i):
            addIn.ReportMessage(_("Network UDA '%s' not existing.") %(i))
            return _error()
    if not Visum.Net.Lines.AttrExists("TN"):
        addIn.ReportMessage(_("Lines UDA 'TN' not existing."))
        return _error()
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
            dialog_1 = MyDialog(None, _("Performance statement"))
            app.SetTopWindow(dialog_1)
            dialog_1.ShowModal()
            if addIn.IsInDebugMode:
                app.MainLoop()
    except:
        addIn.HandleException(addIn.TemplateText.MainApplicationError)
        if not addIn.IsInDebugMode:
            Terminated.set()
