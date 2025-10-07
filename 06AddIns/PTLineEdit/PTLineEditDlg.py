#!/usr/bin/env python
import wx
import wx.grid
import sys
import os
import PTLineEdit_Execute as PTLEE
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
        sizer.Add(wx.StaticText(self,label= _("14.03.2025")),0,wx.LEFT,10)
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
        super(MyDialog, self).__init__(parent, id=-1, title=title, size=(400, 300), style=wx.CAPTION | wx.STAY_ON_TOP)

        self.__InitUI()

    def __InitUI(self):
        self.button_InitFilter = wx.Button(self, -1, _("Init Filter"), name = "InitFilter")
        self.button_PTFilter = wx.Button(self, -1, _("PT Filter"), name = "PTFilter")
        self.Bind(wx.EVT_BUTTON, self.OnProc, self.button_InitFilter)
        self.Bind(wx.EVT_BUTTON, self.OnProc, self.button_PTFilter)
        
        self.button_InitSRtimeBus = wx.Button(self, -1, _("Init bus link times"), name = "InitSRtimeBus")
        self.button_SetSRtimeBus = wx.Button(self, -1, _("Set bus link times on SR"), name = "SetSRtimeBus")
        self.Bind(wx.EVT_BUTTON, self.OnProc, self.button_InitSRtimeBus)
        self.Bind(wx.EVT_BUTTON, self.OnProc, self.button_SetSRtimeBus)

        self.button_export = wx.Button(self, -1, _('Export'), name = "Export")
        self.button_import = wx.Button(self, -1, _('Import'), name = "Import")
        self.button_export_import = wx.Button(self, -1, _('Export / Import'), name = "Export_Import")
        self.button_finish = wx.Button(self, -1, _('Finish'), name = "Finish")
        self.Bind(wx.EVT_BUTTON, self.OnProc, self.button_export)
        self.Bind(wx.EVT_BUTTON, self.OnProc, self.button_import)
        self.Bind(wx.EVT_BUTTON, self.OnProc, self.button_export_import)
        self.Bind(wx.EVT_BUTTON, self.OnProc, self.button_finish)
        
        self.button_help = wx.Button(self, -1, _('Help'))
        self.button_info = wx.Button(self, -1, _('Info'))
        self.button_exit = wx.Button(self, wx.ID_CANCEL, _('Close'))
        self.Bind(wx.EVT_BUTTON, self.OnHelp, self.button_help)
        self.Bind(wx.EVT_BUTTON, self.OnInfo, self.button_info)
        self.Bind(wx.EVT_BUTTON, self.OnExit, self.button_exit)

        image = wx.Image(addIn.DirectoryPath +"gear.png", wx.BITMAP_TYPE_PNG)
        small_image = image.Scale(24, 24, wx.IMAGE_QUALITY_HIGH)
        gear_icon = wx.Bitmap(small_image)
        self.gear_button = wx.BitmapButton(self, bitmap=gear_icon, size=(20, 20))
        self.gear_button.Bind(wx.EVT_BUTTON, self.Open_StopFrame)

        self.__do_layout()
    
    def __do_layout(self):   
        
        sb_filter = wx.StaticBox(self, -1, _("Filter"))
        sb_filter.SetFont(wx.Font(8, wx.DEFAULT, wx.NORMAL, wx.BOLD))
        sbSizer_filter = wx.StaticBoxSizer(sb_filter, wx.VERTICAL)
        sbSizer_filter.AddSpacer(10)
        sbSizer_filter.Add(self.button_InitFilter, flag = wx.ALIGN_CENTER | wx.ALL, border = 2)
        sbSizer_filter.Add(self.button_PTFilter, flag = wx.ALIGN_CENTER | wx.ALL, border = 2)
        
        sb_systemroute = wx.StaticBox(self, -1, _("SystemRoutes (SR)"))
        sb_systemroute.SetFont(wx.Font(8, wx.DEFAULT, wx.NORMAL, wx.BOLD))
        sbSizer_systemroute = wx.StaticBoxSizer(sb_systemroute, wx.VERTICAL)
        sbSizer_systemroute.AddSpacer(10)
        sbSizer_systemroute.Add(self.button_InitSRtimeBus, flag = wx.ALIGN_CENTER | wx.ALL, border = 2)
        sbSizer_systemroute.Add(self.button_SetSRtimeBus, flag = wx.ALIGN_CENTER | wx.ALL, border = 2)

        sizer_export = wx.BoxSizer(wx.HORIZONTAL)
        sizer_export.Add(self.button_export, 0, 5)
        sizer_export.Add(self.gear_button, 0, 5)

        sb_exportimport = wx.StaticBox(self, -1, _("Export / Import"))
        sb_exportimport.SetFont(wx.Font(8, wx.DEFAULT, wx.NORMAL, wx.BOLD))
        sbSizer_exportimport = wx.StaticBoxSizer(sb_exportimport, wx.VERTICAL)
        sbSizer_exportimport.AddSpacer(10)
        sbSizer_exportimport.Add(sizer_export, flag = wx.ALIGN_CENTER | wx.ALL, border = 2)
        sbSizer_exportimport.Add(self.button_import, flag = wx.ALIGN_CENTER | wx.ALL, border = 2)
        sbSizer_exportimport.AddSpacer(5)
        sbSizer_exportimport.Add(self.button_export_import, flag = wx.ALIGN_CENTER | wx.ALL, border = 2)
        sbSizer_exportimport.AddSpacer(5)
        sbSizer_exportimport.Add(self.button_finish, flag = wx.ALIGN_CENTER | wx.ALL, border = 2)

        sb_end = wx.StaticBox(self, -1, _("End"))
        sb_end.SetFont(wx.Font(8, wx.DEFAULT, wx.NORMAL, wx.BOLD))
        sbSizer_end = wx.StaticBoxSizer(sb_end, wx.VERTICAL)
        sbSizer_end.SetMinSize((200, 50))
        sbSizer_end.AddSpacer(10)
        sbSizer_end.Add(self.button_help, flag = wx.ALIGN_CENTER | wx.ALL, border = 2)
        sbSizer_end.Add(self.button_info, flag = wx.ALIGN_CENTER | wx.ALL, border = 2)
        sbSizer_end.Add(self.button_exit, flag = wx.ALIGN_CENTER | wx.ALL, border = 2)
        
        vbox = wx.BoxSizer(wx.VERTICAL)
        vbox.Add(sbSizer_filter, proportion = 0, flag = wx.EXPAND | wx.LEFT | wx.RIGHT, border = 10)
        vbox.AddSpacer(10)
        vbox.Add(sbSizer_exportimport, proportion = 0, flag = wx.EXPAND | wx.LEFT | wx.RIGHT, border = 10)
        vbox.AddSpacer(10)
        vbox.Add(sbSizer_systemroute, proportion = 0, flag = wx.EXPAND | wx.LEFT | wx.RIGHT, border = 10)
        vbox.AddSpacer(10)
        vbox.Add(sbSizer_end, proportion = 1, flag = wx.EXPAND | wx.LEFT | wx.RIGHT, border = 10)
        vbox.AddSpacer(10)

        self.button_import.Enable(False)

        self.SetSizerAndFit(vbox)
        self.Layout()
        self.Centre()
        
    def Open_StopFrame(self, event):
        self.stop_frame = StopFrame(self)
        
    def Get_StopData(self, data):
        self.stop_data = data

    def OnExit(self,event):
        if not addIn.IsInDebugMode:
            Terminated.set()
        self.Destroy()
        
    def OnProc(self,event):
        proc = True
        param, paramOK = self.setParameter()
        ProcName = event.GetEventObject().GetName()
        if ProcName == "InitFilter":
            PTLEE.InitFilter(Visum)
            Visum.Log(20480,_("All filters initialized and AddVal1 of Nodes and LineRoutes set to 0"))
        elif ProcName == "PTFilter":
            PTLEE.PTFilter(Visum)
            Visum.Log(20480,_("All filters ready for export"))
        elif ProcName == "SetSRtimeBus":
            PTLEE.SRtimeBus(Visum, True)
            Visum.Log(20480,_("Set bus link times based on SystemRoutes"))
        elif ProcName == "InitSRtimeBus":
            PTLEE.SRtimeBus(Visum, False)
            Visum.Log(20480,_("Init bus link times based on PT-Speed"))
        elif ProcName == "Finish":
            PTLEE.InitFilter(Visum)
            PTLEE.PTFilter(Visum)
            Visum.Log(20480,_("Finished"))
        elif ProcName == "Export":
            if hasattr(self, "stop_data"): Stops = self.stop_data
            else: Stops = [["0", "0"]]
            self.PTExport_State, self.PTExport, self.Nodes_ChaindedVS = PTLEE.PTExport(Visum, addIn.DirectoryPath, Stops)
            if self.PTExport_State == True:
                self.button_export.Enable(False)
                self.button_import.Enable(True)
                self.button_export_import.Enable(False)
                self.button_finish.Enable(False)
                Visum.Log(20480,_("Export finished"))
            else:
                proc = False
        elif ProcName == "Import":
            if hasattr(self, "PTExport_State"):
                proc = PTLEE.PTImport(Visum, self.Nodes_ChaindedVS, self.PTExport)
            else:
                proc = PTLEE.PTImport(Visum)
            self.button_export.Enable(True)
            self.button_import.Enable(False)
            self.button_export_import.Enable(True)
            self.button_finish.Enable(True)
            Visum.Log(20480,_("Import finished"))
        elif ProcName == "Export_Import":
            if hasattr(self, "stop_data"): Stops = self.stop_data
            else: Stops = [["0", "0"]]
            self.PTExport_State, self.PTExport, self.Nodes_ChaindedVS = PTLEE.PTExport(Visum, addIn.DirectoryPath, Stops)
            proc = PTLEE.PTImport(Visum, self.Nodes_ChaindedVS, self.PTExport)
            Visum.Log(20480,_("Export / Import finished"))
        else:
            param["Proc"] = ProcName
            addInParam.SaveParameter(param)
            if not addIn.IsInDebugMode:
                Terminated.set()
        
        if not proc:
            addIn.ReportMessage(_("Some error occurred"))
        else:
            addIn.ReportMessage(_("Ok"), 2)
        
    def OnHelp(self,event):
        try:
            os.startfile(addIn.DirectoryPath + _("HelpPTLineEdit.htm"))
        except:
            addIn.HandleException()
            
    def OnInfo(self, event):
        title = _("Info")
        frame = InfoFrame(title=title)
        
    def setParameter(self):
        param = dict()
        try:
            param["Proc"] = False
            return param, True
        except:
            addIn.HandleException(_("PTLineEdit, value error: "))
            return param, False

class StopFrame(wx.Frame):
    def __init__(self, parent):
        super(StopFrame, self).__init__(None, id=-1, title="", size=(280, 230), style=wx.CAPTION | wx.STAY_ON_TOP)

        self.parent = parent
        if hasattr(self.parent, "stop_data"):
            self.data = self.parent.stop_data
        else:
            self.data = [['0' for _ in range(2)] for _ in range(5)]

        panel = wx.Panel(self)
        self.grid = wx.grid.Grid(panel)
        self.grid.CreateGrid(5, 2)
        self.grid.SetColLabelValue(0, _("old StopNo"))
        self.grid.SetColLabelValue(1, _("new StopNo"))
        for row in range(5):
            for col in range(2):
                if hasattr(self.parent, "stop_data"):
                    self.grid.SetCellValue(row, col, str(self.parent.stop_data[row][col]))
                else:
                    self.grid.SetCellValue(row, col, "0")
                self.grid.SetCellEditor(row, col, wx.grid.GridCellNumberEditor())  # Restrict to numbers
                self.grid.SetCellAlignment(row, col, wx.ALIGN_CENTER, wx.ALIGN_CENTER)
                

        self.grid.Bind(wx.grid.EVT_GRID_CELL_CHANGED, self.on_cell_change)
        self.button_save = wx.Button(panel, label=_("Save"))
        self.button_close = wx.Button(panel, label=_("Close"))
        self.button_save.Bind(wx.EVT_BUTTON, self._OnSave)
        self.button_close.Bind(wx.EVT_BUTTON, self._OnClose)
        
        sizerhor = wx.BoxSizer(wx.HORIZONTAL)
        sizerhor.Add(self.button_save, 0, wx.ALIGN_CENTER | wx.ALL, 5)
        sizerhor.Add(self.button_close, 0, wx.ALIGN_CENTER | wx.ALL, 5)
        
        sizer = wx.BoxSizer(wx.VERTICAL)
        sizer.Add(self.grid, 1, wx.EXPAND | wx.ALL, 5)
        sizer.Add(sizerhor, 0, wx.ALIGN_CENTER | wx.ALL, 5)
        panel.SetSizer(sizer)
        
        self.Show()
        
    def on_cell_change(self, event):
        row = event.GetRow()
        col = event.GetCol()
        value = self.grid.GetCellValue(row, col)
        self.data[row][col] = value

    def _OnClose(self, event):
        self.Close()

    def _OnSave(self, event):
            self.parent.Get_StopData(self.data)
            self.Close()

def CheckNetwork():
    if 0 in [Visum.Net.LineRoutes.Count]:
        addIn.ReportMessage(_("Current Visum Version doesn't contain PT LineRoutes!"))
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
            dialog_1 = MyDialog(None, _("PT-Line Editor"))
            app.SetTopWindow(dialog_1)
            dialog_1.ShowModal()
            if addIn.IsInDebugMode:
                app.MainLoop()
    except:
        addIn.HandleException(addIn.TemplateText.MainApplicationError)
        if not addIn.IsInDebugMode:
            Terminated.set()
