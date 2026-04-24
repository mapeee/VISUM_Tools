#!/usr/bin/env python
import wx
import wx.lib.mixins.listctrl as listmix
import wx.dataview as dv
import sys
import subprocess
from pathlib import Path
import os
from datetime import datetime, timedelta
import importlib.util
import pandas as pd
import re
from VisumPy.AddIn import AddIn, AddInState, AddInParameter
_ = AddIn.gettext

class InfoFrame(wx.Frame):
    def __init__(self, title):
        super(InfoFrame, self).__init__(None, id=-1, title=title, style=wx.CAPTION | wx.STAY_ON_TOP, size=(190, 190))
        self.Centre()
        
        img_hvv = wx.Image(addIn.DirectoryPath +"logo.png",wx.BITMAP_TYPE_ANY)
        img_hvv = img_hvv.Scale(55,30,wx.IMAGE_QUALITY_BOX_AVERAGE)
        img_hvv = img_hvv.ConvertToBitmap()
        png_hvv = wx.StaticBitmap(self, -1, img_hvv, (0, 0))
        
        self.button = wx.Button(self, -1, _("OK"))
        self.Bind(wx.EVT_BUTTON, self.__OnOK, self.button)
        self.SetBackgroundColour(wx.Colour(wx.NullColour))

        self.label = wx.StaticText(self,label= _("hvv GmbH"))
        self.label.SetFont(wx.Font(8, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.BOLD))
        
        sizer = wx.BoxSizer(wx.VERTICAL)
        sizer.Add(png_hvv,0,wx.LEFT,5)
        sizer.AddSpacer(10)
        sizer.Add(self.label,0,wx.LEFT,10)
        sizer.AddSpacer(2)
        sizer.Add(wx.StaticText(self,label= _("Marcus Peter")),0,wx.LEFT,10)
        sizer.Add(wx.StaticText(self,label= _("19.02.2026")),0,wx.LEFT,10)
        sizer.AddSpacer(2)
        sizer.Add(wx.StaticText(self,label= _("Version 0.9: Beta")),0,wx.LEFT,10)
        sizer.AddSpacer(10)
        sizer.Add(self.button,0,wx.ALIGN_CENTER,5)
        sizer.AddSpacer(5)
        self.SetSizer(sizer)

        self.Show()

    def __OnOK(self,event):
        self.Close(True)

        
class List_ti(wx.ListCtrl, listmix.TextEditMixin):
    '''
    List Widget zur Eingabe der Zeitintervalle
    '''
    def __init__(self, parent):
        wx.ListCtrl.__init__(self, parent, style=wx.LC_REPORT | wx.BORDER_SUNKEN)
        listmix.TextEditMixin.__init__(self)
        self.SetSingleStyle(wx.LC_REPORT, True)
        self.InsertColumn(0, _("Start"), width=120)
        self.InsertColumn(1, _("End"), width=120)
        row = self.InsertItem(self.GetItemCount(), "06:00")
        self.SetItem(row, 1, "21:00")


class MyDialog(wx.Dialog):
    def __init__(self, parent, title):
        super(MyDialog, self).__init__(parent, id=-1, title=title, style=wx.CAPTION | wx.STAY_ON_TOP)
        self.__InitUI()

    def __InitUI(self):
        '''
        Erstelle alle notwendingen Elemente
        '''
        self.label_lines = wx.StaticText(self, -1, _(""))
        self.label_stops = wx.StaticText(self, -1, _(""))
        
        # time intervals
        self.label_time = wx.StaticText(self, -1, _("Time reference"))
        self.label_day = wx.StaticText(self, -1, _("Day"))
        self.combo_day = wx.ComboBox(self, -1, "")
        self.label_ti = wx.StaticText(self, -1, _("Time intervals"))
        self.list_ti = List_ti(self)
        self.list_ti.Bind(wx.EVT_LIST_END_LABEL_EDIT, self.OnEditTi)
        self.button_ti = wx.Button(self, -1, _("Add"))
        self.button_ti_remove = wx.Button(self, -1, _("Remove"))
        self.Bind(wx.EVT_BUTTON, self.OnAddTi, self.button_ti)
        self.Bind(wx.EVT_BUTTON, self.OnRemoveTi, self.button_ti_remove)
        
        # stop types
        self.label_st = wx.StaticText(self, -1, _("Stop types"))
        self.label_mode = wx.StaticText(self, -1, _("Aggregation level"))
        self.combo_mode = wx.ComboBox(self, -1, "")
        self.dvlc_scml = dv.DataViewListCtrl(self, -1, style=wx.BORDER_SUNKEN)
        self.Bind(wx.EVT_COMBOBOX, self.OnModeChanged, self.combo_mode)
        
        # stop categories
        self.label_sc = wx.StaticText(self, -1, _("Stop categories"))
        self.label1_sc = wx.StaticText(self, -1, _("F1: service frequency 1"))
        self.label2_sc = wx.StaticText(self, -1, _("F2: service frequency 2"))
        self.label3_sc = wx.StaticText(self, -1, _("F3: service frequency 3"))
        self.label4_sc = wx.StaticText(self, -1, _("F4: service frequency 4"))
        self.label5_sc = wx.StaticText(self, -1, _("F5: service frequency 5"))
        self.label6_sc = wx.StaticText(self, -1, _("F6: service frequency 6"))
        self.spin1_sc = wx.SpinCtrl(self, id=0, min=0, max=100, initial=24)
        self.spin2_sc = wx.SpinCtrl(self, id=1, min=0, max=100, initial=12)
        self.spin3_sc = wx.SpinCtrl(self, id=2, min=0, max=100, initial=6)
        self.spin4_sc = wx.SpinCtrl(self, id=3, min=0, max=100, initial=4)
        self.spin5_sc = wx.SpinCtrl(self, id=4, min=0, max=100, initial=2)
        self.spin6_sc = wx.SpinCtrl(self, id=5, min=0, max=100, initial=1)
        self.Bind(wx.EVT_SPINCTRL, self.OnSpin_sc, self.spin1_sc)
        self.Bind(wx.EVT_SPINCTRL, self.OnSpin_sc, self.spin2_sc)
        self.Bind(wx.EVT_SPINCTRL, self.OnSpin_sc, self.spin3_sc)
        self.Bind(wx.EVT_SPINCTRL, self.OnSpin_sc, self.spin4_sc)
        self.Bind(wx.EVT_SPINCTRL, self.OnSpin_sc, self.spin5_sc)
        self.Bind(wx.EVT_SPINCTRL, self.OnSpin_sc, self.spin6_sc)
        
        # service areas
        self.label_sa = wx.StaticText(self, -1, _("Service areas"))
        self.label1_sa = wx.StaticText(self, -1, _("E1: distance 1"))
        self.label2_sa = wx.StaticText(self, -1, _("E2: distance 2"))
        self.label3_sa = wx.StaticText(self, -1, _("E3: distance 3"))
        self.label4_sa = wx.StaticText(self, -1, _("E4: distance 4"))
        self.label5_sa = wx.StaticText(self, -1, _("E5: distance 5"))
        self.spin1_sa = wx.SpinCtrl(self, id=6, min=0, max=5000, initial=300)
        self.spin2_sa = wx.SpinCtrl(self, id=7, min=0, max=5000, initial=400)
        self.spin3_sa = wx.SpinCtrl(self, id=8, min=0, max=5000, initial=600)
        self.spin4_sa = wx.SpinCtrl(self, id=9, min=0, max=5000, initial=1000)
        self.spin5_sa = wx.SpinCtrl(self, id=10, min=0, max=5000, initial=1500)
        self.Bind(wx.EVT_SPINCTRL, self.OnSpin_sa, self.spin1_sa)
        self.Bind(wx.EVT_SPINCTRL, self.OnSpin_sa, self.spin2_sa)
        self.Bind(wx.EVT_SPINCTRL, self.OnSpin_sa, self.spin3_sa)
        self.Bind(wx.EVT_SPINCTRL, self.OnSpin_sa, self.spin4_sa)
        self.Bind(wx.EVT_SPINCTRL, self.OnSpin_sa, self.spin5_sa)
        
        # Info PNG
        img_ql = wx.Image(addIn.DirectoryPath +"PTQualityLevels.png",wx.BITMAP_TYPE_ANY)
        img_ql = img_ql.ConvertToBitmap()
        self.png_ql = wx.StaticBitmap(self, -1, img_ql, (0, 0))
        
        # misc
        self.label_le = wx.StaticText(self, -1, _("End of line double"))
        self.cb_le = wx.CheckBox(self, -1, "")
        self.label_bt = wx.StaticText(self, -1, _("Calculation type"))
        self.combo_bt = wx.ComboBox(self, -1, "")
        self.label_scen = wx.StaticText(self, -1, _("Scenario"))
        self.text_scen = wx.TextCtrl(self, -1, value="/")
        self.label_poi = wx.StaticText(self, -1, _("POI category"))
        self.combo_poi = wx.ComboBox(self, -1, "")
        self.label_poidel = wx.StaticText(self, -1, _("Delete existing POI"))
        self.cb_poidel = wx.CheckBox(self, -1, "")
        self.Bind(wx.EVT_TEXT, self.OnPOIBox, self.combo_poi)
        
        # clip
        self.label_clip = wx.StaticText(self, -1, _("Use Clip polygons"))
        self.cb_clip = wx.CheckBox(self, -1, "")
        self.listbox_clip = wx.ListBox(self)
        self.button_add_clip = wx.Button(self, -1, _("Add files..."))
        self.button_remove_clip = wx.Button(self, -1, _("Remove"))
        self.Bind(wx.EVT_CHECKBOX, self.OnCheckbox, self.cb_clip)
        self.Bind(wx.EVT_BUTTON, self.OnAddClip, self.button_add_clip)
        self.Bind(wx.EVT_BUTTON, self.OnRemoveClip, self.button_remove_clip)
        
        # end
        self.button_ok = wx.Button(self, -1, _("OK"))
        self.button_help = wx.Button(self, -1, _("Help"))
        self.button_info = wx.Button(self, -1, _("Info"))
        self.button_exit = wx.Button(self, wx.ID_CANCEL, _("Close"))
        self.Bind(wx.EVT_BUTTON, self.OnOK, self.button_ok)
        self.Bind(wx.EVT_BUTTON, self.OnHelp, self.button_help)
        self.Bind(wx.EVT_BUTTON, self.OnInfo, self.button_info)
        self.Bind(wx.EVT_BUTTON, self.OnExit, self.button_exit)

        self.__do_layout()
        self.__do_parameters()
        self.__do_comboChoice()
        self.__do_stoptypes()
        self.__set_properties()
        
        defaultParam = {"ti" : False, "day" : False, "le" : False, "scml" : False, "clipfiles" : False,
                        "bt" : False, "scen" : False, "poi" : False, "poidel" : False, "clip" : False,
                        "mode" : False, "sa" : False, "sc" : False}
        addInParam.Check(False, defaultParam)

    def __do_layout(self):
        # Time intervals / Stop types
        sb_time = wx.StaticBox(self, -1, "")
        sbSizer_time = wx.StaticBoxSizer(sb_time, wx.VERTICAL)
        fgSizer_time = wx.FlexGridSizer(rows=0, cols=2, vgap=0, hgap=40)
        fgSizer_time.AddGrowableCol(0, 1)
        fgSizer_time.AddGrowableCol(1, 1)
        grid_time_left = wx.GridBagSizer(vgap=7, hgap=10)
        grid_time_left.Add(self.label_time, pos=(0,0), flag = wx.ALIGN_LEFT | wx.ALIGN_CENTER_VERTICAL, border = 10)
        grid_time_left.Add(self.label_day, pos=(1,0), flag = wx.ALIGN_LEFT | wx.ALIGN_CENTER_VERTICAL)
        grid_time_left.Add(self.combo_day, pos=(1,1), flag = wx.EXPAND | wx.ALIGN_CENTER_VERTICAL)
        grid_time_left.Add(self.label_ti, pos=(2,0), span=(1,2), flag = wx.ALIGN_LEFT | wx.ALIGN_CENTER_VERTICAL)
        grid_time_left.Add(self.list_ti, pos=(3,0), span=(1,2), flag = wx.EXPAND)
        grid_time_left.Add(self.button_ti, pos=(4,0), flag = wx.ALIGN_LEFT | wx.ALIGN_CENTER_VERTICAL)
        grid_time_left.Add(self.button_ti_remove, pos=(4,1), flag = wx.ALIGN_LEFT | wx.ALIGN_CENTER_VERTICAL)
        grid_time_left.AddGrowableCol(1, 1)
        grid_time_left.AddGrowableRow(2, 1)
        grid_time_right = wx.GridBagSizer(vgap=7, hgap=10)
        grid_time_right.Add(self.label_st, pos=(0,0), flag = wx.ALIGN_LEFT | wx.ALIGN_CENTER_VERTICAL, border = 10)
        grid_time_right.Add(self.label_mode, pos=(1,0), flag = wx.ALIGN_LEFT | wx.ALIGN_CENTER_VERTICAL, border = 10)
        grid_time_right.Add(self.combo_mode, pos=(1,1), flag = wx.EXPAND | wx.ALIGN_CENTER_VERTICAL)
        grid_time_right.Add(self.dvlc_scml, pos=(2,0), span=(1,2), flag= wx.EXPAND | wx.ALL)
        grid_time_right.AddGrowableCol(1, 1)
        grid_time_right.AddGrowableRow(1, 1) 
        fgSizer_time.Add(grid_time_left, 1, wx.TOP)
        fgSizer_time.Add(grid_time_right, 1, wx.TOP)
        sbSizer_time.Add(fgSizer_time, 1, wx.ALL | wx.EXPAND, 10)
        # Stop cat / Service areas
        sb_st = wx.StaticBox(self, -1, "")
        sbSizer_st = wx.StaticBoxSizer(sb_st, wx.VERTICAL)
        grid_st = wx.GridBagSizer(vgap=0, hgap=40)
        grid_st_left = wx.GridBagSizer(vgap=7, hgap=10)
        grid_st_left.Add(self.label_sc, pos=(0,0), flag = wx.ALIGN_LEFT | wx.ALIGN_CENTER_VERTICAL, border = 10)
        grid_st_left.Add(self.label1_sc, pos=(1,0), flag = wx.ALIGN_LEFT | wx.ALIGN_CENTER_VERTICAL, border = 10)
        grid_st_left.Add(self.spin1_sc, pos=(1,1), flag = wx.EXPAND)
        grid_st_left.Add(wx.StaticText(self, -1, "/h"), pos=(1,2), flag = wx.ALIGN_CENTER_VERTICAL)
        grid_st_left.Add(self.label2_sc, pos=(2,0), flag = wx.ALIGN_LEFT | wx.ALIGN_CENTER_VERTICAL, border = 10)
        grid_st_left.Add(self.spin2_sc, pos=(2,1), flag = wx.EXPAND)
        grid_st_left.Add(wx.StaticText(self, -1, "/h"), pos=(2,2), flag = wx.ALIGN_CENTER_VERTICAL)
        grid_st_left.Add(self.label3_sc, pos=(3,0), flag = wx.ALIGN_LEFT | wx.ALIGN_CENTER_VERTICAL, border = 10)
        grid_st_left.Add(self.spin3_sc, pos=(3,1), flag = wx.EXPAND)
        grid_st_left.Add(wx.StaticText(self, -1, "/h"), pos=(3,2), flag = wx.ALIGN_CENTER_VERTICAL)
        grid_st_left.Add(self.label4_sc, pos=(4,0), flag = wx.ALIGN_LEFT | wx.ALIGN_CENTER_VERTICAL, border = 10)
        grid_st_left.Add(self.spin4_sc, pos=(4,1), flag = wx.EXPAND)
        grid_st_left.Add(wx.StaticText(self, -1, "/h"), pos=(4,2), flag = wx.ALIGN_CENTER_VERTICAL)
        grid_st_left.Add(self.label5_sc, pos=(5,0), flag = wx.ALIGN_LEFT | wx.ALIGN_CENTER_VERTICAL, border = 10)
        grid_st_left.Add(self.spin5_sc, pos=(5,1), flag = wx.EXPAND)
        grid_st_left.Add(wx.StaticText(self, -1, "/h"), pos=(5,2), flag = wx.ALIGN_CENTER_VERTICAL)
        grid_st_left.Add(self.label6_sc, pos=(6,0), flag = wx.ALIGN_LEFT | wx.ALIGN_CENTER_VERTICAL, border = 10)
        grid_st_left.Add(self.spin6_sc, pos=(6,1), flag = wx.EXPAND)
        grid_st_left.Add(wx.StaticText(self, -1, "/h"), pos=(6,2), flag = wx.ALIGN_CENTER_VERTICAL)
        grid_st_right = wx.GridBagSizer(vgap=7, hgap=10)
        grid_st_right.Add(self.label_sa, pos=(0,0), flag = wx.ALIGN_LEFT | wx.ALIGN_CENTER_VERTICAL, border = 10)
        grid_st_right.Add(self.label1_sa, pos=(1,0), flag = wx.ALIGN_LEFT | wx.ALIGN_CENTER_VERTICAL, border = 10)
        grid_st_right.Add(self.spin1_sa, pos=(1,1), flag = wx.EXPAND)
        grid_st_right.Add(wx.StaticText(self, -1, "m"), pos=(1,2), flag = wx.ALIGN_CENTER_VERTICAL)
        grid_st_right.Add(self.label2_sa, pos=(2,0), flag = wx.ALIGN_LEFT | wx.ALIGN_CENTER_VERTICAL, border = 10)
        grid_st_right.Add(self.spin2_sa, pos=(2,1), flag = wx.EXPAND)
        grid_st_right.Add(wx.StaticText(self, -1, "m"), pos=(2,2), flag = wx.ALIGN_CENTER_VERTICAL)
        grid_st_right.Add(self.label3_sa, pos=(3,0), flag = wx.ALIGN_LEFT | wx.ALIGN_CENTER_VERTICAL, border = 10)
        grid_st_right.Add(self.spin3_sa, pos=(3,1), flag = wx.EXPAND)
        grid_st_right.Add(wx.StaticText(self, -1, "m"), pos=(3,2), flag = wx.ALIGN_CENTER_VERTICAL)
        grid_st_right.Add(self.label4_sa, pos=(4,0), flag = wx.ALIGN_LEFT | wx.ALIGN_CENTER_VERTICAL, border = 10)
        grid_st_right.Add(self.spin4_sa, pos=(4,1), flag = wx.EXPAND)
        grid_st_right.Add(wx.StaticText(self, -1, "m"), pos=(4,2), flag = wx.ALIGN_CENTER_VERTICAL)
        grid_st_right.Add(self.label5_sa, pos=(5,0), flag = wx.ALIGN_LEFT | wx.ALIGN_CENTER_VERTICAL, border = 10)
        grid_st_right.Add(self.spin5_sa, pos=(5,1), flag = wx.EXPAND)
        grid_st_right.Add(wx.StaticText(self, -1, "m"), pos=(5,2), flag = wx.ALIGN_CENTER_VERTICAL)
        grid_st.Add(grid_st_left,  pos=(0,0), flag = wx.EXPAND)
        grid_st.Add(grid_st_right, pos=(0,1), flag = wx.EXPAND)
        grid_st.Add((20, 20), pos=(1, 0))
        grid_st.Add(self.png_ql, pos=(2,0), span=(1,2), flag = wx.EXPAND)
        grid_st.AddGrowableCol(0)
        grid_st.AddGrowableCol(1)
        sbSizer_st.Add(grid_st, 1, wx.ALL | wx.EXPAND, 10)
        # Parameters
        sb_para = wx.StaticBox(self, -1, _("Parameters"))
        sb_para.SetFont(wx.Font(8, wx.DEFAULT, wx.NORMAL, wx.BOLD))
        sbSizer_para = wx.StaticBoxSizer(sb_para, wx.VERTICAL)
        grid_para = wx.GridBagSizer(vgap=7, hgap=10)
        grid_para.Add(self.label_bt, pos=(0,0), flag = wx.ALIGN_LEFT | wx.ALIGN_CENTER_VERTICAL)
        grid_para.Add(self.combo_bt, pos=(0,1), flag = wx.ALIGN_LEFT | wx.EXPAND)
        grid_para.Add(self.label_le, pos=(1,0), flag = wx.ALIGN_LEFT | wx.ALIGN_CENTER_VERTICAL)
        grid_para.Add(self.cb_le, pos=(1,1), flag = wx.ALIGN_LEFT | wx.ALIGN_CENTER_VERTICAL)
        grid_para.Add(self.label_scen, pos=(2,0), flag = wx.ALIGN_LEFT | wx.ALIGN_CENTER_VERTICAL)
        grid_para.Add(self.text_scen, pos=(2,1), flag = wx.ALIGN_LEFT | wx.EXPAND)
        grid_para.Add(self.label_poi, pos=(3,0), flag = wx.ALIGN_LEFT | wx.ALIGN_CENTER_VERTICAL)
        grid_para.Add(self.combo_poi, pos=(3,1), flag = wx.ALIGN_LEFT | wx.ALIGN_CENTER_VERTICAL)
        grid_para.Add(self.label_poidel, pos=(4,0), flag = wx.ALIGN_LEFT | wx.ALIGN_CENTER_VERTICAL)
        grid_para.Add(self.cb_poidel, pos=(4,1), flag = wx.ALIGN_LEFT | wx.ALIGN_CENTER_VERTICAL)
        grid_para.Add(self.label_clip, pos=(5,0), flag = wx.ALIGN_LEFT | wx.ALIGN_CENTER_VERTICAL)
        grid_para.Add(self.cb_clip, pos=(5,1), flag = wx.ALIGN_LEFT | wx.ALIGN_CENTER_VERTICAL)
        grid_para.Add(self.listbox_clip, pos=(6,0), span=(1,2), flag= wx.EXPAND | wx.ALL)
        grid_para.Add(self.button_add_clip, pos=(7,0), flag=wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM)
        grid_para.Add(self.button_remove_clip, pos=(7,1), flag=wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM)
        grid_para.AddGrowableCol(0)
        grid_para.AddGrowableCol(1)
        sbSizer_para.Add(grid_para, 1, wx.ALL | wx.EXPAND, 10)
        # End
        sb_end = wx.StaticBox(self, -1, _("End"))
        sb_end.SetFont(wx.Font(8, wx.DEFAULT, wx.NORMAL, wx.BOLD))
        sbSizer_end = wx.StaticBoxSizer(sb_end, wx.VERTICAL)
        grid_end = wx.FlexGridSizer(rows=2, cols=2, hgap=5, vgap=5)
        grid_end.Add(self.button_ok, flag = wx.EXPAND)
        grid_end.Add(self.button_exit, flag = wx.EXPAND)
        grid_end.Add(self.button_help, flag = wx.EXPAND)
        grid_end.Add(self.button_info, flag = wx.EXPAND)
        sbSizer_end.Add(grid_end, 1, wx.ALL | wx.ALIGN_CENTER, 10)
        
        # Outer box
        vbox = wx.BoxSizer(wx.VERTICAL)
        vbox.AddSpacer(15)
        vbox.Add(self.label_lines, flag = wx.EXPAND | wx.LEFT | wx.RIGHT, border = 25)
        vbox.Add(self.label_stops, flag = wx.EXPAND | wx.LEFT | wx.RIGHT, border = 25)
        vbox.AddSpacer(15)
        vbox.Add(sbSizer_time, proportion = 0, flag = wx.EXPAND | wx.LEFT | wx.RIGHT, border = 20)
        vbox.AddSpacer(15)
        vbox.Add(sbSizer_st, proportion = 0, flag = wx.EXPAND | wx.LEFT | wx.RIGHT, border = 20)
        vbox.AddSpacer(15)
        vbox.Add(sbSizer_para, proportion = 0, flag = wx.EXPAND | wx.LEFT | wx.RIGHT, border = 20)
        vbox.AddSpacer(15)
        vbox.Add(sbSizer_end, proportion = 0, flag = wx.EXPAND | wx.LEFT | wx.RIGHT, border = 20)
        vbox.AddSpacer(15)

        self.SetSizerAndFit(vbox)
        self.Layout()
        self.Centre()

    def __do_comboChoice(self):
        # weekday
        self.combo_day.Clear()
        if Visum.Net.CalendarPeriod.AttValue("TYPE") == "CALENDARPERIODWEEK":
            self.combo_day.Append([_("Monday"), _("Tuesday"), _("Wednesday"),
                                  _("Thursday"), _("Friday"), _("Saturday"), _("Sunday")])
        elif Visum.Net.CalendarPeriod.AttValue("TYPE") == "CALENDARPERIODYEAR":
            start_date = datetime.strptime(Visum.Net.CalendarPeriod.AttValue("VALIDFROM"), "%d.%m.%Y")
            end_date = datetime.strptime(Visum.Net.CalendarPeriod.AttValue("VALIDUNTIL"), "%d.%m.%Y")
            dates = []
            current = start_date
            while current <= end_date:
                dates.append(current.strftime("%d.%m.%Y"))
                current += timedelta(days=1)
            self.combo_day.Append(dates)
        else:
            self.combo_day.Append([_("daily")])
            self.combo_day.Enable(False)
        self.combo_day.SetSelection(0)
        # calculation type
        self.combo_bt.Clear()
        self.combo_bt.Append(["HKAT", "HKAT FHH"])
        self.combo_bt.SetSelection(1)
        # poi category
        self.combo_poi.Clear()
        if not Visum.Net.POICategories.Count:
            self.combo_poi.Append(["New Category"])
            self.combo_poi.SetSelection(0)
            font = self.combo_poi.GetFont()
            font.SetStyle(wx.FONTSTYLE_ITALIC)
            self.combo_poi.SetFont(font)
            self.combo_poi.SetForegroundColour(wx.Colour(150, 150, 150))
            self.combo_poi.SetInsertionPointEnd()
        else:
            poi_cat = [i.AttValue("NAME") for i in Visum.Net.POICategories.GetAll]
            self.combo_poi.Append(poi_cat)
            idx = next((i for i, v in enumerate(poi_cat)
                if any(k in v.lower() for k in ("klasse", "sochrone"))), 0) # Standardauswahl von erster Kategorie, die string enthält.
            self.combo_poi.SetSelection(idx)
        self.combo_poi.SetMinSize((150, -1))
        self.combo_poi.SetMaxSize((150, -1))

    def __do_parameters(self):
        # clip
        self.cb_clip.SetValue(False)
        self.button_add_clip.Disable()
        self.button_remove_clip.Disable()
        self.listbox_clip.Enable(False)
        # misc
        self.cb_le.SetValue(True)
        self.cb_poidel.SetValue(True)
        self.label_lines.SetLabel(_("Active lines: %s") %(str(Visum.Net.Lines.CountActive)))
        self.label_stops.SetLabel(_("Active stops: %s") %(str(Visum.Net.Stops.CountActive)))
    
    def __do_stoptypes(self):
        # calculation type
        self.combo_mode.Clear()
        # Select Mainlines as default if not first active Mainline is '' (not all Lines with Mainline)
        if not sorted({row[1] for row in Visum.Net.Lines.GetMultiAttValues("MAINLINENAME", True)})[0]: # True if at least one active Line has no MAINLINENAME
            self.combo_mode.Append([_("TSys")])
            self.combo_mode.SetSelection(0)
            self.combo_mode.Enable(False)
            mode = _("TSys")
            mode2 = "TSYSCODE"
        else:
            self.combo_mode.Append([_("TSys"), _("Mainlines")])
            self.combo_mode.SetSelection(1)
            mode =  _("Mainline")
            mode2 = "MAINLINENAME"
        # dvlc stop categories and mainlines/TSys
        self.dvlc_scml.AppendTextColumn(mode, width=60)
        choices = dv.DataViewChoiceRenderer(["1", "2", "3"], mode=dv.DATAVIEW_CELL_EDITABLE)
        col = dv.DataViewColumn(_("StopType"), choices, 1, width=60)
        self.dvlc_scml.AppendColumn(col)
        self._setDefaultModes(mode2)
        
    def __set_properties(self):
        font = self.label_ti.GetFont()
        font.MakeBold()
        font.SetPointSize(8)
        self.label_ti.SetFont(font)
        self.list_ti.SetMinSize((-1, 100))
        self.dvlc_scml.SetMinSize((-1, 120))
        self.label_time.SetFont(font)
        self.label_st.SetFont(font)
        self.label_sa.SetFont(font)
        self.label_sc.SetFont(font)

    def OnAddClip(self,event):
        start_dir = os.path.dirname(os.path.abspath(__file__))
        start_dir = os.path.join(start_dir, "data") # Standardverzeichnis für Clip-Files
        with wx.FileDialog(self, message="", defaultDir=start_dir,
                           wildcard="GeoJSON files (*.geojson;*.json)|*.geojson;*.json", # Nur GeoJSON und JSON erlaubte Formate
                           style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST | wx.FD_MULTIPLE) as dlg:
            if dlg.ShowModal() == wx.ID_OK:
                paths = dlg.GetPaths() # <-- multiple paths
                for path in paths:
                    if path not in self.listbox_clip.GetItems():
                        self.listbox_clip.Append(path)
    
    def OnAddTi(self,event):
        index = self.list_ti.InsertItem(self.list_ti.GetItemCount(), "00:00")
        self.list_ti.SetItem(index, 1, "00:00")
        
    def OnCheckbox(self,event):
        if self.cb_clip.GetValue():
            self.button_add_clip.Enable()
            self.button_remove_clip.Enable()
            self.listbox_clip.Enable(True) 
        else:
            self.button_add_clip.Disable()
            self.button_remove_clip.Disable()
            self.listbox_clip.Enable(False) 
            
    def OnEditTi(self, event):
        '''
        Texteingabe in Zeitintervalle; Textformat ist vorgegeben
        '''
        new_text = event.GetText()
        pattern = r"^([01]\d|2[0-3]):([0-5]\d)$"
        if not re.match(pattern, new_text):
            event.Veto()
            return
        event.Skip() # accept change
        
    def OnExit(self,event):
        if not addIn.IsInDebugMode:
            Terminated.set()
        self.Destroy()
        
    def OnHelp(self,event):
        try:
            try:
                file = Path(addIn.DirectoryPath + _("HelpPTQualityLevels.htm")).resolve()
                url = file.as_uri() + "#_Toc226985244"
                subprocess.run(["cmd", "/c", "start", "", "msedge", url], check=False)
            except:
                os.startfile(addIn.DirectoryPath + _("HelpPTQualityLevels.htm"))
        except:
            addIn.HandleException() 
            
    def OnInfo(self, event):
        title = _("Info")
        frame = InfoFrame(title=title)
        
    def OnModeChanged(self,event):
        mode = [_("TSys"), _("Mainline")][self.combo_mode.GetSelection()]
        mode2 = ["TSYSCODE", "MAINLINENAME"][self.combo_mode.GetSelection()]
        if mode2 == "MAINLINENAME" and not sorted({row[1] for row in Visum.Net.Lines.GetMultiAttValues("MAINLINENAME", True)})[0]: # True if at least one active Line has MAINLINENAME ''
            addIn.ReportMessage(_("At least one active line has no mainline"))
            self.combo_mode.SetSelection(0)
            return
        self.dvlc_scml.DeleteAllItems()
        self.dvlc_scml.GetColumn(0).SetTitle(mode)
        self._setDefaultModes(mode2)
        
    def OnOK(self,event):
        param, paramOK = self._setParameter()
        if not paramOK:
            return
        addInParam.SaveParameter(param)
        self.OnExit(None)
        
    def OnSpin_sa(self,event):
        list_sa = [self.spin1_sa.GetValue(), self.spin2_sa.GetValue(), self.spin3_sa.GetValue(),
                   self.spin4_sa.GetValue(), self.spin5_sa.GetValue()]
        index = event.GetId()-6
        index_value = list_sa[index]
        for i, v in enumerate(list_sa):
            if i == index:
                continue
            elif i < index:
                list_sa[i] = min(v, index_value)
            else:
                list_sa[i] = max(v, index_value)
        self.spin1_sa.SetValue(list_sa[0])
        self.spin2_sa.SetValue(list_sa[1])
        self.spin3_sa.SetValue(list_sa[2])
        self.spin4_sa.SetValue(list_sa[3])
        self.spin5_sa.SetValue(list_sa[4])
        
    def OnSpin_sc(self,event):
        list_sc = [self.spin1_sc.GetValue(), self.spin2_sc.GetValue(), self.spin3_sc.GetValue(),
                   self.spin4_sc.GetValue(), self.spin5_sc.GetValue(), self.spin6_sc.GetValue()]
        index = event.GetId()
        index_value = list_sc[index]
        for i, v in enumerate(list_sc):
            if i == index:
                continue
            elif i < index:
                list_sc[i] = max(v, index_value)
            else:
                list_sc[i] = min(v, index_value)
        self.spin1_sc.SetValue(list_sc[0])
        self.spin2_sc.SetValue(list_sc[1])
        self.spin3_sc.SetValue(list_sc[2])
        self.spin4_sc.SetValue(list_sc[3])
        self.spin5_sc.SetValue(list_sc[4])
        self.spin6_sc.SetValue(list_sc[5])
        
    def OnPOIBox(self, event):
        text = self.combo_poi.GetValue()
        font = self.combo_poi.GetFont()
        if text == "New Category":
            font.SetStyle(wx.FONTSTYLE_ITALIC)
            self.combo_poi.SetFont(font)
            self.combo_poi.SetForegroundColour(wx.Colour(150, 150, 150))
            self.combo_poi.SetInsertionPointEnd()
        else:
            font = wx.Font(10, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL)
            self.combo_poi.SetFont(font)
            self.combo_poi.SetForegroundColour(wx.Colour(0, 0, 0))
            self.combo_poi.SetInsertionPointEnd()
        self.combo_poi.Refresh()
        event.Skip()
        
    def OnRemoveClip(self,event):
        selections = list(self.listbox_clip.GetSelections())
        selections.reverse()  # remove from bottom to top
        for index in selections:
            self.listbox_clip.Delete(index)
        
    def OnRemoveTi(self,event):
        item = self.list_ti.GetFirstSelected()
        while item != -1:
            self.list_ti.DeleteItem(item)
            item = self.list_ti.GetFirstSelected()
            
    def _get_scml_data(self):
        '''
        Erstelle pandas df mit den Werten aus den Spalten MODE und STOPTYPE
        '''
        rows = []
        for r in range(self.dvlc_scml.GetItemCount()):
            a = self.dvlc_scml.GetTextValue(r, 0)
            b = int(self.dvlc_scml.GetTextValue(r, 1))
            rows.append((a, b))
        data_scml = pd.DataFrame(rows, columns=["MODE", "STOPTYPE"])
        return data_scml
        
    def _get_ti_data(self):
        '''
        Generiere aus Einträgen der Zeitintervalle eine Liste mit Start- und Endzeit (jeweils in Sekunden)
        1. Erstelle erst Liste mit Zeitintervallen (Start Ende)
        2. Prüfe auf Überlappungen
        '''
        data_ti = []
        for row in range(self.list_ti.GetItemCount()):
            h, m = map(int, self.list_ti.GetItemText(row, 0).split(":"))
            start = h*60*60 + m*60 # hours / minutes to seconds
            h, m = map(int, self.list_ti.GetItemText(row, 1).split(":"))
            end = h*60*60 + m*60
            data_ti.append([start, end])
        data_ti = sorted(data_ti)
        prev_end = 0
        for start, end in data_ti:
            if end <= start or start < prev_end: # Fehler, wenn Ende vor Beginn oder Beginn vor Ende aus vorherigem Intervall
                return False
            prev_end = end
        return data_ti
        
    def _setDefaultModes(self, _mode2):
        if _mode2 == "TSYSCODE":
            values_default = dict(zip(["S", "U", "FV", "RV", "SPNV"], 
                              ["1", "1", "1", "1", "1"]))
        else:
            values_default = dict(zip(["IC", "ICE", "ICE(S)", "Nightjet", "RE", "RB", "S-Bahn", "AKN", "U-Bahn",
                                       "XpressBus", "Metrobus", "Nachtbus", "Bus", "Schiff"], 
                              ["1", "1", "1", "1", "1", "1", "1", "1", "1",
                               "3", "3", "3", "3", "3"]))
        mode_values = sorted({row[1] for row in Visum.Net.Lines.GetMultiAttValues(_mode2, True)})
        for m in mode_values:
            self.dvlc_scml.AppendItem([m, values_default.get(m, "3")])
    
    def _setParameter(self):
        param = dict()
        try:
            param["ti"] = self._get_ti_data()
            if not param["ti"]:
                addIn.ReportMessage(_("Overlapping time intervals"))
                return None, False
            if Visum.Net.CalendarPeriod.AttValue("TYPE") == "CALENDARPERIODYEAR":
                param["day"] = self.combo_day.GetValue()
            else:
                param["day"] = self.combo_day.GetSelection() # get index of selection
            param["bt"] = self.combo_bt.GetSelection() # get index of selection
            param["le"] = self.cb_le.GetValue()
            param["clip"] = self.cb_clip.GetValue()
            if param["clip"] and not importlib.util.find_spec("geopandas"):
                addIn.ReportMessage(_("Python module geopandas not installed"))
                return None, False
            param["clipfiles"] = self.listbox_clip.GetItems()
            if param["clip"] and not len(param["clipfiles"]):
                addIn.ReportMessage(_("No clip file selected"))
                return None, False
            param["scml"] = self._get_scml_data()
            param["sa"] = [self.spin1_sa.GetValue(), self.spin2_sa.GetValue(), self.spin3_sa.GetValue(),
                       self.spin4_sa.GetValue(), self.spin5_sa.GetValue()]
            param["sc"] = [self.spin1_sc.GetValue(), self.spin2_sc.GetValue(), self.spin3_sc.GetValue(),
                       self.spin4_sc.GetValue(), self.spin5_sc.GetValue(), self.spin6_sc.GetValue()]
            param["scen"] = self.text_scen.GetValue()
            if not param["scen"]:
                addIn.ReportMessage(_("Scenario must not be empty"))
                return None, False
            param["poi"] = self.combo_poi.GetValue() # get index of selection
            param["poidel"] = self.cb_poidel.GetValue()
            param["mode"] = self.combo_mode.GetSelection() # get index of selection
            return param, True
        except:
            addIn.HandleException(_("PT quality level, value error: "))
            return param, False

def CheckNetwork():
    def _error():
        if not addIn.IsInDebugMode:
            Terminated.set()
        return False
    if not Visum.Net.VehicleJourneys.CountActive:
        addIn.ReportMessage(_("Current Version has no active VehicleJourneys"))
        return _error()
    if not Visum.Net.Stops.CountActive:
        addIn.ReportMessage(_("Current Version has no active Stops"))
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
            dialog_1 = MyDialog(None, _("PT quality levels"))
            app.SetTopWindow(dialog_1)
            dialog_1.ShowModal()
            if addIn.IsInDebugMode:
                app.MainLoop()
    except:
        addIn.HandleException(addIn.TemplateText.MainApplicationError)
        if not addIn.IsInDebugMode:
            Terminated.set()
