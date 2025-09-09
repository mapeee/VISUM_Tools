# -*- coding: utf-8 -*-
"""
Created on Fri Sep  5 13:02:44 2025

@author: peter
"""

import wx
from VisumPy.AddIn import AddIn, AddInState, AddInParameter
_ = AddIn.gettext
import numpy as np
from typing import List, Tuple

import matplotlib
matplotlib.use('WXAgg')
from matplotlib.backends.backend_wxagg import (
    FigureCanvasWxAgg as FigureCanvas,
    NavigationToolbar2WxAgg as NavigationToolbar
)
from matplotlib.figure import Figure
import matplotlib.patches as patches
from matplotlib import colors as mcolors


class PlotFrame(wx.Frame):
    def __init__(
        self,
        parent,
        boxes: List[dict],
        xlim: Tuple[float, float] = (0, 24),
        ylim: Tuple[float, float] = (0, 20),
        title: str = _("Map PT Journeys"),
        throttle_ms: int = 30,  # ~30 FPS throttle for smooth zoom/pan
    ):
        super().__init__(parent, title=title, size=(900, 560))
        self.x_bounds = xlim
        self.y_bounds = ylim
        self._clamping = False
        self._pending = False
        self._throttle_ms = throttle_ms

        # --- Figure & axes ---
        fig = Figure()
        self.ax = fig.add_subplot(111)

        # Base limits & ticks
        self.ax.set_xlim(*xlim)
        self.ax.set_ylim(*ylim)

        self.ax.set_xticks(np.arange(xlim[0], xlim[1] + 1, 2))
        self.ax.set_xlabel(_("Hour"))
        self.ax.set_yticks(np.arange(0, max(ylim[1], 10) + 1, 2))
        self.ax.set_ylabel(_("Count"))

        # Canvas + toolbar
        self.canvas = FigureCanvas(self, -1, fig)
        toolbar = NavigationToolbar(self.canvas); toolbar.Realize()
        sizer = wx.BoxSizer(wx.VERTICAL)
        sizer.Add(toolbar, 0, wx.EXPAND)
        sizer.Add(self.canvas, 1, wx.EXPAND)
        self.SetSizer(sizer)

        # ---- Boxes state ----
        self.box_items = []
        for spec in boxes:
            self._add_box_from_dict(spec)

        # events: clamp + throttled label fit
        self.ax.callbacks.connect("xlim_changed", self._on_view_changed)
        self.ax.callbacks.connect("ylim_changed", self._on_view_changed)
        self.canvas.mpl_connect("resize_event", self._schedule_fit)
        self.canvas.mpl_connect("draw_event",   self._ensure_measured_once)

        # first draw → measure once → fit
        self.canvas.draw_idle()

    # ----- add a box from a plain dict -----
    def _add_box_from_dict(self, spec: dict):
        # required keys
        x0 = float(spec["x0"]); y0 = float(spec["y0"])
        w  = float(spec["w"]);  h  = float(spec["h"])
        label = str(spec.get("LINENAME", ""))

        # style with defaults
        edgecolor = spec.get("edgecolor", "red")
        facecolor = spec.get("facecolor", "none")
        fill_alpha = float(spec.get("fill_alpha", 0.30))
        linewidth = float(spec.get("linewidth", 0.5))
        base_font = float(spec.get("base_font", 10.0))
        textcolor = spec.get("textcolor", "blue")
        weight = spec.get("weight", "bold")

        # normalize face color to RGBA (honor explicit RGBA)
        if facecolor in (None, "none"):
            face_rgba = "none"
        else:
            try:
                # If tuple is RGB or RGBA, this handles both; alpha overrides if provided
                face_rgba = mcolors.to_rgba(facecolor, fill_alpha)
            except Exception:
                face_rgba = "none"

        rect = patches.Rectangle((x0, y0), w, h,
                                 linewidth=linewidth,
                                 edgecolor=edgecolor,
                                 facecolor=face_rgba)
        self.ax.add_patch(rect)

        cx, cy = x0 + w/2, y0 + h/2
        txt = self.ax.text(cx, cy, label,
                           ha='center', va='center',
                           fontsize=base_font,
                           weight=weight, color=textcolor,
                           clip_on=True)

        self.box_items.append({
            "spec": dict(x0=x0, y0=y0, w=w, h=h, label=label,
                         edgecolor=edgecolor, facecolor=facecolor,
                         fill_alpha=fill_alpha, linewidth=linewidth,
                         base_font=base_font, textcolor=textcolor, weight=weight),
            "rect": rect, "text": txt,
            "tw_base": None, "th_base": None
        })

    # ----- measure text once at base size -----
    def _ensure_measured_once(self, event):
        renderer = self.canvas.get_renderer()
        if renderer is None:
            return
        dirty = False
        for item in self.box_items:
            if item["tw_base"] is None:
                t = item["text"]
                t.set_fontsize(item["spec"]["base_font"])
                bb = t.get_window_extent(renderer=renderer)
                item["tw_base"] = max(1, bb.width)
                item["th_base"] = max(1, bb.height)
                dirty = True
        if dirty:
            self._schedule_fit(None)

    # ----- throttle fit to avoid jank during drag -----
    def _schedule_fit(self, event):
        if self._pending:
            return
        self._pending = True
        wx.CallLater(self._throttle_ms, self._do_fit)

    def _do_fit(self):
        self._pending = False
        renderer = self.canvas.get_renderer()
        if renderer is None:
            self.canvas.draw()
            renderer = self.canvas.get_renderer()

        pad = 3  # px inside box
        for item in self.box_items:
            spec = item["spec"]; text = item["text"]

            # keep centered
            cx, cy = spec["x0"] + spec["w"]/2, spec["y0"] + spec["h"]/2
            text.set_position((cx, cy))

            # data→pixel for box size
            p0 = self.ax.transData.transform((spec["x0"],             spec["y0"]))
            p1 = self.ax.transData.transform((spec["x0"] + spec["w"], spec["y0"] + spec["h"]))
            box_w_px = abs(p1[0] - p0[0]); box_h_px = abs(p1[1] - p0[1])
            usable_w = max(1, box_w_px - 2*pad); usable_h = max(1, box_h_px - 2*pad)

            # base text size (measured once)
            if item["tw_base"] is None:
                text.set_fontsize(spec["base_font"])
                bb = text.get_window_extent(renderer=renderer)
                item["tw_base"] = max(1, bb.width)
                item["th_base"] = max(1, bb.height)

            tw = item["tw_base"]; th = item["th_base"]
            scale = min(usable_w / tw, usable_h / th)
            new_size = max(1, min(500, spec["base_font"] * scale))
            text.set_fontsize(new_size)

        self.canvas.draw_idle()

    # ----- clamp view to bounds -----
    @staticmethod
    def _clamp_interval(a, b, lo, hi):
        span = b - a; maxspan = hi - lo
        if span >= maxspan: return lo, hi
        if a < lo: a = lo; b = a + span
        if b > hi: b = hi; a = b - span
        return a, b

    def _on_view_changed(self, ax):
        if self._clamping:
            self._schedule_fit(None)
            return
        self._clamping = True
        try:
            xmin, xmax = self._clamp_interval(*self.ax.get_xlim(), *self.x_bounds)
            ymin, ymax = self._clamp_interval(*self.ax.get_ylim(), *self.y_bounds)
            if (xmin, xmax) != tuple(self.ax.get_xlim()):
                self.ax.set_xlim(xmin, xmax)
            if (ymin, ymax) != tuple(self.ax.get_ylim()):
                self.ax.set_ylim(ymin, ymax)
        finally:
            self._clamping = False
        self._schedule_fit(None)