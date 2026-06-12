import os
import time
import threading
from dataclasses import dataclass, field
from pathlib import Path

import pythoncom
import win32com.client
import numpy as np
import pyqtgraph as pg
from PyQt5.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QComboBox, QLabel, QSpinBox, QApplication,
    QRadioButton, QButtonGroup, QDoubleSpinBox, QCheckBox, QAbstractSpinBox, QFrame,
    QTableWidget, QTableWidgetItem, QHeaderView, QGraphicsOpacityEffect, QStackedWidget,
    QTabBar, QToolTip,
)
from PyQt5.QtCore import QTimer, Qt, pyqtSignal
from PyQt5.QtGui import QFont, QColor, QCursor

try:
    import serial
    from magstim_test import disable_remote, get_parameters, get_temperature
    _SERIAL_AVAILABLE = True
except ImportError:
    _SERIAL_AVAILABLE = False

from labchart_client import LabChartClient

LABCHART_FILE = "MainDQ.adicht"

WINDOW_SECS_DEFAULT = 4
WINDOW_SECS_MIN = 4
WINDOW_SECS_MAX = 15
SCOPE_PRE_DEFAULT = 0.2
SCOPE_POST_DEFAULT = 0.8

UPDATE_MS = 100
STATUS_MS = 1000
MAX_DISPLAY_PTS = 2000

COLOR_LEFT  = "#1565C0"  # blue
COLOR_RIGHT = "#C62828"  # red

DEFAULT_CH_LEFT  = 2  # Ch3 (0-indexed)
DEFAULT_CH_RIGHT = 3  # Ch4 (0-indexed)

TRIGGER_CH_IDX    = 7    # Ch8 "Channel 8" — TMS trigger signal (0-indexed)
TRIGGER_THRESHOLD = 2.5  # V — rising above this = TMS fired
TRIGGER_REFRACTORY = 0.5  # s — minimum gap between consecutive triggers

MAGSTIM_PORT = "COM7"
MAGSTIM_POLL_MS = 1000



@dataclass
class _Side:
    """Per-channel (left/right) state and widget references."""
    ch:    int
    color: str
    # per-trigger state
    hunt_num:  int  = 0
    hunt_den:  int  = 0
    hunt_history:      list = field(default_factory=list)
    hunt_latest_eval:  bool = True
    table_latest_eval: bool = True
    # y-axis state
    yfix:  bool  = False
    yhalf: float = 1.0
    # plot overlay lists
    vlines: list = field(default_factory=list)
    shades: list = field(default_factory=list)
    # widgets — assigned during _setup_ui
    combo:       object = None
    plot:        object = None
    curve:       object = None
    panel:       object = None
    stats:       object = None
    table:       object = None
    hunt_panel:  object = None
    hunt_chk:    object = None
    hunt_num_lbl: object = None
    hunt_den_lbl: object = None
    hunt_clear:  object = None
    hunt_redo:   object = None
    btn_yfix:    object = None
    btn_yauto:   object = None


class _MaxSizeStack(QStackedWidget):
    """QStackedWidget that always reserves space for its largest page."""
    def sizeHint(self):
        from PyQt5.QtCore import QSize
        w = max((self.widget(i).sizeHint().width()  for i in range(self.count())), default=0)
        h = max((self.widget(i).sizeHint().height() for i in range(self.count())), default=0)
        return QSize(w, h)
    def minimumSizeHint(self):
        return self.sizeHint()


class _WheelSpinBox(QSpinBox):
    """QSpinBox: wheel without focus, no arrow buttons, click to select all."""
    def wheelEvent(self, event):
        delta = 1 if event.angleDelta().y() > 0 else -1
        self.setValue(self.value() + delta)
        event.accept()

    def mousePressEvent(self, event):
        super().mousePressEvent(event)
        self.selectAll()


class CalculatorWindow(QWidget):
    """MSO percentage calculator, always on top."""

    _PCTS = list(range(10, 210, 10))  # 10, 20, ..., 200

    def __init__(self, parent=None):
        super().__init__(parent, Qt.Window)
        self.setWindowTitle("MSO Calculator")
        self.setWindowFlags(self.windowFlags() | Qt.WindowStaysOnTopHint)
        outer = QHBoxLayout(self)
        outer.setContentsMargins(12, 12, 12, 12)
        outer.setSpacing(8)

        # Left: input + Cal button (same level as table)
        left = QWidget()
        left.setFixedWidth(88)
        ll = QVBoxLayout(left)
        ll.setContentsMargins(0, 0, 0, 0)
        ll.setSpacing(4)
        ll.addStretch()
        ll.addWidget(QLabel("rMT MSO"))
        self._spin = _WheelSpinBox()
        self._spin.setRange(1, 100)
        self._spin.setValue(50)
        self._spin.setButtonSymbols(QAbstractSpinBox.NoButtons)
        ll.addWidget(self._spin)
        btn_cal = QPushButton("Cal")
        btn_cal.clicked.connect(self._on_cal)
        ll.addWidget(btn_cal)
        ll.addStretch()
        outer.addWidget(left)

        # Right: table — columns = %, single row updated on each Cal
        self._table = QTableWidget(0, len(self._PCTS))
        self._table.setHorizontalHeaderLabels([f"{p}%" for p in self._PCTS])
        self._table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self._table.verticalHeader().setVisible(True)
        self._table.setEditTriggers(QTableWidget.NoEditTriggers)
        self._table.setSelectionMode(QTableWidget.NoSelection)
        self._table.setFocusPolicy(Qt.NoFocus)
        outer.addWidget(self._table)

        self.resize(800, 100)

    def showEvent(self, event):
        super().showEvent(event)
        self._fit_height()

    def _fit_height(self):
        hh = self._table.horizontalHeader().height()
        row_h = self._table.rowHeight(0) if self._table.rowCount() > 0 else 24
        sb_h = self._table.horizontalScrollBar().height()
        self.resize(self.width(), hh + row_h + sb_h + 28)

    def _on_cal(self):
        rmt = self._spin.value()
        self._table.setRowCount(0)
        self._table.insertRow(0)
        self._table.setVerticalHeaderItem(0, QTableWidgetItem(f"rMT {rmt}"))

        first_exceeded = False
        for i, pct in enumerate(self._PCTS):
            mso = round(rmt * pct / 100)
            exceeds = mso > 100
            if exceeds and first_exceeded:
                item = QTableWidgetItem("")
            else:
                item = QTableWidgetItem(str(mso))
                item.setTextAlignment(Qt.AlignCenter)
                if exceeds:
                    item.setForeground(QColor(150, 150, 150))
                    first_exceeded = True
            self._table.setItem(0, i, item)

        self._table.resizeRowsToContents()
        self._fit_height()
        idx_100 = self._PCTS.index(100)
        self._table.scrollTo(
            self._table.model().index(0, idx_100),
            QTableWidget.PositionAtCenter
        )


class PlotWindow(QWidget):
    """Standalone window: scatter plot (top) + tabbed merged data table (bottom)."""

    _COLS = ["Trial", "CH", "Hand", "Paralysis", "Loc_ID", "MSO1", "MSO2", "ISI",
             "MEP_amp", "Prestim_mean", "Temp", "Mode"]

    # ch number → RGBA tuple
    _CH_COLORS = {
        3: (33,  150, 243, 200),   # blue
        4: (244,  67,  54, 200),   # red
        5: ( 76, 175,  80, 200),   # green
        6: (255, 152,   0, 200),   # orange
        7: (156,  39, 176, 200),   # purple
        8: (  0, 188, 212, 200),   # cyan
    }
    _FALLBACK = [(255, 87, 34, 200), (96, 125, 139, 200),
                 (233, 30, 99, 200), (121, 85, 72, 200)]

    # (display name, column index in _COLS)
    _PLOT_COLS = [
        ("MEP_amp",      8),
        ("Prestim_mean", 9),
        ("Trial",        0),
        ("MSO1",         5),
        ("MSO2",         6),
        ("ISI",          7),
        ("Temp",         10),
        ("Loc_ID",       4),
    ]
    _CATEGORICAL_X      = {4, 5, 6, 7}   # Loc_ID, MSO1, MSO2, ISI → categorical x-axis
    _ZERO_PLACEHOLDER   = {8, 9}   # MEP_amp, Prestim_mean start as 0 placeholder

    def __init__(self, viewer, parent=None):
        super().__init__(parent, Qt.Window)
        self._viewer       = viewer
        self._resizing     = False
        self._all_rows     = []   # cached merged+sorted rows
        self._tab_channels = []   # CH int values for tabs index 1,2,...
        self._scatter_items  = {}   # ch_int -> ScatterPlotItem (individual dots)
        self._median_items   = {}   # ch_int -> ScatterPlotItem (median dots)
        self._errbar_items   = {}   # ch_int -> ErrorBarItem (IQR)
        self._median_data    = {}   # ch_int -> list of {med, q1, q3} per x level
        self._scatter_data   = {}   # ch_int -> list of {x_pos, y_val, trial}
        self._legend_chs     = set()
        self._excluded       = set()   # set of (ch, trial) tuples
        self._use_mean       = False
        self._hover_text     = ""
        self._hover_timer    = QTimer(self)
        self._hover_timer.setInterval(300)
        self._hover_timer.timeout.connect(self._refresh_tooltip)
        self._y_col          = 8   # default: MEP_amp
        self._x_col          = 5   # default: MSO1
        self.setWindowTitle("TMSviewer — Plot")
        self.setWindowFlags(self.windowFlags() | Qt.WindowStaysOnTopHint)
        self.resize(992, 558)

        root = QVBoxLayout(self)
        root.setContentsMargins(8, 8, 8, 8)
        root.setSpacing(8)

        # ── Scatter plot (9) + right button panel (1) ───────────────────
        graph_area    = QWidget()
        graph_row     = QHBoxLayout(graph_area)
        graph_row.setContentsMargins(0, 0, 0, 0)
        graph_row.setSpacing(4)

        self.canvas = pg.GraphicsLayoutWidget()
        self._plot_item = self.canvas.addPlot()
        self._plot_item.setLabel("left",   "MEP amp")
        self._plot_item.setLabel("bottom", "MSO1")
        self._plot_item.showGrid(x=True, y=True, alpha=0.3)
        self._plot_item.getAxis("bottom").setStyle(tickTextOffset=6)
        self._plot_item.getAxis("left").enableAutoSIPrefix(False)
        self._plot_item.scene().sigMouseMoved.connect(self._on_plot_mouse_moved)
        graph_row.addWidget(self.canvas, stretch=9)

        graph_btn_panel  = QWidget()
        graph_btn_layout = QVBoxLayout(graph_btn_panel)
        graph_btn_layout.setContentsMargins(2, 0, 2, 0)
        graph_btn_layout.setSpacing(4)

        self._btn_stat_mode = QPushButton("Mean")
        self._btn_stat_mode.clicked.connect(self._on_toggle_stat_mode)
        graph_btn_layout.addWidget(self._btn_stat_mode)

        btn_save_plot = QPushButton("Save plot")
        btn_save_plot.clicked.connect(self._on_save_plot)
        graph_btn_layout.addWidget(btn_save_plot)

        graph_btn_layout.addSpacing(84)

        btn_calc = QPushButton("Calculator")
        btn_calc.clicked.connect(self._on_open_calculator)
        graph_btn_layout.addWidget(btn_calc)

        graph_btn_layout.addStretch()
        graph_row.addWidget(graph_btn_panel, stretch=1)

        root.addWidget(graph_area, stretch=3)

        # ── Legend bar (Y-axis controls left + channel colours centred) ────
        self._legend_bar = QWidget()
        legend_row = QHBoxLayout(self._legend_bar)
        legend_row.setContentsMargins(4, 2, 4, 2)
        legend_row.setSpacing(8)

        legend_row.addWidget(QLabel("Y axis"))
        self._btn_yfix = QPushButton("Fix")
        self._btn_yfix.setFixedWidth(40)
        self._btn_yfix.setCheckable(True)
        self._btn_yfix.clicked.connect(self._on_yaxis_fix)
        legend_row.addWidget(self._btn_yfix)
        self._btn_yauto = QPushButton("Auto")
        self._btn_yauto.setFixedWidth(44)
        self._btn_yauto.setCheckable(True)
        self._btn_yauto.setChecked(True)
        self._btn_yauto.clicked.connect(self._on_yaxis_auto)
        legend_row.addWidget(self._btn_yauto)

        legend_row.addSpacing(14)
        legend_row.addWidget(QLabel("Y:"))
        self._combo_y = QComboBox()
        for name, _ in self._PLOT_COLS:
            self._combo_y.addItem(name)
        self._combo_y.setCurrentIndex(0)   # MEP_amp
        self._combo_y.currentIndexChanged.connect(self._on_y_changed)
        legend_row.addWidget(self._combo_y)

        legend_row.addSpacing(6)
        legend_row.addWidget(QLabel("X:"))
        self._combo_x = QComboBox()
        for name, _ in self._PLOT_COLS:
            self._combo_x.addItem(name)
        self._combo_x.setCurrentIndex(3)   # MSO1
        self._combo_x.currentIndexChanged.connect(self._on_x_changed)
        legend_row.addWidget(self._combo_x)

        legend_row.addStretch()

        # Channel colour items live in a sub-widget so _update_legend
        # can clear them without touching the Y-axis controls above.
        self._ch_legend = QWidget()
        self._legend_layout = QHBoxLayout(self._ch_legend)
        self._legend_layout.setContentsMargins(0, 0, 0, 0)
        self._legend_layout.setSpacing(16)
        legend_row.addWidget(self._ch_legend)

        legend_row.addStretch()

        root.addWidget(self._legend_bar)

        # ── Tab bar + table ─────────────────────────────────────────────
        # ── Bottom area: table (9) + button panel (1) ───────────────────
        bottom_area   = QWidget()
        bottom_row    = QHBoxLayout(bottom_area)
        bottom_row.setContentsMargins(0, 0, 0, 0)
        bottom_row.setSpacing(4)

        # Left: tab bar + table
        table_area = QWidget()
        ta_layout  = QVBoxLayout(table_area)
        ta_layout.setContentsMargins(0, 0, 0, 8)
        ta_layout.setSpacing(0)

        self._tab_bar = QTabBar()
        self._tab_bar.addTab("All")
        self._tab_bar.currentChanged.connect(self._on_tab_changed)
        ta_layout.addWidget(self._tab_bar)

        self.merged_table = self._make_merged_table()
        ta_layout.addWidget(self.merged_table)
        bottom_row.addWidget(table_area, stretch=9)

        # Right: action buttons
        btn_panel  = QWidget()
        btn_layout = QVBoxLayout(btn_panel)
        btn_layout.setContentsMargins(2, 0, 2, 0)
        btn_layout.setSpacing(4)

        self._btn_exclude = QPushButton("Exclude")
        self._btn_exclude.clicked.connect(self._on_exclude_clicked)
        btn_layout.addWidget(self._btn_exclude)

        btn_clear_excl = QPushButton("Excl.Clear")
        btn_clear_excl.clicked.connect(self._on_clear_excluded)
        btn_layout.addWidget(btn_clear_excl)

        btn_layout.addStretch()
        bottom_row.addWidget(btn_panel, stretch=1)

        root.addWidget(bottom_area, stretch=1)

    # ------------------------------------------------------------------

    def _ch_color(self, ch):
        if ch in self._CH_COLORS:
            return self._CH_COLORS[ch]
        return self._FALLBACK[(ch - 9) % len(self._FALLBACK)]

    def _make_merged_table(self):
        cols = self._COLS + ["Excl"]
        tbl = QTableWidget(0, len(cols))
        tbl.setHorizontalHeaderLabels(cols)
        tbl.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        tbl.horizontalHeader().setStretchLastSection(False)
        tbl.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        tbl.verticalHeader().setVisible(False)
        tbl.setEditTriggers(QTableWidget.NoEditTriggers)
        tbl.setSelectionBehavior(QTableWidget.SelectRows)
        tbl.itemSelectionChanged.connect(self._on_table_selection_changed)
        return tbl

    def refresh(self):
        """Rebuild from viewer L + R tables: update scatter plot, tabs, table."""
        rows = []
        for src in (self._viewer.L.table, self._viewer.R.table):
            for r in range(src.rowCount()):
                rows.append([
                    src.item(r, c).text() if src.item(r, c) else ""
                    for c in range(src.columnCount())
                ])

        def _key(row):
            try:    ch    = int(row[1])
            except: ch    = 0
            try:    trial = int(row[0])
            except: trial = 0
            return (ch, trial)

        rows.sort(key=_key)
        self._all_rows = rows

        # Add tabs for newly seen channels (ascending order)
        unique_chs = sorted({int(r[1]) for r in rows if r[1].isdigit()})
        for ch in unique_chs:
            if ch not in self._tab_channels:
                self._tab_channels.append(ch)
                self._tab_bar.addTab(f"Ch{ch}")

        self._refresh_plot()
        self._populate_table()

    # ── Scatter plot ───────────────────────────────────────────────────

    def _refresh_plot(self):
        tab_idx = self._tab_bar.currentIndex()
        if tab_idx <= 0:
            source_rows = self._all_rows
        else:
            ch_filter = str(self._tab_channels[tab_idx - 1])
            source_rows = [r for r in self._all_rows if r[1] == ch_filter]

        points = []
        for row in source_rows:
            try:
                y_val  = float(row[self._y_col])
                if self._y_col in self._ZERO_PLACEHOLDER and y_val == 0.0:
                    continue
                x_val  = float(row[self._x_col])
                ch     = int(row[1])
                trial  = int(row[0])
                if (ch, trial) in self._excluded:
                    continue
                points.append((x_val, y_val, ch, trial))
            except (ValueError, IndexError):
                continue

        if not points:
            for d in (self._scatter_items, self._median_items, self._errbar_items):
                for ch in list(d):
                    self._plot_item.removeItem(d.pop(ch))
            return

        is_cat = self._x_col in self._CATEGORICAL_X
        x_map_inv = {}   # position → original x value (used in tooltip)

        if is_cat:
            unique_x = sorted({p[0] for p in points})
            x_map = {v: i for i, v in enumerate(unique_x)}
            x_map_inv = {i: v for v, i in x_map.items()}
            def _fmt(v):
                try:
                    f = float(v)
                    return f"{f:.1f}" if f != int(f) else str(int(f))
                except (ValueError, TypeError):
                    return str(v)
            self._plot_item.getAxis("bottom").setTicks(
                [[(i, _fmt(v)) for i, v in enumerate(unique_x)]]
            )
            self._plot_item.setXRange(-0.5, len(unique_x) - 0.5, padding=0)
            get_x = lambda p: x_map[p[0]]
        else:
            self._plot_item.getAxis("bottom").setTicks(None)
            self._plot_item.enableAutoRange(axis="x")
            get_x = lambda p: p[0]

        by_ch = {}
        for pt in points:
            ch = pt[2]
            by_ch.setdefault(ch, {"x": [], "y": [], "trial": []})
            by_ch[ch]["x"].append(get_x(pt))
            by_ch[ch]["y"].append(pt[1])
            by_ch[ch]["trial"].append(pt[3])

        for d in (self._scatter_items, self._median_items, self._errbar_items):
            for ch in list(d):
                if ch not in by_ch:
                    self._plot_item.removeItem(d.pop(ch))

        for ch, data in by_ch.items():
            r, g, b, _ = self._ch_color(ch)
            x_arr = np.array(data["x"])
            y_arr = np.array(data["y"])

            # Individual dots
            if is_cat:
                ind_pen   = pg.mkPen(r, g, b, 204)  # outline only, alpha≈0.8
                ind_brush = pg.mkBrush(0, 0, 0, 0)   # transparent fill
            else:
                ind_pen   = pg.mkPen(None)
                ind_brush = pg.mkBrush(r, g, b, 200)

            self._scatter_data[ch] = [
                {"x_pos": x, "y_val": y, "trial": t}
                for x, y, t in zip(data["x"], data["y"], data["trial"])
            ]
            if ch not in self._scatter_items:
                sc = pg.ScatterPlotItem(size=9)
                self._plot_item.addItem(sc)
                self._scatter_items[ch] = sc
                sc.sigClicked.connect(
                    lambda _sc, pts, ev, c=ch: self._on_scatter_clicked(c, pts)
                )
            self._scatter_items[ch].setData(
                x=x_arr.tolist(), y=y_arr.tolist(),
                pen=ind_pen, brush=ind_brush,
            )

            if is_cat:
                # Compute median + IQR per x level
                unique_xs = np.unique(x_arr)
                med_x, med_y, tops, bots, level_data = [], [], [], [], []
                for ux in unique_xs:
                    vals = y_arr[x_arr == ux]
                    med      = float(np.median(vals))
                    q1       = float(np.percentile(vals, 25))
                    q3       = float(np.percentile(vals, 75))
                    mean_val = float(np.mean(vals))
                    sd_val   = float(np.std(vals, ddof=1)) if len(vals) > 1 else 0.0
                    orig = x_map_inv.get(int(ux), ux)
                    level_data.append({
                        "x_pos": float(ux), "x_val": orig,
                        "med": med, "q1": q1, "q3": q3,
                        "mean": mean_val, "sd": sd_val,
                    })
                    cy = mean_val if self._use_mean else med
                    med_x.append(float(ux))
                    med_y.append(cy)
                    if self._use_mean:
                        tops.append(sd_val)
                        bots.append(sd_val)
                    else:
                        tops.append(q3 - med)
                        bots.append(med - q1)
                self._median_data[ch] = level_data

                med_x_arr = np.array(med_x)
                med_y_arr = np.array(med_y)

                # Median dots
                if ch not in self._median_items:
                    sc_med = pg.ScatterPlotItem(size=14)
                    self._plot_item.addItem(sc_med)
                    self._median_items[ch] = sc_med
                self._median_items[ch].setData(
                    x=med_x, y=med_y,
                    pen=pg.mkPen(None),
                    brush=pg.mkBrush(r, g, b, 255),
                )

                # IQR error bars
                if ch not in self._errbar_items:
                    eb = pg.ErrorBarItem(
                        beam=0.15,
                        pen=pg.mkPen(r, g, b, 200, width=2),
                    )
                    self._plot_item.addItem(eb)
                    self._errbar_items[ch] = eb
                self._errbar_items[ch].setData(
                    x=med_x_arr, y=med_y_arr,
                    top=np.array(tops), bottom=np.array(bots),
                )
            else:
                for d in (self._median_items, self._errbar_items):
                    if ch in d:
                        self._plot_item.removeItem(d.pop(ch))

        self._plot_item.enableAutoRange(axis="y")
        self._update_legend(set(by_ch.keys()))
        self._update_plot_title()

    def _on_plot_mouse_moved(self, pos):
        if not self._median_data or not self._plot_item.sceneBoundingRect().contains(pos):
            self._hover_timer.stop()
            self._hover_text = ""
            QToolTip.hideText()
            return
        vb = self._plot_item.getViewBox()
        mp = vb.mapSceneToView(pos)
        mx, my = mp.x(), mp.y()
        vr = self._plot_item.viewRange()
        thresh_x = (vr[0][1] - vr[0][0]) * 0.04
        thresh_y = (vr[1][1] - vr[1][0]) * 0.06
        x_name = self._PLOT_COLS[self._combo_x.currentIndex()][0]
        ch_info = self._viewer.ch_info
        for ch, level_list in self._median_data.items():
            for d in level_list:
                if abs(mx - d["x_pos"]) < thresh_x and abs(my - d["med"]) < thresh_y:
                    x_val = int(d["x_val"]) if float(d["x_val"]) == int(d["x_val"]) else d["x_val"]
                    header = f"ch {ch}  |  {x_name} ({x_val})"
                    if self._use_mean:
                        self._hover_text = (
                            f"{header}\n"
                            f"Mean:   {d['mean']:.3f}\n"
                            f"SD:     {d['sd']:.3f}"
                        )
                    else:
                        self._hover_text = (
                            f"{header}\n"
                            f"Median: {d['med']:.3f}\n"
                            f"Q1:     {d['q1']:.3f}\n"
                            f"Q3:     {d['q3']:.3f}"
                        )
                    QToolTip.showText(QCursor.pos(), self._hover_text)
                    self._hover_timer.start()
                    return
        # Individual dot hover (smaller threshold, lower priority than median)
        thresh_xi = thresh_x * 0.6
        thresh_yi = thresh_y * 0.6
        for ch, pt_list in self._scatter_data.items():
            for pt in pt_list:
                if abs(mx - pt["x_pos"]) < thresh_xi and abs(my - pt["y_val"]) < thresh_yi:
                    self._hover_text = f"ch {ch}  |  trial {pt['trial']}"
                    QToolTip.showText(QCursor.pos(), self._hover_text)
                    self._hover_timer.start()
                    return

        self._hover_timer.stop()
        self._hover_text = ""
        QToolTip.hideText()

    def _refresh_tooltip(self):
        if self._hover_text:
            QToolTip.showText(QCursor.pos(), self._hover_text)

    def _update_plot_title(self):
        tab_idx = self._tab_bar.currentIndex()
        if tab_idx <= 0:
            ch_label = "All"
        else:
            ch = self._tab_channels[tab_idx - 1]
            ch_info = self._viewer.ch_info
            if ch_info and 0 < ch <= len(ch_info):
                ch_label = ch_info[ch - 1][0]
            else:
                ch_label = f"Ch{ch}"
        y_name = self._PLOT_COLS[self._combo_y.currentIndex()][0]
        x_name = self._PLOT_COLS[self._combo_x.currentIndex()][0]
        stat = "mean" if self._use_mean else "median"
        self._plot_item.setTitle(f"({ch_label})  {y_name} vs. {x_name}  ({stat})")

    def _on_scatter_clicked(self, ch, spots):
        if not spots or ch not in self._scatter_data:
            return
        idx   = spots[0].index()
        data  = self._scatter_data[ch]
        if idx >= len(data):
            return
        trial = data[idx]["trial"]
        ch_str    = str(ch)
        trial_str = str(trial)
        for row in range(self.merged_table.rowCount()):
            ti = self.merged_table.item(row, 0)
            ci = self.merged_table.item(row, 1)
            if ti and ci and ti.text() == trial_str and ci.text() == ch_str:
                self.merged_table.selectRow(row)
                self.merged_table.scrollTo(
                    self.merged_table.model().index(row, 0)
                )
                return

    def _on_table_selection_changed(self):
        selected = self.merged_table.selectionModel().selectedRows()
        if not selected:
            self._btn_exclude.setText("Exclude")
            return
        all_excluded = all(
            self.merged_table.item(idx.row(), len(self._COLS)) is not None
            and self.merged_table.item(idx.row(), len(self._COLS)).text() == "1"
            for idx in selected
        )
        self._btn_exclude.setText("Include" if all_excluded else "Exclude")

    def _on_exclude_clicked(self):
        selected = self.merged_table.selectionModel().selectedRows()
        include_mode = self._btn_exclude.text() == "Include"
        for idx in selected:
            row = idx.row()
            try:
                ch    = int(self.merged_table.item(row, 1).text())
                trial = int(self.merged_table.item(row, 0).text())
                if include_mode:
                    self._excluded.discard((ch, trial))
                else:
                    self._excluded.add((ch, trial))
            except (AttributeError, ValueError):
                pass
        self._populate_table()
        self._refresh_plot()

    def _on_clear_excluded(self):
        self._excluded.clear()
        self._populate_table()
        self._refresh_plot()

    def _on_toggle_stat_mode(self):
        self._use_mean = not self._use_mean
        self._btn_stat_mode.setText("Median" if self._use_mean else "Mean")
        self._refresh_plot()

    def _on_save_plot(self):
        from PyQt5.QtWidgets import QFileDialog
        tab_idx = self._tab_bar.currentIndex()
        if tab_idx <= 0:
            ch_label = "All"
        else:
            ch = self._tab_channels[tab_idx - 1]
            ch_info = self._viewer.ch_info
            ch_label = ch_info[ch - 1][0] if ch_info and 0 < ch <= len(ch_info) else f"Ch{ch}"
        y_name    = self._PLOT_COLS[self._combo_y.currentIndex()][0]
        x_name    = self._PLOT_COLS[self._combo_x.currentIndex()][0]
        default   = f"({ch_label}) {y_name} vs. {x_name}.png"
        path, _ = QFileDialog.getSaveFileName(
            self, "Save plot", default, "PNG image (*.png);;All files (*)"
        )
        if not path:
            return
        if not path.lower().endswith(".png"):
            path += ".png"
        self.canvas.grab().save(path)

    def _on_open_calculator(self):
        if not hasattr(self, "_calc_window") or self._calc_window is None:
            self._calc_window = CalculatorWindow(self)
        self._calc_window.show()
        self._calc_window.raise_()
        self._calc_window.activateWindow()

    def _on_y_changed(self, idx):
        name, col = self._PLOT_COLS[idx]
        self._y_col = col
        self._plot_item.setLabel("left", name)
        self._update_axis_exclusion()
        self._on_yaxis_auto()
        self._refresh_plot()

    def _on_x_changed(self, idx):
        name, col = self._PLOT_COLS[idx]
        self._x_col = col
        self._plot_item.setLabel("bottom", name)
        self._update_axis_exclusion()
        self._refresh_plot()

    def _update_axis_exclusion(self):
        y_idx = self._combo_y.currentIndex()
        x_idx = self._combo_x.currentIndex()
        for i in range(len(self._PLOT_COLS)):
            self._combo_y.model().item(i).setEnabled(i != x_idx)
            self._combo_x.model().item(i).setEnabled(i != y_idx)

    def _on_yaxis_fix(self):
        y_min, y_max = self._plot_item.getViewBox().viewRange()[1]
        self._plot_item.enableAutoRange(axis="y", enable=False)
        self._plot_item.setYRange(y_min, y_max, padding=0)
        self._btn_yfix.setChecked(True)
        self._btn_yauto.setChecked(False)

    def _on_yaxis_auto(self):
        self._plot_item.enableAutoRange(axis="y")
        self._btn_yfix.setChecked(False)
        self._btn_yauto.setChecked(True)

    def _update_legend(self, active_chs):
        if active_chs == self._legend_chs:
            return
        self._legend_chs = set(active_chs)

        while self._legend_layout.count():
            item = self._legend_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        for ch in sorted(active_chs):
            r, g, b, _ = self._ch_color(ch)
            dot = QLabel()
            dot.setFixedSize(12, 12)
            dot.setStyleSheet(
                f"background-color: rgb({r},{g},{b}); border-radius: 6px;"
            )
            self._legend_layout.addWidget(dot)
            self._legend_layout.addWidget(QLabel(f"Ch{ch}"))

    # ── Table ──────────────────────────────────────────────────────────

    def _on_tab_changed(self, _idx):
        self._refresh_plot()
        self._populate_table()

    def _populate_table(self):
        idx = self._tab_bar.currentIndex()
        if idx <= 0:
            rows = self._all_rows
        else:
            ch_filter = str(self._tab_channels[idx - 1])
            rows = [r for r in self._all_rows if r[1] == ch_filter]

        self.merged_table.setRowCount(0)
        for row_data in rows:
            r = self.merged_table.rowCount()
            self.merged_table.insertRow(r)
            try:
                is_excl = (int(row_data[1]), int(row_data[0])) in self._excluded
            except (ValueError, IndexError):
                is_excl = False
            for c, val in enumerate(row_data):
                item = QTableWidgetItem(val)
                item.setTextAlignment(Qt.AlignCenter)
                if is_excl:
                    item.setForeground(QColor(150, 150, 150))
                self.merged_table.setItem(r, c, item)
            excl_item = QTableWidgetItem("1" if is_excl else "0")
            excl_item.setTextAlignment(Qt.AlignCenter)
            if is_excl:
                excl_item.setForeground(QColor(150, 150, 150))
            self.merged_table.setItem(r, len(self._COLS), excl_item)
        self.merged_table.scrollToBottom()

    def resizeEvent(self, event):
        super().resizeEvent(event)
        if self._resizing:
            return
        self._resizing = True
        w = event.size().width()
        target_h = round(w * 9 / 16)
        if event.size().height() != target_h:
            self.resize(w, target_h)
        self._resizing = False

    def showEvent(self, event):
        super().showEvent(event)
        self._update_axis_exclusion()
        self.refresh()

    def closeEvent(self, event):
        if hasattr(self, "_calc_window") and self._calc_window is not None:
            self._calc_window.close()
        super().closeEvent(event)


class RealTimeViewer(QMainWindow):
    _tms_detected    = pyqtSignal(object)  # emitted from COM background thread, carries (record, tick)
    _magstim_polled  = pyqtSignal(object)  # emitted from Magstim poll thread, carries params dict

    def __init__(self):
        super().__init__()
        pythoncom.CoInitialize()
        self.client        = None
        self._com_thread   = None
        self.is_running    = False
        self._start_time   = None
        self._first_play   = False
        self._next_trial_num = 1
        self._tms_detected.connect(self._on_tms_trigger)
        self.ch_info       = []
        self.L = _Side(ch=DEFAULT_CH_LEFT,  color=COLOR_LEFT)
        self.R = _Side(ch=DEFAULT_CH_RIGHT, color=COLOR_RIGHT)
        self.trigger_times   = []
        self._last_trigger_t = None
        self.separate = False
        self._magstim_mode = "SM"
        self._magstim_port = None
        self._magstim_stop = threading.Event()
        self._magstim_last_poll_t = 0.0
        self._mso1_val = 0
        self._mso2_val = 0
        self._isi_val  = "—"
        self._temp_val = "—"
        self._magstim_polled.connect(self._on_magstim_polled)
        self._plot_window = None
        self.window_secs   = WINDOW_SECS_DEFAULT
        self.scope_pre     = SCOPE_PRE_DEFAULT
        self.scope_post    = SCOPE_POST_DEFAULT
        self.mode          = "chart"   # "chart" | "scope"

        pg.setConfigOption("background", "w")
        pg.setConfigOption("foreground", "k")

        self._setup_ui()
        self._connect_magstim()

        self.data_timer = QTimer()
        self.data_timer.timeout.connect(self._update)

        self.status_timer = QTimer()
        self.status_timer.timeout.connect(self._refresh_status)
        self.status_timer.start(STATUS_MS)
        self._refresh_status()

    # ------------------------------------------------------------------
    # UI
    # ------------------------------------------------------------------

    def _setup_ui(self):
        self.setWindowTitle("TMSviewer")
        self.setMinimumSize(1800, 1013)

        central = self._make_central_widget()
        self.setCentralWidget(central)

        root = QVBoxLayout(central)
        root.setContentsMargins(10, 10, 10, 10)
        root.setSpacing(8)

        root.addLayout(self._make_mode_row())
        root.addWidget(self._make_analysis_stack())
        root.addLayout(self._make_channel_row())
        root.addLayout(self._make_plots_row(), stretch=1)
        root.addWidget(self._make_scope_below_stack())

        self.lbl_status = QLabel("Connecting to LabChart...")
        self.lbl_status.setAlignment(Qt.AlignCenter)
        _sf = QFont(); _sf.setPointSize(10)
        self.lbl_status.setFont(_sf)
        root.addWidget(self.lbl_status)

        root.addLayout(self._make_btn_row())
        self._on_mode_changed()

    def _make_central_widget(self):
        central = QWidget()
        central.setStyleSheet("""
            background-color: white;
            QPushButton {
                background-color: #a0a0a0;
                border: 1px solid #808080;
                border-radius: 4px;
                padding: 3px 8px;
            }
            QPushButton:hover { background-color: #909090; }
            QPushButton:pressed { background-color: #787878; }
            QPushButton:disabled { background-color: #d0d0d0; color: #aaa; }
            QPushButton:checked { background-color: #787878; }
        """)
        _app_font = QApplication.font()
        _pt = _app_font.pointSize()
        if _pt > 0:
            _scaled = QFont(_app_font)
            _scaled.setPointSize(round(_pt * 1.1))
            central.setFont(_scaled)
        elif _app_font.pixelSize() > 0:
            _scaled = QFont(_app_font)
            _scaled.setPixelSize(round(_app_font.pixelSize() * 1.1))
            central.setFont(_scaled)
        return central

    def _make_mode_row(self):
        mode_row = QHBoxLayout()
        mode_row.setSpacing(12)

        mode_row.addWidget(QLabel("Mode:"))
        self.radio_chart = QRadioButton("Chart")
        self.radio_chart.setEnabled(False)
        self.radio_scope = QRadioButton("Scope")
        self.radio_scope.setChecked(True)
        mode_group = QButtonGroup(self)
        mode_group.addButton(self.radio_chart)
        mode_group.addButton(self.radio_scope)
        self.radio_chart.toggled.connect(self._on_mode_changed)
        mode_row.addWidget(self.radio_chart)
        mode_row.addWidget(self.radio_scope)
        mode_row.addSpacing(20)

        self.widget_chart_params = QWidget()
        chart_p = QHBoxLayout(self.widget_chart_params)
        chart_p.setContentsMargins(0, 0, 0, 0)
        chart_p.setSpacing(6)
        chart_p.addWidget(QLabel("Window:"))
        self.spinbox_window = QSpinBox()
        self.spinbox_window.setMinimum(WINDOW_SECS_MIN)
        self.spinbox_window.setMaximum(WINDOW_SECS_MAX)
        self.spinbox_window.setValue(WINDOW_SECS_DEFAULT)
        self.spinbox_window.setSuffix(" s")
        self.spinbox_window.setFixedWidth(63)
        chart_p.addWidget(self.spinbox_window)
        btn_chart_apply = QPushButton("Apply")
        btn_chart_apply.setFixedWidth(50)
        btn_chart_apply.clicked.connect(self._apply_chart_params)
        chart_p.addWidget(btn_chart_apply)
        mode_row.addWidget(self.widget_chart_params)

        self.widget_scope_params = QWidget()
        scope_p = QHBoxLayout(self.widget_scope_params)
        scope_p.setContentsMargins(0, 0, 0, 0)
        scope_p.setSpacing(6)
        scope_p.addWidget(QLabel("Pre:"))
        self.spin_pre = QDoubleSpinBox()
        self.spin_pre.setRange(0.1, 5.0)
        self.spin_pre.setSingleStep(0.1)
        self.spin_pre.setValue(SCOPE_PRE_DEFAULT)
        self.spin_pre.setSuffix(" s")
        self.spin_pre.setFixedWidth(63)
        scope_p.addWidget(self.spin_pre)
        scope_p.addWidget(QLabel("Post:"))
        self.spin_post = QDoubleSpinBox()
        self.spin_post.setRange(0.1, 10.0)
        self.spin_post.setSingleStep(0.1)
        self.spin_post.setValue(SCOPE_POST_DEFAULT)
        self.spin_post.setSuffix(" s")
        self.spin_post.setFixedWidth(63)
        scope_p.addWidget(self.spin_post)
        btn_scope_apply = QPushButton("Apply")
        btn_scope_apply.setFixedWidth(50)
        btn_scope_apply.clicked.connect(self._apply_scope_params)
        scope_p.addWidget(btn_scope_apply)
        self.widget_scope_params.setVisible(False)
        mode_row.addWidget(self.widget_scope_params)

        mode_row.addSpacing(20)
        mode_row.addWidget(QLabel("Paralysis:"))
        self._paralysis_combo = QComboBox()
        self._paralysis_combo.addItems(["NA", "LEFT", "RIGHT"])
        mode_row.addWidget(self._paralysis_combo)

        mode_row.addStretch()
        return mode_row

    def _make_analysis_stack(self):
        w_scope_analysis = QWidget()
        analysis_row = QHBoxLayout(w_scope_analysis)
        analysis_row.setContentsMargins(0, 0, 0, 0)
        analysis_row.setSpacing(12)

        self.radio_none = QRadioButton("None")
        self.radio_mep  = QRadioButton("MEP")
        self.radio_mep.setChecked(True)
        self.radio_none.setEnabled(False)
        self.radio_mep.setEnabled(False)
        analysis_group = QButtonGroup(self)
        analysis_group.addButton(self.radio_none)
        analysis_group.addButton(self.radio_mep)
        self.radio_mep.toggled.connect(self._on_analysis_mode_changed)
        analysis_row.addWidget(self.radio_none)
        analysis_row.addWidget(self.radio_mep)
        analysis_row.addStretch()

        self.stack_analysis = _MaxSizeStack()
        self.stack_analysis.addWidget(w_scope_analysis)
        self.stack_analysis.addWidget(QWidget())
        return self.stack_analysis

    def _make_channel_row(self):
        combos_row = QHBoxLayout()
        combos_row.setSpacing(8)
        self.L.combo = QComboBox()
        self.R.combo = QComboBox()
        self.L.combo.currentIndexChanged.connect(lambda idx: self._on_combo("left",  idx))
        self.R.combo.currentIndexChanged.connect(lambda idx: self._on_combo("right", idx))
        combos_row.addWidget(QLabel("Left channel:"))
        combos_row.addWidget(self.L.combo, stretch=1)
        combos_row.addSpacing(150)
        combos_row.addWidget(QLabel("Right channel:"))
        combos_row.addWidget(self.R.combo, stretch=1)
        return combos_row

    def _make_centre_column(self):
        self.centre_widget = QWidget()
        centre = QVBoxLayout(self.centre_widget)
        centre.setSpacing(4)
        centre.setContentsMargins(4, 0, 4, 0)

        lbl_pages = QLabel("Trial #")
        lbl_pages.setAlignment(Qt.AlignCenter)
        self.trigger_spin = QSpinBox()
        self.trigger_spin.setRange(0, 9999)
        self.trigger_spin.setValue(0)
        self.trigger_spin.setButtonSymbols(QAbstractSpinBox.NoButtons)
        font_big = QFont()
        font_big.setPointSize(28)
        font_big.setBold(True)
        self.trigger_spin.setFont(font_big)
        self.trigger_spin.setAlignment(Qt.AlignCenter)
        self.trigger_spin.valueChanged.connect(self._on_trigger_selected)
        centre.addWidget(lbl_pages)
        centre.addWidget(self.trigger_spin)
        centre.addSpacing(12)

        self.lbl_arrow_1 = QLabel("<--->")
        self.lbl_arrow_1.setAlignment(Qt.AlignCenter)
        arrow_font = QFont()
        arrow_font.setPointSize(14)
        arrow_font.setBold(True)
        self.lbl_arrow_1.setFont(arrow_font)
        centre.addWidget(self.lbl_arrow_1)
        self.widget_mep_params = self._make_mep_params_widget(second=False)
        centre.addWidget(self.widget_mep_params)
        centre.addSpacing(6)

        btn_reset = QPushButton("Reset")
        btn_reset.clicked.connect(self._reset_mep_params)
        centre.addWidget(btn_reset)

        self.btn_separate = QPushButton("Separate")
        self.btn_separate.setCheckable(True)
        self.btn_separate.toggled.connect(self._on_separate_toggled)
        centre.addWidget(self.btn_separate)
        centre.addSpacing(6)

        self.lbl_arrow_2 = QLabel("--->")
        self.lbl_arrow_2.setAlignment(Qt.AlignCenter)
        self.lbl_arrow_2.setFont(arrow_font)
        centre.addWidget(self.lbl_arrow_2)
        self.widget_mep_params_2 = self._make_mep_params_widget(second=True)
        self.widget_mep_params_2.setEnabled(False)
        centre.addWidget(self.widget_mep_params_2)

        self._fx_arrow_2  = QGraphicsOpacityEffect(); self._fx_arrow_2.setOpacity(0)
        self._fx_params_2 = QGraphicsOpacityEffect(); self._fx_params_2.setOpacity(0)
        self.lbl_arrow_2.setGraphicsEffect(self._fx_arrow_2)
        self.widget_mep_params_2.setGraphicsEffect(self._fx_params_2)

        centre.addStretch()
        return self.centre_widget

    def _make_plots_row(self):
        self.L.plot  = self._make_plot("Left EMG",  COLOR_LEFT)
        self.R.plot  = self._make_plot("Right EMG", COLOR_RIGHT)
        self.L.curve = self.L.plot.plot(pen=pg.mkPen(COLOR_LEFT,  width=1))
        self.R.curve = self.R.plot.plot(pen=pg.mkPen(COLOR_RIGHT, width=1))

        # _make_side_panel assigns s.panel, s.stats, s.btn_yfix, s.btn_yauto into self.L / self.R
        self._make_side_panel("left")
        self._make_side_panel("right")

        self.stack_centre = _MaxSizeStack()
        self.stack_centre.setFixedWidth(135)
        self.stack_centre.addWidget(self._make_centre_column())
        self.stack_centre.addWidget(QWidget())

        plots_row = QHBoxLayout()
        plots_row.setSpacing(8)
        plots_row.addWidget(self.L.panel)
        plots_row.addWidget(self.L.plot,  stretch=1)
        plots_row.addWidget(self.stack_centre)
        plots_row.addWidget(self.R.plot,  stretch=1)
        plots_row.addWidget(self.R.panel)
        return plots_row

    def _make_mso_row(self):
        mso_row  = QHBoxLayout()
        mso_row.setSpacing(12)
        mso_font  = QFont(); mso_font.setPointSize(18)
        btn_font  = QFont(); btn_font.setPointSize(16)
        mode_font = QFont(); mode_font.setPointSize(18); mode_font.setBold(True)

        mso_row.addStretch(1)

        self.btn_sm = QPushButton("SM"); self.btn_sm.setFont(mode_font)
        self.btn_sm.setFixedWidth(64)
        self.btn_sm.setStyleSheet("background-color: #1565C0; color: white;")
        self.btn_sm.clicked.connect(self._on_sm_clicked)
        mso_row.addWidget(self.btn_sm)
        self.btn_dm = QPushButton("DM"); self.btn_dm.setFont(mode_font)
        self.btn_dm.setFixedWidth(64)
        self.btn_dm.clicked.connect(self._on_dm_clicked)
        mso_row.addWidget(self.btn_dm)
        dot_font = QFont(); dot_font.setPointSize(15)
        self.lbl_magstim_dot = QLabel("●"); self.lbl_magstim_dot.setFont(dot_font)
        self.lbl_magstim_dot.setStyleSheet("color: #aaaaaa;")
        self.lbl_magstim_dot.setToolTip(MAGSTIM_PORT)
        mso_row.addWidget(self.lbl_magstim_dot)
        status_font = QFont(); status_font.setPointSize(13); status_font.setBold(True)
        self.lbl_magstim_state = QLabel("—"); self.lbl_magstim_state.setFont(status_font)
        self.lbl_magstim_state.setStyleSheet("color: #aaaaaa;")
        mso_row.addWidget(self.lbl_magstim_state)
        temp_font = QFont(); temp_font.setPointSize(13)
        self.lbl_magstim_temp = QLabel("—"); self.lbl_magstim_temp.setFont(temp_font)
        self.lbl_magstim_temp.setToolTip("Coil temp unavailable")
        mso_row.addWidget(self.lbl_magstim_temp)
        mso_row.addSpacing(8)

        lbl_mso1 = QLabel("MSO1:"); lbl_mso1.setFont(mso_font)
        mso_row.addWidget(lbl_mso1)
        self.lbl_mso_val = QLabel("—"); self.lbl_mso_val.setFont(mso_font)
        self.lbl_mso_val.setMinimumWidth(45)
        mso_row.addWidget(self.lbl_mso_val)
        mso_row.addSpacing(16)

        self.lbl_mso2 = QLabel("MSO2:"); self.lbl_mso2.setFont(mso_font)
        self.lbl_mso2.setEnabled(False)
        mso_row.addWidget(self.lbl_mso2)
        self.lbl_mso2_val = QLabel("—"); self.lbl_mso2_val.setFont(mso_font)
        self.lbl_mso2_val.setMinimumWidth(45)
        self.lbl_mso2_val.setEnabled(False)
        mso_row.addWidget(self.lbl_mso2_val)

        isi_font = QFont(); isi_font.setPointSize(13)
        self.lbl_ipi = QLabel("  ISI:"); self.lbl_ipi.setFont(isi_font)
        self.lbl_ipi.setEnabled(False)
        mso_row.addWidget(self.lbl_ipi)
        self.lbl_isi_val = QLabel("—"); self.lbl_isi_val.setFont(isi_font)
        self.lbl_isi_val.setEnabled(False)
        self.lbl_isi_val.setMinimumWidth(80)
        mso_row.addWidget(self.lbl_isi_val)
        mso_row.addSpacing(16)

        lbl_loc = QLabel("Loc ID:"); lbl_loc.setFont(mso_font)
        mso_row.addWidget(lbl_loc)
        self.spin_location = _WheelSpinBox()
        self.spin_location.setRange(1, 9999); self.spin_location.setValue(1)
        self.spin_location.setFixedWidth(100); self.spin_location.setFont(mso_font)
        self.spin_location.setButtonSymbols(QAbstractSpinBox.UpDownArrows)
        mso_row.addWidget(self.spin_location)

        mso_row.addStretch(1)
        btn_new_best = QPushButton("New Best"); btn_new_best.setFont(btn_font)
        btn_new_best.clicked.connect(self._on_new_best)
        mso_row.addWidget(btn_new_best)
        btn_equal_best = QPushButton("Equal Best"); btn_equal_best.setFont(btn_font)
        btn_equal_best.clicked.connect(self._on_equal_best)
        mso_row.addWidget(btn_equal_best)
        self.combo_best = QComboBox(); self.combo_best.setFont(btn_font)
        self.combo_best.setMinimumWidth(180)
        mso_row.addWidget(self.combo_best)
        return mso_row

    def _make_scope_below_stack(self):
        w_scope_below = QWidget()
        scope_vbox = QVBoxLayout(w_scope_below)
        scope_vbox.setContentsMargins(0, 0, 0, 0)
        scope_vbox.setSpacing(8)

        hunt_row = QHBoxLayout()
        hunt_row.setSpacing(8)
        hunt_row.addSpacing(81)
        hunt_row.addLayout(self._make_hunt_panel("left"))
        hunt_row.addSpacing(135)
        hunt_row.addLayout(self._make_hunt_panel("right"))
        hunt_row.addSpacing(81)
        scope_vbox.addLayout(hunt_row)

        sep = QFrame(); sep.setFrameShape(QFrame.HLine); sep.setFrameShadow(QFrame.Sunken)
        scope_vbox.addWidget(sep)

        scope_vbox.addLayout(self._make_mso_row())

        sep2 = QFrame(); sep2.setFrameShape(QFrame.HLine); sep2.setFrameShadow(QFrame.Sunken)
        scope_vbox.addWidget(sep2)

        tables_row = QHBoxLayout(); tables_row.setSpacing(8)
        self.L.table = self._make_data_table()
        self.R.table = self._make_data_table()

        btn_plot = QPushButton("Plot"); btn_plot.setFixedWidth(130)
        btn_plot.clicked.connect(self._on_plot_clicked)
        btn_save = QPushButton("Save"); btn_save.setFixedWidth(130)
        btn_save.clicked.connect(self._save_tables)
        btn_clear_all = QPushButton("Clear"); btn_clear_all.setFixedWidth(130)
        btn_clear_all.clicked.connect(self._on_clear_all)
        centre_btns = QVBoxLayout()
        centre_btns.setSpacing(6)
        centre_btns.addWidget(btn_plot)
        centre_btns.addWidget(btn_save)
        centre_btns.addWidget(btn_clear_all)

        tables_row.addSpacing(81)
        tables_row.addWidget(self.L.table, stretch=1)
        tables_row.addLayout(centre_btns)
        tables_row.addWidget(self.R.table, stretch=1)
        tables_row.addSpacing(81)
        scope_vbox.addLayout(tables_row)

        self.stack_below = _MaxSizeStack()
        self.stack_below.addWidget(w_scope_below)
        self.stack_below.addWidget(QWidget())
        return self.stack_below

    def _make_btn_row(self):
        btn_row = QHBoxLayout()
        btn_row.setSpacing(8)
        _bf = QFont(); _bf.setPointSize(13)

        self.btn = QPushButton("▶  Play")
        self.btn.setFixedHeight(43)
        self.btn.setFont(_bf)
        self.btn.setFocusPolicy(Qt.NoFocus)
        self.btn.clicked.connect(self._toggle)

        self.btn_sample = QPushButton("Sample")
        self.btn_sample.setFixedHeight(43)
        self.btn_sample.setFont(_bf)
        self.btn_sample.setFocusPolicy(Qt.NoFocus)
        self.btn_sample.setEnabled(False)
        self.btn_sample.setStyleSheet(
            "QPushButton:enabled  { background-color: #2e7d32; color: white; }"
            "QPushButton:disabled { background-color: #cccccc; color: #888888; }"
        )
        self.btn_sample.clicked.connect(self._add_sample_comment)

        btn_row.addWidget(self.btn, stretch=1)
        btn_row.addWidget(self.btn_sample, stretch=1)
        return btn_row

    def _make_mep_params_widget(self, second=False):
        widget = QWidget()
        mp = QVBoxLayout(widget)
        mp.setContentsMargins(0, 0, 0, 0)
        mp.setSpacing(3)

        mep_start = QSpinBox();  mep_start.setRange(1, 500);   mep_start.setValue(10);  mep_start.setButtonSymbols(QAbstractSpinBox.NoButtons)
        mep_end   = QSpinBox();  mep_end.setRange(1, 500);     mep_end.setValue(50);    mep_end.setSuffix(" ms");   mep_end.setButtonSymbols(QAbstractSpinBox.NoButtons)
        threshold = QDoubleSpinBox(); threshold.setRange(0.01, 100); threshold.setSingleStep(0.01)
        threshold.setDecimals(2); threshold.setValue(0.05); threshold.setSuffix(" mV"); threshold.setButtonSymbols(QAbstractSpinBox.NoButtons)
        prestim_start = QSpinBox(); prestim_start.setRange(-1000, -1); prestim_start.setValue(-200); prestim_start.setButtonSymbols(QAbstractSpinBox.NoButtons)
        prestim_end   = QSpinBox(); prestim_end.setRange(-500, -1);   prestim_end.setValue(-50); prestim_end.setSuffix(" ms"); prestim_end.setButtonSymbols(QAbstractSpinBox.NoButtons)

        mp.addWidget(QLabel("MEP window"))
        mep_row = QHBoxLayout(); mep_row.setSpacing(3)
        mep_row.addWidget(mep_start, stretch=1); mep_row.addWidget(QLabel("—")); mep_row.addWidget(mep_end, stretch=1)
        mp.addLayout(mep_row)
        mp.addSpacing(4)
        mp.addWidget(QLabel("Threshold"))
        mp.addWidget(threshold)
        mp.addSpacing(4)
        mp.addWidget(QLabel("Prestim window"))
        pre_row = QHBoxLayout(); pre_row.setSpacing(3)
        pre_row.addWidget(prestim_start, stretch=1); pre_row.addWidget(QLabel("—")); pre_row.addWidget(prestim_end, stretch=1)
        mp.addLayout(pre_row)

        if second:
            self.spin_mep_start_2    = mep_start
            self.spin_mep_end_2      = mep_end
            self.spin_threshold_2    = threshold
            self.spin_prestim_start_2 = prestim_start
            self.spin_prestim_end_2  = prestim_end
        else:
            self.spin_mep_start      = mep_start
            self.spin_mep_end        = mep_end
            self.spin_threshold      = threshold
            self.spin_prestim_start  = prestim_start
            self.spin_prestim_end    = prestim_end

        return widget

    def _make_hunt_panel(self, side):
        panel = QWidget()
        panel.setEnabled(False)  # disabled until MEP mode is on
        layout = QHBoxLayout(panel)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)

        font = QFont()
        font.setPointSize(12)

        chk = QCheckBox("Hunt")
        chk.setFont(font)

        lbl_num   = QLabel("0")
        lbl_slash = QLabel("/")
        lbl_den   = QLabel("0")
        for lbl in (lbl_num, lbl_slash, lbl_den):
            lbl.setFont(font)
            lbl.setAlignment(Qt.AlignCenter)

        btn_clear = QPushButton("Clear")
        btn_redo  = QPushButton("Redo")
        for btn in (btn_clear, btn_redo):
            btn.setFont(font)

        sub_widgets = (lbl_num, lbl_slash, lbl_den, btn_clear, btn_redo)
        for w in sub_widgets:
            w.setEnabled(False)

        chk.toggled.connect(lambda checked, ws=sub_widgets:
                            [w.setEnabled(checked) for w in ws])

        layout.addWidget(chk)
        layout.addWidget(lbl_num)
        layout.addWidget(lbl_slash)
        layout.addWidget(lbl_den)
        layout.addWidget(btn_clear)
        layout.addWidget(btn_redo)

        s = self.L if side == "left" else self.R
        s.hunt_panel   = panel
        s.hunt_chk     = chk
        s.hunt_num_lbl = lbl_num
        s.hunt_den_lbl = lbl_den
        s.hunt_clear   = btn_clear
        s.hunt_redo    = btn_redo
        btn_clear.clicked.connect(lambda: self._on_hunt_clear(side))
        btn_redo.clicked.connect(lambda: self._on_hunt_redo(side))

        outer = QHBoxLayout()
        outer.setContentsMargins(0, 0, 0, 0)
        outer.addWidget(panel)
        outer.addStretch()
        return outer

    def _on_plot_clicked(self):
        if self._plot_window is None or not self._plot_window.isVisible():
            self._plot_window = PlotWindow(self)
        self._plot_window.refresh()
        self._plot_window.show()
        self._plot_window.raise_()
        self._plot_window.activateWindow()

    def _sync_plot_table(self):
        if self._plot_window is not None and self._plot_window.isVisible():
            self._plot_window.refresh()

    def _save_tables(self):
        import csv
        from PyQt5.QtWidgets import QFileDialog
        desktop = os.path.join(os.path.expanduser("~"), "Desktop")
        path, _ = QFileDialog.getSaveFileName(
            self, "Save Data", os.path.join(desktop, "tms_data.csv"),
            "CSV files (*.csv)"
        )
        if not path:
            return
        if not path.lower().endswith(".csv"):
            path += ".csv"

        tbl = self.L.table
        cols = [tbl.horizontalHeaderItem(c).text() for c in range(tbl.columnCount())]
        rows = []
        for src in (self.L.table, self.R.table):
            for r in range(src.rowCount()):
                rows.append([
                    (src.item(r, c).text() if src.item(r, c) else "")
                    for c in range(src.columnCount())
                ])
        rows.sort(key=lambda row: (int(row[0]) if row[0].isdigit() else 0, int(row[1]) if row[1].isdigit() else 0))

        with open(path, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(cols)
            writer.writerows(rows)

    def _on_clear_all(self):
        from PyQt5.QtWidgets import QMessageBox
        msg = QMessageBox(self)
        msg.setWindowTitle("Clear")
        msg.setText("This will clear all data. Are you sure?")
        msg.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
        msg.setDefaultButton(QMessageBox.No)
        if msg.exec_() != QMessageBox.Yes:
            return

        if self.is_running:
            self._stop()

        for s in (self.L, self.R):
            s.table.setRowCount(0)
            s.hunt_num = 0
            s.hunt_den = 0
            s.hunt_history.clear()
            s.hunt_num_lbl.setText("0")
            s.hunt_den_lbl.setText("0")
            s.curve.setData([], [])

        self._clear_trigger_lines()
        self.trigger_times.clear()
        self._next_trial_num = 1
        self.trigger_spin.blockSignals(True)
        self.trigger_spin.setValue(0)
        self.trigger_spin.setMaximum(9999)
        self.trigger_spin.blockSignals(False)

        if self._plot_window is not None:
            self._plot_window.close()
            self._plot_window = None

    def _make_data_table(self):
        cols = ["Trial", "CH", "Hand", "Paralysis", "Loc_ID", "MSO1", "MSO2", "ISI", "MEP_amp", "Prestim_mean", "Temp", "Mode", "Excl"]
        tbl = QTableWidget(0, len(cols))
        tbl.setHorizontalHeaderLabels(cols)
        tbl.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        tbl.horizontalHeader().setStretchLastSection(False)
        tbl.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        tbl.verticalHeader().setVisible(False)
        tbl.setEditTriggers(QTableWidget.NoEditTriggers)
        tbl.setSelectionBehavior(QTableWidget.SelectRows)
        tbl.setFixedHeight(162)
        return tbl

    def _make_side_panel(self, side):
        """Returns (outer_widget, stats_label) — Y-zoom controls + MEP/RMS stats."""
        outer = QWidget()
        outer.setFixedWidth(88)
        vbox = QVBoxLayout(outer)
        vbox.setContentsMargins(2, 4, 2, 0)
        vbox.setSpacing(3)

        btn_font = QFont()
        btn_font.setPointSize(10)
        btn_font.setBold(True)

        small_font = QFont()
        small_font.setPointSize(9)

        btn_plus  = QPushButton("+")
        btn_minus = QPushButton("−")
        btn_fix   = QPushButton("Fix")
        btn_auto  = QPushButton("Auto")

        for btn in (btn_plus, btn_minus):
            btn.setFont(btn_font)
            btn.setFixedHeight(28)
        for btn in (btn_fix, btn_auto):
            btn.setFont(small_font)
            btn.setFixedHeight(24)
            btn.setCheckable(True)

        btn_auto.setChecked(True)

        stats = QLabel("Prestim RMS\n—\n\nMEP amp\n—")
        stats.setAlignment(Qt.AlignCenter)
        stats.setFont(small_font)
        stats.setWordWrap(True)

        vbox.addWidget(btn_plus)
        vbox.addWidget(btn_minus)
        vbox.addSpacing(4)
        vbox.addWidget(btn_fix)
        vbox.addWidget(btn_auto)
        vbox.addSpacing(16)
        vbox.addWidget(stats)
        vbox.addStretch()

        s = self.L if side == "left" else self.R
        s.btn_yfix  = btn_fix
        s.btn_yauto = btn_auto
        s.panel = outer
        s.stats = stats
        btn_plus.clicked.connect(lambda: self._yzoom(side, 1 / 1.25))
        btn_minus.clicked.connect(lambda: self._yzoom(side, 1.25))
        btn_fix.clicked.connect(lambda: self._yfix_clicked(side))
        btn_auto.clicked.connect(lambda: self._yauto_clicked(side))

    # ------------------------------------------------------------------
    # Y-axis zoom / fix / auto
    # ------------------------------------------------------------------

    def _yzoom(self, side, factor):
        s = self.L if side == "left" else self.R
        if not s.yfix:
            vr = s.plot.viewRange()
            s.yhalf = max(abs(vr[1][0]), abs(vr[1][1]), 1e-6)
            s.yfix = True
            s.btn_yfix.setChecked(True)
            s.btn_yauto.setChecked(False)
        s.yhalf = max(s.yhalf * factor, 1e-6)
        s.plot.enableAutoRange(axis="y", enable=False)
        s.plot.setYRange(-s.yhalf, s.yhalf, padding=0)

    def _yfix_clicked(self, side):
        s = self.L if side == "left" else self.R
        vr = s.plot.viewRange()
        s.yhalf = max(abs(vr[1][0]), abs(vr[1][1]), 1e-6)
        s.yfix = True
        s.btn_yfix.setChecked(True)
        s.btn_yauto.setChecked(False)
        s.plot.enableAutoRange(axis="y", enable=False)
        s.plot.setYRange(-s.yhalf, s.yhalf, padding=0)

    def _yauto_clicked(self, side):
        s = self.L if side == "left" else self.R
        s.yfix = False
        s.btn_yfix.setChecked(False)
        s.btn_yauto.setChecked(True)

    # ------------------------------------------------------------------

    def _analyse_epoch(self, arr, second=False):
        """Compute MEP amplitude and prestim RMS for one epoch array.

        Returns {"mep_amp": float, "prestim_rms": float}, or None if arr is empty.
        """
        if arr is None or len(arr) == 0:
            return None
        if second:
            mep_s = self.spin_mep_start_2.value()     / 1000.0
            mep_e = self.spin_mep_end_2.value()       / 1000.0
            pre_s = self.spin_prestim_start_2.value() / 1000.0
            pre_e = self.spin_prestim_end_2.value()   / 1000.0
        else:
            mep_s = self.spin_mep_start.value()     / 1000.0
            mep_e = self.spin_mep_end.value()       / 1000.0
            pre_s = self.spin_prestim_start.value() / 1000.0
            pre_e = self.spin_prestim_end.value()   / 1000.0
        n     = len(arr)
        dur   = self.scope_pre + self.scope_post
        spt   = dur / max(n - 1, 1)
        t_rel = np.arange(n) * spt - self.scope_pre
        mep_data = arr[(t_rel >= mep_s) & (t_rel <= mep_e)]
        pre_data = arr[(t_rel >= pre_s) & (t_rel <= pre_e)]
        mep_amp     = float(np.nanmax(mep_data) - np.nanmin(mep_data)) if mep_data.size > 0 else 0.0
        prestim_rms = float(np.sqrt(np.nanmean(pre_data ** 2)))        if pre_data.size > 0 else 0.0
        return {"mep_amp": mep_amp, "prestim_rms": prestim_rms}

    def _update_stats(self, lbl, arr, second=False):
        r = self._analyse_epoch(arr, second)
        if r is None:
            lbl.setText("Prestim RMS\n—\n\nMEP amp\n—")
            return
        lbl.setText(
            f"Prestim RMS\n{r['prestim_rms']:.3f} mV\n\nMEP amp\n{r['mep_amp']:.3f} mV"
        )

    def _make_plot(self, title, color):
        plot = pg.PlotWidget()
        plot.setTitle(f"<b>{title}</b>", color=color, size="11pt")
        plot.setLabel("left", "mV")
        plot.setLabel("bottom", "Time (s)")
        plot.showGrid(x=False, y=False)
        plot.getAxis("left").setPen(pg.mkPen("k"))
        plot.getAxis("bottom").setPen(pg.mkPen("k"))
        plot.enableAutoRange(axis="y", enable=True)
        plot.setStyleSheet("border: 2px solid black;")
        return plot

    # ------------------------------------------------------------------
    # Mode switching
    # ------------------------------------------------------------------

    # ------------------------------------------------------------------

    def _on_mode_changed(self):
        chart = self.radio_chart.isChecked()
        self.mode = "chart" if chart else "scope"
        page = 1 if chart else 0
        self.stack_analysis.setCurrentIndex(page)
        self.stack_centre.setCurrentIndex(page)
        self.stack_below.setCurrentIndex(page)
        for s in (self.L, self.R):
            s.panel.setVisible(not chart)
            s.plot.enableAutoRange(axis="y", enable=chart)
        self.widget_chart_params.setVisible(chart)
        self.widget_scope_params.setVisible(not chart)
        self.trigger_spin.setEnabled(not chart)

        self.radio_none.setEnabled(not chart)
        self.radio_mep.setEnabled(not chart)
        mep_active = (not chart) and self.radio_mep.isChecked()
        for s in (self.L, self.R):
            s.hunt_panel.setEnabled(mep_active)

        x_label = "Time (s)" if chart else "Time from trigger (s)"
        for s in (self.L, self.R):
            s.plot.setLabel("bottom", x_label)
            s.curve.setData([], [])
        self._clear_trigger_lines()

        if not chart and self.trigger_times:
            channels = list({self.L.ch, self.R.ch})
            self._update_scope(channels, selected_row=self.trigger_spin.value() - 1)

    # ------------------------------------------------------------------
    # Parameter apply
    # ------------------------------------------------------------------

    def _on_analysis_mode_changed(self):
        mep = self.radio_mep.isChecked()
        self.widget_mep_params.setEnabled(mep)
        for s in (self.L, self.R):
            s.hunt_panel.setEnabled(mep)
        if self.radio_scope.isChecked():
            self._update_scope([self.L.ch, self.R.ch])

    def _update_hunt_display(self, side):
        s = self.L if side == "left" else self.R
        s.hunt_num_lbl.setText(str(s.hunt_num))
        s.hunt_den_lbl.setText(str(s.hunt_den))

    def _set_table_cell(self, tbl, row, col, text):
        if row < tbl.rowCount():
            item = QTableWidgetItem(text)
            item.setTextAlignment(Qt.AlignCenter)
            tbl.setItem(row, col, item)

    def _both_hunt_active(self):
        return all(s.hunt_panel.isEnabled() and s.hunt_chk.isChecked() for s in (self.L, self.R))

    def _get_hand(self, s):
        if self.ch_info and s.ch < len(self.ch_info):
            name = self.ch_info[s.ch][0].lower()
            if "right" in name:
                return "RH"
            if "left" in name:
                return "LH"
        return "NA"

    def _get_paralysis(self, s):
        ph_side = self._paralysis_combo.currentText()  # "NA", "LEFT", "RIGHT"
        if ph_side == "NA":
            return "NA"
        hand = self._get_hand(s)
        if hand == "NA":
            return "NA"
        hand_side = "LEFT" if hand == "LH" else "RIGHT"
        return "PH" if hand_side == ph_side else "NPH"

    def _on_hunt_clear(self, side):
        targets = (self.L, self.R) if self._both_hunt_active() else (self.L if side == "left" else self.R,)
        for s in targets:
            s.hunt_num = 0; s.hunt_den = 0
            s.hunt_history.clear()
            s.hunt_latest_eval = True
            self._update_hunt_display("left" if s is self.L else "right")

    def _on_hunt_redo(self, side):
        targets = (self.L, self.R) if self._both_hunt_active() else (self.L if side == "left" else self.R,)
        for s in targets:
            if s.hunt_history:
                s.hunt_num, s.hunt_den = s.hunt_history.pop()
                s.hunt_latest_eval = True
                self._update_hunt_display("left" if s is self.L else "right")

    def _on_new_best(self):
        label = f"{self._mso1_val}%at{self.spin_location.value()}"
        self.combo_best.addItem(label)
        idx = self.combo_best.count() - 1
        self.combo_best.setItemData(idx, QColor("red"), Qt.ForegroundRole)
        self.combo_best.setCurrentIndex(idx)

    def _on_equal_best(self):
        label = f"{self._mso1_val}%at{self.spin_location.value()}"
        self.combo_best.addItem(label)
        idx = self.combo_best.count() - 1
        self.combo_best.setItemData(idx, QColor("#e65100"), Qt.ForegroundRole)
        self.combo_best.setCurrentIndex(idx)

    def _on_sm_clicked(self):
        self._magstim_mode = "SM"
        self.btn_sm.setStyleSheet("background-color: #1565C0; color: white;")
        self.btn_dm.setStyleSheet("")
        self.lbl_mso2.setEnabled(False)
        self.lbl_mso2_val.setEnabled(False)
        self.lbl_mso2_val.setText("—")
        self.lbl_ipi.setEnabled(False)
        self.lbl_isi_val.setEnabled(False)
        self.lbl_isi_val.setText("—")

    def _on_dm_clicked(self):
        self._magstim_mode = "DM"
        self.btn_dm.setStyleSheet("background-color: #1565C0; color: white;")
        self.btn_sm.setStyleSheet("")
        self.lbl_mso2.setEnabled(True)
        self.lbl_mso2_val.setEnabled(True)
        self.lbl_ipi.setEnabled(True)
        self.lbl_isi_val.setEnabled(True)

    def _connect_magstim(self):
        if not _SERIAL_AVAILABLE:
            return
        try:
            port = serial.Serial(
                port=MAGSTIM_PORT, baudrate=9600,
                bytesize=serial.EIGHTBITS, parity=serial.PARITY_NONE,
                stopbits=serial.STOPBITS_ONE, timeout=0.3,
            )
            self._magstim_port = port
            self._magstim_stop.clear()
            t = threading.Thread(target=self._magstim_thread_func, daemon=True)
            t.start()
        except Exception:
            pass

    def _magstim_thread_func(self):
        try:
            disable_remote(self._magstim_port)
        except Exception:
            pass
        while not self._magstim_stop.wait(MAGSTIM_POLL_MS / 1000.0):
            if self._magstim_port is None:
                break
            try:
                params = get_parameters(self._magstim_port)
                temp   = get_temperature(self._magstim_port)
                if params is not None:
                    self._magstim_polled.emit({"params": params, "temp": temp})
            except Exception:
                pass

    def _on_magstim_polled(self, data):
        self._magstim_last_poll_t = time.time()
        params = data["params"]
        temp   = data["temp"]
        f = params["flags"]
        state = "RDY" if f["ready"] else "ARM" if f["armed"] else "SBY"

        self.lbl_magstim_dot.setStyleSheet("color: #2e7d32;")
        self.lbl_magstim_dot.setToolTip(MAGSTIM_PORT)
        self.lbl_magstim_state.setText(state)
        self.lbl_magstim_state.setStyleSheet(
            "color: #2e7d32;" if f["ready"] else
            "color: #e65100;" if f["armed"] else
            "color: #555555;"
        )

        if temp is not None:
            t1 = temp["temp1"] / 10.0
            t2 = temp["temp2"] / 10.0
            avg_c = (t1 + t2) / 2.0
            self._temp_val = f"{avg_c:.1f}"
            self.lbl_magstim_temp.setText(f"{avg_c:.1f}°C")
            self.lbl_magstim_temp.setToolTip(
                f"Coil1: {t1:.1f}°C    Coil2: {t2:.1f}°C"
            )

        self._mso1_val = params["power_a"]
        self.lbl_mso_val.setText(str(params["power_a"]))
        if self._magstim_mode == "DM":
            self._mso2_val = params["power_b"]
            self._isi_val  = f"{params['ipi'] / 10.0:.1f}"
            self.lbl_mso2_val.setText(str(params["power_b"]))
            self.lbl_isi_val.setText(f"{params['ipi'] / 10.0:.1f} ms")
        else:
            self._mso2_val = 0
            self._isi_val  = "—"

    def closeEvent(self, event):
        if self.is_running:
            self._stop()
        self._magstim_stop.set()
        if self._magstim_port is not None:
            try:
                self._magstim_port.close()
            except Exception:
                pass
        if self._plot_window is not None:
            self._plot_window.close()
        super().closeEvent(event)

    def _on_separate_toggled(self, checked):
        self.separate = checked
        self.lbl_arrow_1.setText("<---" if checked else "<--->")
        opacity = 1.0 if checked else 0.0
        self._fx_arrow_2.setOpacity(opacity)
        self._fx_params_2.setOpacity(opacity)
        self.widget_mep_params_2.setEnabled(checked)

    def _reset_mep_params(self):
        for s, v in ((self.spin_mep_start, 10), (self.spin_mep_end, 50),
                     (self.spin_threshold, 0.05), (self.spin_prestim_start, -200),
                     (self.spin_prestim_end, -50)):
            s.setValue(v)
        for s, v in ((self.spin_mep_start_2, 10), (self.spin_mep_end_2, 50),
                     (self.spin_threshold_2, 0.05), (self.spin_prestim_start_2, -200),
                     (self.spin_prestim_end_2, -50)):
            s.setValue(v)

    def _apply_chart_params(self):
        self.window_secs = self.spinbox_window.value()

    def _apply_scope_params(self):
        self.scope_pre  = self.spin_pre.value()
        self.scope_post = self.spin_post.value()

    # ------------------------------------------------------------------
    # Channel dropdowns
    # ------------------------------------------------------------------

    def _on_combo(self, side, idx):
        if not self.ch_info or idx < 0:
            return
        s = self.L if side == "left" else self.R
        s.ch = idx
        name, unit = self.ch_info[idx]
        s.plot.setTitle(f"<b>{name}</b>", color=s.color, size="11pt")
        s.plot.setLabel("left", unit)
        self._update_combo_exclusion()

    def _update_combo_exclusion(self):
        """Disable in each combo the channel currently shown by the other side."""
        if not self.ch_info:
            return
        for own, other in ((self.L, self.R), (self.R, self.L)):
            for i in range(own.combo.count()):
                item = own.combo.model().item(i)
                if item is not None:
                    item.setEnabled(i != other.ch)

    # ------------------------------------------------------------------
    # COM event pump & trigger handler
    # ------------------------------------------------------------------

    def _register_com_events(self):
        if self._com_thread is not None and self._com_thread.is_alive():
            return
        self._com_thread = threading.Thread(target=self._com_thread_func, daemon=True)
        self._com_thread.start()
        print("[COM events] background thread started")

    def _com_thread_func(self):
        """Background thread: independent LabChart connection, pumps COM events."""
        pythoncom.CoInitialize()
        emit = self._tms_detected.emit

        class _Sink:
            def OnCommentAdded(self, *args):
                emit(args)  # forward (record, tick) so main thread can compute exact time
            def OnStartSamplingBlock(self, *_):
                pass
            def OnNewSamples(self, *_):
                pass

        try:
            lc_app = win32com.client.GetActiveObject("ADIChart.Application")
            doc = lc_app.ActiveDocument
            win32com.client.WithEvents(doc, _Sink)
            print("[COM thread] OnCommentAdded registered")
            while True:
                pythoncom.PumpWaitingMessages()
                time.sleep(0.05)
        except Exception as e:
            print(f"[COM thread] exiting: {e}")
        finally:
            pythoncom.CoUninitialize()

    def _on_tms_trigger(self, event_args=()):
        if self.client is None:
            return
        try:
            # OnCommentAdded args: (text, channel, record, tick)
            if len(event_args) >= 4:
                rec  = int(event_args[2])
                tick = int(event_args[3])
                spt  = self.client.doc.GetRecordSecsPerTick(rec)
                t    = tick * spt
            else:
                t = self.client.current_time()
            if (self._last_trigger_t is not None and
                    0 <= t - self._last_trigger_t < TRIGGER_REFRACTORY):
                return
            self._last_trigger_t = t
            self._register_trigger(t, add_table_row=True)
        except Exception as e:
            print(f"[OnCommentAdded] args={event_args} err={e}")

    # ------------------------------------------------------------------
    # Status & reconnect
    # ------------------------------------------------------------------

    def _refresh_status(self):
        if self.client is None:
            try:
                self.client = LabChartClient()
                self._populate_combos()
                self._register_com_events()
                self._first_play = True
                self.btn.setText("▶  Play")
            except Exception:
                self.btn.setText("📂  Open LabChart")
                self.lbl_status.setText("✕  LabChart Not Connected")
                self.lbl_status.setStyleSheet("color: #b71c1c;")
                return

        try:
            sampling = self.client.is_sampling()
            if sampling:
                self.btn_sample.setEnabled(True)
                self.lbl_status.setText("●  LabChart Connected  ·  Recording")
                self.lbl_status.setStyleSheet("color: #2e7d32;")
                # Auto-start viewer if LabChart began recording externally
                if not self.is_running:
                    self.is_running = True
                    self.btn.setText("⏹  Stop")
                    self.data_timer.start(UPDATE_MS)
            else:
                self.btn_sample.setEnabled(False)
                in_grace = (self._start_time is not None and
                            time.monotonic() - self._start_time < 2.0)
                if self.is_running and not in_grace:
                    self.data_timer.stop()
                    self.is_running = False
                    self.btn.setText("▶  Play")
                self.lbl_status.setText("●  LabChart Connected  ·  Stopped")
                self.lbl_status.setStyleSheet("color: #e65100;")
        except Exception:
            self.client = None
            self._com_thread = None
            self.btn_sample.setEnabled(False)
            if self.is_running:
                self.data_timer.stop()
                self.is_running = False
            self.btn.setText("📂  Open LabChart")
            self.lbl_status.setText("✕  LabChart Not Connected")
            self.lbl_status.setStyleSheet("color: #b71c1c;")

        # Magstim connection timeout: grey out if no poll for >3 seconds
        if self._magstim_last_poll_t > 0 and time.time() - self._magstim_last_poll_t > 3.0:
            self.lbl_magstim_dot.setStyleSheet("color: #aaaaaa;")
            self.lbl_magstim_state.setText("—")
            self.lbl_magstim_state.setStyleSheet("color: #aaaaaa;")
            self.lbl_magstim_temp.setText("—")
            self.lbl_magstim_temp.setToolTip("Coil temp unavailable")

    def _populate_combos(self):
        new_info = self.client.get_channel_info()
        if new_info == self.ch_info:
            return
        self.ch_info = new_info
        for s, default in ((self.L, DEFAULT_CH_LEFT), (self.R, DEFAULT_CH_RIGHT)):
            s.combo.blockSignals(True)
            s.combo.clear()
            for i, (name, unit) in enumerate(self.ch_info):
                s.combo.addItem(f"Ch{i + 1}  {name}  [{unit}]")
            s.combo.setCurrentIndex(min(default, len(self.ch_info) - 1))
            s.combo.blockSignals(False)
        self._on_combo("left",  self.L.combo.currentIndex())
        self._on_combo("right", self.R.combo.currentIndex())

    # ------------------------------------------------------------------
    # Play / Stop / Sample / Open
    # ------------------------------------------------------------------

    def _toggle(self):
        if self.client is None:
            self._open_labchart()
        elif self.is_running:
            self._stop()
        else:
            self._start()

    def _start(self):
        if self.client is None:
            return
        self._start_time = time.monotonic()
        self.is_running = True
        self.btn.setText("⏹  Stop")
        if self._first_play:
            self._first_play = False
            try:
                self.client.start_sampling()
            except Exception:
                pass
            QTimer.singleShot(500, self._bounce_stop)
        else:
            self._do_start()

    def _bounce_stop(self):
        try:
            self.client.stop_sampling()
        except Exception:
            pass
        QTimer.singleShot(500, self._do_start)

    def _do_start(self):
        if self.client is None:
            return
        try:
            self.client.start_sampling()
        except Exception as e:
            print(f"[StartSampling] {e}")
        self._start_time = time.monotonic()
        self.is_running = True
        self.btn.setText("⏹  Stop")
        self.data_timer.start(UPDATE_MS)

    def _stop(self):
        self.data_timer.stop()
        if self.client is not None:
            try:
                self.client.stop_sampling()
            except Exception as e:
                print(f"[StopSampling] {e}")
        self.is_running = False
        self.btn.setText("▶  Play")

    def _add_sample_comment(self):
        if self.client is None:
            return
        try:
            self.client.add_comment("Trigger")
        except Exception as e:
            print(f"[AddComment] {e}")

    def _on_trigger_selected(self, value):
        self._next_trial_num = value + 1
        if self.mode == "scope" and value >= 1:
            idx = min(value - 1, len(self.trigger_times) - 1)
            if idx >= 0:
                channels = list({self.L.ch, self.R.ch})
                self._update_scope(channels, selected_row=idx)

    def _register_trigger(self, t, add_table_row=True):
        self.trigger_times.append(t)
        trial_num = self._next_trial_num
        self._next_trial_num += 1
        self.trigger_spin.setMaximum(max(9999, trial_num))
        self.trigger_spin.blockSignals(True)
        self.trigger_spin.setValue(trial_num)
        self.trigger_spin.blockSignals(False)
        self.trigger_spin.setEnabled(self.mode == "scope")
        if not add_table_row:
            return
        for s in (self.L, self.R):
            s.hunt_latest_eval  = False
            s.table_latest_eval = False
        loc_id   = self.spin_location.value()
        mso2_str = str(self._mso2_val) if self._magstim_mode == "DM" else "0"
        isi_str  = self._isi_val if self._magstim_mode == "DM" else "0"
        row_vals = [
            lambda s: str(trial_num),
            lambda s: str(s.ch + 1),
            lambda s: self._get_hand(s),
            lambda s: self._get_paralysis(s),
            lambda s: str(loc_id),
            lambda s: str(self._mso1_val),
            lambda s: mso2_str,
            lambda s: isi_str,
            lambda s: "0",          # MEP_amp  (col 8)
            lambda s: "0",          # Prestim_mean (col 9)
            lambda s: self._temp_val if self._temp_val != "—" else "0",
            lambda s: self._magstim_mode,
        ]
        for s in (self.L, self.R):
            row = s.table.rowCount()
            s.table.insertRow(row)
            for col, fn in enumerate(row_vals):
                item = QTableWidgetItem(fn(s))
                item.setTextAlignment(Qt.AlignCenter)
                s.table.setItem(row, col, item)
            excl_item = QTableWidgetItem("0")
            excl_item.setTextAlignment(Qt.AlignCenter)
            s.table.setItem(row, len(row_vals), excl_item)
            s.table.scrollToBottom()
        for s in (self.L, self.R):
            if s.hunt_panel.isEnabled() and s.hunt_chk.isChecked():
                s.hunt_history.append((s.hunt_num, s.hunt_den))
                s.hunt_den += 1
                self._update_hunt_display("left" if s is self.L else "right")
        self._sync_plot_table()

    def _open_labchart(self):
        search_dir = Path(__file__).parent
        candidates = list(search_dir.glob(f"{Path(LABCHART_FILE).stem}*.adicht"))
        if not candidates:
            self.lbl_status.setText(f"✕  {LABCHART_FILE} not found in TMSviewer folder")
            self.lbl_status.setStyleSheet("color: #b71c1c;")
            return
        os.startfile(str(candidates[0]))

    # ------------------------------------------------------------------
    # Real-time update
    # ------------------------------------------------------------------

    def _update(self):
        if self.client is None:
            return
        channels = list({self.L.ch, self.R.ch})

        if self.mode == "chart":
            self._update_chart(channels)
        else:
            self._update_scope(channels)

    def _update_chart(self, channels):
        try:
            data, t_start, t_end = self.client.get_latest_data(
                window_secs=self.window_secs,
                channels=channels,
            )
        except Exception as e:
            print(f"[chart update] {e}")
            return

        self._draw(self.L.curve, data.get(self.L.ch), t_start, t_end, offset=0)
        self._draw(self.R.curve, data.get(self.R.ch), t_start, t_end, offset=0)
        self._redraw_trigger_lines_chart(t_start, t_end)

    def _update_scope(self, channels, selected_row=None):
        if not self.trigger_times:
            return

        user_selected = selected_row is not None
        if selected_row is None:
            selected_row = self.trigger_spin.value() - 1
        selected_row = min(selected_row, len(self.trigger_times) - 1)
        if selected_row < 0:
            return

        trigger_t = self.trigger_times[selected_row]

        # Guard only during live streaming — skip when user explicitly picks a past trigger
        if not user_selected:
            try:
                now = self.client.current_time()
                if now < trigger_t + self.scope_post:
                    return
            except Exception:
                return

        try:
            data, t_start, t_end = self.client.get_scope_data(
                trigger_time=trigger_t,
                pre_secs=self.scope_pre,
                post_secs=self.scope_post,
                channels=channels,
            )
        except Exception as e:
            print(f"[scope update] {e}")
            return

        total_window = self.scope_pre + self.scope_post
        if total_window <= 1.0:
            t_scale  = 1000.0
            x_label  = "Time from trigger (ms)"
        else:
            t_scale  = 1.0
            x_label  = "Time from trigger (s)"
        for s in (self.L, self.R):
            s.plot.setLabel("bottom", x_label)

        arr_L = data.get(self.L.ch)
        arr_R = data.get(self.R.ch)
        sides = ((self.L, arr_L), (self.R, arr_R))
        for s, a in sides:
            self._draw(s.curve, a, t_start, t_end, offset=trigger_t, t_scale=t_scale)
        self._redraw_trigger_lines_scope(t_scale=t_scale)

        # Y-range: respect Fix/Auto state per plot
        for s, a in sides:
            s.plot.enableAutoRange(axis="y", enable=False)
            if s.yfix:
                s.plot.setYRange(-s.yhalf, s.yhalf, padding=0)
            else:
                rng = self._artifact_free_range(a, t_start, t_end, trigger_t)
                if rng is not None:
                    ymin, ymax = rng
                    pad = (ymax - ymin) * 0.1 if ymax != ymin else 0.1
                    s.plot.setYRange(ymin - pad, ymax + pad, padding=0)

        # Evaluate both sides once; reuse for display, hunt, and table
        is_latest = (selected_row == len(self.trigger_times) - 1)
        if self.radio_mep.isChecked():
            r_L = self._analyse_epoch(arr_L, second=False)
            r_R = self._analyse_epoch(arr_R, second=self.separate)
            self._update_stats(self.L.stats, arr_L, second=False)
            self._update_stats(self.R.stats, arr_R, second=self.separate)
        else:
            r_L = r_R = None
            for s in (self.L, self.R):
                s.stats.setText("Prestim RMS\n—\n\nMEP amp\n—")

        # Hunt + table fill — latest trigger only, once
        if is_latest and self.radio_mep.isChecked():
            for s, r in ((self.L, r_L), (self.R, r_R)):
                if s.hunt_panel.isEnabled() and s.hunt_chk.isChecked() and not s.hunt_latest_eval and r is not None:
                    thr = (self.spin_threshold_2.value() if (s is self.R and self.separate)
                           else self.spin_threshold.value())
                    if r["mep_amp"] > thr:
                        s.hunt_num += 1
                        self._update_hunt_display("left" if s is self.L else "right")
                    s.hunt_latest_eval = True

            row = selected_row
            updated = False
            for s, r in ((self.L, r_L), (self.R, r_R)):
                if not s.table_latest_eval and r is not None:
                    self._set_table_cell(s.table, row, 8, f"{r['mep_amp']:.3f}")
                    self._set_table_cell(s.table, row, 9, f"{r['prestim_rms']:.3f}")
                    s.table_latest_eval = True
                    updated = True
            if updated:
                self._sync_plot_table()

    # ------------------------------------------------------------------
    # Trigger lines
    # ------------------------------------------------------------------

    def _clear_trigger_lines(self):
        for s in (self.L, self.R):
            for line in s.vlines: s.plot.removeItem(line)
            for item in s.shades: s.plot.removeItem(item)
            s.vlines.clear()
            s.shades.clear()

    def _redraw_trigger_lines_chart(self, t_start, t_end):
        self._clear_trigger_lines()
        pen = pg.mkPen("#ff6f00", width=3, style=Qt.DashLine)
        for t in self.trigger_times:
            if t_start <= t <= t_end:
                for s in (self.L, self.R):
                    line = pg.InfiniteLine(pos=t, angle=90, pen=pen)
                    s.plot.addItem(line)
                    s.vlines.append(line)

    def _redraw_trigger_lines_scope(self, t_scale=1.0):
        self._clear_trigger_lines()
        pen = pg.mkPen("#ff6f00", width=3, style=Qt.DashLine)
        for s in (self.L, self.R):
            line = pg.InfiniteLine(pos=0, angle=90, pen=pen)
            s.plot.addItem(line)
            s.vlines.append(line)

        if self.radio_mep.isChecked():
            def _add_shades(s, pre_s_spin, pre_e_spin, mep_s_spin, mep_e_spin):
                for a, b in (
                    (pre_s_spin.value() / 1000.0 * t_scale, pre_e_spin.value() / 1000.0 * t_scale),
                    (mep_s_spin.value() / 1000.0 * t_scale, mep_e_spin.value() / 1000.0 * t_scale),
                ):
                    region = pg.LinearRegionItem(
                        values=(a, b),
                        brush=pg.mkBrush(255, 220, 0, 60),
                        pen=pg.mkPen(None),
                        movable=False,
                    )
                    s.plot.addItem(region)
                    s.shades.append(region)

            _add_shades(self.L, self.spin_prestim_start, self.spin_prestim_end,
                        self.spin_mep_start, self.spin_mep_end)
            if self.separate:
                _add_shades(self.R, self.spin_prestim_start_2, self.spin_prestim_end_2,
                            self.spin_mep_start_2, self.spin_mep_end_2)
            else:
                _add_shades(self.R, self.spin_prestim_start, self.spin_prestim_end,
                            self.spin_mep_start, self.spin_mep_end)

    # ------------------------------------------------------------------
    # Draw helper
    # ------------------------------------------------------------------

    @staticmethod
    def _artifact_free_range(arr, t_start, t_end, trigger_t, exclude_s=0.005):
        """Return (ymin, ymax) ignoring data within ±exclude_s seconds of the trigger."""
        if arr is None or len(arr) == 0:
            return None
        t_rel = np.linspace(t_start - trigger_t, t_end - trigger_t, len(arr))
        mask  = (t_rel < -exclude_s) | (t_rel > exclude_s)
        valid = arr[mask]
        if len(valid) == 0:
            return None
        return float(valid.min()), float(valid.max())

    @staticmethod
    def _draw(curve, arr, t_start, t_end, offset=0, t_scale=1.0):
        if arr is None or len(arr) == 0:
            return
        t = np.linspace((t_start - offset) * t_scale, (t_end - offset) * t_scale, len(arr))
        if len(arr) > MAX_DISPLAY_PTS:
            step = len(arr) // MAX_DISPLAY_PTS
            arr = arr[::step]
            t   = t[::step]
        curve.setData(t, arr)
