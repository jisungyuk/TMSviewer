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
    QTabBar, QToolTip, QLineEdit, QShortcut, QDialog, QTextBrowser,
)
from PyQt5.QtGui import QFont, QColor, QCursor, QKeySequence
from PyQt5.QtCore import QTimer, Qt, pyqtSignal

try:
    import serial
    import serial.tools.list_ports
    from magstim_test import enable_remote, disable_remote, get_parameters, get_temperature
    _SERIAL_AVAILABLE = True
except ImportError:
    _SERIAL_AVAILABLE = False

from labchart_client import LabChartClient


WINDOW_SECS_DEFAULT = 4
WINDOW_SECS_MIN = 4
WINDOW_SECS_MAX = 15
SCOPE_PRE_DEFAULT = 0.15
SCOPE_POST_DEFAULT = 0.35

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

TRIGGERBOX_PORT    = "COM5"
TRIGGERBOX_BAUD    = 115200
TRIGGERBOX_CHANNEL = 1      # ch1 = 0x01; wired to LabChart ch5
TRIGGERBOX_PULSE_MS = 100



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


class ActiveWindow(QWidget):
    """Always-on-top popup: two streaming graphs with channel selectors and middle controls."""

    _DEFAULT_LEFT  = 10   # CH11 (0-indexed)
    _DEFAULT_RIGHT = 11   # CH12 (0-indexed)

    def __init__(self, viewer, parent=None):
        super().__init__(parent, Qt.Window)
        self._viewer   = viewer
        self._resizing = False
        self.setWindowTitle("TMSviewer — Active")
        self.setWindowFlags(self.windowFlags() | Qt.WindowStaysOnTopHint)
        self.resize(992, 558)

        root = QVBoxLayout(self)
        root.setContentsMargins(8, 8, 8, 8)
        root.setSpacing(4)

        self.lbl_window_title = QLabel("Maximum Voluntary Contraction")
        title_font = QFont(); title_font.setPointSize(11); title_font.setBold(True)
        self.lbl_window_title.setFont(title_font)
        self.lbl_window_title.setAlignment(Qt.AlignCenter)
        root.addWidget(self.lbl_window_title)

        # ── top row: channel selectors + arrow ────────────────────────────
        top_row = QHBoxLayout()
        top_row.setContentsMargins(0, 0, 0, 0)

        self.combo_ch_left  = QComboBox()
        self.combo_ch_right = QComboBox()
        top_row.addWidget(self.combo_ch_left,  stretch=5)
        top_row.addStretch(1)
        top_row.addStretch(1)
        top_row.addWidget(self.combo_ch_right, stretch=5)
        root.addLayout(top_row)
        self.populate_ch_combos()

        # ── main row: left graph | middle buttons | right graph ────────────
        main_row = QHBoxLayout()
        main_row.setSpacing(8)

        self.plot_left = pg.PlotWidget()
        self.plot_right = pg.PlotWidget()
        for plt in (self.plot_left, self.plot_right):
            plt.setBackground("k")
            plt.setXRange(0, 1)
            plt.setYRange(0, 1)
            plt.getAxis("bottom").hide()
            plt.getAxis("left").setTextPen("w")
            plt.showGrid(x=False, y=True, alpha=0.2)
            plt.setMouseEnabled(x=False, y=False)

        self._ball_left  = pg.ScatterPlotItem([0.5], [0.0],
            pen=None, brush=pg.mkBrush(140, 140, 140), size=26, symbol='o')
        self._ball_right = pg.ScatterPlotItem([0.5], [0.0],
            pen=None, brush=pg.mkBrush(140, 140, 140), size=26, symbol='o')
        self.plot_left.addItem(self._ball_left)
        self.plot_right.addItem(self._ball_right)

        txt_font = QFont(); txt_font.setPointSize(18); txt_font.setBold(True)
        self._status_txt_l   = pg.TextItem("WAIT", color="w", anchor=(0.5, 0))
        self._status_txt_r   = pg.TextItem("WAIT", color="w", anchor=(0.5, 0))
        self._countdown_txt_l = pg.TextItem("",    color="w", anchor=(0.5, 0))
        self._countdown_txt_r = pg.TextItem("",    color="w", anchor=(0.5, 0))
        for t in (self._status_txt_l, self._status_txt_r,
                  self._countdown_txt_l, self._countdown_txt_r):
            t.setFont(txt_font)
        self._status_txt_l.setPos(0.5,    self._STATUS_YFRAC)
        self._status_txt_r.setPos(0.5,    self._STATUS_YFRAC)
        self._countdown_txt_l.setPos(0.5, self._COUNTDOWN_YFRAC)
        self._countdown_txt_r.setPos(0.5, self._COUNTDOWN_YFRAC)
        self.plot_left.addItem(self._status_txt_l)
        self.plot_left.addItem(self._countdown_txt_l)
        self.plot_right.addItem(self._status_txt_r)
        self.plot_right.addItem(self._countdown_txt_r)

        self._val_left   = 0.0
        self._val_right  = 0.0
        self._streaming  = False
        self._switch_mode = 1   # 0=both, 1=left only, 2=right only
        self._band_left      = None
        self._band_line_left = None
        self._band_right     = None
        self._band_line_right= None
        self._history_left  = []   # [(max, min), ...]
        self._history_right = []
        self._cur_max_left  = None
        self._cur_min_left  = None
        self._cur_max_right = None
        self._cur_min_right = None
        self._task_end_t = 0.0
        self._fade_start_val_l = 0.0
        self._fade_start_val_r = 0.0
        self._fade_elapsed     = 0.0

        self._stream_timer = QTimer(self)
        self._stream_timer.setInterval(50)
        self._stream_timer.timeout.connect(self._on_stream_tick)

        self._fade_timer = QTimer(self)
        self._fade_timer.setInterval(50)
        self._fade_timer.timeout.connect(self._on_fade_tick)

        self._lc_status_timer = QTimer(self)
        self._lc_status_timer.setInterval(1000)
        self._lc_status_timer.timeout.connect(self._poll_lc_status)
        self._lc_status_timer.start()

        mid_col = QVBoxLayout()
        mid_col.setSpacing(8)
        mid_col.setContentsMargins(4, 0, 4, 0)

        self.lbl_lc_status = QLabel("● —")
        self.lbl_lc_status.setAlignment(Qt.AlignCenter)
        lc_font = QFont(); lc_font.setPointSize(9); lc_font.setBold(True)
        self.lbl_lc_status.setFont(lc_font)
        self.lbl_lc_status.setStyleSheet("color: gray;")

        self.arrow_lbl = QLabel("<---")
        self.arrow_lbl.setAlignment(Qt.AlignCenter)
        arrow_font = QFont(); arrow_font.setPointSize(14); arrow_font.setBold(True)
        self.arrow_lbl.setFont(arrow_font)

        self.btn_switch = QPushButton("Switch")
        self.btn_switch.setFixedWidth(90)
        self.btn_switch.setToolTip(
            "Click: toggle left / right channel\n"
            "Shift + Click: both channels"
        )
        self.btn_switch.clicked.connect(self._on_switch_clicked)
        self.btn_mvc    = QPushButton("MVC")
        self.btn_mvc.setFixedWidth(90)
        self.btn_mvc.clicked.connect(self._on_mvc_clicked)
        dur_lbl = QLabel("Duration")
        dur_lbl.setAlignment(Qt.AlignCenter)
        self.spin_duration = QSpinBox()
        self.spin_duration.setRange(0, 999)
        self.spin_duration.setValue(10)
        self.spin_duration.setSuffix(" s")
        self.spin_duration.setFixedWidth(72)

        self._lbl_yaxis_header = QLabel("Y axis  (<-)")
        self._lbl_yaxis_header.setAlignment(Qt.AlignCenter)

        self._lbl_ydir_right_hdr = QLabel("->")
        self._lbl_ydir_right_hdr.setAlignment(Qt.AlignCenter)

        self.spin_yaxis_left  = QDoubleSpinBox()
        self.spin_yaxis_right = QDoubleSpinBox()
        for sp in (self.spin_yaxis_left, self.spin_yaxis_right):
            sp.setRange(0.01, 1000.0)
            sp.setValue(1.0)
            sp.setDecimals(2)
            sp.setSingleStep(0.1)
            sp.setFixedWidth(72)
        self.spin_yaxis_left.valueChanged.connect(
            lambda v: (self.plot_left.setYRange(0, v),
                       self._reposition_texts("left", v)))
        self.spin_yaxis_right.valueChanged.connect(
            lambda v: (self.plot_right.setYRange(0, v),
                       self._reposition_texts("right", v)))

        mid_col.addWidget(self.lbl_lc_status,  alignment=Qt.AlignHCenter)
        mid_col.addWidget(self.arrow_lbl,       alignment=Qt.AlignHCenter)
        mid_col.addSpacing(4)
        mid_col.addWidget(self.btn_switch,      alignment=Qt.AlignHCenter)
        mid_col.addWidget(self.btn_mvc,         alignment=Qt.AlignHCenter)
        mid_col.addSpacing(6)
        mid_col.addWidget(dur_lbl,              alignment=Qt.AlignHCenter)
        mid_col.addWidget(self.spin_duration,   alignment=Qt.AlignHCenter)
        mid_col.addSpacing(6)
        mid_col.addWidget(self._lbl_yaxis_header,    alignment=Qt.AlignHCenter)
        mid_col.addWidget(self.spin_yaxis_left,      alignment=Qt.AlignHCenter)
        mid_col.addWidget(self._lbl_ydir_right_hdr,  alignment=Qt.AlignHCenter)
        mid_col.addWidget(self.spin_yaxis_right,     alignment=Qt.AlignHCenter)
        mid_col.addStretch()
        self._update_yaxis_ui()

        QShortcut(QKeySequence("F1"), self, activated=self.btn_mvc.click)
        QShortcut(QKeySequence("F2"), self, activated=self.btn_switch.click)

        main_row.addWidget(self.plot_left,  stretch=5)
        main_row.addLayout(mid_col)
        main_row.addWidget(self.plot_right, stretch=5)
        root.addLayout(main_row, stretch=1)

        # ── stats / avg rows ───────────────────────────────────────────────
        stats_font = QFont(); stats_font.setPointSize(9)
        _sel = Qt.TextSelectableByMouse | Qt.TextSelectableByKeyboard
        _REDO_W = 46

        def _stat_label(text=""):
            lbl = QLabel(text)
            lbl.setFont(stats_font)
            lbl.setAlignment(Qt.AlignCenter)
            lbl.setTextInteractionFlags(_sel)
            return lbl

        self.lbl_stats_left  = _stat_label("n= —    MAX: —    MIN: —")
        self.lbl_stats_right = _stat_label("n= —    MAX: —    MIN: —")
        self._avg_max_left  = None
        self._avg_min_left  = None
        self._avg_max_right = None
        self._avg_min_right = None

        self.btn_redo_left  = QPushButton("Redo")
        self.btn_redo_right = QPushButton("Redo")
        for btn in (self.btn_redo_left, self.btn_redo_right):
            btn.setFixedWidth(_REDO_W)
            btn.setEnabled(False)
        self.btn_redo_left.clicked.connect(lambda: self._on_redo_clicked("left"))
        self.btn_redo_right.clicked.connect(lambda: self._on_redo_clicked("right"))

        stats_row = QHBoxLayout()
        stats_row.setContentsMargins(0, 0, 0, 0)
        stats_row.setSpacing(4)
        stats_row.addWidget(self.lbl_stats_left,  stretch=5)
        stats_row.addWidget(self.btn_redo_left)
        stats_row.addStretch(1)
        stats_row.addWidget(self.btn_redo_right)
        stats_row.addWidget(self.lbl_stats_right, stretch=5)
        root.addLayout(stats_row)

        _btn_avg_style = (
            "QPushButton { border: 1px solid #555; border-radius: 3px; "
            "padding: 1px 6px; background: #252525; color: #ccc; font-size: 9pt; }"
            "QPushButton:hover { background: #333; color: #fff; }"
        )
        _edit_avg_style = (
            "QLineEdit { border: 1px solid #555; border-radius: 3px; "
            "padding: 1px 4px; background: #1a1a1a; color: #eee; font-size: 9pt; }"
        )

        def _avg_side_layout(max_btn_attr, max_edit_attr, min_btn_attr, min_edit_attr, side):
            lay = QHBoxLayout()
            lay.setContentsMargins(0, 0, 0, 0)
            lay.setSpacing(3)
            lay.addStretch()
            for btn_attr, edit_attr, kind in (
                (max_btn_attr, max_edit_attr, "max"),
                (min_btn_attr, min_edit_attr, "min"),
            ):
                lbl_kind = "MAX" if kind == "max" else "MIN"
                btn = QPushButton(lbl_kind)
                btn.setFixedWidth(46)
                btn.setStyleSheet(_btn_avg_style)
                btn.setToolTip(f"Click to copy {lbl_kind} value")
                edit = QLineEdit()
                edit.setPlaceholderText("—")
                edit.setStyleSheet(_edit_avg_style)
                edit.setFixedWidth(72)
                edit.setAlignment(Qt.AlignCenter)
                btn.clicked.connect(lambda checked=False, e=edit: self._copy_edit(e))
                setattr(self, btn_attr,  btn)
                setattr(self, edit_attr, edit)
                lay.addWidget(btn)
                lay.addWidget(edit)
            lay.addStretch()
            return lay

        avg_row = QHBoxLayout()
        avg_row.setContentsMargins(0, 0, 0, 0)
        avg_row.setSpacing(4)
        avg_row.addLayout(_avg_side_layout(
            "btn_avg_max_left",  "edit_avg_max_left",
            "btn_avg_min_left",  "edit_avg_min_left",  "left"),  stretch=5)
        avg_row.addSpacing(_REDO_W)
        avg_row.addStretch(1)
        avg_row.addSpacing(_REDO_W)
        avg_row.addLayout(_avg_side_layout(
            "btn_avg_max_right", "edit_avg_max_right",
            "btn_avg_min_right", "edit_avg_min_right", "right"), stretch=5)
        root.addLayout(avg_row)

        # ── divider ────────────────────────────────────────────────────────
        divider = QFrame()
        divider.setFrameShape(QFrame.HLine)
        divider.setFrameShadow(QFrame.Sunken)
        root.addWidget(divider)

        # ── Hold task checkbox ─────────────────────────────────────────────
        hold_row = QHBoxLayout()
        hold_row.setContentsMargins(0, 0, 0, 0)
        self.chk_hold_task = QCheckBox("Hold task")
        hold_font = QFont(); hold_font.setPointSize(9); hold_font.setBold(True)
        self.chk_hold_task.setFont(hold_font)
        self.chk_hold_task.setEnabled(False)
        self.chk_hold_task.toggled.connect(self._on_hold_task_toggled)
        hold_row.addStretch()
        hold_row.addWidget(self.chk_hold_task)
        hold_row.addStretch()
        root.addLayout(hold_row)

        # ── Target row ─────────────────────────────────────────────────────
        _tgt_style = (
            "QDoubleSpinBox { border: 1px solid #555; border-radius: 3px; "
            "padding: 1px 3px; background: #1a1a1a; color: #eee; font-size: 9pt; }"
        )
        _tgt_btn_style = (
            "QPushButton { border: 1px solid #555; border-radius: 3px; "
            "padding: 2px 8px; background: #252525; color: #ccc; font-size: 9pt; }"
            "QPushButton:hover { background: #333; color: #fff; }"
            "QPushButton:disabled { color: #555; background: #1a1a1a; }"
        )

        def _make_target_side(pct_attr, tol_attr, btn_attr):
            w = QWidget()
            lay = QHBoxLayout(w)
            lay.setContentsMargins(0, 0, 0, 0)
            lay.setSpacing(4)
            spin_pct = QDoubleSpinBox()
            spin_pct.setRange(0.0, 200.0)
            spin_pct.setValue(50.0)
            spin_pct.setDecimals(1)
            spin_pct.setSuffix(" %")
            spin_pct.setFixedWidth(68)
            spin_pct.setStyleSheet(_tgt_style)
            spin_pct.setButtonSymbols(QAbstractSpinBox.NoButtons)
            lbl_pm = QLabel("±")
            lbl_pm.setAlignment(Qt.AlignCenter)
            spin_tol = QDoubleSpinBox()
            spin_tol.setRange(0.0, 100.0)
            spin_tol.setValue(10.0)
            spin_tol.setDecimals(1)
            spin_tol.setSuffix(" %")
            spin_tol.setFixedWidth(62)
            spin_tol.setStyleSheet(_tgt_style)
            spin_tol.setButtonSymbols(QAbstractSpinBox.NoButtons)
            btn = QPushButton("APPLY")
            btn.setFixedWidth(68)
            btn.setStyleSheet(_tgt_btn_style)
            setattr(self, pct_attr, spin_pct)
            setattr(self, tol_attr, spin_tol)
            setattr(self, btn_attr,  btn)
            lay.addWidget(spin_pct)
            lay.addWidget(lbl_pm)
            lay.addWidget(spin_tol)
            lay.addWidget(btn)
            w.setEnabled(False)
            return w

        self._tgt_widget_left  = _make_target_side(
            "spin_tgt_pct_left",  "spin_tgt_tol_left",  "btn_apply_left")
        self._tgt_widget_right = _make_target_side(
            "spin_tgt_pct_right", "spin_tgt_tol_right", "btn_apply_right")
        self.btn_apply_left.clicked.connect(lambda: self._on_apply_clicked("left"))
        self.btn_apply_right.clicked.connect(lambda: self._on_apply_clicked("right"))

        lbl_targets = QLabel("targets")
        lbl_targets.setAlignment(Qt.AlignCenter)
        lbl_targets_font = QFont(); lbl_targets_font.setPointSize(9)
        lbl_targets.setFont(lbl_targets_font)

        def _centered_col(widget):
            col = QHBoxLayout()
            col.setContentsMargins(0, 0, 0, 0)
            col.addStretch()
            col.addWidget(widget)
            col.addStretch()
            return col

        target_row = QHBoxLayout()
        target_row.setContentsMargins(0, 2, 0, 0)
        target_row.setSpacing(0)
        target_row.addLayout(_centered_col(self._tgt_widget_left),  stretch=1)
        target_row.addLayout(_centered_col(lbl_targets),            stretch=1)
        target_row.addLayout(_centered_col(self._tgt_widget_right), stretch=1)
        root.addLayout(target_row)

    def populate_ch_combos(self):
        ch_info = getattr(self._viewer, "ch_info", None)
        for cb, default in ((self.combo_ch_left,  self._DEFAULT_LEFT),
                            (self.combo_ch_right, self._DEFAULT_RIGHT)):
            cb.blockSignals(True)
            cb.clear()
            if ch_info:
                for i, (name, unit) in enumerate(ch_info):
                    cb.addItem(f"Ch{i + 1}  {name}  [{unit}]")
            else:
                for i in range(1, 17):
                    cb.addItem(f"Ch{i}")
            cb.setCurrentIndex(min(default, cb.count() - 1))
            cb.blockSignals(False)

    def _set_status(self, text, color):
        if self._switch_mode != 2:
            self._status_txt_l.setText(text)
            self._status_txt_l.setColor(color)
        else:
            self._status_txt_l.setText("")
        if self._switch_mode != 1:
            self._status_txt_r.setText(text)
            self._status_txt_r.setColor(color)
        else:
            self._status_txt_r.setText("")

    def _set_countdown(self, text):
        if self._switch_mode != 2:
            self._countdown_txt_l.setText(text)
        else:
            self._countdown_txt_l.setText("")
        if self._switch_mode != 1:
            self._countdown_txt_r.setText(text)
        else:
            self._countdown_txt_r.setText("")

    def _set_ball_color(self, r, g, b):
        brush = pg.mkBrush(r, g, b)
        self._ball_left.setBrush(brush)
        self._ball_right.setBrush(brush)

    _STATUS_YFRAC    = 0.96
    _COUNTDOWN_YFRAC = 0.82

    def _reposition_texts(self, side, y_max):
        if side == "left":
            self._status_txt_l.setPos(0.5,    y_max * self._STATUS_YFRAC)
            self._countdown_txt_l.setPos(0.5, y_max * self._COUNTDOWN_YFRAC)
        else:
            self._status_txt_r.setPos(0.5,    y_max * self._STATUS_YFRAC)
            self._countdown_txt_r.setPos(0.5, y_max * self._COUNTDOWN_YFRAC)

    def _update_yaxis_ui(self):
        mode = self._switch_mode
        _headers = {0: "Y axis  <-", 1: "Y axis  <-", 2: "Y axis  ->"}
        self._lbl_yaxis_header.setText(_headers[mode])
        self.spin_yaxis_left.setVisible(mode == 0 or mode == 1)
        self._lbl_ydir_right_hdr.setVisible(mode == 0)
        self.spin_yaxis_right.setVisible(mode == 0 or mode == 2)

    def _on_switch_clicked(self):
        if QApplication.keyboardModifiers() & Qt.ShiftModifier:
            new_mode = 0
        else:
            new_mode = 2 if self._switch_mode == 1 else 1

        if self.chk_hold_task.isChecked():
            from PyQt5.QtWidgets import QMessageBox
            has_l = bool(self.edit_avg_max_left.text().strip())
            has_r = bool(self.edit_avg_max_right.text().strip())
            if new_mode == 1 and not has_l:
                QMessageBox.warning(self, "Hold Task",
                    "Left side MVC not measured.\nPlease complete left side MVC first.")
                return
            if new_mode == 2 and not has_r:
                QMessageBox.warning(self, "Hold Task",
                    "Right side MVC not measured.\nPlease complete right side MVC first.")
                return
            if new_mode == 0 and not (has_l and has_r):
                QMessageBox.warning(self, "Hold Task",
                    "Both sides need MVC measurements to use both-channel mode.")
                return

        self._switch_mode = new_mode
        arrows = ["<--->", "<---", "--->"]
        self.arrow_lbl.setText(arrows[self._switch_mode])
        # freeze the inactive side immediately
        if self._switch_mode == 1:   # left only → reset right
            self._val_right = 0.0
            self._ball_right.setData([0.5], [0.0])
        elif self._switch_mode == 2: # right only → reset left
            self._val_left = 0.0
            self._ball_left.setData([0.5], [0.0])
        self._update_yaxis_ui()
        self._check_hold_task_availability()
        self._update_target_enabled()

    def _commit_trial(self):
        if self._cur_max_left is not None:
            self._history_left.append((self._cur_max_left, self._cur_min_left))
        if self._cur_max_right is not None:
            self._history_right.append((self._cur_max_right, self._cur_min_right))
        self._cur_max_left = self._cur_min_left = None
        self._cur_max_right = self._cur_min_right = None
        self.btn_redo_left.setEnabled(bool(self._history_left))
        self.btn_redo_right.setEnabled(bool(self._history_right))

    def _on_mvc_clicked(self):
        _hold = self.chk_hold_task.isChecked()
        if self._streaming:
            # Stop pressed — immediately go to RELAX
            self._stream_timer.stop()
            self._streaming = False
            if not _hold:
                self._commit_trial()
            self._fade_start_val_l = self._val_left
            self._fade_start_val_r = self._val_right
            self._fade_elapsed     = 0.0
            self.btn_mvc.setText("HOLD" if _hold else "MVC")
            self._set_ball_color(140, 140, 140)
            self._set_status("RELAX", "w")
            self._set_countdown("")
            self._fade_timer.start()
            return
        client = getattr(self._viewer, "client", None)
        if client is None:
            return
        self._streaming  = True
        self._cur_max_left  = None
        self._cur_min_left  = None
        self._cur_max_right = None
        self._cur_min_right = None
        _dur = self.spin_duration.value()
        self._task_end_t = float('inf') if _dur == 0 else time.time() + _dur
        self._fade_timer.stop()
        self.btn_mvc.setText("Stop")
        self._set_ball_color(220, 30, 30)
        if not _hold:
            self._set_status("GO", (0, 200, 80))
        else:
            self._set_status("", "w")
        self._stream_timer.start()

    def _on_stream_tick(self):
        client = getattr(self._viewer, "client", None)
        if client is None or not self._streaming:
            self._stream_timer.stop()
            return

        now = time.time()
        if self._task_end_t != float('inf'):
            secs_left = max(0.0, self._task_end_t - now)
            self._set_countdown(f"{secs_left:.1f}")
        else:
            self._set_countdown("")

        ch_l = self.combo_ch_left.currentIndex()
        ch_r = self.combo_ch_right.currentIndex()
        try:
            data, _, _ = client.get_latest_data(window_secs=4.0,
                                                  channels=[ch_l, ch_r])
            for ch_idx, attr, active in (
                (ch_l, "_val_left",  self._switch_mode != 2),
                (ch_r, "_val_right", self._switch_mode != 1),
            ):
                if not active:
                    continue
                arr = data.get(ch_idx)
                if arr is not None and len(arr):
                    finite = arr[np.isfinite(arr)]
                    n = len(finite)
                    if n:
                        nonzero_idx = np.flatnonzero(finite)
                        if len(nonzero_idx):
                            real_end  = int(nonzero_idx[-1]) + 1
                            tail_size = max(1, n // 40)
                            tail      = finite[max(0, real_end - tail_size):real_end]
                            v = float(np.sqrt(np.mean(tail ** 2)))
                        else:
                            v = 0.0
                        setattr(self, attr, v)
        except Exception:
            pass

        _hold = self.chk_hold_task.isChecked()
        if self._switch_mode != 2:
            if _hold:
                try:
                    mx = float(self.edit_avg_max_left.text())
                    ball_y_l = (self._val_left / mx * 100) if mx > 0 else 0.0
                except ValueError:
                    ball_y_l = self._val_left
                # color: green inside band, red outside
                if self._band_left is not None:
                    pct = self.spin_tgt_pct_left.value()
                    tol = self.spin_tgt_tol_left.value()
                    in_band = (pct - tol) <= ball_y_l <= (pct + tol)
                    self._ball_left.setBrush(
                        pg.mkBrush(30, 200, 50) if in_band else pg.mkBrush(220, 30, 30))
            else:
                ball_y_l = self._val_left
                v = self._val_left
                self._cur_max_left = v if self._cur_max_left is None else max(self._cur_max_left, v)
                self._cur_min_left = v if self._cur_min_left is None else min(self._cur_min_left, v)
            self._ball_left.setData([0.5], [ball_y_l])
        if self._switch_mode != 1:
            if _hold:
                try:
                    mx = float(self.edit_avg_max_right.text())
                    ball_y_r = (self._val_right / mx * 100) if mx > 0 else 0.0
                except ValueError:
                    ball_y_r = self._val_right
                if self._band_right is not None:
                    pct = self.spin_tgt_pct_right.value()
                    tol = self.spin_tgt_tol_right.value()
                    in_band = (pct - tol) <= ball_y_r <= (pct + tol)
                    self._ball_right.setBrush(
                        pg.mkBrush(30, 200, 50) if in_band else pg.mkBrush(220, 30, 30))
            else:
                ball_y_r = self._val_right
                v = self._val_right
                self._cur_max_right = v if self._cur_max_right is None else max(self._cur_max_right, v)
                self._cur_min_right = v if self._cur_min_right is None else min(self._cur_min_right, v)
            self._ball_right.setData([0.5], [ball_y_r])
        if not _hold:
            self._update_stats()

        if now >= self._task_end_t:
            _hold = self.chk_hold_task.isChecked()
            self._stream_timer.stop()
            self._streaming = False
            if not _hold:
                self._commit_trial()
            self._fade_start_val_l = self._val_left
            self._fade_start_val_r = self._val_right
            self._fade_elapsed     = 0.0
            self.btn_mvc.setText("HOLD" if _hold else "MVC")
            self._set_ball_color(140, 140, 140)
            self._set_status("RELAX", "w")
            self._set_countdown("")
            self._fade_timer.start()

    _FADE_SECS = 1.0

    def _on_fade_tick(self):
        self._fade_elapsed += 0.05
        t = min(self._fade_elapsed / self._FADE_SECS, 1.0)
        self._val_left  = self._fade_start_val_l * (1.0 - t)
        self._val_right = self._fade_start_val_r * (1.0 - t)
        if self.chk_hold_task.isChecked():
            try:
                mx_l = float(self.edit_avg_max_left.text())
                bl = (self._val_left  / mx_l * 100) if mx_l > 0 else 0.0
            except ValueError:
                bl = self._val_left
            try:
                mx_r = float(self.edit_avg_max_right.text())
                br = (self._val_right / mx_r * 100) if mx_r > 0 else 0.0
            except ValueError:
                br = self._val_right
        else:
            bl, br = self._val_left, self._val_right
        self._ball_left.setData( [0.5], [bl])
        self._ball_right.setData([0.5], [br])
        if t >= 1.0:
            self._val_left  = 0.0
            self._val_right = 0.0
            self._fade_timer.stop()
            self._set_status("WAIT", "w")

    def _on_redo_clicked(self, side):
        if self.chk_hold_task.isChecked():
            return
        if side == "left" and self._history_left:
            self._history_left.pop()
        elif side == "right" and self._history_right:
            self._history_right.pop()
        self.btn_redo_left.setEnabled(bool(self._history_left))
        self.btn_redo_right.setEnabled(bool(self._history_right))
        self._update_stats()

    def _update_stats(self):
        def _fmt(history, cur_max, cur_min):
            n = len(history)
            if self._streaming and cur_max is not None:
                mx, mn = cur_max, cur_min
                stats = f"n={n+1}*  MAX: {mx:.4f}  MIN: {mn:.4f}"
            elif n > 0:
                mx, mn = history[-1]
                stats = f"n={n}   MAX: {mx:.4f}  MIN: {mn:.4f}"
            else:
                stats = "n= —    MAX: —    MIN: —"

            all_entries = list(history)
            if self._streaming and cur_max is not None:
                all_entries.append((cur_max, cur_min))
            if all_entries:
                avg_mx = sum(h[0] for h in all_entries) / len(all_entries)
                avg_mn = sum(h[1] for h in all_entries) / len(all_entries)
            else:
                avg_mx = avg_mn = None
            return stats, avg_mx, avg_mn

        s_l, amx_l, amn_l = _fmt(self._history_left,  self._cur_max_left,  self._cur_min_left)
        s_r, amx_r, amn_r = _fmt(self._history_right, self._cur_max_right, self._cur_min_right)
        self.lbl_stats_left.setText(s_l)
        self.lbl_stats_right.setText(s_r)
        self._avg_max_left  = amx_l
        self._avg_min_left  = amn_l
        self._avg_max_right = amx_r
        self._avg_min_right = amn_r
        if amx_l is not None:
            self.edit_avg_max_left.setText(f"{amx_l:.4f}")
        if amn_l is not None:
            self.edit_avg_min_left.setText(f"{amn_l:.4f}")
        if amx_r is not None:
            self.edit_avg_max_right.setText(f"{amx_r:.4f}")
        if amn_r is not None:
            self.edit_avg_min_right.setText(f"{amn_r:.4f}")
        self._check_hold_task_availability()

    def _copy_edit(self, edit_widget):
        text = edit_widget.text().strip()
        if text:
            QApplication.clipboard().setText(text)

    def _check_hold_task_availability(self):
        has_l = bool(self.edit_avg_max_left.text().strip())
        has_r = bool(self.edit_avg_max_right.text().strip())
        mode = self._switch_mode
        if mode == 0:
            available = has_l and has_r
        elif mode == 1:
            available = has_l
        else:
            available = has_r
        if not available and self.chk_hold_task.isChecked():
            self.chk_hold_task.setChecked(False)
        self.chk_hold_task.setEnabled(available)

    def _update_target_enabled(self):
        if not self.chk_hold_task.isChecked():
            return
        left_active  = self._switch_mode != 2
        right_active = self._switch_mode != 1
        has_l = bool(self.edit_avg_max_left.text().strip())
        has_r = bool(self.edit_avg_max_right.text().strip())
        self.spin_tgt_pct_left.setEnabled(left_active)
        self.spin_tgt_tol_left.setEnabled(left_active)
        self.btn_apply_left.setEnabled(left_active and has_l)
        self.spin_tgt_pct_right.setEnabled(right_active)
        self.spin_tgt_tol_right.setEnabled(right_active)
        self.btn_apply_right.setEnabled(right_active and has_r)

    def _on_hold_task_toggled(self, checked):
        if checked:
            self.lbl_window_title.setText("Hold Task")
            self.btn_mvc.setText("HOLD")
            self.combo_ch_left.setEnabled(False)
            self.combo_ch_right.setEnabled(False)
            self.spin_yaxis_left.setEnabled(False)
            self.spin_yaxis_right.setEnabled(False)
            self.btn_redo_left.setEnabled(False)
            self.btn_redo_right.setEnabled(False)
            # %MVC mode: Y axis → 0–100 for both plots
            self.plot_left.setYRange(0, 100)
            self.plot_right.setYRange(0, 100)
            self._reposition_texts("left",  100)
            self._reposition_texts("right", 100)
            self._tgt_widget_left.setEnabled(True)
            self._tgt_widget_right.setEnabled(True)
            self._update_target_enabled()
            # auto-apply targets for sides with MVC set
            if bool(self.edit_avg_max_left.text().strip()):
                self._on_apply_clicked("left")
            if bool(self.edit_avg_max_right.text().strip()):
                self._on_apply_clicked("right")
            self.spin_duration.setValue(0)
        else:
            self.lbl_window_title.setText("Maximum Voluntary Contraction")
            self.btn_mvc.setText("MVC")
            self.combo_ch_left.setEnabled(True)
            self.combo_ch_right.setEnabled(True)
            self.spin_yaxis_left.setEnabled(True)
            self.spin_yaxis_right.setEnabled(True)
            self.btn_redo_left.setEnabled(bool(self._history_left))
            self.btn_redo_right.setEnabled(bool(self._history_right))
            # restore Y range from spinboxes
            v_l = self.spin_yaxis_left.value()
            self.plot_left.setYRange(0, v_l)
            self._reposition_texts("left", v_l)
            v_r = self.spin_yaxis_right.value()
            self.plot_right.setYRange(0, v_r)
            self._reposition_texts("right", v_r)
            self._tgt_widget_left.setEnabled(False)
            self._tgt_widget_right.setEnabled(False)
            self._clear_bands()

    def _on_apply_clicked(self, side):
        pct = getattr(self, f"spin_tgt_pct_{side}").value()
        tol = getattr(self, f"spin_tgt_tol_{side}").value()
        low  = max(0.0,   pct - tol)
        high = min(100.0, pct + tol)
        self._clear_band(side)
        plot = self.plot_left if side == "left" else self.plot_right
        band = pg.LinearRegionItem(
            values=(low, high),
            orientation="horizontal",
            movable=False,
            brush=pg.mkBrush(160, 160, 160, 55),
            pen=pg.mkPen(180, 180, 180, 160, width=2),
        )
        band.setZValue(-10)
        line = pg.InfiniteLine(
            pos=pct, angle=0, movable=False,
            pen=pg.mkPen(220, 220, 220, 230, width=3),
        )
        plot.addItem(band)
        plot.addItem(line)
        setattr(self, f"_band_{side}",      band)
        setattr(self, f"_band_line_{side}", line)

    def _clear_band(self, side):
        plot = self.plot_left if side == "left" else self.plot_right
        band = getattr(self, f"_band_{side}", None)
        line = getattr(self, f"_band_line_{side}", None)
        if band is not None:
            plot.removeItem(band)
            setattr(self, f"_band_{side}", None)
        if line is not None:
            plot.removeItem(line)
            setattr(self, f"_band_line_{side}", None)

    def _clear_bands(self):
        self._clear_band("left")
        self._clear_band("right")

    def _poll_lc_status(self):
        client = getattr(self._viewer, "client", None)
        if client is None:
            self.lbl_lc_status.setText("● —")
            self.lbl_lc_status.setStyleSheet("color: gray;")
            return
        try:
            if client.is_sampling():
                self.lbl_lc_status.setText("● PLAY")
                self.lbl_lc_status.setStyleSheet("color: #00cc44;")
            else:
                self.lbl_lc_status.setText("● STOP")
                self.lbl_lc_status.setStyleSheet("color: #ff4444;")
        except Exception:
            self.lbl_lc_status.setText("● —")
            self.lbl_lc_status.setStyleSheet("color: gray;")

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
        if getattr(self, "_force_close", False):
            if hasattr(self, "_calc_window") and self._calc_window is not None:
                self._calc_window.close()
            super().closeEvent(event)
        else:
            if self._streaming:
                self._on_mvc_clicked()
            event.ignore()
            self.hide()


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
        self._fro_single_pending = False
        self._fro_firing = False
        self._fro_last_config = None
        self._fro_bounced = False
        self._com_thread_stop = threading.Event()
        self._magstim_port = None
        self._magstim_stop = threading.Event()
        self._magstim_last_poll_t = 0.0
        self._mso1_val = 0
        self._mso2_val = 0
        self._isi_val  = "—"
        self._temp_val = "—"
        self._magstim_polled.connect(self._on_magstim_polled)
        self._plot_window   = None
        self._active_window = None
        self.window_secs   = WINDOW_SECS_DEFAULT
        self.scope_pre     = SCOPE_PRE_DEFAULT
        self.scope_post    = SCOPE_POST_DEFAULT
        self.mode          = "chart"   # "chart" | "scope"

        pg.setConfigOption("background", "w")
        pg.setConfigOption("foreground", "k")

        self._setup_ui()

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

        _sf = QFont(); _sf.setPointSize(10)
        status_row = QHBoxLayout()
        self.lbl_status = QLabel("Connecting to LabChart...")
        self.lbl_status.setAlignment(Qt.AlignCenter)
        self.lbl_status.setFont(_sf)
        status_row.addWidget(self.lbl_status, stretch=1)
        _sep = QLabel("|")
        _sep.setAlignment(Qt.AlignCenter)
        _sep.setFont(_sf)
        _sep.setStyleSheet("color: #aaaaaa;")
        status_row.addWidget(_sep)
        self.lbl_magstim_status = QLabel("MAGIC: Not connected")
        self.lbl_magstim_status.setAlignment(Qt.AlignCenter)
        self.lbl_magstim_status.setFont(_sf)
        self.lbl_magstim_status.setStyleSheet("color: #aaaaaa;")
        status_row.addWidget(self.lbl_magstim_status, stretch=1)

        self.lbl_triggerbox_status = QLabel("TriggerBox: Not connected")
        self.lbl_triggerbox_status.setAlignment(Qt.AlignCenter)
        self.lbl_triggerbox_status.setFont(_sf)
        self.lbl_triggerbox_status.setStyleSheet("color: #aaaaaa;")
        self.lbl_triggerbox_status.setVisible(False)
        status_row.addWidget(self.lbl_triggerbox_status, stretch=1)

        root.addLayout(status_row)

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
        mode_row.addSpacing(8)
        btn_active = QPushButton("Active")
        btn_active.setFixedWidth(70)
        btn_active.clicked.connect(self._on_active_clicked)
        mode_row.addWidget(btn_active)

        mode_row.addStretch()
        return mode_row

    def _make_analysis_stack(self):
        w_scope_analysis = QWidget()
        analysis_row = QHBoxLayout(w_scope_analysis)
        analysis_row.setContentsMargins(0, 0, 0, 0)
        analysis_row.setSpacing(12)

        self.radio_none = QRadioButton("None")
        self.radio_mep  = QRadioButton("rMT")
        self.radio_sici = QRadioButton("Paired-Pulse (FRO)")
        self.radio_mep.setChecked(True)
        self.radio_none.setEnabled(False)
        self.radio_mep.setEnabled(False)
        self.radio_sici.setEnabled(False)
        analysis_group = QButtonGroup(self)
        analysis_group.addButton(self.radio_none)
        analysis_group.addButton(self.radio_mep)
        analysis_group.addButton(self.radio_sici)
        self.radio_mep.toggled.connect(self._on_analysis_mode_changed)
        self.radio_sici.toggled.connect(self._on_analysis_mode_changed)
        self._btn_setup_guide = QPushButton("SETUP-GUIDE")
        self._btn_setup_guide.setStyleSheet(
            "QPushButton { background-color: #c62828; color: white; font-weight: bold; }"
            "QPushButton:hover { background-color: #b71c1c; }"
            "QPushButton:pressed { background-color: #7f0000; }"
        )
        self._btn_setup_guide.setFixedHeight(int(self._btn_setup_guide.sizeHint().height() * 0.9))
        self._btn_setup_guide.setVisible(False)
        self._btn_setup_guide.clicked.connect(self._on_setup_guide_clicked)

        analysis_row.addWidget(self.radio_none)
        analysis_row.addWidget(self.radio_mep)
        analysis_row.addWidget(self.radio_sici)
        analysis_row.addWidget(self._btn_setup_guide)

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

        reconnect_font = QFont(); reconnect_font.setPointSize(11)
        self.btn_magstim_reconnect = QPushButton("↺")
        self.btn_magstim_reconnect.setFont(reconnect_font)
        self.btn_magstim_reconnect.setFixedWidth(28)
        self.btn_magstim_reconnect.setToolTip("Connect Magstim and use MAGIC")
        self.btn_magstim_reconnect.clicked.connect(self._on_magstim_btn_clicked)
        mso_row.addWidget(self.btn_magstim_reconnect)
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
        self.lbl_magstim_dot.setToolTip("Not connected")
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
        self.spin_mso1 = _WheelSpinBox()
        self.spin_mso1.setRange(0, 100); self.spin_mso1.setValue(0)
        self.spin_mso1.setFont(mso_font); self.spin_mso1.setFixedWidth(60)
        self.spin_mso1.setButtonSymbols(QAbstractSpinBox.NoButtons)
        self.spin_mso1.setToolTip("↑ Master TMS")
        mso_row.addWidget(self.spin_mso1)
        mso_row.addSpacing(16)

        self.lbl_mso2 = QLabel("MSO2:"); self.lbl_mso2.setFont(mso_font)
        self.lbl_mso2.setStyleSheet("color: transparent;")
        mso_row.addWidget(self.lbl_mso2)
        self.spin_mso2 = _WheelSpinBox()
        self.spin_mso2.setRange(0, 100); self.spin_mso2.setValue(0)
        self.spin_mso2.setFont(mso_font); self.spin_mso2.setFixedWidth(60)
        self.spin_mso2.setButtonSymbols(QAbstractSpinBox.NoButtons)
        self.spin_mso2.setToolTip("↓ Slave TMS")
        self.spin_mso2.setStyleSheet("color: transparent; background: transparent; border: transparent;")
        mso_row.addWidget(self.spin_mso2)

        isi_font = QFont(); isi_font.setPointSize(13)
        self.lbl_ipi = QLabel("  ISI:"); self.lbl_ipi.setFont(isi_font)
        self.lbl_ipi.setStyleSheet("color: transparent;")
        mso_row.addWidget(self.lbl_ipi)
        self.spin_isi = QDoubleSpinBox()
        self.spin_isi.setRange(0.0, 999.9); self.spin_isi.setValue(0.0)
        self.spin_isi.setDecimals(1); self.spin_isi.setSuffix(" ms")
        self.spin_isi.setFont(isi_font); self.spin_isi.setFixedWidth(90)
        self.spin_isi.setButtonSymbols(QAbstractSpinBox.NoButtons)
        self.spin_isi.setStyleSheet("color: transparent; background: transparent; border: transparent;")
        mso_row.addWidget(self.spin_isi)
        mso_row.addSpacing(16)

        self._lbl_loc = QLabel("Loc ID:"); self._lbl_loc.setFont(mso_font)
        mso_row.addWidget(self._lbl_loc)
        self.spin_location = _WheelSpinBox()
        self.spin_location.setRange(1, 9999); self.spin_location.setValue(1)
        self.spin_location.setFixedWidth(100); self.spin_location.setFont(mso_font)
        self.spin_location.setButtonSymbols(QAbstractSpinBox.UpDownArrows)
        mso_row.addWidget(self.spin_location)

        mso_row.addStretch(1)
        self._btn_new_best = QPushButton("New Best"); self._btn_new_best.setFont(btn_font)
        self._btn_new_best.clicked.connect(self._on_new_best)
        mso_row.addWidget(self._btn_new_best)
        self.combo_best = QComboBox(); self.combo_best.setFont(btn_font)
        self.combo_best.setMinimumWidth(180)
        mso_row.addWidget(self.combo_best)
        return mso_row

    def _make_scope_below_stack(self):
        w_scope_below = QWidget()
        scope_vbox = QVBoxLayout(w_scope_below)
        scope_vbox.setContentsMargins(0, 0, 0, 0)
        scope_vbox.setSpacing(8)

        self._hunt_row_widget = QWidget()
        hunt_row = QHBoxLayout(self._hunt_row_widget)
        hunt_row.setContentsMargins(0, 0, 0, 0)
        hunt_row.setSpacing(8)
        hunt_row.addSpacing(81)
        hunt_row.addLayout(self._make_hunt_panel("left"))
        hunt_row.addSpacing(135)
        hunt_row.addLayout(self._make_hunt_panel("right"))
        hunt_row.addSpacing(81)
        # SICI row (shown instead of hunt row in SICI mode)
        self._sici_row_widget = QWidget()
        self._sici_row_widget.setVisible(False)
        sici_row = QHBoxLayout(self._sici_row_widget)
        sici_row.setContentsMargins(0, 0, 0, 0)
        sici_row.setSpacing(10)

        sici_font = QFont(); sici_font.setPointSize(12)

        self.btn_single = QPushButton("Single Pulse")
        self.btn_single.setFont(sici_font)
        self.btn_single.clicked.connect(self._on_single_pulse_clicked)
        self.btn_double = QPushButton("Double Pulse")
        self.btn_double.setFont(sici_font)
        self.btn_double.clicked.connect(self._on_double_pulse_clicked)

        self._sici_isi_spin = QDoubleSpinBox()
        self._sici_isi_spin.setRange(0.1, 999.9)
        self._sici_isi_spin.setValue(2.5)
        self._sici_isi_spin.setSingleStep(0.1)
        self._sici_isi_spin.setDecimals(1)
        self._sici_isi_spin.setSuffix("")
        self._sici_isi_spin.setButtonSymbols(QAbstractSpinBox.NoButtons)
        self._sici_isi_spin.setFixedWidth(80)
        self._sici_isi_spin.setFont(sici_font)

        tb_btn_font = QFont(); tb_btn_font.setPointSize(11)
        self.btn_triggerbox_connect = QPushButton("↺")
        self.btn_triggerbox_connect.setFont(tb_btn_font)
        self.btn_triggerbox_connect.setFixedWidth(28)
        self.btn_triggerbox_connect.setToolTip("Trigger Box와 연결하기")
        self.btn_triggerbox_connect.clicked.connect(self._on_triggerbox_btn_clicked)

        self._lbl_fro_countdown = QLabel("")
        self._lbl_fro_countdown.setFont(sici_font)
        self._lbl_fro_countdown.setFixedWidth(36)
        self._lbl_fro_countdown.setAlignment(Qt.AlignCenter)
        self._lbl_fro_countdown.setStyleSheet("color: #e65100; font-weight: bold;")

        self._fro_countdown_timer = QTimer(self)
        self._fro_countdown_timer.setInterval(1000)
        self._fro_countdown_timer.timeout.connect(self._tick_fro_countdown)
        self._fro_countdown_val = 0

        sici_row.addSpacing(81)
        sici_row.addWidget(self.btn_triggerbox_connect)
        sici_row.addWidget(self.btn_single)
        sici_row.addWidget(self.btn_double)
        sici_row.addSpacing(12)
        lbl_isi_sici = QLabel("ISI (ms):")
        lbl_isi_sici.setFont(sici_font)
        sici_row.addWidget(lbl_isi_sici)
        sici_row.addWidget(self._sici_isi_spin)
        sici_row.addWidget(self._lbl_fro_countdown)
        sici_row.addStretch()

        # Stack hunt and sici rows in the same vertical space
        row_stack = QWidget()
        row_stack_layout = QVBoxLayout(row_stack)
        row_stack_layout.setContentsMargins(0, 0, 0, 0)
        row_stack_layout.setSpacing(0)
        row_stack_layout.addWidget(self._hunt_row_widget)
        row_stack_layout.addWidget(self._sici_row_widget)
        scope_vbox.addWidget(row_stack)

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
        prestim_start = QSpinBox(); prestim_start.setRange(-1000, -1); prestim_start.setValue(-150); prestim_start.setButtonSymbols(QAbstractSpinBox.NoButtons)
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

    def _on_active_clicked(self):
        if self._active_window is None:
            self._active_window = ActiveWindow(self)
        self._active_window.populate_ch_combos()
        self._active_window.show()
        self._active_window.raise_()
        self._active_window.activateWindow()

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

        self._do_clear_all()

    def _do_clear_all(self):
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

        self.radio_mep.setEnabled(not chart)
        self.radio_sici.setEnabled(not chart)
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
        is_rmt  = self.radio_mep.isChecked()
        is_sici = self.radio_sici.isChecked()
        is_any  = is_rmt or is_sici

        def _has_data():
            return (any(s.table.rowCount() > 0 for s in (self.L, self.R))
                    or bool(self.trigger_times))

        if (is_sici or is_rmt) and _has_data():
            from PyQt5.QtWidgets import QMessageBox
            target = "Paired-Pulse (FRO)" if is_sici else "rMT"
            msg = QMessageBox(self)
            msg.setWindowTitle("Switch Mode")
            msg.setText(f"Switching to {target} mode will clear all current session data.")
            msg.setStandardButtons(QMessageBox.Ok | QMessageBox.Cancel)
            msg.button(QMessageBox.Cancel).setText("Return")
            msg.setDefaultButton(QMessageBox.Cancel)
            if msg.exec_() != QMessageBox.Ok:
                self.radio_mep.blockSignals(True)
                self.radio_sici.blockSignals(True)
                if is_sici:
                    self.radio_mep.setChecked(True)
                else:
                    self.radio_sici.setChecked(True)
                self.radio_mep.blockSignals(False)
                self.radio_sici.blockSignals(False)
                return
            self._do_clear_all()

        self._btn_setup_guide.setVisible(is_sici)
        self.widget_mep_params.setEnabled(is_any)

        # Hunt row: visible only for rMT; SICI row: visible only for SICI
        self._hunt_row_widget.setVisible(is_rmt)
        self._sici_row_widget.setVisible(is_sici)
        for s in (self.L, self.R):
            s.hunt_panel.setEnabled(is_rmt)

        # MAGIC / TriggerBox status toggle
        self.btn_magstim_reconnect.setEnabled(not is_sici)
        self.lbl_magstim_status.setVisible(not is_sici)
        self.lbl_triggerbox_status.setVisible(is_sici)
        if is_sici:
            self._stop_com_thread()
        else:
            self._stop_triggerbox_retry()
            self._register_com_events()

        # SICI constraints
        self.btn_sm.setEnabled(not is_sici)
        _isi_style = "color: transparent;" if is_sici else ""
        self.lbl_ipi.setStyleSheet(_isi_style)
        self.spin_isi.setStyleSheet(_isi_style)
        sici_location_enabled = not is_sici
        self._lbl_loc.setEnabled(sici_location_enabled)
        self.spin_location.setEnabled(sici_location_enabled)
        self._btn_new_best.setEnabled(sici_location_enabled)
        self.combo_best.setEnabled(sici_location_enabled)

        if is_sici:
            if self.is_running:
                self._stop()
            # Force DM mode
            self._on_dm_clicked()
            # Auto-scan and connect TriggerBox
            QTimer.singleShot(100, self._on_triggerbox_btn_clicked)

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
        mso1 = self.spin_mso1.value()
        loc  = self.spin_location.value()
        if self._magstim_mode == "DM":
            mso2 = self.spin_mso2.value()
            isi  = self.spin_isi.value()
            label = f"{mso1}_{mso2}_{isi:.1f}_{loc}"
        else:
            label = f"{mso1}_{loc}"

        # collect existing items, add new one, sort by MSO1 (first number before _)
        items = [self.combo_best.itemText(i) for i in range(self.combo_best.count())]
        items.append(label)

        def _mso1_key(s):
            try:
                return int(s.split("_")[0])
            except ValueError:
                return 0

        items.sort(key=_mso1_key)
        self.combo_best.clear()
        for item in items:
            self.combo_best.addItem(item)
        self.combo_best.setCurrentText(label)

    def _on_sm_clicked(self):
        self._magstim_mode = "SM"
        self.btn_sm.setStyleSheet("background-color: #1565C0; color: white;")
        self.btn_dm.setStyleSheet("")
        self.lbl_mso2.setStyleSheet("color: transparent;")
        self.spin_mso2.setStyleSheet("color: transparent; background: transparent; border: transparent;")
        self.lbl_ipi.setStyleSheet("color: transparent;")
        self.spin_isi.setStyleSheet("color: transparent; background: transparent; border: transparent;")
        self.combo_best.clear()

    def _on_dm_clicked(self):
        self._magstim_mode = "DM"
        self.btn_dm.setStyleSheet("background-color: #1565C0; color: white;")
        self.btn_sm.setStyleSheet("")
        self.lbl_mso2.setStyleSheet("")
        self.spin_mso2.setStyleSheet("")
        self.combo_best.clear()
        self.lbl_ipi.setStyleSheet("")
        self.spin_isi.setStyleSheet("")

    def _on_setup_guide_clicked(self):
        dlg = QDialog(self, Qt.Popup)
        dlg.setFixedSize(520, 520)

        layout = QVBoxLayout(dlg)
        layout.setContentsMargins(20, 20, 20, 20)

        img_path = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                "reference", "photo", "pic1.jpeg")
        img_src = "file:///" + img_path.replace("\\", "/")

        browser = QTextBrowser()
        browser.setOpenExternalLinks(False)
        browser.setFrameShape(QFrame.NoFrame)
        browser.setHtml(f"""
<h3>Paired-Pulse (FRO) Mode</h3>
<p>This mode allows TMSviewer to control TMS stimulation directly from the PC by sending TTL
pulses via the Trigger Box (ch5). The signal chain is:</p>
<p><b>PC &rarr; Trigger Box &rarr; PowerLab / LabChart &rarr; TMS</b></p>
<p>LabChart registers the assigned stimulation command (Single Pulse or Paired Pulse with a
configurable ISI) using Fast Response Output (FRO), then delivers TTL pulses through
Output 1 and Output 2. The TMS receives these signals and fires accordingly.</p>

<hr>

<h4>Hardware Setup</h4>
<p>Set the BiStim to <b>BiStim mode</b> and enable <b>Independent Triggering Mode</b>.</p>

<p><img src="{img_src}" width="360"></p>

<p>In PowerLab...</p>
<p><b>A. Conditioning &rarr; Testing pulse:</b><br>
&nbsp;&nbsp;&nbsp;&nbsp;connect J BNC to Output 1, Y BNC to Output 2</p>
<p><b>B. Testing &rarr; Conditioning pulse:</b><br>
&nbsp;&nbsp;&nbsp;&nbsp;connect Y BNC to Output 1, J BNC to Output 2</p>

<hr>

<h4>&#9888; If you were using MAGIC in rMT mode</h4>
<ol>
  <li>Turn off the TMS machine, PowerLab, and LabChart.</li>
  <li>Disengage all MAGIC-related cable setup.</li>
  <li>Restart the TMS machine, PowerLab, and LabChart.</li>
</ol>
""")
        layout.addWidget(browser)
        dlg.exec_()

    # ------------------------------------------------------------------
    # FRO helpers (SICI Single/Double Pulse)
    # ------------------------------------------------------------------

    def _load_vbs_hex(self, rel_path):
        """Extract PlayMessage hex string from a .vbs file."""
        import re
        base = os.path.dirname(os.path.abspath(__file__))
        path = os.path.join(base, rel_path)
        try:
            with open(path, "r") as f:
                content = f.read()
        except OSError:
            return None
        m = re.search(r'PlayMessage\s*\("(0x[0-9A-Fa-f]+)"\)', content)
        return m.group(1) if m else None

    _FRO_TWO_STEP = False  # True: template → 50ms → modified; False: modified only

    def _set_fro_output_delay(self, hex_str, output_num, delay_s):
        """Replace the PulseDelay of a specific output and fix checksum (bytes 20-23)."""
        raw = bytearray(bytes.fromhex(hex_str[2:]))
        orig_sum      = sum(raw)
        output_marker = f"Output = {output_num}\n".encode("utf-16-le")
        key           = "PulseDelay = ".encode("utf-16-le")
        newline_le    = b"\x0a\x00"

        pos = raw.find(bytes(output_marker))
        if pos == -1:
            return hex_str
        pd_pos = raw.find(bytes(key), pos)
        if pd_pos == -1:
            return hex_str

        val_start = pd_pos + len(key)
        val_end   = raw.find(newline_le, val_start)
        if val_end == -1:
            return hex_str

        old_len = val_end - val_start
        new_val = f"{delay_s:.4f}".encode("utf-16-le")
        new_len = len(new_val)

        if new_len < old_len:
            new_val = new_val + b"\x20\x00" * ((old_len - new_len) // 2)
        elif new_len > old_len:
            new_val = new_val[:old_len]

        raw[val_start:val_end] = new_val

        # fix checksum at bytes 20-23 (sum of all bytes, uint32 little-endian)
        delta    = sum(raw) - orig_sum
        old_csum = int.from_bytes(raw[20:24], "little")
        new_csum = (old_csum + delta) % 0x100000000
        raw[20:24] = new_csum.to_bytes(4, "little")

        return "0x" + raw.hex().upper()

    # ------------------------------------------------------------------
    # TriggerBox helpers
    # ------------------------------------------------------------------

    def _ensure_triggerbox(self):
        if not _SERIAL_AVAILABLE:
            return False
        tb = getattr(self, "_triggerbox_port", None)
        if tb and tb.is_open:
            return True
        try:
            self._triggerbox_port = serial.Serial(
                TRIGGERBOX_PORT, TRIGGERBOX_BAUD, timeout=0.1
            )
            self._triggerbox_port.write(bytes([0]))  # reset on open
            self._triggerbox_port.flush()
            return True
        except Exception:
            self._triggerbox_port = None
            return False

    def _on_triggerbox_btn_clicked(self):
        if not _SERIAL_AVAILABLE:
            self.lbl_triggerbox_status.setText("TriggerBox: pyserial not installed")
            self.lbl_triggerbox_status.setStyleSheet("color: #c62828;")
            return

        # Close any existing connection before rescanning
        tb = getattr(self, "_triggerbox_port", None)
        if tb and tb.is_open:
            try:
                tb.write(bytes([0]))
                tb.flush()
                tb.close()
            except Exception:
                pass
        self._triggerbox_port = None

        self.lbl_triggerbox_status.setText("⟳  Scanning COM ports for TriggerBox...")
        self.lbl_triggerbox_status.setStyleSheet("color: #1565C0;")
        self.btn_triggerbox_connect.setEnabled(False)
        QApplication.processEvents()

        found_port = None
        for port_info in serial.tools.list_ports.comports():
            desc = (port_info.description or "").lower()
            mfr  = (port_info.manufacturer or "").lower()
            if "brain products" in desc or "brain products" in mfr or "triggerbox" in desc:
                found_port = port_info.device
                break

        # Fallback: try each port at 115200
        if found_port is None:
            for port_info in serial.tools.list_ports.comports():
                try:
                    s = serial.Serial(port_info.device, TRIGGERBOX_BAUD, timeout=0.1)
                    s.write(bytes([0]))
                    s.flush()
                    s.close()
                    found_port = port_info.device
                    break
                except Exception:
                    continue

        self.btn_triggerbox_connect.setEnabled(True)

        if found_port:
            try:
                self._triggerbox_port = serial.Serial(found_port, TRIGGERBOX_BAUD, timeout=0.1)
                self._triggerbox_port.write(bytes([0]))
                self._triggerbox_port.flush()
                self.lbl_triggerbox_status.setText(f"TriggerBox: Connected ({found_port})")
                self.lbl_triggerbox_status.setStyleSheet("color: #2e7d32;")
                self._stop_triggerbox_retry()
                return
            except Exception:
                self.lbl_triggerbox_status.setText(f"TriggerBox: Failed to open {found_port} — retrying...")
                self.lbl_triggerbox_status.setStyleSheet("color: #c62828;")
        else:
            self.lbl_triggerbox_status.setText("TriggerBox: Not found — retrying...")
            self.lbl_triggerbox_status.setStyleSheet("color: #c62828;")

        self._start_triggerbox_retry()

    def _start_triggerbox_retry(self):
        if not hasattr(self, "_triggerbox_retry_timer"):
            self._triggerbox_retry_timer = QTimer(self)
            self._triggerbox_retry_timer.timeout.connect(self._on_triggerbox_btn_clicked)
        if not self._triggerbox_retry_timer.isActive():
            self._triggerbox_retry_timer.start(3000)

    def _stop_triggerbox_retry(self):
        if hasattr(self, "_triggerbox_retry_timer"):
            self._triggerbox_retry_timer.stop()

    def _fire_triggerbox(self):
        if not self._ensure_triggerbox():
            return
        channel_byte = 1 << (TRIGGERBOX_CHANNEL - 1)
        try:
            self._triggerbox_port.write(bytes([channel_byte]))
            self._triggerbox_port.flush()
            QTimer.singleShot(TRIGGERBOX_PULSE_MS, self._reset_triggerbox)
        except Exception:
            pass
        if self.radio_sici.isChecked():
            t = getattr(self, '_fro_pending_trigger_t', None)
            try:
                self._register_trigger(t if t is not None else 0.0)
            except Exception as e:
                print(f"[FRO] register_trigger failed: {e}")

    def _reset_triggerbox(self):
        tb = getattr(self, "_triggerbox_port", None)
        if tb and tb.is_open:
            try:
                tb.write(bytes([0]))
                tb.flush()
            except Exception:
                pass

    def _fro_firing_done(self):
        self._fro_firing = False
        if self.is_running:
            self.data_timer.start(UPDATE_MS)

    def _start_fro_countdown(self):
        self._fro_countdown_val = 5
        self._lbl_fro_countdown.setText("5")
        self.btn_single.setEnabled(False)
        self.btn_double.setEnabled(False)
        self._fro_countdown_timer.start()

    def _tick_fro_countdown(self):
        self._fro_countdown_val -= 1
        if self._fro_countdown_val <= 0:
            self._fro_countdown_timer.stop()
            self._lbl_fro_countdown.setText("")
            self.btn_single.setEnabled(True)
            self.btn_double.setEnabled(True)
        else:
            self._lbl_fro_countdown.setText(str(self._fro_countdown_val))

    def _on_single_pulse_clicked(self):
        if self.client is None:
            return
        try:
            self._fro_pending_trigger_t = self.client.current_time() + 0.2001
        except Exception:
            self._fro_pending_trigger_t = None
        self._fro_single_pending = True
        if self._fro_last_config != "SP":
            sp_hex = self._load_vbs_hex(os.path.join("reference", "FRO", "SinglePulse.vbs"))
            if not sp_hex:
                return
            self._fro_firing = True
            self.data_timer.stop()
            try:
                self.client.play_message(sp_hex)
                self._fro_last_config = "SP"
            except Exception as e:
                print(f"[SinglePulse] PlayMessage failed: {e}")
                self._fro_firing = False
                if self.is_running:
                    self.data_timer.start(UPDATE_MS)
                return
        else:
            self._fro_firing = True
            self.data_timer.stop()
        QTimer.singleShot(150, self._fire_triggerbox)
        if self.is_running:
            QTimer.singleShot(500, self._fro_firing_done)
        else:
            self._fro_firing = False
        self._start_fro_countdown()

    def _on_double_pulse_clicked(self):
        if self.client is None:
            return
        isi_s  = self._sici_isi_spin.value() / 1000.0
        out2_s = 0.0501 + isi_s
        config_key = f"DP_{out2_s:.4f}"
        try:
            self._fro_pending_trigger_t = self.client.current_time() + 0.2001
        except Exception:
            self._fro_pending_trigger_t = None
        if self._fro_last_config != config_key:
            dp_hex = self._load_vbs_hex(os.path.join("reference", "FRO", "DoublePulse.vbs"))
            if not dp_hex:
                return
            modified = self._set_fro_output_delay(dp_hex, 2, out2_s)
            self._fro_firing = True
            self.data_timer.stop()
            try:
                self.client.play_message(modified)
                self._fro_last_config = config_key
            except Exception as e:
                print(f"[DoublePulse] PlayMessage failed: {e}")
                self._fro_firing = False
                if self.is_running:
                    self.data_timer.start(UPDATE_MS)
                return
        else:
            self._fro_firing = True
            self.data_timer.stop()
        QTimer.singleShot(150, self._fire_triggerbox)
        if self.is_running:
            QTimer.singleShot(500, self._fro_firing_done)
        else:
            self._fro_firing = False
        self._start_fro_countdown()

    def _refresh_com_ports(self):
        pass  # combo removed; port stored in _magstim_selected_port

    def _auto_find_magstim_port(self):
        """Scan all COM ports and return the first one that responds to Magstim Q@ command."""
        if not _SERIAL_AVAILABLE:
            return None
        ports = sorted(p.device for p in serial.tools.list_ports.comports())
        for port_name in ports:
            try:
                with serial.Serial(port_name, 9600, bytesize=serial.EIGHTBITS,
                                   parity=serial.PARITY_NONE, stopbits=serial.STOPBITS_ONE,
                                   timeout=0.2) as test_port:
                    if enable_remote(test_port):
                        disable_remote(test_port)
                        return port_name
            except Exception:
                continue
        return None

    def _on_magstim_btn_clicked(self):
        if self._magstim_port is not None and self._magstim_port.is_open:
            self._disconnect_magstim()
        else:
            self._connect_magstim_auto()

    def _disconnect_magstim(self):
        self._magstim_stop.set()
        if self._magstim_port is not None:
            try:
                self._magstim_port.close()
            except Exception:
                pass
            self._magstim_port = None
        self._magstim_last_poll_t = 0.0
        self._magstim_stop = threading.Event()
        self.lbl_magstim_dot.setStyleSheet("color: #aaaaaa;")
        self.lbl_magstim_dot.setToolTip("Not connected")
        self.lbl_magstim_state.setText("—")
        self.lbl_magstim_state.setStyleSheet("color: #aaaaaa;")
        self.lbl_magstim_status.setText("MAGIC: Not connected")
        self.lbl_magstim_status.setStyleSheet("color: #aaaaaa;")
        self.btn_magstim_reconnect.setText("↺")
        self.btn_magstim_reconnect.setToolTip("Connect Magstim and use MAGIC")
        self.spin_mso1.setReadOnly(False)
        self.spin_mso2.setReadOnly(False)
        self.spin_isi.setReadOnly(False)

    def _connect_magstim_auto(self):
        from PyQt5.QtWidgets import QMessageBox, QApplication
        self._magstim_stop.set()
        if self._magstim_port is not None:
            try:
                self._magstim_port.close()
            except Exception:
                pass
            self._magstim_port = None
        self._magstim_last_poll_t = 0.0
        self._magstim_stop = threading.Event()
        self.lbl_magstim_status.setText("⟳  Scanning COM ports for Magstim device...")
        self.lbl_magstim_status.setStyleSheet("color: #1565C0;")
        QApplication.processEvents()
        found = self._auto_find_magstim_port()
        if found:
            self._magstim_selected_port = found
            self._connect_magstim()
            QMessageBox.information(self, "Magstim", f"Magstim detected on {found}.")
        else:
            self.lbl_magstim_status.setText("MAGIC: Not connected")
            self.lbl_magstim_status.setStyleSheet("color: #aaaaaa;")
            QMessageBox.warning(self, "Magstim", "Magstim not detected.\nPlease check that the device is connected.")

    def _connect_magstim(self, port=None):
        if not _SERIAL_AVAILABLE:
            return
        selected_port = port or getattr(self, "_magstim_selected_port", MAGSTIM_PORT)
        if not selected_port:
            return
        try:
            ser = serial.Serial(
                port=selected_port, baudrate=9600,
                bytesize=serial.EIGHTBITS, parity=serial.PARITY_NONE,
                stopbits=serial.STOPBITS_ONE, timeout=0.3,
            )
            self._magstim_port = ser
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

        port = getattr(self, "_magstim_selected_port", MAGSTIM_PORT)
        self.lbl_magstim_dot.setStyleSheet("color: #2e7d32;")
        self.lbl_magstim_dot.setToolTip(port)
        self.lbl_magstim_state.setText(state)
        self.lbl_magstim_state.setStyleSheet(
            "color: #2e7d32;" if f["ready"] else
            "color: #e65100;" if f["armed"] else
            "color: #555555;"
        )
        self.lbl_magstim_status.setText(f"MAGIC: Connected - {port}")
        self.lbl_magstim_status.setStyleSheet("color: #2e7d32;")
        self.btn_magstim_reconnect.setText("✕")
        self.btn_magstim_reconnect.setToolTip("Disconnect MAGIC and control manually")

        if temp is not None:
            t1 = temp["temp1"] / 10.0
            t2 = temp["temp2"] / 10.0
            avg_c = (t1 + t2) / 2.0
            self._temp_val = f"{avg_c:.1f}"
            self.lbl_magstim_temp.setText(f"{avg_c:.1f}°C")
            self.lbl_magstim_temp.setToolTip(
                f"Coil1: {t1:.1f}°C    Coil2: {t2:.1f}°C"
            )

        self.spin_mso1.setValue(params["power_a"])
        self.spin_mso1.setReadOnly(True)
        if self._magstim_mode == "DM":
            self.spin_mso2.setValue(params["power_b"])
            self.spin_isi.setValue(params["ipi"] / 10.0)
            self.spin_mso2.setReadOnly(True)
            self.spin_isi.setReadOnly(True)
        else:
            self.spin_mso2.setReadOnly(True)
            self.spin_isi.setReadOnly(True)

    def closeEvent(self, event):
        if self.is_running:
            self._stop()
        self._magstim_stop.set()
        if self._magstim_port is not None:
            try:
                disable_remote(self._magstim_port)
            except Exception:
                pass
            try:
                self._magstim_port.close()
            except Exception:
                pass
        if self._plot_window is not None:
            self._plot_window.close()
        if self._active_window is not None:
            self._active_window._force_close = True
            self._active_window.close()
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
        self._com_thread_stop.clear()
        self._com_thread = threading.Thread(target=self._com_thread_func, daemon=True)
        self._com_thread.start()
        print("[COM events] background thread started")

    def _stop_com_thread(self):
        if self._com_thread is not None and self._com_thread.is_alive():
            self._com_thread_stop.set()
            self._com_thread.join(timeout=0.5)
            self._com_thread = None
            print("[COM events] background thread stopped")

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
            while not self._com_thread_stop.is_set():
                if not self._fro_firing:
                    pythoncom.PumpWaitingMessages()
                time.sleep(0.05)
            print("[COM thread] stopped cleanly")
        except Exception as e:
            print(f"[COM thread] exiting: {e}")
        finally:
            pythoncom.CoUninitialize()

    def _on_tms_trigger(self, event_args=()):
        if self.client is None:
            return
        if self.radio_sici.isChecked():
            return  # FRO mode: trigger registered directly via Sample button
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
            self.lbl_magstim_dot.setToolTip("Not connected")
            self.lbl_magstim_state.setText("—")
            self.lbl_magstim_state.setStyleSheet("color: #aaaaaa;")
            self.lbl_magstim_temp.setText("—")
            self.lbl_magstim_temp.setToolTip("Coil temp unavailable")
            self.lbl_magstim_status.setText("MAGIC: Not connected")
            self.lbl_magstim_status.setStyleSheet("color: #aaaaaa;")
            self.btn_magstim_reconnect.setText("↺")
            self.btn_magstim_reconnect.setToolTip("Connect Magstim and use MAGIC")
            self.spin_mso1.setReadOnly(False)
            self.spin_mso2.setReadOnly(False)
            self.spin_isi.setReadOnly(False)

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
        if self.radio_sici.isChecked():
            self._on_triggerbox_btn_clicked()
        if self.radio_sici.isChecked():
            self._first_play = False
            self._fro_preload()
        elif self._first_play:
            self._first_play = False
            try:
                self.client.start_sampling()
            except Exception:
                pass
            QTimer.singleShot(500, self._bounce_stop)
        else:
            self._do_start()

    def _fro_preload(self):
        """Send SP + DP PlayMessages BEFORE sampling starts to prime PowerLab FRO."""
        self.btn_single.setEnabled(False)
        self.btn_double.setEnabled(False)
        sp_hex = self._load_vbs_hex(os.path.join("reference", "FRO", "SinglePulse.vbs"))
        if sp_hex:
            try:
                self.client.play_message(sp_hex)
                self._fro_last_config = "SP"
                print("[FRO preload] SP sent")
            except Exception as e:
                print(f"[FRO preload] SP failed: {e}")
        QTimer.singleShot(300, self._fro_preload_dp)

    def _fro_preload_dp(self):
        if self.client is None:
            self._fro_preload_done()
            return
        dp_hex = self._load_vbs_hex(os.path.join("reference", "FRO", "DoublePulse.vbs"))
        if dp_hex:
            isi_s  = self._sici_isi_spin.value() / 1000.0
            out2_s = 0.0501 + isi_s
            modified = self._set_fro_output_delay(dp_hex, 2, out2_s)
            try:
                self.client.play_message(modified)
                self._fro_last_config = f"DP_{out2_s:.4f}"
                print("[FRO preload] DP sent")
            except Exception as e:
                print(f"[FRO preload] DP failed: {e}")
        QTimer.singleShot(300, self._fro_preload_done)

    def _fro_preload_done(self):
        print("[FRO preload] complete — starting sampling")
        try:
            self.client.start_sampling()
        except Exception as e:
            print(f"[StartSampling] {e}")
        self._start_time = time.monotonic()

        if not self._fro_bounced:
            # First pass: auto-bounce to initialize PowerLab FRO
            self._fro_last_config = None
            QTimer.singleShot(500, self._fro_auto_bounce)
        else:
            # Second pass: fully initialized
            self._fro_bounced = False
            self.btn_single.setEnabled(True)
            self.btn_double.setEnabled(True)
            self.data_timer.start(UPDATE_MS)

    def _fro_auto_bounce(self):
        if self.client is None:
            self.btn_single.setEnabled(True)
            self.btn_double.setEnabled(True)
            self.data_timer.start(UPDATE_MS)
            return
        print("[FRO auto-bounce] stop → restart")
        try:
            self.client.stop_sampling()
        except Exception as e:
            print(f"[FRO auto-bounce] stop failed: {e}")
        self._fro_bounced = True
        QTimer.singleShot(300, self._fro_preload)

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

    def _fro_warmup(self):
        self.btn_single.setEnabled(False)
        self.btn_double.setEnabled(False)
        QTimer.singleShot(500, self._fro_warmup_sp)

    def _fro_warmup_sp(self):
        if self.client is None:
            self._fro_warmup_done()
            return
        sp_hex = self._load_vbs_hex(os.path.join("reference", "FRO", "SinglePulse.vbs"))
        if sp_hex:
            try:
                self.client.play_message(sp_hex)
                self._fro_last_config = "SP"
                print("[FRO warmup] SP sent")
            except Exception as e:
                print(f"[FRO warmup] SP failed: {e}")
        QTimer.singleShot(300, self._fro_warmup_dp)

    def _fro_warmup_dp(self):
        if self.client is None:
            self._fro_warmup_done()
            return
        dp_hex = self._load_vbs_hex(os.path.join("reference", "FRO", "DoublePulse.vbs"))
        if dp_hex:
            isi_s  = self._sici_isi_spin.value() / 1000.0
            out2_s = 0.0501 + isi_s
            modified = self._set_fro_output_delay(dp_hex, 2, out2_s)
            try:
                self.client.play_message(modified)
                self._fro_last_config = f"DP_{out2_s:.4f}"
                print("[FRO warmup] DP sent")
            except Exception as e:
                print(f"[FRO warmup] DP failed: {e}")
        QTimer.singleShot(300, self._fro_warmup_done)

    def _fro_warmup_done(self):
        self.btn_single.setEnabled(True)
        self.btn_double.setEnabled(True)
        self.data_timer.start(UPDATE_MS)
        print("[FRO warmup] complete")

    def _stop(self):
        self.data_timer.stop()
        if self.client is not None:
            try:
                self.client.stop_sampling()
            except Exception as e:
                print(f"[StopSampling] {e}")
        self.is_running = False
        self._fro_last_config = None
        self._fro_bounced = False
        self.btn.setText("▶  Play")

    def _add_sample_comment(self):
        if self.client is None:
            return
        if self.radio_sici.isChecked():
            # FRO mode: COM thread is paused, register trigger directly
            try:
                self.client.add_comment("Trigger")
            except Exception:
                pass
            self._register_trigger(self.client.current_time())
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
        _single_pulse = self._fro_single_pending
        self._fro_single_pending = False
        mso2_str = "0" if _single_pulse else (str(self.spin_mso2.value()) if self._magstim_mode == "DM" else "0")
        isi_str  = f"{self.spin_isi.value():.1f}" if self._magstim_mode == "DM" else "0"
        row_vals = [
            lambda s: str(trial_num),
            lambda s: str(s.ch + 1),
            lambda s: self._get_hand(s),
            lambda s: self._get_paralysis(s),
            lambda s: str(loc_id),
            lambda s: str(self.spin_mso1.value()),
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
        import sys
        search_dir = Path(sys.executable).parent if getattr(sys, "frozen", False) else Path(__file__).parent
        adiset = sorted(search_dir.glob("*.adiset"))
        if adiset:
            os.startfile(str(adiset[0]))
            return
        adicht = sorted(search_dir.glob("*.adicht"))
        if adicht:
            os.startfile(str(adicht[0]))
            return
        self.lbl_status.setText("✕  No .adiset or .adicht file found in app folder")
        self.lbl_status.setStyleSheet("color: #b71c1c;")

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
                        brush=pg.mkBrush(160, 160, 160, 60),
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
        ymin, ymax = float(np.nanmin(valid)), float(np.nanmax(valid))
        if not (np.isfinite(ymin) and np.isfinite(ymax)):
            return None
        return ymin, ymax

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
