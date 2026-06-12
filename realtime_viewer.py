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
)
from PyQt5.QtCore import QTimer, Qt, pyqtSignal
from PyQt5.QtGui import QFont, QColor

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
        self._magstim_polled.connect(self._on_magstim_polled)
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
        btn_save = QPushButton("Save"); btn_save.setFixedWidth(130)
        btn_save.clicked.connect(self._save_tables)
        tables_row.addSpacing(81)
        tables_row.addWidget(self.L.table, stretch=1)
        tables_row.addWidget(btn_save, alignment=Qt.AlignCenter)
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
        self.btn.clicked.connect(self._toggle)

        self.btn_sample = QPushButton("Sample")
        self.btn_sample.setFixedHeight(43)
        self.btn_sample.setFont(_bf)
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
        layout.addStretch()

        s = self.L if side == "left" else self.R
        s.hunt_panel  = panel
        s.hunt_chk    = chk
        s.hunt_num_lbl = lbl_num
        s.hunt_den_lbl = lbl_den
        s.hunt_clear  = btn_clear
        s.hunt_redo   = btn_redo
        btn_clear.clicked.connect(lambda: self._on_hunt_clear(side))
        btn_redo.clicked.connect(lambda: self._on_hunt_redo(side))

        outer = QHBoxLayout()
        outer.setContentsMargins(0, 0, 0, 0)
        outer.addWidget(panel)
        outer.addStretch()
        return outer

    def _save_tables(self):
        from PyQt5.QtWidgets import QFileDialog
        desktop = os.path.join(os.path.expanduser("~"), "Desktop")
        path, _ = QFileDialog.getSaveFileName(
            self, "Save Table", os.path.join(desktop, "tms_data.txt"),
            "Text files (*.txt)"
        )
        if not path:
            return

        def table_to_text(tbl, title):
            cols = [tbl.horizontalHeaderItem(c).text() for c in range(tbl.columnCount())]
            lines = [title, "\t".join(cols)]
            for r in range(tbl.rowCount()):
                row_data = []
                for c in range(tbl.columnCount()):
                    item = tbl.item(r, c)
                    row_data.append(item.text() if item else "")
                lines.append("\t".join(row_data))
            return "\n".join(lines)

        with open(path, "w", encoding="utf-8") as f:
            f.write(table_to_text(self.L.table, "=== Left EMG ==="))
            f.write("\n\n")
            f.write(table_to_text(self.R.table, "=== Right EMG ==="))

    def _make_data_table(self):
        cols = ["Channel", "Loc ID", "%MSO1", "Trial", "MEP amp", "Prestim RMS"]
        tbl = QTableWidget(0, len(cols))
        tbl.setHorizontalHeaderLabels(cols)
        tbl.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
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
            self.lbl_magstim_temp.setText(f"{avg_c:.1f}°C")
            self.lbl_magstim_temp.setToolTip(
                f"Coil1: {t1:.1f}°C    Coil2: {t2:.1f}°C"
            )

        self._mso1_val = params["power_a"]
        self.lbl_mso_val.setText(str(params["power_a"]))
        if self._magstim_mode == "DM":
            self.lbl_mso2_val.setText(str(params["power_b"]))
            self.lbl_isi_val.setText(f"{params['ipi'] / 10.0:.1f} ms")

    def closeEvent(self, event):
        self._magstim_stop.set()
        if self._magstim_port is not None:
            try:
                self._magstim_port.close()
            except Exception:
                pass
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
        loc_id  = self.spin_location.value()
        mso_pct = self._mso1_val
        for s in (self.L, self.R):
            row = s.table.rowCount()
            s.table.insertRow(row)
            for col, val in enumerate([str(s.ch + 1), str(loc_id), str(mso_pct), str(trial_num), "—", "—"]):
                item = QTableWidgetItem(val)
                item.setTextAlignment(Qt.AlignCenter)
                s.table.setItem(row, col, item)
            s.table.scrollToBottom()
        for s in (self.L, self.R):
            if s.hunt_panel.isEnabled() and s.hunt_chk.isChecked():
                s.hunt_history.append((s.hunt_num, s.hunt_den))
                s.hunt_den += 1
                self._update_hunt_display("left" if s is self.L else "right")

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
            for s, r in ((self.L, r_L), (self.R, r_R)):
                if not s.table_latest_eval and r is not None:
                    self._set_table_cell(s.table, row, 4, f"{r['mep_amp']:.3f} mV")
                    self._set_table_cell(s.table, row, 5, f"{r['prestim_rms']:.3f} mV")
                    s.table_latest_eval = True

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
