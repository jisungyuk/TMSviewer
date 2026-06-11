import os
import time
import threading
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
    _tms_detected = pyqtSignal(object)  # emitted from COM background thread, carries (record, tick)

    def __init__(self):
        super().__init__()
        pythoncom.CoInitialize()
        self.client        = None
        self._event_sink   = None
        self._com_thread   = None
        self.is_running    = False
        self._start_time   = None
        self._first_play   = False
        self._next_trial_num = 1
        self._tms_detected.connect(self._on_tms_trigger)
        self.ch_info       = []
        self.ch_left       = DEFAULT_CH_LEFT
        self.ch_right      = DEFAULT_CH_RIGHT
        self.trigger_times   = []
        self._last_trigger_t = None
        self.separate = False
        self.hunt_num_val_left  = 0;  self.hunt_den_val_left  = 0
        self.hunt_num_val_right = 0;  self.hunt_den_val_right = 0
        self.hunt_history_left  = []  # list of (num, den) before each sample
        self.hunt_history_right = []
        self.hunt_latest_eval_left  = True
        self.hunt_latest_eval_right = True
        self.table_latest_eval_left  = True
        self.table_latest_eval_right = True
        self._yfix_left    = False
        self._yfix_right   = False
        self._yhalf_left   = 1.0
        self._yhalf_right  = 1.0
        self._vlines_left  = []
        self._vlines_right = []
        self._shades_left  = []
        self._shades_right = []
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
        self.setCentralWidget(central)

        # Scale default (unsized) font by 10%
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

        root = QVBoxLayout(central)
        root.setContentsMargins(10, 10, 10, 10)
        root.setSpacing(8)

        # ── Mode selector row ───────────────────────────────────────────
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

        # Chart params (window size)
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

        # Scope params (pre / post trigger)
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
        root.addLayout(mode_row)

        # ── Analysis row — radio buttons only ───────────────────────────
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
        self.stack_analysis.addWidget(w_scope_analysis)  # index 0: scope
        self.stack_analysis.addWidget(QWidget())          # index 1: chart
        root.addWidget(self.stack_analysis)

        # MEP param spinboxes — placed in centre column (built here, inserted below)
        self.spin_mep_start = QSpinBox()
        self.spin_mep_start.setRange(1, 500)
        self.spin_mep_start.setValue(10)

        self.spin_mep_end = QSpinBox()
        self.spin_mep_end.setRange(1, 500)
        self.spin_mep_end.setValue(50)
        self.spin_mep_end.setSuffix(" ms")

        self.spin_threshold = QDoubleSpinBox()
        self.spin_threshold.setRange(0.01, 100)
        self.spin_threshold.setSingleStep(0.01)
        self.spin_threshold.setDecimals(2)
        self.spin_threshold.setValue(0.05)
        self.spin_threshold.setSuffix(" mV")

        self.spin_prestim_start = QSpinBox()
        self.spin_prestim_start.setRange(-1000, -1)
        self.spin_prestim_start.setValue(-200)

        self.spin_prestim_end = QSpinBox()
        self.spin_prestim_end.setRange(-500, -1)
        self.spin_prestim_end.setValue(-50)
        self.spin_prestim_end.setSuffix(" ms")

        # ── Channel dropdowns ───────────────────────────────────────────
        combos_row = QHBoxLayout()
        combos_row.setSpacing(8)
        self.combo_left  = QComboBox()
        self.combo_right = QComboBox()
        self.combo_left.currentIndexChanged.connect(self._on_combo_left)
        self.combo_right.currentIndexChanged.connect(self._on_combo_right)
        combos_row.addWidget(QLabel("Left channel:"))
        combos_row.addWidget(self.combo_left, stretch=1)
        combos_row.addSpacing(150)   # matches centre column width
        combos_row.addWidget(QLabel("Right channel:"))
        combos_row.addWidget(self.combo_right, stretch=1)
        root.addLayout(combos_row)

        # ── Two plots + trigger list in the middle ──────────────────────
        plots_row = QHBoxLayout()
        plots_row.setSpacing(8)

        self.plot_left  = self._make_plot("Left EMG",  COLOR_LEFT)
        self.plot_right = self._make_plot("Right EMG", COLOR_RIGHT)
        self.curve_left  = self.plot_left.plot(pen=pg.mkPen(COLOR_LEFT,  width=1))
        self.curve_right = self.plot_right.plot(pen=pg.mkPen(COLOR_RIGHT, width=1))

        # Centre column — trigger counter + MEP params
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

        # MEP params set 1 (left graph always)
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

        # MEP params set 2 (right graph — always visible, enabled only when Separate is ON)
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

        self.panel_left,  self.stats_left  = self._make_side_panel("left")
        self.panel_right, self.stats_right = self._make_side_panel("right")

        self.stack_centre = _MaxSizeStack()
        self.stack_centre.setFixedWidth(135)
        self.stack_centre.addWidget(self.centre_widget)  # index 0: scope
        self.stack_centre.addWidget(QWidget())            # index 1: chart

        plots_row.addWidget(self.panel_left)
        plots_row.addWidget(self.plot_left,  stretch=1)
        plots_row.addWidget(self.stack_centre)
        plots_row.addWidget(self.plot_right, stretch=1)
        plots_row.addWidget(self.panel_right)
        root.addLayout(plots_row, stretch=1)

        # ── Scope-only bottom section (stacked with chart page) ─────────
        w_scope_below = QWidget()
        scope_vbox = QVBoxLayout(w_scope_below)
        scope_vbox.setContentsMargins(0, 0, 0, 0)
        scope_vbox.setSpacing(8)

        # Hunt row
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

        # MSO / Location row
        mso_row = QHBoxLayout()
        mso_row.setSpacing(12)
        mso_font = QFont(); mso_font.setPointSize(23)
        btn_font = QFont(); btn_font.setPointSize(16)

        mso_row.addStretch(1)
        lbl_mso = QLabel("MSO %:"); lbl_mso.setFont(mso_font)
        mso_row.addWidget(lbl_mso)
        self.spin_mso = _WheelSpinBox()
        self.spin_mso.setRange(0, 100); self.spin_mso.setValue(0)
        self.spin_mso.setFixedWidth(81); self.spin_mso.setFont(mso_font)
        self.spin_mso.setButtonSymbols(QAbstractSpinBox.NoButtons)
        mso_row.addWidget(self.spin_mso)
        mso_row.addSpacing(32)
        lbl_loc = QLabel("Location ID:"); lbl_loc.setFont(mso_font)
        mso_row.addWidget(lbl_loc)
        self.spin_location = _WheelSpinBox()
        self.spin_location.setRange(0, 9999); self.spin_location.setValue(0)
        self.spin_location.setFixedWidth(81); self.spin_location.setFont(mso_font)
        self.spin_location.setButtonSymbols(QAbstractSpinBox.NoButtons)
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
        scope_vbox.addLayout(mso_row)

        sep2 = QFrame(); sep2.setFrameShape(QFrame.HLine); sep2.setFrameShadow(QFrame.Sunken)
        scope_vbox.addWidget(sep2)

        # Data tables
        tables_row = QHBoxLayout(); tables_row.setSpacing(8)
        self.table_left  = self._make_data_table()
        self.table_right = self._make_data_table()
        btn_save = QPushButton("Save"); btn_save.setFixedWidth(130)
        btn_save.clicked.connect(self._save_tables)
        tables_row.addSpacing(81)
        tables_row.addWidget(self.table_left,  stretch=1)
        tables_row.addWidget(btn_save, alignment=Qt.AlignCenter)
        tables_row.addWidget(self.table_right, stretch=1)
        tables_row.addSpacing(81)
        scope_vbox.addLayout(tables_row)

        self.stack_below = _MaxSizeStack()
        self.stack_below.addWidget(w_scope_below)  # index 0: scope
        self.stack_below.addWidget(QWidget())       # index 1: chart
        root.addWidget(self.stack_below)

        # ── LabChart status label ───────────────────────────────────────
        self.lbl_status = QLabel("Connecting to LabChart...")
        self.lbl_status.setAlignment(Qt.AlignCenter)
        f = QFont()
        f.setPointSize(10)
        self.lbl_status.setFont(f)
        root.addWidget(self.lbl_status)

        # ── Play / Stop + Sample buttons ───────────────────────────────
        btn_row = QHBoxLayout()
        btn_row.setSpacing(8)
        f2 = QFont()
        f2.setPointSize(13)

        self.btn = QPushButton("▶  Play")
        self.btn.setFixedHeight(43)
        self.btn.setFont(f2)
        self.btn.clicked.connect(self._toggle)

        self.btn_sample = QPushButton("Sample")
        self.btn_sample.setFixedHeight(43)
        self.btn_sample.setFont(f2)
        self.btn_sample.setEnabled(False)
        self.btn_sample.setStyleSheet(
            "QPushButton:enabled  { background-color: #2e7d32; color: white; }"
            "QPushButton:disabled { background-color: #cccccc; color: #888888; }"
        )
        self.btn_sample.clicked.connect(self._add_sample_comment)

        btn_row.addWidget(self.btn, stretch=1)
        btn_row.addWidget(self.btn_sample, stretch=1)
        root.addLayout(btn_row)

        self._on_mode_changed()

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

        if side == "left":
            self.hunt_panel_left  = panel
            self.hunt_chk_left    = chk
            self.hunt_num_left    = lbl_num
            self.hunt_den_left    = lbl_den
            self.hunt_clear_left  = btn_clear
            self.hunt_redo_left   = btn_redo
            btn_clear.clicked.connect(lambda: self._on_hunt_clear("left"))
            btn_redo.clicked.connect(lambda: self._on_hunt_redo("left"))
        else:
            self.hunt_panel_right = panel
            self.hunt_chk_right   = chk
            self.hunt_num_right   = lbl_num
            self.hunt_den_right   = lbl_den
            self.hunt_clear_right = btn_clear
            self.hunt_redo_right  = btn_redo
            btn_clear.clicked.connect(lambda: self._on_hunt_clear("right"))
            btn_redo.clicked.connect(lambda: self._on_hunt_redo("right"))

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
            f.write(table_to_text(self.table_left,  "=== Left EMG ==="))
            f.write("\n\n")
            f.write(table_to_text(self.table_right, "=== Right EMG ==="))

    def _make_data_table(self):
        cols = ["Channel", "Loc ID", "%MSO", "Trial", "MEP amp", "Prestim RMS"]
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

        if side == "left":
            self._btn_yfix_left  = btn_fix
            self._btn_yauto_left = btn_auto
            btn_plus.clicked.connect(lambda: self._yzoom("left", 1 / 1.25))
            btn_minus.clicked.connect(lambda: self._yzoom("left", 1.25))
            btn_fix.clicked.connect(lambda: self._yfix_clicked("left"))
            btn_auto.clicked.connect(lambda: self._yauto_clicked("left"))
        else:
            self._btn_yfix_right  = btn_fix
            self._btn_yauto_right = btn_auto
            btn_plus.clicked.connect(lambda: self._yzoom("right", 1 / 1.25))
            btn_minus.clicked.connect(lambda: self._yzoom("right", 1.25))
            btn_fix.clicked.connect(lambda: self._yfix_clicked("right"))
            btn_auto.clicked.connect(lambda: self._yauto_clicked("right"))

        return outer, stats

    # ------------------------------------------------------------------
    # Y-axis zoom / fix / auto
    # ------------------------------------------------------------------

    def _yzoom(self, side, factor):
        if side == "left":
            plot = self.plot_left
            if not self._yfix_left:
                vr = plot.viewRange()
                self._yhalf_left = max(abs(vr[1][0]), abs(vr[1][1]), 1e-6)
                self._yfix_left = True
                self._btn_yfix_left.setChecked(True)
                self._btn_yauto_left.setChecked(False)
            self._yhalf_left = max(self._yhalf_left * factor, 1e-6)
            plot.enableAutoRange(axis="y", enable=False)
            plot.setYRange(-self._yhalf_left, self._yhalf_left, padding=0)
        else:
            plot = self.plot_right
            if not self._yfix_right:
                vr = plot.viewRange()
                self._yhalf_right = max(abs(vr[1][0]), abs(vr[1][1]), 1e-6)
                self._yfix_right = True
                self._btn_yfix_right.setChecked(True)
                self._btn_yauto_right.setChecked(False)
            self._yhalf_right = max(self._yhalf_right * factor, 1e-6)
            plot.enableAutoRange(axis="y", enable=False)
            plot.setYRange(-self._yhalf_right, self._yhalf_right, padding=0)

    def _yfix_clicked(self, side):
        if side == "left":
            vr = self.plot_left.viewRange()
            self._yhalf_left = max(abs(vr[1][0]), abs(vr[1][1]), 1e-6)
            self._yfix_left = True
            self._btn_yfix_left.setChecked(True)
            self._btn_yauto_left.setChecked(False)
            self.plot_left.enableAutoRange(axis="y", enable=False)
            self.plot_left.setYRange(-self._yhalf_left, self._yhalf_left, padding=0)
        else:
            vr = self.plot_right.viewRange()
            self._yhalf_right = max(abs(vr[1][0]), abs(vr[1][1]), 1e-6)
            self._yfix_right = True
            self._btn_yfix_right.setChecked(True)
            self._btn_yauto_right.setChecked(False)
            self.plot_right.enableAutoRange(axis="y", enable=False)
            self.plot_right.setYRange(-self._yhalf_right, self._yhalf_right, padding=0)

    def _yauto_clicked(self, side):
        if side == "left":
            self._yfix_left = False
            self._btn_yfix_left.setChecked(False)
            self._btn_yauto_left.setChecked(True)
        else:
            self._yfix_right = False
            self._btn_yfix_right.setChecked(False)
            self._btn_yauto_right.setChecked(True)

    # ------------------------------------------------------------------

    def _update_stats(self, lbl, arr, t_start, trigger_t, t_scale, second=False):
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

        n       = len(arr)
        dur     = self.scope_pre + self.scope_post
        spt     = dur / max(n - 1, 1)
        t_rel   = np.arange(n) * spt - self.scope_pre  # seconds relative to trigger

        mep_mask = (t_rel >= mep_s) & (t_rel <= mep_e)
        pre_mask = (t_rel >= pre_s) & (t_rel <= pre_e)

        mep_data = arr[mep_mask]
        pre_data = arr[pre_mask]

        if mep_data.size > 0:
            mep_amp = float(np.nanmax(mep_data) - np.nanmin(mep_data))
            mep_str = f"{mep_amp:.3f} mV"
        else:
            mep_str = "—"

        if pre_data.size > 0:
            rms = float(np.sqrt(np.nanmean(pre_data ** 2)))
            rms_str = f"{rms:.3f} mV"
        else:
            rms_str = "—"

        lbl.setText(f"Prestim RMS\n{rms_str}\n\nMEP amp\n{mep_str}")

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
        self.panel_left.setVisible(not chart)
        self.panel_right.setVisible(not chart)
        if chart:
            self.plot_left.enableAutoRange(axis="y", enable=True)
            self.plot_right.enableAutoRange(axis="y", enable=True)
        else:
            self.plot_left.enableAutoRange(axis="y", enable=False)
            self.plot_right.enableAutoRange(axis="y", enable=False)
        self.widget_chart_params.setVisible(chart)
        self.widget_scope_params.setVisible(not chart)
        self.trigger_spin.setEnabled(not chart)

        self.radio_none.setEnabled(not chart)
        self.radio_mep.setEnabled(not chart)
        mep_active = (not chart) and self.radio_mep.isChecked()
        self.hunt_panel_left.setEnabled(mep_active)
        self.hunt_panel_right.setEnabled(mep_active)

        x_label = "Time (s)" if chart else "Time from trigger (s)"
        self.plot_left.setLabel("bottom",  x_label)
        self.plot_right.setLabel("bottom", x_label)

        # Clear plots on mode switch
        self.curve_left.setData([], [])
        self.curve_right.setData([], [])
        self._clear_trigger_lines()

        # Scope로 돌아올 때 즉시 재그리기 (data_timer 상태와 무관하게)
        if not chart and self.trigger_times:
            channels = list({self.ch_left, self.ch_right})
            self._update_scope(channels, selected_row=self.trigger_spin.value() - 1)

    # ------------------------------------------------------------------
    # Parameter apply
    # ------------------------------------------------------------------

    def _on_analysis_mode_changed(self):
        mep = self.radio_mep.isChecked()
        self.widget_mep_params.setEnabled(mep)
        self.hunt_panel_left.setEnabled(mep)
        self.hunt_panel_right.setEnabled(mep)
        if self.radio_scope.isChecked():
            channels = [self.ch_left, self.ch_right]
            self._update_scope(channels)

    def _update_hunt_display(self, side):
        if side == "left":
            self.hunt_num_left.setText(str(self.hunt_num_val_left))
            self.hunt_den_left.setText(str(self.hunt_den_val_left))
        else:
            self.hunt_num_right.setText(str(self.hunt_num_val_right))
            self.hunt_den_right.setText(str(self.hunt_den_val_right))

    def _compute_mep_amp(self, arr, second=False):
        if second:
            mep_s = self.spin_mep_start_2.value() / 1000.0
            mep_e = self.spin_mep_end_2.value()   / 1000.0
        else:
            mep_s = self.spin_mep_start.value() / 1000.0
            mep_e = self.spin_mep_end.value()   / 1000.0
        n     = len(arr)
        dur   = self.scope_pre + self.scope_post
        spt   = dur / max(n - 1, 1)
        t_rel = np.arange(n) * spt - self.scope_pre
        mep_data = arr[(t_rel >= mep_s) & (t_rel <= mep_e)]
        if mep_data.size > 0:
            return float(np.nanmax(mep_data) - np.nanmin(mep_data))
        return 0.0

    def _compute_prestim_rms(self, arr, second=False):
        if second:
            pre_s = self.spin_prestim_start_2.value() / 1000.0
            pre_e = self.spin_prestim_end_2.value()   / 1000.0
        else:
            pre_s = self.spin_prestim_start.value() / 1000.0
            pre_e = self.spin_prestim_end.value()   / 1000.0
        n     = len(arr)
        dur   = self.scope_pre + self.scope_post
        spt   = dur / max(n - 1, 1)
        t_rel = np.arange(n) * spt - self.scope_pre
        pre_data = arr[(t_rel >= pre_s) & (t_rel <= pre_e)]
        if pre_data.size > 0:
            return float(np.sqrt(np.nanmean(pre_data ** 2)))
        return 0.0

    def _set_table_cell(self, tbl, row, col, text):
        if row < tbl.rowCount():
            item = QTableWidgetItem(text)
            item.setTextAlignment(Qt.AlignCenter)
            tbl.setItem(row, col, item)

    def _both_hunt_active(self):
        return (self.hunt_panel_left.isEnabled()  and self.hunt_chk_left.isChecked() and
                self.hunt_panel_right.isEnabled() and self.hunt_chk_right.isChecked())

    def _on_hunt_clear(self, side):
        sides = ["left", "right"] if self._both_hunt_active() else [side]
        for s in sides:
            if s == "left":
                self.hunt_num_val_left = 0;  self.hunt_den_val_left = 0
                self.hunt_history_left.clear()
                self.hunt_latest_eval_left = True
                self._update_hunt_display("left")
            else:
                self.hunt_num_val_right = 0; self.hunt_den_val_right = 0
                self.hunt_history_right.clear()
                self.hunt_latest_eval_right = True
                self._update_hunt_display("right")

    def _on_hunt_redo(self, side):
        sides = ["left", "right"] if self._both_hunt_active() else [side]
        for s in sides:
            if s == "left" and self.hunt_history_left:
                self.hunt_num_val_left, self.hunt_den_val_left = self.hunt_history_left.pop()
                self.hunt_latest_eval_left = True
                self._update_hunt_display("left")
            elif s == "right" and self.hunt_history_right:
                self.hunt_num_val_right, self.hunt_den_val_right = self.hunt_history_right.pop()
                self.hunt_latest_eval_right = True
                self._update_hunt_display("right")

    def _on_new_best(self):
        label = f"{self.spin_mso.value()}%at{self.spin_location.value()}"
        self.combo_best.addItem(label)
        idx = self.combo_best.count() - 1
        self.combo_best.setItemData(idx, QColor("red"), Qt.ForegroundRole)
        self.combo_best.setCurrentIndex(idx)

    def _on_equal_best(self):
        label = f"{self.spin_mso.value()}%at{self.spin_location.value()}"
        self.combo_best.addItem(label)
        idx = self.combo_best.count() - 1
        self.combo_best.setItemData(idx, QColor("#e65100"), Qt.ForegroundRole)
        self.combo_best.setCurrentIndex(idx)

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

    def _on_combo_left(self, idx):
        if not self.ch_info or idx < 0:
            return
        self.ch_left = idx
        name, unit = self.ch_info[idx]
        self.plot_left.setTitle(f"<b>{name}</b>", color=COLOR_LEFT, size="11pt")
        self.plot_left.setLabel("left", unit)

    def _on_combo_right(self, idx):
        if not self.ch_info or idx < 0:
            return
        self.ch_right = idx
        name, unit = self.ch_info[idx]
        self.plot_right.setTitle(f"<b>{name}</b>", color=COLOR_RIGHT, size="11pt")
        self.plot_right.setLabel("left", unit)

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
            self._event_sink = None
            self._com_thread = None
            self.btn_sample.setEnabled(False)
            if self.is_running:
                self.data_timer.stop()
                self.is_running = False
            self.btn.setText("📂  Open LabChart")
            self.lbl_status.setText("✕  LabChart Not Connected")
            self.lbl_status.setStyleSheet("color: #b71c1c;")

    def _populate_combos(self):
        new_info = self.client.get_channel_info()
        if new_info == self.ch_info:
            return
        self.ch_info = new_info
        for combo, default in (
            (self.combo_left,  DEFAULT_CH_LEFT),
            (self.combo_right, DEFAULT_CH_RIGHT),
        ):
            combo.blockSignals(True)
            combo.clear()
            for i, (name, unit) in enumerate(self.ch_info):
                combo.addItem(f"Ch{i + 1}  {name}  [{unit}]")
            combo.setCurrentIndex(min(default, len(self.ch_info) - 1))
            combo.blockSignals(False)
        self._on_combo_left(self.combo_left.currentIndex())
        self._on_combo_right(self.combo_right.currentIndex())

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
                channels = list({self.ch_left, self.ch_right})
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
        self.hunt_latest_eval_left  = False
        self.hunt_latest_eval_right = False
        self.table_latest_eval_left  = False
        self.table_latest_eval_right = False
        loc_id  = self.spin_location.value()
        mso_pct = self.spin_mso.value()
        for tbl, ch_idx in ((self.table_left, self.ch_left), (self.table_right, self.ch_right)):
            row = tbl.rowCount()
            tbl.insertRow(row)
            for col, val in enumerate([str(ch_idx + 1), str(loc_id), str(mso_pct), str(trial_num), "—", "—"]):
                item = QTableWidgetItem(val)
                item.setTextAlignment(Qt.AlignCenter)
                tbl.setItem(row, col, item)
            tbl.scrollToBottom()
        if self.hunt_panel_left.isEnabled() and self.hunt_chk_left.isChecked():
            self.hunt_history_left.append((self.hunt_num_val_left, self.hunt_den_val_left))
            self.hunt_den_val_left += 1
            self._update_hunt_display("left")
        if self.hunt_panel_right.isEnabled() and self.hunt_chk_right.isChecked():
            self.hunt_history_right.append((self.hunt_num_val_right, self.hunt_den_val_right))
            self.hunt_den_val_right += 1
            self._update_hunt_display("right")

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

    def _check_trigger_channel(self):
        try:
            data, t_start, t_end = self.client.get_latest_data(
                window_secs=0.3, channels=[TRIGGER_CH_IDX])
        except Exception:
            return
        arr = data.get(TRIGGER_CH_IDX)
        if arr is None or len(arr) == 0:
            return
        above = arr > TRIGGER_THRESHOLD
        if not above.any():
            return
        transitions = np.diff(above.astype(int))
        starts = np.where(transitions == 1)[0] + 1
        if len(starts) == 0:
            if above[0]:
                starts = np.array([0])
            else:
                return
        first_idx = int(starts[0])
        t_trigger = t_start + (t_end - t_start) * first_idx / max(len(arr) - 1, 1)
        if self._last_trigger_t is not None and t_trigger - self._last_trigger_t < TRIGGER_REFRACTORY:
            return
        self._last_trigger_t = t_trigger
        self._register_trigger(t_trigger, add_table_row=True)

    def _update(self):
        if self.client is None:
            return
        channels = list({self.ch_left, self.ch_right})

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

        self._draw(self.curve_left,  data.get(self.ch_left),  t_start, t_end, offset=0)
        self._draw(self.curve_right, data.get(self.ch_right), t_start, t_end, offset=0)
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
        self.plot_left.setLabel("bottom",  x_label)
        self.plot_right.setLabel("bottom", x_label)

        arr_left  = data.get(self.ch_left)
        arr_right = data.get(self.ch_right)
        self._draw(self.curve_left,  arr_left,  t_start, t_end, offset=trigger_t, t_scale=t_scale)
        self._draw(self.curve_right, arr_right, t_start, t_end, offset=trigger_t, t_scale=t_scale)
        self._redraw_trigger_lines_scope(t_scale=t_scale)

        # Y-range: respect Fix/Auto state per plot
        for plot, arr, yfix, yhalf in (
            (self.plot_left,  arr_left,  self._yfix_left,  self._yhalf_left),
            (self.plot_right, arr_right, self._yfix_right, self._yhalf_right),
        ):
            plot.enableAutoRange(axis="y", enable=False)
            if yfix:
                plot.setYRange(-yhalf, yhalf, padding=0)
            else:
                r = self._artifact_free_range(arr, t_start, t_end, trigger_t)
                if r is not None:
                    ymin, ymax = r
                    pad = (ymax - ymin) * 0.1 if ymax != ymin else 0.1
                    plot.setYRange(ymin - pad, ymax + pad, padding=0)

        if self.radio_mep.isChecked():
            if arr_left  is not None:
                self._update_stats(self.stats_left,  arr_left,  t_start, trigger_t, t_scale, second=False)
            if arr_right is not None:
                self._update_stats(self.stats_right, arr_right, t_start, trigger_t, t_scale, second=self.separate)
        else:
            self.stats_left.setText("Prestim RMS\n—\n\nMEP amp\n—")
            self.stats_right.setText("Prestim RMS\n—\n\nMEP amp\n—")

        # Hunt evaluation — only for the latest trigger, only once
        is_latest = (selected_row == len(self.trigger_times) - 1)
        if is_latest and self.radio_mep.isChecked():
            if (self.hunt_panel_left.isEnabled() and self.hunt_chk_left.isChecked()
                    and not self.hunt_latest_eval_left and arr_left is not None):
                amp = self._compute_mep_amp(arr_left, second=False)
                if amp > self.spin_threshold.value():
                    self.hunt_num_val_left += 1
                    self._update_hunt_display("left")
                self.hunt_latest_eval_left = True

            if (self.hunt_panel_right.isEnabled() and self.hunt_chk_right.isChecked()
                    and not self.hunt_latest_eval_right and arr_right is not None):
                threshold = self.spin_threshold_2.value() if self.separate else self.spin_threshold.value()
                amp = self._compute_mep_amp(arr_right, second=self.separate)
                if amp > threshold:
                    self.hunt_num_val_right += 1
                    self._update_hunt_display("right")
                self.hunt_latest_eval_right = True

        # Table MEP/RMS fill — latest trigger only, once
        if is_latest:
            row = selected_row  # table row index matches trigger index
            if not self.table_latest_eval_left and arr_left is not None and self.radio_mep.isChecked():
                amp = self._compute_mep_amp(arr_left, second=False)
                rms = self._compute_prestim_rms(arr_left, second=False)
                self._set_table_cell(self.table_left, row, 4, f"{amp:.3f} mV")
                self._set_table_cell(self.table_left, row, 5, f"{rms:.3f} mV")
                self.table_latest_eval_left = True
            if not self.table_latest_eval_right and arr_right is not None and self.radio_mep.isChecked():
                amp = self._compute_mep_amp(arr_right, second=self.separate)
                rms = self._compute_prestim_rms(arr_right, second=self.separate)
                self._set_table_cell(self.table_right, row, 4, f"{amp:.3f} mV")
                self._set_table_cell(self.table_right, row, 5, f"{rms:.3f} mV")
                self.table_latest_eval_right = True

    # ------------------------------------------------------------------
    # Trigger lines
    # ------------------------------------------------------------------

    def _clear_trigger_lines(self):
        for line in self._vlines_left:
            self.plot_left.removeItem(line)
        for line in self._vlines_right:
            self.plot_right.removeItem(line)
        self._vlines_left.clear()
        self._vlines_right.clear()
        for item in self._shades_left:
            self.plot_left.removeItem(item)
        for item in self._shades_right:
            self.plot_right.removeItem(item)
        self._shades_left.clear()
        self._shades_right.clear()

    def _redraw_trigger_lines_chart(self, t_start, t_end):
        self._clear_trigger_lines()
        pen = pg.mkPen("#ff6f00", width=3, style=Qt.DashLine)
        for t in self.trigger_times:
            if t_start <= t <= t_end:
                l = pg.InfiniteLine(pos=t,   angle=90, pen=pen)
                r = pg.InfiniteLine(pos=t,   angle=90, pen=pen)
                self.plot_left.addItem(l)
                self.plot_right.addItem(r)
                self._vlines_left.append(l)
                self._vlines_right.append(r)

    def _redraw_trigger_lines_scope(self, t_scale=1.0):
        self._clear_trigger_lines()
        pen = pg.mkPen("#ff6f00", width=3, style=Qt.DashLine)
        # Trigger marker at 0
        l = pg.InfiniteLine(pos=0, angle=90, pen=pen)
        r = pg.InfiniteLine(pos=0, angle=90, pen=pen)
        self.plot_left.addItem(l)
        self.plot_right.addItem(r)
        self._vlines_left.append(l)
        self._vlines_right.append(r)

        # Analysis window shading (only when MEP mode is on)
        if self.radio_mep.isChecked():
            def _add_shades(plot, shade_list, spin_pre_s, spin_pre_e, spin_mep_s, spin_mep_e):
                ps = spin_pre_s.value() / 1000.0 * t_scale
                pe = spin_pre_e.value() / 1000.0 * t_scale
                ms = spin_mep_s.value() / 1000.0 * t_scale
                me = spin_mep_e.value() / 1000.0 * t_scale
                for vals in ((ps, pe), (ms, me)):
                    region = pg.LinearRegionItem(
                        values=vals,
                        brush=pg.mkBrush(255, 220, 0, 60),
                        pen=pg.mkPen(None),
                        movable=False,
                    )
                    plot.addItem(region)
                    shade_list.append(region)

            # Left always uses first set
            _add_shades(self.plot_left, self._shades_left,
                        self.spin_prestim_start, self.spin_prestim_end,
                        self.spin_mep_start, self.spin_mep_end)
            # Right uses second set if Separate is ON, otherwise first
            if self.separate:
                _add_shades(self.plot_right, self._shades_right,
                            self.spin_prestim_start_2, self.spin_prestim_end_2,
                            self.spin_mep_start_2, self.spin_mep_end_2)
            else:
                _add_shades(self.plot_right, self._shades_right,
                            self.spin_prestim_start, self.spin_prestim_end,
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
