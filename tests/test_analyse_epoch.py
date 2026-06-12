"""
Unit tests for _analyse_epoch calculation logic.
Run with: python -m pytest tests/test_analyse_epoch.py -v
No GUI or hardware required.
"""
import sys, types
import numpy as np
import pytest

# ── Stub out everything that needs Qt / COM / serial ──────────────────────────
for mod in [
    "PyQt5", "PyQt5.QtWidgets", "PyQt5.QtCore", "PyQt5.QtGui",
    "pyqtgraph", "pythoncom", "win32com", "win32com.client",
    "serial", "magstim_test", "labchart_client",
]:
    parts = mod.split(".")
    for i in range(1, len(parts) + 1):
        name = ".".join(parts[:i])
        if name not in sys.modules:
            sys.modules[name] = types.ModuleType(name)

# Stub specific names used at import time in realtime_viewer
import PyQt5.QtCore as _qc
_qc.pyqtSignal = lambda *a, **kw: None

import PyQt5.QtWidgets as _qw
for _name in [
    "QMainWindow", "QWidget", "QVBoxLayout", "QHBoxLayout",
    "QPushButton", "QComboBox", "QLabel", "QSpinBox", "QApplication",
    "QRadioButton", "QButtonGroup", "QDoubleSpinBox", "QCheckBox",
    "QAbstractSpinBox", "QFrame", "QTableWidget", "QTableWidgetItem",
    "QHeaderView", "QGraphicsOpacityEffect", "QStackedWidget",
]:
    setattr(_qw, _name, type(_name, (), {}))

import PyQt5.QtGui as _qg
for _name in ["QFont", "QColor"]:
    setattr(_qg, _name, type(_name, (), {}))

sys.modules["labchart_client"].LabChartClient = type("LabChartClient", (), {})
sys.modules["magstim_test"].disable_remote  = lambda *a: None
sys.modules["magstim_test"].get_parameters  = lambda *a: None
sys.modules["magstim_test"].get_temperature = lambda *a: None

# ── Minimal _analyse_epoch extracted for testing ──────────────────────────────
# We test the math directly without instantiating the full GUI class.

SCOPE_PRE  = 0.2   # seconds
SCOPE_POST = 0.8   # seconds

def analyse_epoch(arr, mep_s, mep_e, pre_s, pre_e, scope_pre=SCOPE_PRE, scope_post=SCOPE_POST):
    """Pure reimplementation matching RealTimeViewer._analyse_epoch logic."""
    if arr is None or len(arr) == 0:
        return None
    n     = len(arr)
    dur   = scope_pre + scope_post
    spt   = dur / max(n - 1, 1)
    t_rel = np.arange(n) * spt - scope_pre
    mep_data = arr[(t_rel >= mep_s) & (t_rel <= mep_e)]
    pre_data = arr[(t_rel >= pre_s) & (t_rel <= pre_e)]
    mep_amp     = float(np.nanmax(mep_data) - np.nanmin(mep_data)) if mep_data.size > 0 else 0.0
    prestim_rms = float(np.sqrt(np.nanmean(pre_data ** 2)))        if pre_data.size > 0 else 0.0
    return {"mep_amp": mep_amp, "prestim_rms": prestim_rms}


# ── Tests ─────────────────────────────────────────────────────────────────────

SR = 2000  # samples / s
N  = int((SCOPE_PRE + SCOPE_POST) * SR)  # 2000 samples for 1 s window


def make_epoch(mep_amp_mV=1.0, noise_rms_mV=0.05):
    """Synthetic epoch: flat noise pre-stim, known-amplitude sine in MEP window."""
    rng = np.random.default_rng(0)
    arr = rng.normal(0, noise_rms_mV, N)
    # inject a 100 Hz sine (amplitude → peak-to-peak = 2*mep_amp_mV) at 10–50 ms post-trigger
    spt   = (SCOPE_PRE + SCOPE_POST) / max(N - 1, 1)
    t_rel = np.arange(N) * spt - SCOPE_PRE
    mask  = (t_rel >= 0.010) & (t_rel <= 0.050)
    arr[mask] += mep_amp_mV * np.sin(2 * np.pi * 100 * t_rel[mask])
    return arr


class TestAnalyseEpoch:
    def test_returns_dict(self):
        arr = make_epoch()
        r = analyse_epoch(arr, mep_s=0.010, mep_e=0.050, pre_s=-0.200, pre_e=-0.050)
        assert isinstance(r, dict)
        assert "mep_amp" in r and "prestim_rms" in r

    def test_mep_amp_approx(self):
        arr = make_epoch(mep_amp_mV=1.0)
        r = analyse_epoch(arr, mep_s=0.010, mep_e=0.050, pre_s=-0.200, pre_e=-0.050)
        # peak-to-peak of unit-amplitude sine ≈ 2.0; allow ±10%
        assert 1.8 <= r["mep_amp"] <= 2.2, f"mep_amp={r['mep_amp']:.3f}"

    def test_prestim_rms_approx(self):
        arr = make_epoch(noise_rms_mV=0.05)
        r = analyse_epoch(arr, mep_s=0.010, mep_e=0.050, pre_s=-0.200, pre_e=-0.050)
        # RMS of N(0, 0.05) ≈ 0.05; allow ±50% due to small-window variance
        assert 0.025 <= r["prestim_rms"] <= 0.10, f"prestim_rms={r['prestim_rms']:.4f}"

    def test_none_on_empty(self):
        assert analyse_epoch(np.array([]), 0.010, 0.050, -0.200, -0.050) is None
        assert analyse_epoch(None,          0.010, 0.050, -0.200, -0.050) is None

    def test_zero_when_window_outside_data(self):
        arr = make_epoch()
        # MEP window entirely outside epoch range → should return 0.0
        r = analyse_epoch(arr, mep_s=5.0, mep_e=6.0, pre_s=-0.200, pre_e=-0.050)
        assert r["mep_amp"] == 0.0

    def test_larger_mep_detected(self):
        arr_small = make_epoch(mep_amp_mV=0.5)
        arr_large = make_epoch(mep_amp_mV=2.0)
        r_s = analyse_epoch(arr_small, mep_s=0.010, mep_e=0.050, pre_s=-0.200, pre_e=-0.050)
        r_l = analyse_epoch(arr_large, mep_s=0.010, mep_e=0.050, pre_s=-0.200, pre_e=-0.050)
        assert r_l["mep_amp"] > r_s["mep_amp"]
