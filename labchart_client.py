import win32com.client
import numpy as np


class LabChartClient:
    _FLAGS = 1  # flags=1 works for all channel types (raw + computed)

    def __init__(self):
        self.app = win32com.client.GetActiveObject("ADIChart.Application")
        self.doc = self.app.ActiveDocument
        self._secs_per_tick = None
        self._last_rec = None

    def _active_rec(self):
        rec = self.doc.SamplingRecord
        if rec is not None and rec > 0:
            self._last_rec = rec
            return rec
        if self._last_rec is not None:
            return self._last_rec
        return self.doc.NumberOfRecords

    @property
    def fs(self):
        return 1.0 / self.secs_per_tick

    @property
    def secs_per_tick(self):
        if self._secs_per_tick is None:
            rec = self._active_rec()
            self._secs_per_tick = self.doc.GetRecordSecsPerTick(rec)
        return self._secs_per_tick

    @property
    def n_channels(self):
        return self.doc.NumberOfChannels

    def get_channel_info(self):
        """Returns list of (name, unit) tuples, one per channel (0-indexed)."""
        rec = self._active_rec()
        return [
            (self.doc.GetChannelName(ch), self.doc.GetUnits(ch, rec))
            for ch in range(1, self.n_channels + 1)
        ]

    def is_sampling(self):
        return bool(self.doc.IsSampling)

    def current_time(self):
        rec = self._active_rec()
        rec_len = self.doc.GetRecordLength(rec)
        spt = self.doc.GetRecordSecsPerTick(rec)
        return rec_len * spt

    def add_comment(self, text):
        self.doc.AppendComment(text)

    def play_message(self, hex_str: str):
        """Send an FRO configuration to LabChart via PlayMessage."""
        app = win32com.client.GetActiveObject("ADIChart.Application")
        app.ActiveDocument.PlayMessage(hex_str)

    def start_sampling(self):
        self.doc.StartSampling(0, False, 0)

    def stop_sampling(self):
        self.doc.StopSampling()

    def get_scope_data(self, trigger_time, pre_secs=0.5, post_secs=2.0, channels=None):
        """Fetch data in a fixed window around a trigger timestamp."""
        rec = self._active_rec()
        spt = self.doc.GetRecordSecsPerTick(rec)
        t_start = max(0.0, trigger_time - pre_secs)
        t_end   = trigger_time + post_secs

        self.doc.SetSelectionTime(rec, t_start, rec, t_end)

        if channels is None:
            channels = list(range(self.n_channels))

        data = {}
        for idx in channels:
            ch  = idx + 1
            raw = self.doc.GetSelectedData(self._FLAGS, ch)
            if raw:
                arr = np.array(raw, dtype=float)
            else:
                n   = max(1, int((t_end - t_start) / spt))
                arr = np.full(n, np.nan)
            data[idx] = arr

        return data, t_start, t_end

    def get_latest_data(self, window_secs=2.0, channels=None):
        """
        Fetch the last `window_secs` of data for the requested channels.

        Parameters
        ----------
        window_secs : float
        channels : list[int] | None  — 0-based channel indices

        Returns
        -------
        data : dict[int, np.ndarray]  — NaN where LabChart returns None
        t_start, t_end : float        — absolute record time (seconds)
        """
        rec = self._active_rec()
        rec_len = self.doc.GetRecordLength(rec)
        spt = self.doc.GetRecordSecsPerTick(rec)
        t_end = rec_len * spt
        t_start = max(0.0, t_end - window_secs)

        self.doc.SetSelectionTime(rec, t_start, rec, t_end)

        if channels is None:
            channels = list(range(self.n_channels))

        data = {}
        for idx in channels:
            ch = idx + 1  # COM API is 1-based
            raw = self.doc.GetSelectedData(self._FLAGS, ch)
            if raw:
                arr = np.array(raw, dtype=float)  # None entries become NaN
            else:
                n = max(1, int((t_end - t_start) / spt))
                arr = np.full(n, np.nan)
            data[idx] = arr

        return data, t_start, t_end
