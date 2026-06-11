import win32com.client
import numpy as np

app = win32com.client.gencache.EnsureDispatch("ADIChart.Application")
doc = app.ActiveDocument
rec = doc.SamplingRecord
rec_len = doc.GetRecordLength(rec)
print(f"rec={rec}, rec_len={rec_len}")

# Try GetChannelData with flags=1 near end of record
print("\n--- GetChannelData(flags=1, ...) ---")
start = rec_len - 2000
for ch in [1, 3]:
    for n in [100, 1000, 2000]:
        try:
            result = doc.GetChannelData(1, ch, rec, start, n)
            if result is not None:
                arr = np.array(result)
                print(f"  Ch{ch}, n={n}: shape={arr.shape}, first={arr[0]:.4f}")
            else:
                print(f"  Ch{ch}, n={n}: None")
        except Exception as e:
            print(f"  Ch{ch}, n={n}: ERROR: {e}")

# Confirm GetSelectedData approach is best
print("\n--- Timing: GetSelectedData approach ---")
import time
secs_per_tick = doc.GetRecordSecsPerTick(rec)
t_end = rec_len * secs_per_tick
t_start = t_end - 2.0

t0 = time.perf_counter()
doc.SetSelectionTime(rec, t_start, rec, t_end)
for ch in [1, 2, 3, 4, 5]:
    result = doc.GetSelectedData(1, ch)
t1 = time.perf_counter()
print(f"  5 channels, 2 sec window: {(t1-t0)*1000:.1f} ms")
print(f"  Result size: {len(result)} samples (expected ~{int(2.0/secs_per_tick)})")
