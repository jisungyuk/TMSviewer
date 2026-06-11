import win32com.client
import pythoncom

app = win32com.client.gencache.EnsureDispatch("ADIChart.Application")
doc = app.ActiveDocument
rec = doc.SamplingRecord
rec_len = doc.GetRecordLength(rec)
secs_per_tick = doc.GetRecordSecsPerTick(rec)
rec_secs = rec_len * secs_per_tick

print(f"rec={rec}, len={rec_len}, secs={rec_secs:.1f}")

# SetSelectionRange + GetSelectedData
print("\n--- SetSelectionRange + GetSelectedData ---")
start = max(0, rec_len - 2000)
try:
    doc.SetSelectionRange(rec, start, rec, rec_len - 1)
    for ch in [1, 3, 5]:
        result = doc.GetSelectedData(0, ch)
        print(f"  Ch{ch}: type={type(result)}, val={str(result)[:150]}")
except Exception as e:
    print(f"ERROR: {e}")
    import traceback; traceback.print_exc()

# SetSelectionTime
print("\n--- SetSelectionTime + GetSelectedData ---")
try:
    doc.SetSelectionTime(rec, rec_secs - 0.1, rec, rec_secs)
    for ch in [1, 3]:
        result = doc.GetSelectedData(0, ch)
        print(f"  Ch{ch}: type={type(result)}, val={str(result)[:150]}")
except Exception as e:
    print(f"ERROR: {e}")
    import traceback; traceback.print_exc()
