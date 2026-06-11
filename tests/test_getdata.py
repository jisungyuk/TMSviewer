import win32com.client
import numpy as np

app = win32com.client.gencache.EnsureDispatch("ADIChart.Application")
doc = app.ActiveDocument
rec = doc.SamplingRecord

secs_per_tick = doc.GetRecordSecsPerTick(rec)
fs = 1.0 / secs_per_tick
print(f"Record: {rec}, Sampling rate: {fs:.0f} Hz")

rec_len = doc.GetRecordLength(rec)
print(f"Record length (ticks): {rec_len}")

# GetChannelData(flags, channelNumber, blockNumber, startSample, numSamples)
print("\n--- GetChannelData tests ---")
for flags in [0, 1, 2]:
    for start in [0, max(0, rec_len - 1000)]:
        try:
            result = doc.GetChannelData(flags, 3, rec, start, 100)
            print(f"  flags={flags}, start={start}: type={type(result)}, result={str(result)[:100]}")
        except Exception as e:
            print(f"  flags={flags}, start={start}: ERROR: {e}")

# GetScopeChannelData(flags, channelNumber, pageNumber, subPageNumber, secsBeforeZero, secsAfterZero)
print("\n--- GetScopeChannelData tests ---")
for args in [
    (0, 3, 0, 0, 0.1, 0.1),
    (0, 3, 1, 0, 0.1, 0.1),
    (0, 3, 0, 0, 0.05, 0.05),
]:
    try:
        result = doc.GetScopeChannelData(*args)
        print(f"  args={args}: type={type(result)}, result={str(result)[:100]}")
    except Exception as e:
        print(f"  args={args}: ERROR: {e}")
