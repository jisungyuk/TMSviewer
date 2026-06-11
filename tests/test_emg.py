import win32com.client

app = win32com.client.gencache.EnsureDispatch("ADIChart.Application")
doc = app.ActiveDocument
rec = doc.SamplingRecord
rec_len = doc.GetRecordLength(rec)

# Check sampling rates per channel
print("--- Sampling rates per channel ---")
secs_per_tick = doc.GetRecordSecsPerTick(rec)
print(f"  Default: {1/secs_per_tick:.0f} Hz (secsPerTick={secs_per_tick})")

# Select last 0.1 seconds
doc.SetSelectionTime(rec, rec_len * secs_per_tick - 0.1, rec, rec_len * secs_per_tick)

# Try all channels with all flags
print("\n--- All channels, flags 0-3 ---")
for ch in range(1, 14):
    results = {}
    for flags in [0, 1, 2, 3]:
        try:
            result = doc.GetSelectedData(flags, ch)
            if result and result[0] is not None:
                results[flags] = f"OK({len(result)} samples, first={result[0]:.4f})"
            elif result:
                results[flags] = f"None*{len(result)}"
            else:
                results[flags] = "empty"
        except Exception as e:
            results[flags] = f"ERR:{e}"
    ch_name = doc.GetChannelName(ch)
    print(f"  Ch{ch} {ch_name}: {results}")
