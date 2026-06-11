import win32com.client

app = win32com.client.Dispatch("ADIChart.Application")
doc = app.ActiveDocument

n_channels = doc.NumberOfChannels
print(f"Channels: {n_channels}")
print(f"Sampling: {doc.IsSampling}")
print(f"Records: {doc.NumberOfRecords}")
print(f"Current record: {doc.SamplingRecord}")

print("\n--- Channel names & units ---")
for ch in range(1, n_channels + 1):
    try:
        name = doc.GetChannelName(ch)
        units = doc.GetUnits(ch, doc.SamplingRecord)
        print(f"  Ch{ch}: {name!r}  [{units}]")
    except Exception as e:
        print(f"  Ch{ch}: ERROR({e})")

print("\n--- Sampling rate ---")
for ch in range(1, min(4, n_channels + 1)):
    try:
        secs_per_tick = doc.GetRecordSecsPerTick(ch, doc.SamplingRecord)
        fs = 1.0 / secs_per_tick
        print(f"  Ch{ch}: {fs:.1f} Hz")
    except Exception as e:
        print(f"  Ch{ch}: ERROR({e})")

print("\n--- GetChannelData test (Ch1, last 100 ticks) ---")
try:
    rec = doc.SamplingRecord
    data = doc.GetChannelData(1, rec, -100, 100)
    print(f"  Type: {type(data)}")
    print(f"  Data: {data}")
except Exception as e:
    print(f"  ERROR: {e}")
