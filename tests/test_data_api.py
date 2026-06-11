import win32com.client

app = win32com.client.Dispatch("ADIChart.Application")
doc = app.ActiveDocument
rec = doc.SamplingRecord

# Try GetRecordSecsPerTick with 1 arg
print("--- GetRecordSecsPerTick ---")
for args in [(1,), (1, rec)]:
    try:
        val = doc.GetRecordSecsPerTick(*args)
        print(f"  args={args} -> {val}")
    except Exception as e:
        print(f"  args={args} -> ERROR: {e}")

# Try GetChannelData with various arg combos
print("\n--- GetChannelData ---")
for args in [
    (1,),
    (1, rec),
    (1, rec, 0),
    (1, 0),
    (1, rec, -100),
    (1, rec, 0, -1),
    (1, rec, 0, 100),
]:
    try:
        val = doc.GetChannelData(*args)
        print(f"  args={args} -> type={type(val)}, val={val}")
    except Exception as e:
        print(f"  args={args} -> ERROR: {e}")

# Try GetScopeChannelData
print("\n--- GetScopeChannelData ---")
for args in [(1,), (1, 100), (1, 0, 100)]:
    try:
        val = doc.GetScopeChannelData(*args)
        print(f"  args={args} -> type={type(val)}, val={val}")
    except Exception as e:
        print(f"  args={args} -> ERROR: {e}")

# Try GetSelectedData
print("\n--- GetSelectedData ---")
for args in [(1,), (1, rec), (1, rec, 0, 100)]:
    try:
        val = doc.GetSelectedData(*args)
        print(f"  args={args} -> type={type(val)}, val={val}")
    except Exception as e:
        print(f"  args={args} -> ERROR: {e}")
