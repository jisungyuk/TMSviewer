import win32com.client

app = win32com.client.gencache.EnsureDispatch("ADIChart.Application")
doc = app.ActiveDocument

# ── Application object methods ──────────────────────────────────────
print("=== Application methods ===")
try:
    typeinfo = app._oleobj_.GetTypeInfo()
    typeattr = typeinfo.GetTypeAttr()
    for i in range(typeattr.cFuncs):
        fd    = typeinfo.GetFuncDesc(i)
        names = typeinfo.GetNames(fd.memid)
        if any(k in names[0].lower() for k in ("comment", "annot", "event", "marker", "trigger")):
            print(f"  [{i}] {names[0]}({', '.join(names[1:])})")
except Exception as e:
    print(f"  Error: {e}")

# ── Try ADIChart.ChartData ──────────────────────────────────────────
print("\n=== ADIChart.ChartData methods ===")
try:
    cd = win32com.client.Dispatch("ADIChart.ChartData.1")
    typeinfo = cd._oleobj_.GetTypeInfo()
    typeattr = typeinfo.GetTypeAttr()
    print(f"Total functions: {typeattr.cFuncs}")
    for i in range(typeattr.cFuncs):
        fd    = typeinfo.GetFuncDesc(i)
        names = typeinfo.GetNames(fd.memid)
        print(f"  [{i}] {names[0]}({', '.join(names[1:])})")
except Exception as e:
    print(f"  Error: {e}")

# ── Events channel (Ch13) - try reading raw data ────────────────────
print("\n=== Events channel (Ch13) raw data ===")
rec = doc.SamplingRecord
rec_len = doc.GetRecordLength(rec)
spt = doc.GetRecordSecsPerTick(rec)
t_end = rec_len * spt

doc.SetSelectionTime(rec, 0, rec, t_end)
for flags in [0, 1, 2, 3]:
    try:
        result = doc.GetSelectedData(flags, 13)
        if result:
            non_nan = [x for x in result if x is not None and str(x) != 'nan']
            print(f"  flags={flags}: {len(result)} samples, non-null={len(non_nan)}, sample={non_nan[:5] if non_nan else []}")
    except Exception as e:
        print(f"  flags={flags}: Error: {e}")
