import win32com.client
import pythoncom

# Generate type library wrapper for proper binding
print("Generating type library...")
try:
    app = win32com.client.gencache.EnsureDispatch("ADIChart.Application")
    doc = app.ActiveDocument
    print("Early binding OK")
except Exception as e:
    print(f"gencache failed: {e}")
    app = win32com.client.Dispatch("ADIChart.Application")
    doc = app.ActiveDocument

# Inspect type info for GetChannelData
print("\n--- Type info for document methods ---")
try:
    typeinfo = doc._oleobj_.GetTypeInfo()
    typeattr = typeinfo.GetTypeAttr()
    for i in range(typeattr.cFuncs):
        fd = typeinfo.GetFuncDesc(i)
        name = typeinfo.GetNames(fd.memid)[0]
        if "Channel" in name or "Data" in name or "Scope" in name:
            params = typeinfo.GetNames(fd.memid)
            print(f"  {name}: params={params}, nParams={fd.cParams}")
except Exception as e:
    print(f"Type info error: {e}")

# Try GetChannelData via Invoke directly
print("\n--- Direct Invoke test ---")
try:
    import win32com.client as wc
    app2 = wc.Dispatch("ADIChart.Application")
    doc2 = app2.ActiveDocument

    # Try MatLabPutChannelData to see format
    result = doc2.MatLabPutChannelData(1)
    print(f"MatLabPutChannelData(1): {type(result)} = {result}")
except Exception as e:
    print(f"MatLabPutChannelData error: {e}")
