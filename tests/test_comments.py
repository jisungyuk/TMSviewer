import win32com.client

app = win32com.client.gencache.EnsureDispatch("ADIChart.Application")
doc = app.ActiveDocument

# ── Services object methods ─────────────────────────────────────────
print("=== Services methods ===")
svc = doc.Services
try:
    typeinfo = svc._oleobj_.GetTypeInfo()
    typeattr = typeinfo.GetTypeAttr()
    print(f"Total functions: {typeattr.cFuncs}")
    for i in range(typeattr.cFuncs):
        fd    = typeinfo.GetFuncDesc(i)
        names = typeinfo.GetNames(fd.memid)
        print(f"  [{i}] {names[0]}({', '.join(names[1:])})")
except Exception as e:
    print(f"  Error: {e}")

# ── SelectionObject methods ─────────────────────────────────────────
print("\n=== SelectionObject methods ===")
sel = doc.SelectionObject
try:
    typeinfo = sel._oleobj_.GetTypeInfo()
    typeattr = typeinfo.GetTypeAttr()
    print(f"Total functions: {typeattr.cFuncs}")
    for i in range(typeattr.cFuncs):
        fd    = typeinfo.GetFuncDesc(i)
        names = typeinfo.GetNames(fd.memid)
        print(f"  [{i}] {names[0]}({', '.join(names[1:])})")
except Exception as e:
    print(f"  Error: {e}")
