import win32com.client
import pythoncom

app = win32com.client.gencache.EnsureDispatch("ADIChart.Application")
doc = app.ActiveDocument

# Get type info
try:
    typeinfo = doc._oleobj_.GetTypeInfo()
    typeattr = typeinfo.GetTypeAttr()
    print(f"Total functions: {typeattr.cFuncs}")
    for i in range(typeattr.cFuncs):
        fd = typeinfo.GetFuncDesc(i)
        names = typeinfo.GetNames(fd.memid)
        name = names[0]
        param_names = names[1:]
        print(f"  [{i}] {name}({', '.join(param_names)})")
except Exception as e:
    print(f"Error: {e}")
    import traceback
    traceback.print_exc()
