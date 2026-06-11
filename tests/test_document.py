import win32com.client

app = win32com.client.Dispatch("ADIChart.Application")
doc = app.ActiveDocument

print("--- Document properties ---")
for attr in dir(doc):
    if not attr.startswith("_"):
        try:
            val = getattr(doc, attr)
            if not callable(val):
                print(f"  {attr}: {val}")
        except Exception as e:
            print(f"  {attr}: ERROR({e})")

print("\n--- Document methods ---")
for attr in dir(doc):
    if not attr.startswith("_"):
        try:
            val = getattr(doc, attr)
            if callable(val):
                print(f"  {attr}()")
        except Exception as e:
            pass
