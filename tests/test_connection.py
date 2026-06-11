import win32com.client

print("Connecting to LabChart...")

try:
    app = win32com.client.Dispatch("ADIChart.Application")
    print(f"Connected: {app}")
except Exception as e:
    print(f"ADIChart.Application failed: {e}")
    app = None

if app is None:
    try:
        app = win32com.client.GetActiveObject("ADIChart.Application")
        print(f"GetActiveObject connected: {app}")
    except Exception as e:
        print(f"GetActiveObject failed: {e}")

if app:
    print("\n--- App properties ---")
    for attr in dir(app):
        if not attr.startswith("_"):
            try:
                val = getattr(app, attr)
                print(f"  {attr}: {val}")
            except Exception as e:
                print(f"  {attr}: ERROR({e})")
