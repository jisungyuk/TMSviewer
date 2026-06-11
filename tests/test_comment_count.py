import win32com.client

app = win32com.client.Dispatch("ADIChart.Application")
doc = app.ActiveDocument
rec = doc.NumberOfRecords

print(f"Records: {rec}")

# doc에서 Comment 관련 메서드/속성 전부 출력
print("\n--- Comment-related on doc ---")
for attr in dir(doc):
    if "comment" in attr.lower() or "annot" in attr.lower():
        try:
            val = getattr(doc, attr)
            print(f"  {attr}: {'(method)' if callable(val) else val}")
        except Exception as e:
            print(f"  {attr}: ERROR({e})")

# 흔한 이름으로 직접 시도
print("\n--- Direct attempts ---")
candidates = [
    ("GetNumberOfComments", (rec,)),
    ("GetCommentText",      (rec, 1)),
    ("GetCommentTick",      (rec, 1)),
    ("NumberOfComments",    ()),
]
for name, args in candidates:
    try:
        fn = getattr(doc, name)
        result = fn(*args) if args else fn
        print(f"  doc.{name}{args} -> {result}")
    except AttributeError:
        print(f"  doc.{name} -> 없음")
    except Exception as e:
        print(f"  doc.{name}{args} -> ERROR: {e}")
