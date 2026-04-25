import win32com.client
try:
    app = win32com.client.Dispatch('CorelDRAW.Application')
    doc = app.ActiveDocument
    if doc:
        page = doc.ActivePage
        print("Page methods:", [m for m in dir(page) if 'dup' in m.lower()])
        # Try duplicating
        new_page = doc.Pages.Item(1).Duplicate()
        print("Duplicated OK")
    else:
        print("No active document.")
except Exception as e:
    print("Error:", e)
