import win32com.client
import traceback
print('Trying generic...')
try:
    app=win32com.client.Dispatch('CorelDRAW.Application')
    print('Generic OK', app.Version)
except Exception as e:
    print('Generic error:')
    traceback.print_exc()

print('Trying specific 22...')
try:
    app=win32com.client.Dispatch('CorelDRAW.Application.22')
    print('Specific 22 OK', app.Version)
except Exception as e:
    print('Specific error:')
    traceback.print_exc()
