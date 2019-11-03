import winreg
import sys

while True:
    seconds = input('Heyho! Input delay for services (seconds) or "q" to exit:')
    if seconds == 'q':
        sys.exit()
    elif seconds.isdecimal:
        seconds = int(seconds)
        if seconds >= 30 and seconds <= 500:
            break
        else:
            print('value must be between 30 and 500')
    else:
        print('enter a numerical value ')

    
seconds *= 1000        
cntrl_key = winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, 'SYSTEM\\CurrentControlSet\\Control')
winreg.SetValueEx(cntrl_key, 'ServicesPipeTimeout', 0, winreg.REG_DWORD, seconds)