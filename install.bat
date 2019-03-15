@ echo off
copy /y keys.key config.ini
echo.>>config.ini
echo %cd%>>config.ini
copy /y config.ini c:\windows\syswow64\
monyze.exe install
monyze.exe start