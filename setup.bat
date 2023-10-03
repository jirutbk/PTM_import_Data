@echo off
pip install PyPDF2
pip install openpyxl
pip install pywin32

FOR /f %%p in ('where pythonw') do SET PYTHONPATH=%%p
ECHO %PYTHONPATH%
REM Assoc .pywc=Python.NoConFile.CompiledFile
REM Ftype Python.NoConFile.CompiledFile=%PYTHONPATH% %1

set TARGET='%CD%\PTM_importData_GUI.pyzw'
set SHORTCUT='%userprofile%\desktop\PTM_importData.lnk'
set PWS=powershell.exe -ExecutionPolicy Bypass -NoLogo -NonInteractive -NoProfile

%PWS% -Command "$ws = New-Object -ComObject WScript.Shell; $S = $ws.CreateShortcut(%SHORTCUT%); $S.TargetPath = %TARGET%; $S.WorkingDirectory = '%CD%'; $S.IconLocation = '%CD%\ptm_icon.ico'; $S.Save()"
echo Create "PTM_importData" shortcut to Desktop: Success

pause
