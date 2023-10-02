@echo off
set TARGET='%CD%\PTM_importData_GUI.pyc'
set SHORTCUT='%userprofile%\desktop\PTM_importData.lnk'
set PWS=powershell.exe -ExecutionPolicy Bypass -NoLogo -NonInteractive -NoProfile

%PWS% -Command "$ws = New-Object -ComObject WScript.Shell; $S = $ws.CreateShortcut(%SHORTCUT%); $S.TargetPath = %TARGET%; $S.WorkingDirectory = '%CD%'; $S.IconLocation = '%CD%\ptm_icon.ico'; $S.Save()"
echo Create "PTM_importData" shortcut to Desktop: Success

pause
