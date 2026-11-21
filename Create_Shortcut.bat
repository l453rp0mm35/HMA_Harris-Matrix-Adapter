@echo off
rem Creates a shortcut "HMA" with icon on the desktop and in the Start menu
rem (an old shortcut "HMCA" is removed in the process).
set "ROOT=%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -Command "$r=$env:ROOT.TrimEnd('\'); foreach($d in @([Environment]::GetFolderPath('Desktop'),[Environment]::GetFolderPath('Programs'))){ Remove-Item ($d+'\HMCA.lnk') -ErrorAction SilentlyContinue; $s=(New-Object -ComObject WScript.Shell).CreateShortcut($d+'\HMA.lnk'); $s.TargetPath=$r+'\HMA.bat'; $s.WorkingDirectory=$r; $s.IconLocation=$r+'\icon\HMA_icon.ico'; $s.WindowStyle=7; $s.Description='HMA - Harris Matrix Adapter'; $s.Save() }"
if errorlevel 1 goto failed
rem Refresh the Windows icon cache so that the new icon shows up right away.
ie4uinit.exe -show >nul 2>&1
echo.
echo The shortcut "HMA" was created on the desktop and in the Start menu.
echo In the Start menu: Start, "All apps" (or simply type "HMA").
echo If the old icon is still shown: press F5 on the desktop, or restart Windows Explorer (Task Manager).
echo.
pause
exit /b
:failed
echo.
echo The shortcut could not be created.
echo Alternative: right-click HMA.bat, "Send to", "Desktop (create shortcut)".
echo For the Start menu copy the shortcut into the folder that Windows+R opens with shell:programs.
echo In the shortcut properties choose the file icon\HMA_icon.ico under "Change icon".
echo.
pause
