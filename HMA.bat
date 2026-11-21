@echo off
rem HMA launcher (Windows): starts the newest HMA_v*.py (draft files are ignored) from the folder "script" (or "skript").
setlocal
set "HMA_DIR=%~dp0script"
if not exist "%HMA_DIR%" set "HMA_DIR=%~dp0skript"
set "HMA_SCRIPT="
for /f "delims=" %%F in ('dir /b /o:n "%HMA_DIR%\HMA_v*.py" 2^>nul ^| findstr /v /i "draft"') do set "HMA_SCRIPT=%%F"
if not defined HMA_SCRIPT (
  echo.
  echo HMA script not found in "%HMA_DIR%".
  echo.
  pause
  exit /b 1
)
cd /d "%HMA_DIR%"
set "PY="
for %%P in (pythonw.exe) do set "PY=%%~$PATH:P"
if defined PY (
  start "" "%PY%" "%HMA_DIR%\%HMA_SCRIPT%"
  exit /b
)
for %%P in (pyw.exe) do set "PY=%%~$PATH:P"
if defined PY (
  start "" "%PY%" "%HMA_DIR%\%HMA_SCRIPT%"
  exit /b
)
for %%P in (python.exe) do set "PY=%%~$PATH:P"
if defined PY (
  "%PY%" "%HMA_DIR%\%HMA_SCRIPT%"
  if errorlevel 1 pause
  exit /b
)
echo.
echo Python was not found. HMA needs Python 3.10 or newer (with tkinter).
echo Installation: https://www.python.org/downloads/  (tick "Add python.exe to PATH")
echo.
pause
