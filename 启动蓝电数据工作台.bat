@echo off
setlocal
set "APP_DIR=%~dp0"
set "PYTHONW_EXE="

for /d %%D in ("%LocalAppData%\Programs\Python\Python3*") do (
  if exist "%%~fD\pythonw.exe" set "PYTHONW_EXE=%%~fD\pythonw.exe"
)

if not defined PYTHONW_EXE (
  for %%P in (pythonw.exe) do set "PYTHONW_EXE=%%~$PATH:P"
)

if not defined PYTHONW_EXE (
  echo No usable Python installation was found.
  echo Please install Python 3.11 or newer, then run:
  echo   pip install -r "%APP_DIR%requirements.txt"
  pause
  exit /b 1
)

start "LAND Workbench" "%PYTHONW_EXE%" "%APP_DIR%landt_workbench.py" %*
endlocal
