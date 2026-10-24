@echo off
setlocal
rem Double-click to start DriveKeeper.
rem First run (or after requirements.txt changes): creates .venv and installs packages.
cd /d "%~dp0"

set "VPY=.venv\Scripts\python.exe"

rem --- 1. Make sure .venv exists and works (recreate if Python was upgraded/removed) ---
if exist "%VPY%" (
  "%VPY%" -c "import tkinter" >nul 2>nul || (
    echo Existing .venv is broken. Recreating...
    rmdir /s /q .venv
  )
)
if not exist "%VPY%" call :create_venv || goto :fail

rem --- 2. Install packages if not installed yet or requirements.txt changed ---
fc /b requirements.txt .venv\installed-requirements.txt >nul 2>nul
if errorlevel 1 goto :install
"%VPY%" -c "import googleapiclient, google_auth_oauthlib, google_auth_httplib2, watchdog" >nul 2>nul || goto :install
goto :run

:install
echo Installing packages (1-2 minutes, first time only)...
"%VPY%" -m pip install --disable-pip-version-check -r requirements.txt || goto :fail
copy /y requirements.txt .venv\installed-requirements.txt >nul

:run
start "" ".venv\Scripts\pythonw.exe" "%~dp0app.py"
exit /b 0

:create_venv
set "PY="
where py >nul 2>nul && set "PY=py -3"
if not defined PY if exist "%LOCALAPPDATA%\Programs\Python\Launcher\py.exe" set "PY="%LOCALAPPDATA%\Programs\Python\Launcher\py.exe" -3"
if not defined PY (
  python -c "import sys, tkinter; sys.exit(0 if sys.version_info >= (3, 9) else 1)" >nul 2>nul && set "PY=python"
)
if not defined PY (
  echo.
  echo Python was not found.
  echo Install it from https://www.python.org/downloads/ and check "Add python.exe to PATH".
  exit /b 1
)
echo First run: creating .venv ...
%PY% -m venv .venv || exit /b 1
exit /b 0

:fail
echo.
echo Setup failed. Check your internet connection and that Python is installed
echo (python.org, "Add python.exe to PATH"), then run this file again.
pause
exit /b 1
