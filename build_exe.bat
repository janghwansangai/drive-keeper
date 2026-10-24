@echo off
setlocal
rem Builds dist\DriveKeeper\DriveKeeper.exe (no Python needed to run it).
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
  echo Run run.bat once first to create .venv.
  pause
  exit /b 1
)
".venv\Scripts\python.exe" -m pip install --disable-pip-version-check -r requirements.txt -r requirements-build.txt || goto :fail
".venv\Scripts\python.exe" build_exe.py || goto :fail
echo.
echo Done: dist\DriveKeeper\DriveKeeper.exe
pause
exit /b 0

:fail
echo.
echo Build failed.
pause
exit /b 1
