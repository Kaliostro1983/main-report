@echo off
REM scripts\run_automizer.bat — launch Automizer UI via main.py
setlocal

REM Go to repo root (this .bat expected to live in scripts/)
pushd "%~dp0\.."

REM Activate venv if present
if exist ".venv\Scripts\activate.bat" call ".venv\Scripts\activate.bat"

REM Run Automizer via main.py mode
python main.py --mode automizer --config config.yml
set ERR=%ERRORLEVEL%

if not "%ERR%"=="0" (
  echo.
  echo [ERROR] Automizer exited with code %ERR%
  echo Press any key to close...
  pause >nul
)

popd
endlocal
