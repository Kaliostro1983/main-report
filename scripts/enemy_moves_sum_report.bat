@echo off
setlocal enabledelayedexpansion

REM === перейти з scripts\ у корінь main-report ===
cd /d "%~dp0\.."

echo.
echo ===== RUN ENEMY MOVES REPORT =====
echo Workdir: %cd%
echo.

REM === перевірки ===
if not exist "main.py" (
  echo [ERROR] main.py not found
  pause
  exit /b 1
)

if not exist "config.yml" (
  echo [ERROR] config.yml not found
  pause
  exit /b 1
)

if not exist ".venv\Scripts\activate.bat" (
  echo [ERROR] .venv not found
  pause
  exit /b 1
)

REM === timestamp стабільний ===
for /f %%i in ('powershell -NoProfile -Command "Get-Date -Format yyyyMMdd_HHmmss"') do set TS=%%i

if not exist "logs" mkdir "logs"
set LOG=logs\enemy_moves_%TS%.log

echo [INFO] Start %TS% > "%LOG%"
echo [INFO] Workdir: %cd% >> "%LOG%"

REM === активуємо venv ===
call ".venv\Scripts\activate.bat" >> "%LOG%" 2>&1

REM === запуск ===
python -X faulthandler main.py --config config.yml --mode enemy-moves-sum >> "%LOG%" 2>&1
set EC=%errorlevel%

echo.>> "%LOG%"
echo [INFO] ExitCode=%EC% >> "%LOG%"

echo.
echo ===== LOG =====
type "%LOG%"
echo.
echo [DONE] ExitCode=%EC%
pause
exit /b %EC%