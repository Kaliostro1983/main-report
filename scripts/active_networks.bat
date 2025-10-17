@echo off
setlocal ENABLEDELAYEDEXPANSION

REM === UTF-8 консоль (щоб кирилиця в логах була коректна)
chcp 65001 >nul

REM === Перехід у корінь проєкту (бат лежить у папці scripts)
set "SCRIPT_DIR=%~dp0"
pushd "%SCRIPT_DIR%\.."

REM === Python з venv, якщо є; інакше системний
set "VENV_PY=%CD%\.venv\Scripts\python.exe"
if exist "%VENV_PY%" (
  set "PY=%VENV_PY%"
) else (
  set "PY=python"
)

REM === Конфіг — аргумент №1 або config.yml за замовчуванням
set "CFG=%~1"
if "%CFG%"=="" set "CFG=config.yml"

REM === Рівень логів — аргумент №2 або INFO
set "LOGLEVEL=%~2"
if "%LOGLEVEL%"=="" set "LOGLEVEL=INFO"

echo.
echo === Запуск звіту "Активні мережі" ===
echo Конфіг: %CFG%
echo Рівень логів: %LOGLEVEL%
echo.

"%PY%" main.py --config "%CFG%" --mode active-freqs --log-level %LOGLEVEL%
set "RC=%ERRORLEVEL%"

echo.
if "%RC%"=="0" (
  echo [OK] Готово. Перевіряй теку build/.
  REM Відкрити теку з результатом (за бажанням розкоментуй):
  REM explorer ".\build"
) else (
  echo [ERR] Процес завершився з кодом %RC%.
)

popd
endlocal
