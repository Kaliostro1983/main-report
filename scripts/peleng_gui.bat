@echo off
setlocal enabledelayedexpansion

REM === 1) Перейти у корінь репозиторію (папка на рівень вище за scripts) ===
pushd "%~dp0\.."  || (echo [ERROR] Can't cd to repo root & pause & exit /b 1)

REM === 2) Активувати віртуальне середовище ===
if exist ".venv\Scripts\activate" (
    call .\.venv\Scripts\activate
) else (
    echo [WARN] .venv not found. Trying system Python...
)

REM === 3) Запустити головний скрипт у потрібному режимі ===
if exist "main.py" (
    python main.py --mode peleng-gui --config config.yml
) else (
    echo [ERROR] main.py not found in: %CD%
    echo        Expected: %~dp0\..\main.py
    popd
    pause
    exit /b 1
)

REM === 4) Повернутися назад і не закривати вікно ===
popd
pause
