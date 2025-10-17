@echo off
setlocal EnableExtensions
chcp 65001 >nul
title Звіт по ворожих радіомережах (63 омсбр)

REM Перейти з scripts\ у корінь репозиторію
pushd "%~dp0\.."  || goto :fail

REM Активувати venv, якщо є
if exist ".venv\Scripts\activate.bat" call ".venv\Scripts\activate.bat"

echo.
echo === Генерація звіту по ворожих радіомережах ===
python ".\main.py" --mode enemies --log-level INFO
set "RC=%ERRORLEVEL%"
echo.
echo ================================================
if not "%RC%"=="0" echo [Помилка] Код завершення: %RC%

echo.
echo Натисни будь-яку клавішу, щоб закрити вікно...
pause >nul
popd
exit /b %RC%

:fail
echo Не вдалося перейти у корінь репозиторію.
echo Поточний файл: %~f0
pause
exit /b 1
