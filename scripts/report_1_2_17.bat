@echo off
setlocal ENABLEDELAYEDEXPANSION

rem === 1) Перейти в корінь проєкту (папка вище від /scripts) ===
pushd "%~dp0.."

rem === 2) Переконатися, що є Python ===
where python >nul 2>&1
if errorlevel 1 (
  echo [ERROR] Python не знайдено у PATH.
  echo Встановіть Python 3.10+ і повторіть.
  goto :end
)

rem === 3) Створити/активувати віртуальне середовище ===
if not exist ".venv\" (
  echo [INFO] Створюю віртуальне середовище .venv ...
  python -m venv .venv
  if errorlevel 1 (
    echo [ERROR] Не вдалося створити .venv
    goto :end
  )
)

call ".venv\Scripts\activate.bat"
if errorlevel 1 (
  echo [ERROR] Не вдалося активувати .venv
  goto :end
)

rem === 4) Поставити залежності, якщо є requirements.txt ===
if exist "requirements.txt" (
  echo [INFO] Перевірка/встановлення залежностей...
  pip install --upgrade pip >nul
  pip install -r requirements.txt
)

rem === 5) Запуск генерації звіту ===
set CFG="config.yml"
if not exist %CFG% (
  echo [ERROR] Не знайдено %CFG% у корені проєкту.
  goto :end
)

echo [RUN] Генерую DOCX-звіт...
python main.py --config %CFG% --mode draft-docx --log-level INFO
set ERR=%ERRORLEVEL%
if not "%ERR%"=="0" (
  echo [ERROR] Python завершився з кодом %ERR%.
  goto :end
)

rem === 6) Показати результат і відкрити папку build ===
echo.
echo [OK] Готово. Перевірте файли у: "%cd%\build"
if exist ".\build\" (
  start "" ".\build\"
)

:end
popd
endlocal
