@echo off
setlocal

REM 1. Перейти в папку проєкту (на випадок запуску з іншого місця)
cd /d %~dp0

REM 2. Якщо віртуального оточення нема – створити
if not exist ".venv" (
    echo [INFO] Створюю віртуальне середовище .venv ...
    py -m venv .venv
)

REM 3. Активувати venv
call ".venv\Scripts\activate.bat"

REM 4. Оновити pip (один раз не завадить)
echo [INFO] Оновлюю pip ...
python -m pip install --upgrade pip

REM 5. Встановити залежності
echo [INFO] Встановлюю залежності з requirements.txt ...
python -m pip install -r requirements.txt

REM 6. Запустити твій main
echo [INFO] Запускаю main.py ...
python main.py

endlocal
pause
