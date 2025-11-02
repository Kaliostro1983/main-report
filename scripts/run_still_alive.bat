@echo off
REM Переходимо в корінь проєкту (на рівень вище за scripts)
cd /d "%~dp0.."

REM Активуємо віртуальне середовище
call .venv\Scripts\activate

REM Генеруємо звіт активності радіомереж
python main.py --mode still-alive --config config.yml

pause
