@echo off
REM %~dp0 = абсолютний шлях до теки, де лежить цей .bat (scripts\)
REM Переходимо в корінь проєкту (на рівень вище за scripts)
cd /d "%~dp0.."

REM Активуємо віртуальне середовище (воно теж лежить у корені поряд з main.py)
call .venv\Scripts\activate

REM Генеруємо звіт (тепер робоча директорія вже корінь, тож build піде у правильну build\)
python main.py --mode simple-report --config config.yml

pause
