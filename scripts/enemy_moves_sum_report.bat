@echo off
REM Генерація звіту "Переміщення ворога" на основі moves.xlsx

REM Переходимо в корінь проєкту (з папки scripts на рівень вище)
cd /d "%~dp0.."

REM За потреби заміни python на py або шлях до venv, наприклад:
REM call venv\Scripts\activate

python main.py --config config.yml --mode enemy-moves-sum --log-level INFO

echo.
echo [OK] Звіт про переміщення ворога згенеровано (див. папку build).
echo.
pause
