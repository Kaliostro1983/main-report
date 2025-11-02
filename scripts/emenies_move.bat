@echo off
cd /d "%~dp0.."

call .venv\Scripts\activate

python main.py --mode move_enemies --config config.yml

pause
