@echo off
cd /d "%~dp0.."

call .venv\Scripts\activate

python main.py --mode freq-lists --config config.yml

pause
