@echo off
pushd "%~dp0\.."  || (echo [ERROR] Can't cd to repo root & pause & exit /b 1)
if exist ".venv\Scripts\activate" ( call .\.venv\Scripts\activate )
python main.py --mode artyleria-report --config config.yml
popd
pause
