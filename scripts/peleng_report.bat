@echo off
setlocal
REM Перейти в корінь репо (папка вище за scripts)
pushd "%~dp0\.."  || (echo [ERROR] Can't cd to repo root & pause & exit /b 1)

REM Активувати venv, якщо є
if exist ".venv\Scripts\activate" (
    call .\.venv\Scripts\activate
) else (
    echo [WARN] .venv not found. Using system Python...
)

REM Виклик з опціональним шляхом: peleng_report.bat [txt-file]
if "%~1"=="" (
    python -m src.pelengreport.runner
) else (
    python -m src.pelengreport.runner "%~1"
)

popd
pause
