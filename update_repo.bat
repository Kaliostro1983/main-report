@echo off
REM update_repo.bat (simple, stable). Assumes the .bat is in the REPO ROOT.
setlocal ENABLEDELAYEDEXPANSION

pushd "%~dp0"

for /f %%i in ('powershell -NoProfile -Command "Get-Date -Format yyyy-MM-dd_HH-mm-ss"') do set "STAMP=%%i"
set "LOGDIR=%CD%\logs"
if not exist "%LOGDIR%" mkdir "%LOGDIR%"
set "LOG=%LOGDIR%\%STAMP%_update_repo.log"

for /f "delims=" %%b in ('git rev-parse --abbrev-ref HEAD') do set "BRANCH=%%b"
for /f "delims=" %%r in ('git remote get-url origin 2^>NUL') do set "REMOTE=%%r"
for /f "delims=" %%h in ('git rev-parse HEAD') do set "OLD=%%h"

echo Repo:   %CD%
echo Remote: %REMOTE%
echo Branch: %BRANCH%
echo HEAD(before): %OLD%

echo Repo:   %CD%>>"%LOG%"
echo Remote: %REMOTE%>>"%LOG%"
echo Branch: %BRANCH%>>"%LOG%"
echo HEAD(before): %OLD%>>"%LOG%"

git fetch --all --prune>>"%LOG%" 2>&1
git pull --rebase --autostash>>"%LOG%" 2>&1

for /f "delims=" %%h in ('git rev-parse HEAD') do set "NEW=%%h"

if /I "%OLD%"=="%NEW%" (
  echo Already up to date.
  echo Already up to date.>>"%LOG%"
) else (
  echo Updated: %OLD% -> %NEW%
  echo Updated: %OLD% -> %NEW%>>"%LOG%"
)

git --no-pager log --oneline -n 5 --decorate --graph
git --no-pager log --oneline -n 5 --decorate --graph>>"%LOG%"

echo Log: "%LOG%"
echo.
pause
popd
endlocal
