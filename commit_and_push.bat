@echo off
REM commit_and_push.bat (simple, stable). Assumes the .bat is in the REPO ROOT.
setlocal ENABLEDELAYEDEXPANSION

REM Go to this script's folder (repo root)
pushd "%~dp0"

REM Timestamp for logs and commit message (uses PowerShell only to format time string)
for /f %%i in ('powershell -NoProfile -Command "Get-Date -Format yyyy-MM-dd_HH-mm-ss"') do set "STAMP=%%i"

REM Prepare log file
set "LOGDIR=%CD%\logs"
if not exist "%LOGDIR%" mkdir "%LOGDIR%"
set "LOG=%LOGDIR%\%STAMP%_commit_and_push.log"

REM Basic context
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

REM Stage and commit with auto message
set "MSG=Auto commit %STAMP%"
git add -A>>"%LOG%" 2>&1
git diff --cached --quiet
if %ERRORLEVEL% NEQ 0 (
  git commit -m "%MSG%">>"%LOG%" 2>&1
) else (
  echo No changes to commit.>>"%LOG%"
  echo No changes to commit.
)

REM Rebase-pull and push
git pull --rebase --autostash>>"%LOG%" 2>&1
git push>>"%LOG%" 2>&1

for /f "delims=" %%h in ('git rev-parse HEAD') do set "NEW=%%h"

echo HEAD(after): %NEW%
git --no-pager log --oneline -n 5 --decorate --graph

echo HEAD(after): %NEW%>>"%LOG%"
git --no-pager log --oneline -n 5 --decorate --graph>>"%LOG%"

echo Log: "%LOG%"
echo.
pause
popd
endlocal
