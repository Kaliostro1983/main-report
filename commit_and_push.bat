@echo off
REM commit_and_push.bat — Stage, commit (with timestamp), and push to current branch.
setlocal enabledelayedexpansion

REM Move to the directory where this script lives (assumed repo root or inside it)
pushd "%~dp0"

REM Ensure Git is available
git --version >NUL 2>&1 || (echo [ERROR] Git is not installed or not in PATH.& goto :end)

REM Stage all changes (including deletions)
git add -A

REM If nothing staged, skip commit
git diff --cached --quiet && (
  echo [INFO] Nothing to commit. Will attempt to push latest.
  goto push_only
)

REM Build robust timestamp via PowerShell (independent of system locale)
for /f %%i in ('powershell -NoProfile -Command "Get-Date -Format yyyy-MM-dd_HH-mm-ss"') do set TS=%%i

REM Commit message: use args if provided; otherwise default
set MSG=%*
if "%MSG%"=="" set MSG=Auto commit %TS%

echo [INFO] Committing: %MSG%
git commit -m "%MSG%"

:push_only
REM Detect current branch
for /f "tokens=*" %%b in ('git rev-parse --abbrev-ref HEAD') do set BRANCH=%%b

echo [INFO] Pushing to origin/!BRANCH! ...
git push origin !BRANCH!

echo.
echo [INFO] Recent commits on !BRANCH!:
git --no-pager log --oneline -n 5 --decorate --graph

:end
popd
endlocal
