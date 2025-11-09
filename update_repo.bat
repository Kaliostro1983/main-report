@echo off
setlocal
pushd "%~dp0\.."

git --version >NUL 2>&1 || (echo [ERROR] Git not found.& goto :end)

for /f "tokens=*" %%b in ('git rev-parse --abbrev-ref HEAD') do set BRANCH=%%b
echo [INFO] Current branch: %BRANCH%

echo [INFO] Fetching...
git fetch origin

echo [INFO] Resetting to origin/%BRANCH% ...
git reset --hard origin/%BRANCH%

echo [INFO] HEAD now at:
git --no-pager log -1 --oneline --decorate

:end
popd
endlocal
