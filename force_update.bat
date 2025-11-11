@echo off
setlocal

REM =================== Find repo root (walk up until .git found) ===================
set "DIR=%~dp0"
for %%# in ("%DIR%") do set "DIR=%%~f#"
:find_root
if exist "%DIR%\.git" (
  set "REPO=%DIR%"
) else (
  set "PARENT=%DIR%.."
  for %%# in ("%PARENT%") do set "DIR=%%~f#"
  if /I "%DIR%"=="%PARENT%" (
    echo [ERROR] .git не знайдено починаючи з: %~dp0
    goto end
  ) else (
    goto find_root
  )
)

pushd "%REPO%"

echo [WARN] Жорстке вирівнювання під origin/main (локальні незакомічені зміни буде втрачено).
git fetch --all --prune
git reset --hard origin/main
git --no-pager log -1 --oneline --decorate

:end
echo.
pause
popd
endlocal
