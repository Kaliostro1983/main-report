@echo off
REM Використання:
REM   scripts\ingest_whatsapp.bat "D:\path\export.txt" Ocheret
set FILE=%~1
set CHAT=%~2
if "%FILE%"=="" echo Need path to exported TXT & exit /b 1
if "%CHAT%"=="" set CHAT=Ocheret
.\.venv\Scripts\python.exe -m whatsapp_ingest.ingest --file "%FILE%" --chat "%CHAT%" --db "src\data\data.db"
