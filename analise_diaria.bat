@echo off
REM Wrapper chamado pelo Windows Task Scheduler.
REM Roda analise_diaria.py com log em logs\YYYY-MM-DD.log.

cd /d "%~dp0"
if not exist logs mkdir logs

for /f "tokens=2 delims==" %%a in ('wmic os get localdatetime /value ^| find "="') do set DT=%%a
set LOG=logs\%DT:~0,4%-%DT:~4,2%-%DT:~6,2%.log

echo === %DATE% %TIME% === >> "%LOG%"
python analise_diaria.py >> "%LOG%" 2>&1
echo === fim (exit %ERRORLEVEL%) === >> "%LOG%"
echo. >> "%LOG%"

exit /b %ERRORLEVEL%
