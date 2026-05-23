@echo off
REM Agenda a análise diária no Windows Task Scheduler.
REM Roda uma vez como administrador (ou usuário normal — depende da política).
REM Padrão: todo dia às 07:00. Edite /ST aqui pra mudar.

cd /d "%~dp0"
set HORA=07:00
set NOME=FindMe Analise Diaria

schtasks /Create /SC DAILY /TN "%NOME%" /TR "%~dp0analise_diaria.bat" /ST %HORA% /F

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo ERRO: falha ao agendar. Tente rodar como administrador.
    pause
    exit /b 1
)

echo.
echo Tarefa agendada com sucesso:
echo   Nome:  %NOME%
echo   Hora:  %HORA% (diaria)
echo   Acao:  %~dp0analise_diaria.bat
echo.
echo Pra ver/editar: abra "Agendador de Tarefas" no Windows e procure "%NOME%".
echo Pra desagendar: schtasks /Delete /TN "%NOME%" /F
echo.
pause
