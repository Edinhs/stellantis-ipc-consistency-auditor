@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"

echo ============================================================================
echo   STELLANTIS - IPC CONSISTENCY AUDITOR [INICIALIZADOR DO SISTEMA]
echo ============================================================================
echo.

:: 1. Se o ambiente virtual nao existir, executa o instalador automaticamente
if exist ".\.venv\Scripts\python.exe" goto :check_mock
echo [AVISO] Ambiente virtual .venv nao encontrado.
echo Iniciando instalacao automatica dos requisitos...
call "%~dp0instalar_requisitos.bat" --no-pause
if %ERRORLEVEL% neq 0 (
    echo [ERRO] Nao foi possivel concluir a instalacao dos requisitos.
    pause
    exit /b 1
)

:check_mock
:: 2. Garantir que a pasta de dados ficticios exista
if exist ".\mock_data" goto :start_server
echo Gerando arquivos de exemplo J3U - 5870...
".\.venv\Scripts\python.exe" scripts\generate_mock_data.py

:start_server
:: 3. Abrir o navegador automaticamente e iniciar o servidor FastAPI
echo Iniciando servidor web em http://127.0.0.1:8000 ...
echo Mantenha esta janela aberta enquanto estiver utilizando o sistema.
echo.
start "" "http://127.0.0.1:8000"
".\.venv\Scripts\python.exe" -m uvicorn app.main:app --host 127.0.0.1 --port 8000

pause
