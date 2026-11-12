@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"

echo ============================================================================
echo   STELLANTIS - IPC CONSISTENCY AUDITOR [INSTALADOR DE DEPENDENCIAS]
echo ============================================================================
echo.

:: 1. Verificar se o Python esta instalado no sistema
where python >nul 2>nul
if %ERRORLEVEL% equ 0 goto :found_python
where py >nul 2>nul
if %ERRORLEVEL% equ 0 goto :found_py

echo [ERRO] Python nao encontrado no sistema!
echo Por favor, instale o Python 3.10+ em https://www.python.org/downloads/
echo marcando a opcao "Add Python to PATH".
pause
exit /b 1

:found_python
set "SYS_PYTHON=python"
goto :check_venv

:found_py
set "SYS_PYTHON=py"

:check_venv
echo [1/4] Python detectado com sucesso:
%SYS_PYTHON% --version
echo.

:: 2. Criar ambiente virtual .venv caso nao exista
if exist ".\.venv\Scripts\python.exe" goto :venv_ok
echo [2/4] Criando ambiente virtual isolado [.venv]...
%SYS_PYTHON% -m venv .venv --system-site-packages
if %ERRORLEVEL% neq 0 (
    echo [ERRO] Falha ao criar o ambiente virtual .venv.
    pause
    exit /b 1
)
goto :install_pkgs

:venv_ok
echo [2/4] Ambiente virtual [.venv] ja existente.

:install_pkgs
echo.
echo [3/4] Instalando bibliotecas necessarias do arquivo requirements.txt...
".\.venv\Scripts\python.exe" -m pip install -r requirements.txt
echo.

:: 4. Gerar massa de dados ficticia inicial (J3U - 5870) e validar instalacao
echo [4/4] Gerando arquivos de exemplo [J3U - 5870] e validando o motor...
".\.venv\Scripts\python.exe" scripts\generate_mock_data.py
".\.venv\Scripts\python.exe" -m pytest -q
if %ERRORLEVEL% neq 0 (
    echo [ERRO] Falha na validacao dos pacotes instalados. Verifique as mensagens acima.
    pause
    exit /b 1
)

echo.
echo ============================================================================
echo   INSTALACAO CONCLUIDA COM SUCESSO!
echo   Agora basta dar dois cliques em "iniciar_sistema.bat" para abrir o sistema.
echo ============================================================================
echo.
if "%~1" neq "--no-pause" pause
