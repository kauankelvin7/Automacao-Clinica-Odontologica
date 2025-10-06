@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul 2>&1
title 🦷 Automação Porto Seguro - Automação v1.0
color 0B
mode con: cols=100 lines=35

:: ===============================================================
:: BANNER INICIAL
:: ===============================================================
cls
echo.
echo    ╔════════════════════════════════════════════════════════════════════════════════════════╗
echo    ║                                                                                        ║
echo    ║     ██████╗  ██████╗ ██████╗ ████████╗ ██████╗     ███████╗███████╗ ██████╗ ██╗   ██╗██████╗  ██████╗    ║
echo    ║     ██╔══██╗██╔═══██╗██╔══██╗╚══██╔══╝██╔═══██╗    ██╔════╝██╔════╝██╔════╝ ██║   ██║██╔══██╗██╔═══██╗   ║
echo    ║     ██████╔╝██║   ██║██████╔╝   ██║   ██║   ██║    ███████╗█████╗  ██║  ███╗██║   ██║██████╔╝██║   ██║   ║
echo    ║     ██╔═══╝ ██║   ██║██╔══██╗   ██║   ██║   ██║    ╚════██║██╔══╝  ██║   ██║██║   ██║██╔══██╗██║   ██║   ║
echo    ║     ██║     ╚██████╔╝██║  ██║   ██║   ╚██████╔╝    ███████║███████╗╚██████╔╝╚██████╔╝██║  ██║╚██████╔╝   ║
echo    ║     ╚═╝      ╚═════╝ ╚═╝  ╚═╝   ╚═╝    ╚═════╝     ╚══════╝╚══════╝ ╚═════╝  ╚═════╝ ╚═╝  ╚═╝ ╚═════╝    ║
echo    ║                                                                                        ║
echo    ║                              🦷 SISTEMA DE AUTOMAÇÃO 🦷                               ║
echo    ║                                   Versão 1.0 - 2025                                    ║
echo    ║                                                                                        ║
echo    ╚════════════════════════════════════════════════════════════════════════════════════════╝
echo.
echo                               ⚡ Verificando dependências do sistema...
timeout /t 2 >nul
echo.

:: ===============================================================
:: VERIFICA PYTHON
:: ===============================================================
cls
call :PRINT_HEADER
echo    [1/4] 🔍 Verificando instalação do Python...
python --version >nul 2>&1
if errorlevel 1 (
    color 0C
    echo.
    echo    ✗ ERRO: Python não foi encontrado no sistema!
    echo.
    echo    📋 Instruções:
    echo       1. Acesse: https://www.python.org/downloads/
    echo       2. Instale Python 3.8 ou superior
    echo       3. Marque a opção "Add Python to PATH"
    echo.
    pause
    exit /b 1
)
for /f "tokens=*" %%i in ('python --version 2^>^&1') do set PYTHON_VERSION=%%i
echo    ✓ Python detectado: %PYTHON_VERSION%
timeout /t 1 >nul
echo.

:: ===============================================================
:: VERIFICA PIP
:: ===============================================================
cls
call :PRINT_HEADER
echo    [2/4] 🔍 Verificando gerenciador de pacotes (pip)...
pip --version >nul 2>&1
if errorlevel 1 (
    echo.
    echo    ✗ pip não encontrado — tentando reparar automaticamente...
    python -m ensurepip --default-pip
    if errorlevel 1 (
        color 0C
        echo.
        echo    ✗ Falha ao reparar pip!
        pause
        exit /b 1
    )
)
for /f "tokens=*" %%i in ('pip --version 2^>^&1') do set PIP_VERSION=%%i
echo    ✓ pip detectado: %PIP_VERSION%
timeout /t 1 >nul
echo.

:: ===============================================================
:: INSTALA DEPENDÊNCIAS
:: ===============================================================
cls
call :PRINT_HEADER
echo    [3/4] 📦 Instalando/atualizando dependências...
echo.
echo    → Atualizando pip, setuptools e wheel...
python -m pip install --upgrade pip setuptools wheel --quiet --disable-pip-version-check
if errorlevel 1 (
    echo    ⚠ Não foi possível atualizar completamente.
)
echo.
echo    → Instalando bibliotecas principais...
pip install selenium openpyxl webdriver-manager --quiet --disable-pip-version-check
if errorlevel 1 (
    color 0C
    echo.
    echo    ✗ FALHA ao instalar dependências!
    echo    💡 Sugestões:
    echo       - Verifique conexão com a internet
    echo       - Execute este arquivo como Administrador
    echo       - Tente manualmente: pip install selenium openpyxl webdriver-manager
    pause
    exit /b 1
)
echo    ✓ Dependências instaladas com sucesso!
timeout /t 1 >nul
echo.

:: ===============================================================
:: VERIFICA ARQUIVOS
:: ===============================================================
cls
call :PRINT_HEADER
echo    [4/4] 📋 Verificando estrutura de arquivos...
echo.
if not exist "main.py" (
    color 0C
    echo    ✗ ERRO: Arquivo principal 'main.py' não encontrado!
    echo.
    echo    Coloque este script na mesma pasta do arquivo Python principal.
    pause
    exit /b 1
)
echo    ✓ Arquivo main.py ............. [OK]

if not exist "arquivos" (
    mkdir arquivos >nul 2>&1
    echo    ✓ Pasta 'arquivos' criada automaticamente
) else (
    echo    ✓ Pasta 'arquivos' ........ [OK]
)
if not exist "logs" (
    mkdir logs >nul 2>&1
)
echo    ✓ Pasta 'logs' ............ [OK]
echo.
timeout /t 2 >nul

:: ===============================================================
:: INICIA O SISTEMA
:: ===============================================================
color 0A
cls
call :PRINT_HEADER
echo.
echo    ╔════════════════════════════════════════════════════════════════════════════════════════╗
echo    ║                                                                                        ║
echo    ║                            ✅ SISTEMA PRONTO PARA EXECUÇÃO ✅                          ║
echo    ║                                                                                        ║
echo    ╚════════════════════════════════════════════════════════════════════════════════════════╝
echo.
echo                           ⚡ Iniciando automação em 3 segundos...
timeout /t 3 >nul

color 0B
cls
echo.
echo    ════════════════════════════════════════════════════════════════════════════════════════
echo                                 🚀 EXECUTANDO SISTEMA 🚀
echo    ════════════════════════════════════════════════════════════════════════════════════════
echo.

python main.py
set EXIT_CODE=%ERRORLEVEL%

echo.
if %EXIT_CODE% EQU 0 (
    color 0A
    echo    ✅ Sistema encerrado com sucesso!
) else (
    color 0C
    echo    ⚠ Sistema encerrado com erros (código: %EXIT_CODE%)
)
echo.
echo    ✅ Obrigado por utilizar!
echo.
pause
exit /b %EXIT_CODE%

:: ===============================================================
:: FUNÇÃO DE CABEÇALHO
:: ===============================================================
:PRINT_HEADER
echo.
echo    ════════════════════════════════════════════════════════════════════════════════════════
echo                            🦷 AUTOMAÇÃO DO SISTEMA DE CREDENCIAMENTO - PORTO SEGURO 🦷
echo    ════════════════════════════════════════════════════════════════════════════════════════
echo.
goto :eof
