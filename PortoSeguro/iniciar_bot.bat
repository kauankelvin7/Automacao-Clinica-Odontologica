@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul 2>&1
title 🦷 Automação Porto Seguro v1.0
mode con: cols=90 lines=30

:: Habilita cores ANSI no Windows 10+
reg add HKCU\Console /v VirtualTerminalLevel /t REG_DWORD /d 1 /f >nul 2>&1

:: Definição de cores ANSI
set "RESET=[0m"
set "BOLD=[1m"
set "CYAN=[96m"
set "BLUE=[94m"
set "GREEN=[92m"
set "YELLOW=[93m"
set "RED=[91m"
set "MAGENTA=[95m"
set "WHITE=[97m"

:: Arquivo de controle de instalação
set "SETUP_FILE=.setup_complete"

:: ===============================================================
:: BANNER INICIAL
:: ===============================================================
cls
echo.
echo %CYAN%   ╔══════════════════════════════════════════════════════════════════════════════╗%RESET%
echo %CYAN%   ║%RESET%                                                                              %CYAN%║%RESET%
echo %CYAN%   ║%BLUE%    ██████╗  ██████╗ ██████╗ ████████╗ ██████╗     ███████╗ ██████╗         %CYAN%║%RESET%
echo %CYAN%   ║%BLUE%    ██╔══██╗██╔═══██╗██╔══██╗╚══██╔══╝██╔═══██╗    ██╔════╝██╔════╝         %CYAN%║%RESET%
echo %CYAN%   ║%BLUE%    ██████╔╝██║   ██║██████╔╝   ██║   ██║   ██║    ███████╗██║  ███╗        %CYAN%║%RESET%
echo %CYAN%   ║%BLUE%    ██╔═══╝ ██║   ██║██╔══██╗   ██║   ██║   ██║    ╚════██║██║   ██║        %CYAN%║%RESET%
echo %CYAN%   ║%BLUE%    ██║     ╚██████╔╝██║  ██║   ██║   ╚██████╔╝    ███████║╚██████╔╝        %CYAN%║%RESET%
echo %CYAN%   ║%BLUE%    ╚═╝      ╚═════╝ ╚═╝  ╚═╝   ╚═╝    ╚═════╝     ╚══════╝ ╚═════╝         %CYAN%║%RESET%
echo %CYAN%   ║%RESET%                                                                              %CYAN%║%RESET%
echo %CYAN%   ║%MAGENTA%              🦷 SISTEMA DE AUTOMAÇÃO DE CREDENCIAMENTO 🦷                   %CYAN%║%RESET%
echo %CYAN%   ║%YELLOW%                          Versão 2.0 - 2025                                  %CYAN%║%RESET%
echo %CYAN%   ║%RESET%                                                                              %CYAN%║%RESET%
echo %CYAN%   ╚══════════════════════════════════════════════════════════════════════════════╝%RESET%
echo.

:: Verifica se já foi feito o setup completo
if exist "%SETUP_FILE%" (
    echo %GREEN%   ✓ Dependências já verificadas anteriormente%RESET%
    echo %CYAN%   ⚡ Modo rápido ativado - Pulando verificações...%RESET%
    echo.
    echo %YELLOW%   💡 Para forçar nova verificação, delete o arquivo: %SETUP_FILE%%RESET%
    timeout /t 2 >nul
    goto :QUICK_START
)

echo %YELLOW%                      ⚡ Verificando dependências do sistema...%RESET%
timeout /t 2 >nul
echo.

:: ===============================================================
:: VERIFICA PYTHON
:: ===============================================================
cls
call :PRINT_HEADER
echo %CYAN%   [1/4]%RESET% %YELLOW%🔍 Verificando instalação do Python...%RESET%
python --version >nul 2>&1
if errorlevel 1 (
    echo.
    echo %RED%   ✗ ERRO: Python não foi encontrado no sistema!%RESET%
    echo.
    echo %YELLOW%   📋 Instruções:%RESET%
    echo %WHITE%      1. Acesse: https://www.python.org/downloads/%RESET%
    echo %WHITE%      2. Instale Python 3.8 ou superior%RESET%
    echo %WHITE%      3. Marque a opção "Add Python to PATH"%RESET%
    echo.
    pause
    exit /b 1
)
for /f "tokens=*" %%i in ('python --version 2^>^&1') do set PYTHON_VERSION=%%i
echo %GREEN%   ✓ Python detectado: %PYTHON_VERSION%%RESET%
timeout /t 1 >nul
echo.

:: ===============================================================
:: VERIFICA PIP
:: ===============================================================
cls
call :PRINT_HEADER
echo %CYAN%   [2/4]%RESET% %YELLOW%🔍 Verificando gerenciador de pacotes (pip)...%RESET%
pip --version >nul 2>&1
if errorlevel 1 (
    echo.
    echo %YELLOW%   ✗ pip não encontrado — tentando reparar automaticamente...%RESET%
    python -m ensurepip --default-pip
    if errorlevel 1 (
        echo.
        echo %RED%   ✗ Falha ao reparar pip!%RESET%
        pause
        exit /b 1
    )
)
for /f "tokens=*" %%i in ('pip --version 2^>^&1') do set PIP_VERSION=%%i
echo %GREEN%   ✓ pip detectado: %PIP_VERSION%%RESET%
timeout /t 1 >nul
echo.

:: ===============================================================
:: VERIFICA E INSTALA DEPENDÊNCIAS (SE NECESSÁRIO)
:: ===============================================================
cls
call :PRINT_HEADER
echo %CYAN%   [3/4]%RESET% %YELLOW%📦 Verificando dependências Python...%RESET%
echo.

set "NEED_INSTALL=0"

:: Verifica cada pacote individualmente
echo %CYAN%   → Verificando selenium...%RESET%
python -c "import selenium" >nul 2>&1
if errorlevel 1 (
    echo %YELLOW%      • selenium não encontrado%RESET%
    set "NEED_INSTALL=1"
) else (
    echo %GREEN%      ✓ selenium já instalado%RESET%
)

echo %CYAN%   → Verificando openpyxl...%RESET%
python -c "import openpyxl" >nul 2>&1
if errorlevel 1 (
    echo %YELLOW%      • openpyxl não encontrado%RESET%
    set "NEED_INSTALL=1"
) else (
    echo %GREEN%      ✓ openpyxl já instalado%RESET%
)

echo %CYAN%   → Verificando webdriver-manager...%RESET%
python -c "import webdriver_manager" >nul 2>&1
if errorlevel 1 (
    echo %YELLOW%      • webdriver-manager não encontrado%RESET%
    set "NEED_INSTALL=1"
) else (
    echo %GREEN%      ✓ webdriver-manager já instalado%RESET%
)

echo %CYAN%   → Verificando python-dotenv...%RESET%
python -c "import dotenv" >nul 2>&1
if errorlevel 1 (
    echo %YELLOW%      • python-dotenv não encontrado%RESET%
    set "NEED_INSTALL=1"
) else (
    echo %GREEN%      ✓ python-dotenv já instalado%RESET%
)

echo.

if "%NEED_INSTALL%"=="1" (
    echo %YELLOW%   📥 Instalando pacotes faltantes...%RESET%
    echo.
    pip install selenium openpyxl webdriver-manager python-dotenv --quiet --disable-pip-version-check
    if errorlevel 1 (
        echo.
        echo %RED%   ✗ FALHA ao instalar dependências!%RESET%
        echo %YELLOW%   💡 Sugestões:%RESET%
        echo %WHITE%      - Verifique conexão com a internet%RESET%
        echo %WHITE%      - Execute este arquivo como Administrador%RESET%
        echo %WHITE%      - Tente manualmente: pip install selenium openpyxl webdriver-manager python-dotenv%RESET%
        pause
        exit /b 1
    )
    echo %GREEN%   ✓ Dependências instaladas com sucesso!%RESET%
) else (
    echo %GREEN%   ✓ Todas as dependências já estão instaladas!%RESET%
)

timeout /t 1 >nul
echo.

:: ===============================================================
:: VERIFICA ARQUIVOS
:: ===============================================================
cls
call :PRINT_HEADER
echo %CYAN%   [4/4]%RESET% %YELLOW%📋 Verificando estrutura de arquivos...%RESET%
echo.
if not exist "src\main.py" (
    echo %RED%   ✗ ERRO: Arquivo principal 'src\main.py' não encontrado!%RESET%
    echo.
    echo %WHITE%   Coloque este script na mesma pasta do arquivo Python principal.%RESET%
    pause
    exit /b 1
)
echo %GREEN%   ✓ Arquivo main.py ............. [OK]%RESET%

if not exist "arquivos" (
    mkdir arquivos >nul 2>&1
    echo %GREEN%   ✓ Pasta 'arquivos' criada automaticamente%RESET%
) else (
    echo %GREEN%   ✓ Pasta 'arquivos' ........ [OK]%RESET%
)
if not exist "logs" (
    mkdir logs >nul 2>&1
)
echo %GREEN%   ✓ Pasta 'logs' ............ [OK]%RESET%
echo.

:: Cria arquivo de controle para pular verificações na próxima vez
echo Setup completo em %date% %time% > "%SETUP_FILE%"
echo %GREEN%   ✓ Configuração salva - próximas execuções serão mais rápidas!%RESET%
echo.

timeout /t 2 >nul

:: ===============================================================
:: INICIA O SISTEMA (MODO RÁPIDO)
:: ===============================================================
:QUICK_START
cls
call :PRINT_HEADER
echo.
echo %GREEN%   ╔══════════════════════════════════════════════════════════════════════════════╗%RESET%
echo %GREEN%   ║%RESET%                                                                              %GREEN%║%RESET%
echo %GREEN%   ║%BOLD%%WHITE%                      ✅ SISTEMA PRONTO PARA EXECUÇÃO ✅                      %RESET%%GREEN%║%RESET%
echo %GREEN%   ║%RESET%                                                                              %GREEN%║%RESET%
echo %GREEN%   ╚══════════════════════════════════════════════════════════════════════════════╝%RESET%
echo.
echo %YELLOW%                      ⚡ Iniciando automação em 2 segundos...%RESET%
timeout /t 2 >nul

cls
echo.
echo %CYAN%   ══════════════════════════════════════════════════════════════════════════════%RESET%
echo %BOLD%%MAGENTA%                            🚀 EXECUTANDO SISTEMA 🚀%RESET%
echo %CYAN%   ══════════════════════════════════════════════════════════════════════════════%RESET%
echo.

python src/main.py
set EXIT_CODE=%ERRORLEVEL%

echo.
if %EXIT_CODE% EQU 0 (
    echo %GREEN%   ✅ Sistema encerrado com sucesso!%RESET%
) else (
    echo %RED%   ⚠ Sistema encerrado com erros (código: %EXIT_CODE%)%RESET%
)
echo.
echo %CYAN%   ✅ Obrigado por utilizar!%RESET%
echo.
pause
exit /b %EXIT_CODE%

:: ===============================================================
:: FUNÇÃO DE CABEÇALHO
:: ===============================================================
:PRINT_HEADER
echo.
echo %CYAN%   ══════════════════════════════════════════════════════════════════════════════%RESET%
echo %BOLD%%MAGENTA%           🦷 AUTOMAÇÃO DE CREDENCIAMENTO - PORTO SEGURO 🦷%RESET%
echo %CYAN%   ══════════════════════════════════════════════════════════════════════════════%RESET%
echo.
goto :eof
