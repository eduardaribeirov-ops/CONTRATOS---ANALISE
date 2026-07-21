@echo off
REM =====================================================================
REM AnalisarContrato.bat - Lancador de UM CLIQUE para Windows
REM
REM Analisa um contrato .docx ou .pdf sem precisar abrir o terminal nem
REM digitar comandos. Use de qualquer uma destas formas:
REM
REM   1) ARRASTE o arquivo do contrato (.docx ou .pdf) para cima deste
REM      arquivo .bat e solte. Pode arrastar varios de uma vez.
REM   2) DE UM DUPLO CLIQUE neste arquivo e informe o caminho do contrato
REM      quando for pedido.
REM   3) Coloque um atalho deste .bat na pasta "Enviar para" do Windows
REM      (veja INSTALACAO_WINDOWS.md) e use o botao direito >
REM      "Enviar para" > "Analisar Contrato" em qualquer contrato.
REM
REM Na PRIMEIRA vez ele prepara sozinho um ambiente Python isolado
REM (.venv) e instala as dependencias - isso pode demorar alguns minutos.
REM Nas vezes seguintes e quase instantaneo.
REM
REM A analise e 100%% offline (sem internet, sem chave de API, sem custo).
REM O relatorio e salvo como <nome_do_contrato>_analise.docx na mesma
REM pasta do contrato e aberto automaticamente no Word ao final.
REM =====================================================================

setlocal enabledelayedexpansion
chcp 65001 >nul
cd /d "%~dp0"

set "RAIZ=%~dp0.."
set "PYDIR=%RAIZ%\python"
set "VENV=%RAIZ%\.venv"
set "PYEXE=%VENV%\Scripts\python.exe"

REM --- Localiza um Python instalado no sistema (py launcher ou python) ---
set "PYSYS="
where py >nul 2>nul && set "PYSYS=py"
if not defined PYSYS (
    where python >nul 2>nul && set "PYSYS=python"
)
if not defined PYSYS (
    echo.
    echo [ERRO] Python nao foi encontrado no computador.
    echo Instale o Python 3 em https://www.python.org/downloads/
    echo e marque a opcao "Add Python to PATH" durante a instalacao.
    echo.
    pause
    goto :eof
)

REM --- Prepara o ambiente isolado na primeira execucao ---
if not exist "%PYEXE%" (
    echo.
    echo Preparando o ambiente pela primeira vez. Isso pode demorar alguns
    echo minutos e so acontece uma vez...
    echo.
    %PYSYS% -m venv "%VENV%"
    if errorlevel 1 (
        echo [ERRO] Nao foi possivel criar o ambiente Python ^(.venv^).
        pause
        goto :eof
    )
    "%PYEXE%" -m pip install --upgrade pip
    "%PYEXE%" -m pip install -r "%PYDIR%\requirements.txt"
    if errorlevel 1 (
        echo [ERRO] Falha ao instalar as dependencias.
        pause
        goto :eof
    )
    echo.
    echo Ambiente pronto.
)

REM --- Se nada foi arrastado, pergunta o caminho do contrato ---
if "%~1"=="" (
    echo.
    echo Arraste um contrato .docx ou .pdf para cima deste arquivo, ou
    set /p "ARQ=informe aqui o caminho do contrato: "
    if not defined ARQ goto :fim
    REM remove aspas caso o usuario cole um caminho entre aspas
    set "ARQ=!ARQ:"=!"
    call :analisar "!ARQ!"
    goto :fim
)

REM --- Analisa cada arquivo arrastado (aceita varios) ---
:loop
if "%~1"=="" goto :fim
call :analisar "%~1"
shift
goto :loop

:fim
echo.
echo Concluido. Feche esta janela quando quiser.
pause
goto :eof

REM ---------------------------------------------------------------------
:analisar
if not exist "%~1" (
    echo.
    echo [AVISO] Arquivo nao encontrado: %~1
    goto :eof
)
echo.
echo ============================================================
echo Analisando: %~nx1
echo ============================================================
"%PYEXE%" "%PYDIR%\analisar_contrato.py" "%~1"
if errorlevel 1 (
    echo [AVISO] Nao foi possivel analisar: %~nx1
    goto :eof
)
REM Abre o relatorio gerado (mesmo nome + _analise.docx, na mesma pasta)
set "RELATORIO=%~dpn1_analise.docx"
if exist "%RELATORIO%" (
    echo Abrindo o relatorio...
    start "" "%RELATORIO%"
)
goto :eof
