@echo off
setlocal
cd /d %~dp0

echo =========================================
echo BUILD EXE - Invoice AutoHandler v1.0
echo =========================================

REM -----------------------------------------
REM [1] Garantir .venv local
REM -----------------------------------------
if not exist .venv (
    echo A criar ambiente virtual...
    python -m venv .venv
)

REM -----------------------------------------
REM [2] Atualizar pip
REM -----------------------------------------
echo.
echo [1/4] Atualizar pip...
.\.venv\Scripts\python.exe -m pip install --upgrade pip

REM -----------------------------------------
REM [3] Instalar dependencias
REM -----------------------------------------
echo.
echo [2/4] Instalar dependencias...
.\.venv\Scripts\python.exe -m pip install pandas pypdf openpyxl pywin32 pyinstaller

REM -----------------------------------------
REM [4] Limpar builds anteriores
REM -----------------------------------------
echo.
echo [3/4] Limpar builds anteriores...
rmdir /s /q build 2>nul
rmdir /s /q dist 2>nul
del /q *.spec 2>nul

REM -----------------------------------------
REM [5] Compilar EXE
REM -----------------------------------------
echo.
echo [4/4] Compilar EXE standalone...
.\.venv\Scripts\python.exe -m PyInstaller ^
--noconfirm ^
--clean ^
--onefile ^
--windowed ^
--name Invoice_AutoHandler_v1_0 ^
Invoice_AutoHandler_v1_0.py

IF %ERRORLEVEL% NEQ 0 (
    echo.
    echo ERRO: compilacao falhou.
    pause
    exit /b 1
)

echo.
echo =========================================
echo BUILD TERMINADO COM SUCESSO
echo EXE criado em: dist\Invoice_AutoHandler_v1_0.exe
echo =========================================
pause