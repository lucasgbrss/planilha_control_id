@echo off
setlocal

echo ========================================
echo  Criando instalador do Control ID Reader
echo ========================================
echo.

if not exist "dist\Control ID Reader.exe" (
    echo ERRO: dist\Control ID Reader.exe nao encontrado.
    echo Gere o executavel primeiro usando atualizar_exe.bat.
    echo.
    pause
    exit /b 1
)

set "ISCC="
for /f "delims=" %%I in ('where ISCC.exe 2^>nul') do (
    if not defined ISCC set "ISCC=%%I"
)

if not defined ISCC (
    set "ISCC=C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
)

if not exist "%ISCC%" (
    set "ISCC=C:\Program Files\Inno Setup 6\ISCC.exe"
)

if not exist "%ISCC%" (
    echo ERRO: Inno Setup nao encontrado.
    echo Instale o Inno Setup 6 e execute este script novamente.
    echo Download: https://jrsoftware.org/isdl.php
    echo.
    pause
    exit /b 1
)

echo Compilando instalador...
"%ISCC%" "instalador.iss"
if errorlevel 1 (
    echo.
    echo ERRO: falha ao criar o instalador.
    pause
    exit /b 1
)

echo.
echo ========================================
echo  Concluido!
echo  Instalador criado em:
echo  installer\Control ID Reader Setup 0.11.6.exe
echo ========================================
pause
