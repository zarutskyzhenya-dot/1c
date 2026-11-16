@echo off
setlocal enabledelayedexpansion
cd /d %~dp0

rem ───────────────────────────────────────────────────────────────────
rem  Mode 1: drag file onto bat
rem ───────────────────────────────────────────────────────────────────
if not "%~1"=="" (
    C:\Python311-32\python.exe create_contractor_odata.py "%~1" --db Vlada2
    echo.
    set /p "DEL=Delete file? 1=yes / 0=no: "
    if "!DEL!"=="1" if exist "%~1" del "%~1"
    pause
    goto :eof
)

rem ───────────────────────────────────────────────────────────────────
rem  Mode 2: process all files in input\ folder
rem ───────────────────────────────────────────────────────────────────
set "INPUT_DIR=%~dp0input"

if not exist "%INPUT_DIR%\" (
    echo Folder not found: %INPUT_DIR%
    pause
    goto :eof
)

set "FOUND=0"
for %%F in ("%INPUT_DIR%\*.*") do (
    if /i not "%%~xF"==".log" set "FOUND=1"
)

if "!FOUND!"=="0" (
    echo.
    echo  input\ folder is empty.
    echo  Put files here: %INPUT_DIR%
    echo.
    pause
    goto :eof
)

echo.
echo Processing files from input\ ...
echo.

for %%F in ("%INPUT_DIR%\*.*") do (
    if /i not "%%~xF"==".log" (
        echo ========================================
        echo  File: %%~nxF
        echo ========================================
        C:\Python311-32\python.exe create_contractor_odata.py "%%F" --db Vlada2 --yes
        echo.
        set /p "DEL=Delete files from input? 1=yes / 0=no: "
        if "!DEL!"=="1" del "%%F"
        echo.
    )
)

echo.
echo Done.
pause
