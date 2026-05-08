@echo off
cd /d "%~dp0"
title Simple Option Chain Launcher
echo ==========================================
echo   SIMPLE OPTION CHAIN LAUNCHER
echo ==========================================
echo.
set /p token=Paste today's ACCESS TOKEN here: 

if "%token%"=="" (
    echo.
    echo No token was entered. Press any key to exit.
    pause >nul
    exit /b 1
)

> access_token.txt <nul set /p=%token%

echo.
echo Access token saved to access_token.txt
echo Launching the simple option chain workbook...
echo.
call start_simple_option_chain_updater.bat
