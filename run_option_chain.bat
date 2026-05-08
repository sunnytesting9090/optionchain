@echo off
cd /d "%~dp0"
title Daily Option Chain Launcher
echo ==========================================
echo   DAILY OPTION CHAIN LAUNCHER
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
echo Launching the updater and opening Excel...
echo.
call start_option_chain_updater.bat
