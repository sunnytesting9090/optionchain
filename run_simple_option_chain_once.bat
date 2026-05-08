@echo off
cd /d "%~dp0"
title Simple Option Chain One-Time Refresh
echo ==========================================
echo   SIMPLE OPTION CHAIN ONE-TIME REFRESH
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
echo Opening simple_option_chain.xlsx and running one refresh...
echo.
python update_simple_option_chain.py --once
set exit_code=%errorlevel%

if not "%exit_code%"=="0" (
    echo.
    echo One-time refresh failed with exit code %exit_code%.
    echo Press any key to close this window.
    pause >nul
    exit /b %exit_code%
)

echo.
echo One-time refresh completed. Closing this launcher.
timeout /t 2 /nobreak >nul
exit /b 0
