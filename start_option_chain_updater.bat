@echo off
cd /d "%~dp0"
title Option Chain Updater
echo ==========================================
echo   OPTION CHAIN UPDATER
echo ==========================================
echo.
echo Starting Python updater...
echo Excel will open automatically and refresh live.
echo Keep this window open while you want updates.
echo.
python update_option_chain.py
echo.
echo Updater stopped. Press any key to close this window.
pause >nul
