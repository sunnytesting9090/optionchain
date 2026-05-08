@echo off
cd /d "%~dp0"
title Simple Option Chain Updater
echo ==========================================
echo   SIMPLE OPTION CHAIN UPDATER
echo ==========================================
echo.
echo Starting Python updater...
echo Excel will open automatically and refresh all 3 indices.
echo Keep this window open while you want live updates.
echo.
python update_simple_option_chain.py
echo.
echo Updater stopped. Press any key to close this window.
pause >nul
