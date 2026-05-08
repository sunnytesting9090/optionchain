@echo off
cd /d "%~dp0"
title Zerodha Option Chain WebSocket Server
echo ==========================================
echo   ZERODHA OPTION CHAIN WEBSOCKET SERVER
echo ==========================================
echo.
echo Starting local bridge at ws://127.0.0.1:8765
echo Open option_chain.html and click Connect.
echo Keep this window open while you want live browser updates.
echo.
python zerodha_option_chain_server.py --host 127.0.0.1 --port 8765
echo.
echo Server stopped. Press any key to close this window.
pause >nul
