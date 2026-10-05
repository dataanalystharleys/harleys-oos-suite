@echo off
title Harley's OOS Automation Suite
echo ========================================================
echo          HARLEY'S OOS AUTOMATION SUITE
echo ========================================================
echo.
echo Starting Streamlit background server...
echo.

cd /d "C:\Users\harle\Downloads\oos"
start /B python -m streamlit run app.py --server.port=8501 --server.headless=true

echo Waiting for server to initialize...
timeout /t 3 /nobreak >nul

echo Opening browser at http://localhost:8501 ...
start "" http://localhost:8501

echo.
echo ========================================================
echo  Server is running! You can keep this window open or
echo  minimize it. Press Ctrl+C in this window to stop.
echo ========================================================
echo.

:: Keep window alive while python is running
:loop
timeout /t 10 >nul
goto loop
