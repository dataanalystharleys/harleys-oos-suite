@echo off
title Harley's OOS Automation Suite
cd /d "C:\Users\harle\Downloads\oos"
echo =====================================================================
echo Launching Harley's OOS Automation Suite...
echo Opening in your web browser at http://localhost:8501 ...
echo =====================================================================

start "" http://localhost:8501

python -m streamlit run app.py --server.port 8501
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo Tool stopped or exited with an error.
    pause
)

