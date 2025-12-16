@echo off
title Smart PDF Data Extractor Pro
echo Starting Smart PDF Data Extractor...
python main.py
if errorlevel 1 (
    echo.
    echo [ERROR] Failed to run! Check if Python and requirements are installed.
    echo Run install.bat first!
    pause
)