@echo off
title AutoCAD Nesting Sheet Builder
echo ============================================================
echo   AutoCAD Nesting Sheet Builder - Launcher
echo ============================================================
python "%~dp0nesting.py"
if %ERRORLEVEL% NEQ 0 (
    pause
)
