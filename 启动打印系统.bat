@echo off
chcp 65001 >nul
cd /d "%~dp0"
title LinkPrint - close this window to stop the service
set PYEXE=%LOCALAPPDATA%\Programs\Python\Python312\python.exe
if not exist "%PYEXE%" set PYEXE=py
"%PYEXE%" print_system.py
echo.
echo Stopped. Press any key to close.
pause >nul
