@echo off
setlocal
cd /d "%~dp0"
title HuggingFace Model BOM

echo Starting HuggingFace Model BOM...
echo.

REM Try the standard 'python' command first.
python hf_bom_server.py
if %errorlevel%==0 goto :end

REM 'python' not found or failed to start - try the Windows 'py' launcher.
echo.
echo 'python' did not start. Trying 'py'...
py hf_bom_server.py
if %errorlevel%==0 goto :end

echo.
echo ------------------------------------------------------------
echo Could not start the tool. Python does not appear to be on PATH.
echo Open Command Prompt in this folder and run:  python hf_bom_server.py
echo ------------------------------------------------------------

:end
echo.
pause
