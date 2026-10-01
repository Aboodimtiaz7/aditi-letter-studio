@echo off
title ADITI Letter Studio
cd /d "%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Run-Letter-Studio.ps1"
if errorlevel 1 pause
