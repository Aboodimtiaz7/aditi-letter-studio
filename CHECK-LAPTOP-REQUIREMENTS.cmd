@echo off
title ADITI Letter Studio - Laptop Check
cd /d "%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Check-Laptop-Requirements.ps1"
pause
