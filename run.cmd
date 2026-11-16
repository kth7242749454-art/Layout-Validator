@echo off
cd /d "%~dp0"
python lee.py
if errorlevel 1 pause
