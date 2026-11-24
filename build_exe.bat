@echo off
rem Build the portable package. All the real work is in build_exe.py
cd /d "%~dp0"
set "PY=D:\Python\python.exe"
if not exist "%PY%" set "PY=python"
"%PY%" build_exe.py
pause
