@echo off
cd /d "%~dp0"
rem 測試用：直接開 Python 視窗版（不會打包成 exe）
set "PY=%LOCALAPPDATA%\Programs\Python\Python313\python.exe"
if not exist "%PY%" set "PY=py -3.13"
%PY% "%~dp0OverdueList_v1.7.1.py"
if errorlevel 1 pause
