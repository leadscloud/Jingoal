::chcp 936

@echo off & title Write Weekly Log
:do
"python" "%~dp0weekly_worklog.py"
pause
call :do
