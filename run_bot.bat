@echo off
REM Started by Windows Task Scheduler every day at 08:30.
REM It re-opens itself MINIMISED with a clear title, so nobody closes the window by mistake.
if /i not "%~1"=="go" (
    start "Bookings bot - running - do NOT close" /min "%~f0" go
    exit /b
)
title Bookings bot - running - do NOT close (it closes itself when done)
cd /d "%~dp0"
set PYTHONIOENCODING=utf-8
if not exist logs mkdir logs
echo Bookings bot is running. Do NOT close this window - it closes itself when the day is done.
python src\mail_bot.py >> logs\bat_run.log 2>&1
exit
