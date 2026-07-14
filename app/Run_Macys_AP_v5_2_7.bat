@echo off
title Macy's AP Accountability System - Beta Version v5.2.7
cd /d "%~dp0"
where py >nul 2>nul
if %errorlevel%==0 (
  py -3 macys_ap_v5_2_7.py
  goto end
)
where python >nul 2>nul
if %errorlevel%==0 (
  python macys_ap_v5_2_7.py
  goto end
)
echo Python was not found. Install Python 3 and enable Add Python to PATH.
pause
exit /b 1
:end
if errorlevel 1 (
 echo App closed with an error.
 pause
)
