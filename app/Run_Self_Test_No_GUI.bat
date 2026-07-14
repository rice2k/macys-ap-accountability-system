@echo off
title Macy's AP Beta v5.2.7 Self Test
cd /d "%~dp0"
where py >nul 2>nul
if %errorlevel%==0 (
  py -3 macys_ap_v5_2_7.py --self-test
  goto end
)
python macys_ap_v5_2_7.py --self-test
:end
echo.
pause
