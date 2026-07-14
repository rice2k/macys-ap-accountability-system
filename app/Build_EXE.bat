@echo off
title Build Macy's AP Beta v5.2.7 EXE
cd /d "%~dp0"
where py >nul 2>nul
if %errorlevel%==0 (
 py -3 -m pip install --upgrade pyinstaller
 py -3 -m PyInstaller --onefile --windowed --icon macys_star_icon.ico --name Macys_AP_Accountability_Beta_v5_2_7 macys_ap_v5_2_7.py
 goto done
)
python -m pip install --upgrade pyinstaller
python -m PyInstaller --onefile --windowed --icon macys_star_icon.ico --name Macys_AP_Accountability_Beta_v5_2_7 macys_ap_v5_2_7.py
:done
echo Build complete. Check dist folder.
pause
