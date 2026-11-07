@echo off
rem Sglobyava EtiketiApteka.exe i arhiva EtiketiApteka.zip za GitHub Release.
cd /d "%~dp0"

venv\Scripts\python.exe -m PyInstaller --noconfirm --clean --windowed ^
  --name EtiketiApteka --icon assets\logo.ico ^
  --add-data "assets;assets" --collect-all tkinterdnd2 ^
  main.py
if errorlevel 1 exit /b 1

copy /y SumatraPDF.exe dist\EtiketiApteka\ >nul
copy /y nzok_reimbursements.json dist\EtiketiApteka\ >nul

if exist dist\EtiketiApteka.zip del dist\EtiketiApteka.zip
powershell -NoProfile -Command "Compress-Archive -Path 'dist\EtiketiApteka' -DestinationPath 'dist\EtiketiApteka.zip'"
echo Gotovo: dist\EtiketiApteka.zip
