@echo off
cd /d "%~dp0"

echo Suzdavam sreda za programata...
python -m venv venv >nul 2>nul

if exist "venv\Scripts\python.exe" goto envready

py -m venv venv >nul 2>nul

if exist "venv\Scripts\python.exe" goto envready

echo.
echo Python ne e nameren (ili ne e instaliran pravilno) na tozi kompiuter.
echo.
echo 1. Otvori https://www.python.org/downloads/ v brauzara
echo 2. Svali i instaliraj Python (otbelyazhi "Add python.exe to PATH" v instalatora)
echo 3. Restartiray kompiutara sled instalatsiyata
echo 4. Pusni tozi fail (Install.bat) otnovo
echo.
pause
exit /b 1

:envready

echo Instaliram nuzhnite biblioteki (mozhe da otneme minuta-dve)...
venv\Scripts\python.exe -m pip install --upgrade pip --quiet
venv\Scripts\python.exe -m pip install -r requirements.txt

echo.
echo Gotovo! Ottuk natatuk pyskay programata sus Start.bat
echo.
pause
