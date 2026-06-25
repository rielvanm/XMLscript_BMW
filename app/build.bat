@echo off
echo ============================================
echo  BMW XML Verwerker -- EXE bouwen
echo ============================================
echo.

:: Controleer of Python beschikbaar is
python --version >nul 2>&1
if errorlevel 1 (
    echo FOUT: Python niet gevonden. Installeer Python 3.10+ en voeg het toe aan PATH.
    pause
    exit /b 1
)

:: Installeer benodigde packages
echo Installeren van benodigde packages...
pip install pyinstaller pandas openpyxl --quiet
if errorlevel 1 (
    echo FOUT: pip install mislukt.
    pause
    exit /b 1
)

:: Verwijder vorige build-mappen
if exist build   rmdir /s /q build
if exist dist    rmdir /s /q dist
if exist BMW_XML.spec del /q BMW_XML.spec

:: Bouw de EXE
echo.
echo Bouwen van EXE (dit kan een minuut duren)...
pyinstaller ^
    --onefile ^
    --windowed ^
    --name BMW_XML ^
    --icon=BMW_XML.ico ^
    --add-data "BMW_XML.png;." ^
    --add-data "BMW_XML.ico;." ^
    XML_BMW_EXE.py

if errorlevel 1 (
    echo.
    echo FOUT: PyInstaller mislukt. Zie de uitvoer hierboven.
    pause
    exit /b 1
)

echo.
echo ============================================
echo  Klaar!  EXE staat in:  dist\BMW_XML.exe
echo ============================================
echo.
pause
