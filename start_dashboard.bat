@echo off
echo ========================================
echo Navision Backorder Analyzer - Dashboard
echo ========================================
echo.

REM Controleer of Python geïnstalleerd is
python --version >nul 2>&1
if errorlevel 1 (
    echo FOUT: Python is niet geïnstalleerd of niet gevonden in PATH
    echo Installeer Python 3.8+ en probeer opnieuw
    pause
    exit /b 1
)

echo Python gevonden, starten van dashboard...
echo.

REM Voer het dashboard uit
python simple_dashboard.py

echo.
echo Dashboard afgesloten.
pause